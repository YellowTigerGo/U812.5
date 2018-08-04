VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSAWorkFlowSrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim SaveAfterOk As Boolean

'业务单据提交审批调用
Public Function DoSubmit(bizObjectID As String, eventID As String, VoucherId As String, context As String, callerCtx As CalledContext, m_ufts As String, isWfcontrolled As Boolean, ByRef strErr As String, login As Object, Optional cbilltype As String) As Boolean
    On Error GoTo ErrHandler
    Dim strVoucherInfo As String
    Dim dom As New DOMDocument
    Dim strTblName As String
    Dim strMainKeyName As String
    Dim effeLines As Long
    Dim isWfCtl As Boolean
    Dim ccode As String
    Dim m_ireturncount As Integer
    Dim m_ccode As String
    Dim AuditServiceProxy As Object
    Dim rst As New ADODB.Recordset
    Dim Conn As ADODB.Connection
    
    
    
     Set AuditServiceProxy = CreateObject("UFIDA.U8.Audit.ServiceProxy.AuditServiceProxy")
    Dim obj As Object
    Set obj = CreateObject("UFLTMService.clsService")
    Set Conn = obj.RequestConnection
    
    strVoucherInfo = GetVoucherInfo(bizObjectID, Conn)
    dom.LoadXML strVoucherInfo
    strTblName = dom.documentElement.Attributes.getNamedItem("maintbl").nodeValue
    strMainKeyName = dom.documentElement.Attributes.getNamedItem("mainkey").nodeValue
    ccode = dom.documentElement.Attributes.getNamedItem("ccode").nodeValue
    
    writeLog "DoSubmit judge ufts"
    rst.Open "select ireturncount," + ccode + " from " & strTblName & " where Convert(char,Convert(money,Ufts),2)='" & m_ufts & "' and " & strMainKeyName & "=" & VoucherId, Conn, adOpenDynamic, adLockReadOnly
    If Not (rst.BOF And rst.EOF) Then
        If IsNull(rst.fields(0).Value) Then
            m_ireturncount = 0
        Else
            m_ireturncount = Val(rst.fields(0).Value)
        End If
        If IsNull(rst.fields(1).Value) Then
            m_ccode = ""
        Else
            m_ccode = rst.fields(1).Value
        End If
    Else
        strErr = GetString("U8.SA.USSASERVER.clsvouchsave.00542") 'zh-CN：该单据已经不存在或已被其他人修改
        DoSubmit = False
        Exit Function
    End If
    
'    Set AuditServiceProxy = GetObjectContext.CreateInstance("UFIDA.U8.Audit.ServiceProxy.AuditServiceProxy")
'    Set Conn = New ADODB.Connection
'    Conn.Open login.UfDbName
'871 edited
   
    
    
    If m_ireturncount = 0 Then
        '提交后调用自动服务 会发生此情况 将此代码提前 假设提交成功 并进入工作流
        If LCase(strTblName) = "sa_settlevouch" Then '委托代缴结算单 保存 专票普票
            Conn.Execute "update " & strTblName & " set cbilltype='" + cbilltype + "', iverifystate=1 where Convert(char,Convert(money,Ufts),2)='" & m_ufts & "' and " & strMainKeyName & "=" & VoucherId, effeLines
        Else
            Conn.Execute "update " & strTblName & " set iverifystate=1 where Convert(char,Convert(money,Ufts),2)='" & m_ufts & "' and " & strMainKeyName & "=" & VoucherId, effeLines
        End If
        'eventID = "Submit"    'shanlan add
        DoSubmit = AuditServiceProxy.SubmitApplicationMessage_2(bizObjectID, eventID, VoucherId, callerCtx, isWfCtl, strErr)
        If DoSubmit Then
            '如果提交成功并且不进入工作流处理以下
            If isWfCtl = False Then
                If LCase(strTblName) = "sa_settlevouch" Then '委托代缴结算单 保存 专票普票
                    Conn.Execute "update " & strTblName & " set cbilltype='" + cbilltype + "', iverifystate=0,iswfcontrolled=0 where " & strMainKeyName & "=" & VoucherId, effeLines
                Else
                    Conn.Execute "update " & strTblName & " set iverifystate=0,iswfcontrolled=0 where  " & strMainKeyName & "=" & VoucherId, effeLines
                End If
            End If
            If effeLines = 0 Then
                strErr = GetString("U8.SA.USSASERVER.clsvouchsave.00542") 'zh-CN：该单据已经不存在或已被其他人修改
                DoSubmit = False
'                GetObjectContext.SetAbort
                Exit Function
            Else
'                GetObjectContext.SetComplete
            End If
            writeLog "DoSubmit succeeded"
        Else
'            GetObjectContext.SetAbort
            Set Conn = Nothing
            DoSubmit = False
            writeLog "DoSubmit failed"
            Exit Function
        End If
    Else
        eventID = "Submit"  'shanlan add
        DoSubmit = AuditServiceProxy.SubmitRejectedMessage(bizObjectID, eventID, VoucherId, bizObjectID, m_ccode, callerCtx, strErr)
        If DoSubmit Then
            Conn.Execute "update " & strTblName & " set ireturncount=ireturncount-1 where Convert(char,Convert(money,Ufts),2)='" & m_ufts & "' and " & strMainKeyName & "=" & VoucherId, effeLines
            If effeLines = 0 Then
                strErr = GetString("U8.SA.USSASERVER.clsvouchsave.00542") 'zh-CN：该单据已经不存在或已被其他人修改
                DoSubmit = False
                Set Conn = Nothing
'                GetObjectContext.SetAbort
                Exit Function
            Else
                Set Conn = Nothing
'                GetObjectContext.SetComplete
            End If
            writeLog "DoSubmit succeeded"
        Else
            Set Conn = Nothing
            DoSubmit = False
'            GetObjectContext.SetAbort
            writeLog "DoSubmit failed"
            Exit Function
        End If
    End If
    Set dom = Nothing
    Set obj = Nothing
    Exit Function
ErrHandler:
    strErr = VBA.Err.Description
    writeLog "DoSubmit ErrHandler:" & VBA.Err.Description
    Set obj = Nothing
    Set Conn = Nothing
    
    Set dom = Nothing
'    GetObjectContext.SetAbort
    DoSubmit = False
End Function
'业务单据提交撤销调用
Public Function UndoSubmit(bizObjectID As String, eventID As String, VoucherId As String, VoucherType As String, callerCtx As CalledContext, m_ufts As String, isWfcontrolled As Boolean, ByRef strErr As String, cVoucherCode As String, login As Object) As Boolean
    On Error GoTo ErrHandler
    Dim strVoucherInfo As String
    Dim dom As New DOMDocument
    Dim strTblName As String
    Dim strMainKeyName As String
    Dim AuditServiceProxy As Object
    Dim sBizdata As String
    Dim effeLines As Long
    Dim Conn As ADODB.Connection
    Dim rst As ADODB.Recordset
'    Set AuditServiceProxy = GetObjectContext.CreateInstance("UFIDA.U8.Audit.ServiceProxy.AuditServiceProxy")
    Dim obj As Object
    Set obj = CreateObject("UFLTMService.clsService")
    Set Conn = obj.RequestConnection
    
    strVoucherInfo = GetVoucherInfo(VoucherType, Conn)
    dom.LoadXML strVoucherInfo
    strTblName = dom.documentElement.Attributes.getNamedItem("maintbl").nodeValue
    strMainKeyName = dom.documentElement.Attributes.getNamedItem("mainkey").nodeValue
    

'    ccode = "ccode"
'    Set Conn = New ADODB.Connection
'    Conn.Open login.UfDbName
    Set AuditServiceProxy = CreateObject("UFIDA.U8.Audit.ServiceProxy.AuditServiceProxy")
    '871 edited
    
    Set rst = New ADODB.Recordset
    
    rst.Open "select Ufts,isnull(iswfcontrolled,N'') as iswfcontrolled,isnull(iverifystate,N'') as iverifystate from " & strTblName & " where Convert(char,Convert(money,Ufts),2)='" & m_ufts & "' and " & strMainKeyName & "=" & VoucherId, Conn, adOpenDynamic, adLockReadOnly
    writeLog "UndoSubmit judge ufts"
    If rst.BOF And rst.EOF Then
        strErr = GetString("U8.SA.USSASERVER.clsvouchsave.00542") 'zh-CN：该单据已经不存在或已被其他人修改
        UndoSubmit = False
        Set Conn = Nothing
        Set dom = Nothing
        Set rst = Nothing
'        GetObjectContext.SetAbort
        Exit Function
    Else
        If rst.fields(1).Value = "1" And rst.fields(2).Value = "2" Then
            strErr = GetString("U8.SA.xsglsql_2.saworkflowsrv.018") 'zh-CN：该单据应经终审不能撤销
            UndoSubmit = False
            Set Conn = Nothing
            Set dom = Nothing
            Set rst = Nothing
'            GetObjectContext.SetAbort
            Exit Function
        End If
    End If
    Set rst = Nothing
    
    writeLog ("UndoSubmit :bizObjectID|" & bizObjectID & "|eventID|" & eventID & "|VoucherId|" & VoucherId & "|VoucherType|" & VoucherType & "|m_ufts|" & m_ufts & "|strErr|" & strErr & "|cVoucherCode|" & cVoucherCode)
   ' eventID = "Submit"  'shanlan add
    UndoSubmit = AuditServiceProxy.RevokeAudition(bizObjectID, eventID, VoucherId, VoucherType, cVoucherCode, sBizdata, callerCtx, strErr)
    If UndoSubmit Then
        Conn.Execute "update " & strTblName & " set iverifystate=0,ireturncount=0 where " & strMainKeyName & "=" & VoucherId, effeLines
        If effeLines = 0 Then
            strErr = GetString("U8.SA.xsglsql_2.saworkflowsrv.010") '"更新审核标志失败"
            UndoSubmit = False
            Set Conn = Nothing
'            Set dom = Nothing
'            GetObjectContext.SetAbort
            
            Exit Function
        End If
        Set Conn = Nothing
'        GetObjectContext.SetComplete
        writeLog ("UndoSubmitAuditServiceProxySetComplete")
    Else
        Set Conn = Nothing
'        GetObjectContext.SetAbort
        writeLog ("UndoSubmitAuditServiceProxySetAbort")
        UndoSubmit = False
        Exit Function
    End If
    Set dom = Nothing
    Set obj = Nothing
    Exit Function
ErrHandler:
    strErr = VBA.Err.Description
    writeLog ("UndoSubmit ErrHandler:") & VBA.Err.Description
    Set Conn = Nothing
    Set obj = Nothing
    UndoSubmit = False
    Set dom = Nothing
'    GetObjectContext.SetAbort
End Function
'审批流审核调用
Public Function DoAudition(keyData As String, auditAction As String, ByRef bizData As String, calledCtx As Object, ByRef errMsg As String) As Boolean
    Dim strAction As String
    Dim objAction As New AuditActionMap
    Dim KeyDataMap As New KeyDataMap
    Dim VoucherId As Long
    Dim VoucherType As String
    Dim Vufts As String
    Dim objBizData As New BizDataMap
    Dim login As U8Login.clsLogin
    Dim Conn As ADODB.Connection
    DoAudition = True
    On Error GoTo ErrHandler
    objAction.LoadXML auditAction
    strAction = objAction.Value("Action")
    KeyDataMap.LoadXML keyData
    VoucherId = CLng(KeyDataMap.Value("VoucherId"))
    VoucherType = KeyDataMap.Value("VoucherType")
    Vufts = KeyDataMap.Value("Ufts")
    objBizData.LoadXML bizData
'    writeLog ("DoAudition-keyData:" & keyData)
    
    Set login = New U8Login.clsLogin
    login.ConstructLogin calledCtx.token
    login.TaskId = calledCtx.TaskId
    login.login "SA"
    
    Set Conn = New Connection
    Conn.Open login.UfDbName
    Dim obj As Object
'    Set obj = CreateObject("UFLTMService.clsService")
'    Set Conn = obj.RequestConnection
    
    If strAction = "Agree" Then
        DoAudition = VerifyVoucherAgree(login, Conn, VoucherType, VoucherId, objBizData, Vufts, keyData)
    ElseIf strAction = "Reject" Then
        DoAudition = RejectVoucher(login, Conn, VoucherType, VoucherId, objBizData, Vufts)
    ElseIf strAction = "Disagree" Then
        DoAudition = VerifyVoucherDisagree(login, Conn, VoucherType, VoucherId, objBizData, Vufts)
    End If
    writeLog ("DoAudition-DoAudition" & CStr(DoAudition))
    writeLog ("DoAudition-" & "after Agree")
    bizData = objBizData.XML
    KeyDataMap.Value("Ufts") = Vufts
'    writeLog ("DoAudition-" & KeyDataMap.XML)
    keyData = KeyDataMap.XML
    Set login = Nothing
    errMsg = objBizData.Value("errmsg")
    writeLog ("DoAudition OK" & errMsg)
    

    Set Conn = Nothing

    
'    writeLog ("DoAudition-" & "after DoAudition keyData:" & keyData)
'    writeLog ("DoAudition-" & "after DoAudition bizData:" & bizData)
   
'    If errMsg <> "" Then Err.Raise vbObjectError + 1000, "", errMsg
    Set KeyDataMap = Nothing
    Set obj = Nothing
    Exit Function
    
ErrHandler:
    errMsg = VBA.Err.Description
    Set login = Nothing
    Set Conn = Nothing
    Set obj = Nothing
    DoAudition = False
    objBizData.Value("errmsg") = errMsg
    writeLog ("DoAudition-" & "after ErrHandler" & errMsg)
    bizData = objBizData.XML
    Set KeyDataMap = Nothing
End Function



'审批流弃审调用
Public Function UndoAudition(keyData As String, auditAction As String, ByRef bizData As String, calledCtx As Object, ByRef errMsg As String) As Boolean ', classnames As String
On Error GoTo ErrHandler
    Dim strVoucherInfo As String
    Dim dom As New DOMDocument
    Dim KeyDataMap As New KeyDataMap
    Dim VoucherId As Long
    Dim VoucherType As String
    Dim strMainViewName As String
    Dim strMainKeyName As String
    Dim strSaVouchType As String
    Dim strTblName As String
    Dim domHead As New DOMDocument
    Dim m_ufts As String
    Dim objBizData As New BizDataMap
    Dim rst As New Recordset
    Dim intVerifyState As String
    Dim lEffectedCount As Integer
    Dim m_rst As ADODB.Recordset
    Dim login As U8Login.clsLogin
    Dim Conn As ADODB.Connection
    UndoAudition = False
    writeLog "UndoAudition:Entered UndoAudition"
'    Dim objVoucherCO As EFVoucherCo.ClsVoucherCO
    Dim objVoucherCO As Object
    KeyDataMap.LoadXML keyData
    VoucherId = CLng(KeyDataMap.Value("VoucherId"))
    VoucherType = KeyDataMap.Value("VoucherType")
    m_ufts = KeyDataMap.Value("Ufts")
    objBizData.LoadXML bizData
    
    Set login = New U8Login.clsLogin
    login.ConstructLogin calledCtx.token
    login.TaskId = calledCtx.TaskId
    login.login "SA"
    
    Set Conn = New Connection
    Conn.Open login.UfDbName
    
    Dim obj As Object
'
'    Set obj = CreateObject("UFLTMService.clsService")
'     writeLog "UndoAudition:Before Set Conn = obj.RequestConnection"
'    Set Conn = obj.RequestConnection
    writeLog "UndoAudition:Conn =" & Conn.ConnectionString
    strVoucherInfo = GetVoucherInfo(VoucherType, Conn)
    dom.LoadXML strVoucherInfo
    strMainViewName = dom.documentElement.Attributes.getNamedItem("mainview").nodeValue
    strMainKeyName = dom.documentElement.Attributes.getNamedItem("mainkey").nodeValue
    strSaVouchType = dom.documentElement.Attributes.getNamedItem("vouchtype").nodeValue
    strTblName = dom.documentElement.Attributes.getNamedItem("maintbl").nodeValue
    
'    writeLog ("UndoAudition-keyData:" & keyData)
    
    Set dom = Nothing
    If VoucherType = "01" Then
        Set m_rst = New ADODB.Recordset
        m_rst.Open "select cvouchtype from dispatchlist where dlid=" & VoucherId, Conn, adOpenForwardOnly, adLockReadOnly
        If m_rst.fields(0).Value = "06" Or m_rst.fields(0).Value = "00" Then
            If m_ufts = "" Then
                rst.Open "select * from Sales_DXFH_T where " & strMainKeyName & "=" & VoucherId, Conn, adOpenForwardOnly, adLockReadOnly
            Else
                rst.Open "select * from Sales_DXFH_T where Convert(char,Convert(money,Ufts),2)='" & m_ufts & "' and " & strMainKeyName & "=" & VoucherId, Conn, adOpenForwardOnly, adLockReadOnly
            End If
        Else
            If m_ufts = "" Then
                rst.Open "select * from " & strMainViewName & " where " & strMainKeyName & "=" & VoucherId, Conn, adOpenForwardOnly, adLockReadOnly
            Else
                rst.Open "select * from " & strMainViewName & " where Convert(char,Convert(money,Ufts),2)='" & m_ufts & "' and " & strMainKeyName & "=" & VoucherId, Conn, adOpenForwardOnly, adLockReadOnly
            End If
        End If
        m_rst.Close
        Set m_rst = Nothing
    Else
        If m_ufts = "" Then
            rst.Open "select * from " & strMainViewName & " where  " & strMainKeyName & "=" & VoucherId, Conn, adOpenForwardOnly, adLockReadOnly
        Else
            rst.Open "select * from " & strMainViewName & " where Convert(char,Convert(money,Ufts),2)='" & m_ufts & "' and " & strMainKeyName & "=" & VoucherId, Conn, adOpenForwardOnly, adLockReadOnly
        End If
    End If
    writeLog ("UndoAudition-rst.Source:" & rst.Source)
    If rst.BOF And rst.EOF Then
        errMsg = GetString("U8.SA.USSASERVER.clsvouchsave.00542") 'zh-CN：该单据已经不存在或已被其他人修改
        objBizData.Value("errMsg") = errMsg
        VBA.Err.Raise vbObjectError + 1000, "", errMsg
        UndoAudition = False
        Exit Function
    End If
    If IsNull(rst.fields("iverifystate")) Then
        intVerifyState = ""
    Else
        intVerifyState = rst.fields("iverifystate")
    End If
    rst.Save domHead, adPersistXML
    rst.Close
    Set rst = Nothing
    Dim ele As IXMLDOMElement
    Set ele = domHead.selectSingleNode("//z:row")
     '''''''''''''''''''''''''返回信息处理
    If objBizData.Value("SaveAfterOk") <> "" Then
       ele.setAttribute "saveafterok", objBizData.Value("SaveAfterOk")
    Else
       If KeyDataMap.Value("SaveAfterOk") <> "" Then
        ele.setAttribute "saveafterok", KeyDataMap.Value("SaveAfterOk")
       End If
    End If
    If objBizData.Value("ccrechpname") <> "" Then
        ele.setAttribute "ccrechpname", objBizData.Value("ccrechpname")
     Else
       If KeyDataMap.Value("ccrechpname") <> "" Then
        ele.setAttribute "ccrechpname", KeyDataMap.Value("ccrechpname")
       End If
    End If
    If objBizData.Value("ccrechppass") <> "" Then
        ele.setAttribute "ccrechppass", objBizData.Value("ccrechppass")
    Else
       If KeyDataMap.Value("ccrechppass") <> "" Then
        ele.setAttribute "ccrechppass", KeyDataMap.Value("ccrechppass")
       End If
    End If
    If objBizData.Value("bCreditCheck") <> "" Then
        ele.setAttribute "bCreditCheck", objBizData.Value("bCreditCheck")
    Else
       If KeyDataMap.Value("bCreditCheck") <> "" Then
        ele.setAttribute "bCreditCheck", KeyDataMap.Value("bCreditCheck")
       End If
    End If
    If objBizData.Value("AfterCheckCredit") <> "" Then
        ele.setAttribute "AfterCheckCredit", objBizData.Value("AfterCheckCredit")
    Else
       If KeyDataMap.Value("AfterCheckCredit") <> "" Then
        ele.setAttribute "AfterCheckCredit", KeyDataMap.Value("AfterCheckCredit")
       End If
    End If
    If objBizData.Value("bcontinue") <> "" Then
        ele.setAttribute "bcontinue", objBizData.Value("bcontinue")
    Else
       If KeyDataMap.Value("ccrechpname") <> "" Then
        ele.setAttribute "bcontinue", KeyDataMap.Value("bcontinue")
       End If
    End If
    ''''''''''''''''''''''''
    writeLog ("UndoAudition-intVerifyState" & intVerifyState)
    If intVerifyState = "2" Then
        writeLog ("UndoAudition:" & "new objVoucherCO:")
'        Set objVoucherCO = New EFVoucherCo.ClsVoucherCO
       
 
            writeLog ("UndoAudition:" & "InitSys")
            Set objVoucherCO = InitColtrol(strSaVouchType, login, Conn)
            'objVoucherCO.Init strSaVouchType, login, Conn, "bs"
            'objVoucherCO.InitSys login, strSaVouchType, clsSAWeb
 
        writeLog ("UndoAudition:" & "objVoucherCO.VerifyVouch")
        objVoucherCO.clsSystemWeb.bManualTrans = True
        objVoucherCO.bNeedLock = False
        errMsg = objVoucherCO.VerifyVouch(domHead, False)
         writeLog ("UndoAudition:" & "errMsg" & errMsg)
        Set objVoucherCO = Nothing
        If errMsg = "" Then
            Conn.Execute "update " + strTblName + " set iverifystate=1 where " + strMainKeyName + " =" & VoucherId, lEffectedCount
            If lEffectedCount = 0 Then
                VBA.Err.Raise vbObjectError + 1000, "", GetString("U8.SA.xsglsql_2.saworkflowsrv.010") '"更新审核标志失败"
            End If
            Set rst = New ADODB.Recordset
            rst.Open "select Convert(char,Convert(money,Ufts),2) as ufts from " + strTblName + " where " + strMainKeyName + " =" & VoucherId, Conn, adOpenStatic, adLockReadOnly
            m_ufts = rst.fields(0).Value '传回新的时间戳
            writeLog ("UndoAudition:" & "m_ufts")
            KeyDataMap.Value("ccrechpname") = ""
            KeyDataMap.Value("ccrechppass") = ""
            KeyDataMap.Value("SaveAfterOk") = ""
            KeyDataMap.Value("bCreditCheck") = ""
            KeyDataMap.Value("AfterCheckCredit") = ""
            KeyDataMap.Value("bcontinue") = ""
            Set rst = Nothing
            UndoAudition = True
        Else
            '错误处理

            objBizData.Value("errMsg") = errMsg
            objBizData.Value("domHead") = domHead.XML
            VBA.Err.Raise vbObjectError + 1000, "", errMsg
            UndoAudition = False
            writeLog ("UndoAudition:错误处理" & errMsg & "UndoAudition:" & CStr(UndoAudition))
        End If
    ElseIf intVerifyState = "-1" Then
        Conn.Execute "update " + strTblName + " set iverifystate=1 where " + strMainKeyName + " =" & VoucherId, lEffectedCount
        If lEffectedCount = 0 Then
            VBA.Err.Raise vbObjectError + 1000, "", GetString("U8.SA.xsglsql_2.saworkflowsrv.010") '"更新审核标志失败"
            UndoAudition = False
        Else
            Set rst = New ADODB.Recordset
            rst.Open "select Convert(char,Convert(money,Ufts),2) as ufts from " + strTblName + " where " + strMainKeyName + " =" & VoucherId, Conn, adOpenStatic, adLockReadOnly
            m_ufts = rst.fields(0).Value '传回新的时间戳
            Set rst = Nothing
            UndoAudition = True
        End If
    End If
     writeLog ("UndoAudition:" & "ok" & errMsg)
    Set login = Nothing
    Set Conn = Nothing
    KeyDataMap.Value("Ufts") = m_ufts
    keyData = KeyDataMap.XML
'    writeLog ("UndoAudition:" & keyData)
    bizData = objBizData.XML
'    writeLog ("UndoAudition:" & bizData)
    Set ele = Nothing
    Set domHead = Nothing
    Set objVoucherCO = Nothing
    Set KeyDataMap = Nothing
    Set obj = Nothing
    Exit Function
ErrHandler:
    errMsg = VBA.Err.Description
    writeLog "UndoAudition:ErrHandler =" & errMsg
    Set login = Nothing
    Set Conn = Nothing
    Set obj = Nothing
    UndoAudition = False
    objBizData.Value("errmsg") = errMsg
    Set domHead = Nothing
    Set objVoucherCO = Nothing
    bizData = objBizData.XML
     writeLog ("UndoAudition:" & "ErrHandler:" & errMsg)
     Set KeyDataMap = Nothing
End Function

'审批流错误处理
Public Function HandleBizException(keyData As String, auditAction As String, ByRef handleResult As String, ByRef bizData As String, calledCtx As Object, ByRef errMsg As String) As Boolean
    Dim objBizData As New BizDataMap
    Dim objKeyData As New KeyDataMap
    Dim nod As IXMLDOMElement
    Dim lst As IXMLDOMNodeList
    Dim strError As String
    Dim ReterrMsg As String
    Dim domHead As DOMDocument
    
    writeLog "entered HandleBizException"
    Set domHead = New DOMDocument
    Dim login As U8Login.clsLogin
    Dim Conn As New ADODB.Connection
    Dim clsSAWeb As Object
    Set login = New U8Login.clsLogin
    HandleBizException = True
    objBizData.LoadXML bizData
    objKeyData.LoadXML keyData
    login.ConstructLogin calledCtx.token
    login.TaskId = calledCtx.TaskId
    login.login "SA"
    
    writeLog "login OK"
    
    Conn.Open login.UfDbName
    
    writeLog "Conn OK"
'    Set frmCheckCredit.Conn = Conn
'    Set frmCheckCredit.login = login
    writeLog ("HandleBizException:Conn" & Conn.ConnectionString)
    'Set clsSAWeb = CreateObject("VoucherModelEFGSPHC.clsSystem")
    'clsSAWeb.Init login
'    clsSAWeb.INIMyInfor
'    Set frmCheckCredit.clsSAWeb = clsSAWeb
    ReterrMsg = objBizData.Value("errMsg")
    domHead.LoadXML objBizData.Value("domHead")
    If ReterrMsg <> "" Then
        If domHead Is Nothing Then
            Dim tmpDOM As New DOMDocument
            Dim tmpErrString As String, strXml As String
            Dim I As Integer
            I = InStr(1, ReterrMsg, "<", vbTextCompare)
            If I <> 0 Then
                tmpErrString = Mid(ReterrMsg, 1, I - 1)
                strXml = Mid(ReterrMsg, I)
                If tmpDOM.LoadXML(strXml) = False Then
'                    MsgBox GetString("U8.SA.xsglsql.01.frmbillvouch.00417"), vbExclamation 'zh-CN：在错误处理中无法生成错误生成DOM对象！
'                    MsgBox ReterrMsg
                    handleResult = "UnHandle"
                    GoTo exitFun
'                    Set login = Nothing
'                    Set Conn = Nothing
'                    Set clsSAWeb = Nothing'
'                    Exit Function
                End If
                Set tmpDOM = Nothing
            Else
                ''正常的错误
                SaveAfterOk = False
'                MsgBox ReterrMsg, vbExclamation
                handleResult = "UnHandle"
                GoTo exitFun
'                Set login = Nothing
'                Set Conn = Nothing
'                Set clsSAWeb = Nothing
'                Exit Function
            End If
        Else
            If domHead.selectNodes("//信用检查不通过").length > 0 Then
            
            ElseIf domHead.selectNodes("//可用量检查不过").length > 0 Then
                
                    objBizData.Value("bcontinue") = 1
                    handleResult = "UnHandle"
                    GoTo exitFun
'                    Set login = Nothing
'                    Set Conn = Nothing
'                    Set clsSAWeb = Nothing
'                    Exit Function
            Else
                I = InStr(1, ReterrMsg, "<", vbTextCompare)
                If I <> 0 Then
                    tmpErrString = Mid(ReterrMsg, 1, I - 1)
                    strXml = Mid(ReterrMsg, I)
                    If tmpDOM.LoadXML(strXml) = False Then
'                        MsgBox GetString("U8.SA.xsglsql.01.frmbillvouch.00417"), vbExclamation 'zh-CN：在错误处理中无法生成错误生成DOM对象！
    '                    MsgBox ReterrMsg
                        handleResult = "UnHandle"
                        GoTo exitFun
'                        Set login = Nothing
'                        Set Conn = Nothing
'                        Set clsSAWeb = Nothing
'                        Exit Function
                    End If
                    Set tmpDOM = Nothing
                Else
'                    MsgBox ReterrMsg, vbExclamation
                    handleResult = "UnHandle"
                    GoTo exitFun
'                    Set login = Nothing
'                    Set Conn = Nothing
'                    Set clsSAWeb = Nothing
'                    Exit Function
                End If
                
            End If
        End If
    End If
exitFun:
    writeLog ("HandleBizException:OK")
    keyData = objKeyData.XML
    bizData = objBizData.XML
    HandleBizException = True
    Set login = Nothing
    Set Conn = Nothing
    Set clsSAWeb = Nothing
End Function

Private Function VerifyVoucherAgree(login As clsLogin, Conn As ADODB.Connection, ByVal VoucherType As String, ByVal VoucherId As String, objBizData As BizDataMap, ByRef m_ufts As String, keyset As String) As Boolean
    On Error GoTo ErrHandler
    Dim objVoucherCO As EFVoucherCo.ClsVoucherCO
    Dim lEffectedCount As Long
    Dim strErr As String
    Dim domHead As New DOMDocument
    Dim dom As New DOMDocument
    Dim strVoucherInfo As String
    Dim strMainViewName As String
    Dim strMainKeyName As String
    Dim rst As New ADODB.Recordset
    Dim m_rst As ADODB.Recordset
    Dim strSaVouchType As String
    Dim strTblName As String
    Dim ele As IXMLDOMElement
    Dim KeyDataMap As New KeyDataMap
    KeyDataMap.LoadXML keyset
    rst.CursorLocation = adUseClient
    VerifyVoucherAgree = False
    writeLog "VerifyVoucherAgree-VoucherType:" + VoucherType
    
    strVoucherInfo = GetVoucherInfo(VoucherType, Conn)
    dom.LoadXML strVoucherInfo
    writeLog "VerifyVoucherAgree-strVoucherInfo:" + strVoucherInfo
    strMainViewName = dom.documentElement.Attributes.getNamedItem("mainview").nodeValue
    strMainKeyName = dom.documentElement.Attributes.getNamedItem("mainkey").nodeValue
    strSaVouchType = dom.documentElement.Attributes.getNamedItem("vouchtype").nodeValue
    strTblName = dom.documentElement.Attributes.getNamedItem("maintbl").nodeValue
    Set dom = Nothing

    writeLog "VerifyVoucherAgree-strVoucherInfo IS OK"
    Dim strVouchClosed As String
    strVouchClosed = ErrVoucherClosed(VoucherId, VoucherType, Conn)
    If strVouchClosed <> "" Then
        strErr = strVouchClosed
        objBizData.Value("errMsg") = strErr
        strVouchClosed = ""
        VBA.Err.Raise vbObjectError + 1000, "", strErr
        Exit Function
    End If
    If VoucherType = "01" Then
        Set m_rst = New ADODB.Recordset
        m_rst.Open "select cvouchtype from dispatchlist where dlid=" & VoucherId, Conn, adOpenForwardOnly, adLockReadOnly
        If m_rst.fields(0).Value = "06" Or m_rst.fields(0).Value = "00" Then
            If m_ufts = "" Then
                rst.Open "select * from Sales_DXFH_T where " & strMainKeyName & "=" & VoucherId, Conn, adOpenForwardOnly, adLockReadOnly
            Else
                rst.Open "select * from Sales_DXFH_T where Convert(char,Convert(money,Ufts),2)='" & m_ufts & "' and " & strMainKeyName & "=" & VoucherId, Conn, adOpenForwardOnly, adLockReadOnly
            End If
        Else
            If m_ufts = "" Then
                rst.Open "select * from " & strMainViewName & " where " & strMainKeyName & "=" & VoucherId, Conn, adOpenForwardOnly, adLockReadOnly
            Else
                rst.Open "select * from " & strMainViewName & " where Convert(char,Convert(money,Ufts),2)='" & m_ufts & "' and " & strMainKeyName & "=" & VoucherId, Conn, adOpenForwardOnly, adLockReadOnly
            End If
        End If
        m_rst.Close
        Set m_rst = Nothing
    Else
        If m_ufts = "" Then
            rst.Open "select * from " & strMainViewName & " where " & strMainKeyName & "=" & VoucherId, Conn, adOpenForwardOnly, adLockReadOnly
        Else
            rst.Open "select * from " & strMainViewName & " where Convert(char,Convert(money,Ufts),2)='" & m_ufts & "' and " & strMainKeyName & "=" & VoucherId, Conn, adOpenForwardOnly, adLockReadOnly
        End If
    End If
    writeLog ("VerifyVoucherAgree rst.Source:" & rst.Source)
    If rst.BOF And rst.EOF Then
        strErr = GetString("U8.SA.USSASERVER.clsvouchsave.00542") 'zh-CN：该单据已经不存在或已被其他人修改
        
        objBizData.Value("errMsg") = strErr
        VBA.Err.Raise vbObjectError + 1000, "", strErr
        Exit Function
    End If
    rst.Save domHead, adPersistXML
    rst.Close
    Set rst = Nothing
    
    Set ele = domHead.selectSingleNode("//z:row")
    If VoucherType = "02" Then
        Set rst = New ADODB.Recordset
        rst.Open "select cbilltype from " & strTblName & " where " & strMainKeyName & "=" & VoucherId, Conn, adOpenForwardOnly, adLockReadOnly
        If Not IsNull(rst.fields(0).Value) Then
            ele.setAttribute "billvouchtype", rst.fields(0).Value
        End If
        rst.Close
        Set rst = Nothing
    End If
    
    '''''''''''''''''''''''''返回信息处理
    
    If objBizData.Value("SaveAfterOk") <> "" Then
       ele.setAttribute "saveafterok", objBizData.Value("SaveAfterOk")
    Else
       If KeyDataMap.Value("SaveAfterOk") <> "" Then
        ele.setAttribute "saveafterok", KeyDataMap.Value("SaveAfterOk")
       End If
    End If
    If objBizData.Value("ccrechpname") <> "" Then
        ele.setAttribute "ccrechpname", objBizData.Value("ccrechpname")
     Else
       If KeyDataMap.Value("ccrechpname") <> "" Then
        ele.setAttribute "ccrechpname", KeyDataMap.Value("ccrechpname")
       End If
    End If
    If objBizData.Value("ccrechppass") <> "" Then
        ele.setAttribute "ccrechppass", objBizData.Value("ccrechppass")
    Else
       If KeyDataMap.Value("ccrechppass") <> "" Then
        ele.setAttribute "ccrechppass", KeyDataMap.Value("ccrechppass")
       End If
    End If
    If objBizData.Value("bCreditCheck") <> "" Then
        ele.setAttribute "bCreditCheck", objBizData.Value("bCreditCheck")
    Else
       If KeyDataMap.Value("bCreditCheck") <> "" Then
        ele.setAttribute "bCreditCheck", KeyDataMap.Value("bCreditCheck")
       End If
    End If
    If objBizData.Value("AfterCheckCredit") <> "" Then
        ele.setAttribute "AfterCheckCredit", objBizData.Value("AfterCheckCredit")
    Else
       If KeyDataMap.Value("AfterCheckCredit") <> "" Then
        ele.setAttribute "AfterCheckCredit", KeyDataMap.Value("AfterCheckCredit")
       End If
    End If
    If objBizData.Value("bcontinue") <> "" Then
        ele.setAttribute "bcontinue", objBizData.Value("bcontinue")
    Else
       If KeyDataMap.Value("ccrechpname") <> "" Then
        ele.setAttribute "bcontinue", KeyDataMap.Value("bcontinue")
       End If
    End If
    ''''''''''''''''''''''''
    writeLog ("VerifyVoucherAgree new objVoucherCO ")
    Set objVoucherCO = New EFVoucherCo.ClsVoucherCO
    
    If LCase(strSaVouchType) = "sa18" Or LCase(strSaVouchType) = "sa19" Or LCase(strSaVouchType) = "sa99" Then
        objVoucherCO.init "SA99", login, Conn, "bs"
        'objVoucherCO.InitSys login, , clsSAWeb
        Set ele = domHead.selectSingleNode("//z:row")
        ele.setAttribute "cvouchtype", strSaVouchType
    Else
        Set objVoucherCO = InitColtrol(strSaVouchType, login, Conn)
        'objVoucherCO.InitSys login, strSaVouchType, clsSAWeb
        writeLog ("VerifyVoucherAgree InitSys")
    End If
    objVoucherCO.clsSystemWeb.bManualTrans = True
    objVoucherCO.bNeedLock = False
    writeLog ("VerifyVoucherAgree objVoucherCO.VerifyVouch(domHead, True)")
    strErr = objVoucherCO.VerifyVouch(domHead, True)
    Set ele = Nothing
    writeLog ("VerifyVoucherAgree objVoucherCO strErr:" & strErr)
    If strErr = "" Then
        Conn.Execute "update " + strTblName + " set iverifystate=2 where " + strMainKeyName + " =" & VoucherId, lEffectedCount
        If lEffectedCount = 0 Then
'            VBA.Err.Raise vbObjectError + 1000, "", GetString("U8.SA.xsglsql_2.saworkflowsrv.010") '"更新审核标志失败"
            objBizData.Value("errMsg") = GetString("U8.SA.xsglsql_2.saworkflowsrv.010") '"更新审核标志失败"
            VBA.Err.Raise vbObjectError + 1000, "", GetString("U8.SA.xsglsql_2.saworkflowsrv.010")
            VerifyVoucherAgree = False
        Else
            Set rst = New ADODB.Recordset
            rst.Open "select Convert(char,Convert(money,Ufts),2) as ufts from " + strTblName + " where " + strMainKeyName + " =" & VoucherId, Conn, adOpenStatic, adLockReadOnly
            m_ufts = rst.fields(0).Value '传回新的时间戳
            VerifyVoucherAgree = True
            Set rst = Nothing
        End If
    Else

        objBizData.Value("errMsg") = strErr
        objBizData.Value("domHead") = domHead.XML
        writeLog ("VerifyVoucherAgree:错误处理" & strErr)
        VBA.Err.Raise vbObjectError + 1000, "", strErr
    End If
    writeLog ("VerifyVoucherAgree succeed:" & m_ufts)
    Set domHead = Nothing
    Set objVoucherCO = Nothing
    Set KeyDataMap = Nothing
    Exit Function
ErrHandler:
    strErr = VBA.Err.Description
    Set domHead = Nothing
    Set objVoucherCO = Nothing
    writeLog ("VerifyVoucherAgree:错误处理" & strErr)
   objBizData.Value("errmsg") = strErr
   VerifyVoucherAgree = False
   Set KeyDataMap = Nothing
End Function
Private Function VerifyVoucherDisagree(login As clsLogin, Conn As Connection, ByVal VoucherType As String, ByVal VoucherId As String, objBizData As BizDataMap, ByRef m_ufts As String) As Boolean
    On Error GoTo ErrHandler
    Dim dom As New DOMDocument
    Dim strVoucherInfo As String
    Dim strMainKeyName As String
    Dim strTblName As String
    Dim lEffectedCount As Long
    Dim strErr As String
    VerifyVoucherDisagree = False
    writeLog VoucherType
    strVoucherInfo = GetVoucherInfo(VoucherType, Conn)
    dom.LoadXML strVoucherInfo
    writeLog strVoucherInfo
    strMainKeyName = dom.documentElement.Attributes.getNamedItem("mainkey").nodeValue
    strTblName = dom.documentElement.Attributes.getNamedItem("maintbl").nodeValue
    Set dom = Nothing
    writeLog ("VerifyVoucherDisagree m_ufts" & m_ufts)
    Dim strVouchClosed As String
    strVouchClosed = ErrVoucherClosed(VoucherId, VoucherType, Conn)
    If strVouchClosed <> "" Then
        VBA.Err.Raise vbObjectError + 1000, "", strVouchClosed
        VerifyVoucherDisagree = False
        strVouchClosed = ""
        Exit Function
    End If
    If m_ufts = "" Then
        Conn.Execute "update " + strTblName + " set iverifystate=-1 where iverifystate<>-1 and " + strMainKeyName + " =" & VoucherId, lEffectedCount
    Else
        Conn.Execute "update " + strTblName + " set iverifystate=-1 where iverifystate<>-1 and Convert(char,Convert(money,Ufts),2)='" & m_ufts & "' and " + strMainKeyName + " =" & VoucherId, lEffectedCount
    End If
    If lEffectedCount = 0 Then
       VBA.Err.Raise vbObjectError + 1000, "", GetString("U8.SA.USSASERVER.clsvouchsave.00542")  'zh-CN：该单据已经不存在或已被其他人修改'GetString("U8.SA.xsglsql_2.saworkflowsrv.010") '"更新审核标志失败"
       VerifyVoucherDisagree = False
    Else
        VerifyVoucherDisagree = True
    End If
    writeLog ("VerifyVoucherDisagree ok")
    Exit Function
ErrHandler:
    strErr = VBA.Err.Description
    writeLog ("VerifyVoucherDisagree ErrHandler:" & Err.Description)
   objBizData.Value("errmsg") = strErr
   VerifyVoucherDisagree = False
End Function
Private Function RejectVoucher(login As clsLogin, Conn As Connection, ByVal VoucherType As String, ByVal VoucherId As String, objBizData As BizDataMap, ByRef m_ufts As String) As Boolean
    On Error GoTo ErrHandler
    Dim dom As New DOMDocument
    Dim strVoucherInfo As String
    Dim strMainKeyName As String
    Dim strTblName As String
    Dim lEffectedCount As Long
    Dim strErr As String
    
    RejectVoucher = False
    
    strVoucherInfo = GetVoucherInfo(VoucherType, Conn)
    dom.LoadXML strVoucherInfo
    strMainKeyName = dom.documentElement.Attributes.getNamedItem("mainkey").nodeValue
    strTblName = dom.documentElement.Attributes.getNamedItem("maintbl").nodeValue
    Set dom = Nothing
    If m_ufts = "" Then
        Conn.Execute "update " + strTblName + " set ireturncount=ireturncount+1 where " + strMainKeyName + " =" & VoucherId, lEffectedCount
    Else
        Conn.Execute "update " + strTblName + " set ireturncount=ireturncount+1 where Convert(char,Convert(money,Ufts),2)='" & m_ufts & "' and " + strMainKeyName + " =" & VoucherId, lEffectedCount
    End If
    If lEffectedCount = 0 Then
       VBA.Err.Raise vbObjectError + 1000, "", GetString("U8.SA.USSASERVER.clsvouchsave.00542") 'zh-CN：该单据已经不存在或已被其他人修改'GetString("U8.SA.xsglsql_2.saworkflowsrv.010") '"更新审核标志失败"
    Else
        RejectVoucher = True
    End If
    
    Exit Function
ErrHandler:
   strErr = VBA.Err.Description
   objBizData.Value("errmsg") = strErr
   RejectVoucher = False
End Function

Public Function InitColtrol(VoucherType As String, login As U8Login.clsLogin, Conn As ADODB.Connection, Optional classname As String = "EFVoucherCo.ClsVoucherCO") As Object
    Dim M_VoucherType As String
    M_VoucherType = "MT66"
    Dim objVoucherCO As Object
'    Set objVoucherCO = New EFVoucherCo.ClsVoucherCO
    Set objVoucherCO = CreateObject(classname)
      objVoucherCO.init VoucherType, login, Conn, "bs"
'        objVoucherCO.GetVouchInf M_VoucherType, VoucherType, "", False, False, ""
'        objVoucherCO.init "MT01", login, Conn, "bs"
     Set InitColtrol = objVoucherCO
End Function
