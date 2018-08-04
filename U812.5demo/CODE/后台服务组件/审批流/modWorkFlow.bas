Attribute VB_Name = "modWorkFlow"
'Public login As U8Login.clsLogin
'Public Conn As ADODB.Connection
'Public SaveAfterOk As Boolean
'Public myinfo As USSAServer.MyInformation
'Public clsSAWeb As Object
Public Declare Function ReplyMessage Lib "user32" (ByVal lReply As Long) As Long
Public Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)

Public Function GetVoucherInfo(strVouchType As String, Conn As ADODB.Connection) As String

'by ahzzd
'根据单据类型得到 业务单据主子表信息

    Select Case LCase(strVouchType)
    
        Case "efzz0404"
            GetVoucherInfo = "<Data maintbl='EFBWGL_SelDeclare' detailtbl='efzz_usertbl' mainkey='id' mainview='efzz_v_fitemss97' vouchtype='efzz04041' ccode='ccode' savouchtype='efzz04041' bctrlcredit='' credittype='' >"
        Case LCase("EFBWGL020301")
             GetVoucherInfo = "<Data maintbl='EFBWGL_SelDeclare' detailtbl='' mainkey='id' mainview='EFBWGL_v_SelDeclareT' vouchtype='EFBWGL020301' ccode='ccode' savouchtype='EFBWGL020301' bctrlcredit='' credittype='' >"
        Case Else
            Dim strTblName As String
            Dim strBodyTable As String
            Dim strMainViewName As String
            Dim strErr As String
            Dim cls_Public As Object
            Set cls_Public = CreateObject("UF_Public_base.clsSysInfo")
            cls_Public.GetVoucherTable strVouchType, strErr, Conn, strTblName, strBodyTable, strMainViewName
            strMainKeyName = "id"
            strSaVouchType = "VoucherType"
            GetVoucherInfo = "<Data maintbl='" & strTblName & "' detailtbl='" & strBodyTable & "' mainkey='id' mainview='" & strMainViewName & "' vouchtype='" & strVouchType & "' ccode='ccode' savouchtype='" & strVouchType & "' bctrlcredit='' credittype='' >"
    End Select
    GetVoucherInfo = GetVoucherInfo & " </Data>"
End Function
Public Function GetVouchInfoEx(Cnn As ADODB.Connection, strVouchType As String, strVouchID As String) As String
    Dim rst As New ADODB.Recordset
    Dim strType As String
    
    rst.CursorLocation = adUseClient
    strType = ""
    GetVouchInfoEx = GetVoucherInfo(strVouchType, Cnn)
    If strVouchType = "01" Then
        rst.Open "select cvouchtype from dispatchlist where dlid=" & strVouchID, Cnn, adOpenForwardOnly, adLockReadOnly
        If Not rst.EOF Then
            strType = rst(0)
        End If
        rst.Close
        GetVouchInfoEx = Replace(GetVouchInfoEx, "savouchtype='05'", "savouchtype='" + strType + "'")
        If strType = "06" Then
            GetVouchInfoEx = Replace(GetVouchInfoEx, "bctrlcredit='bCreditDisp'", "bctrlcredit='bCreditDispWT'")
        End If
    End If
    Set rst = Nothing
End Function
'Public Sub CreateBaseObject(calledCtx As UFSoft_U8_Framework_LoginContext.calledContext, ByRef login As U8Login.clsLogin, ByRef Conn As ADODB.Connection)
'        If login Is Nothing Then
'            Set login = New U8Login.clsLogin
'            login.ConstructLogin calledCtx.token
'            login.TaskId = calledCtx.TaskId
'            login.login "SA"
'        End If
'        Set Conn = New Connection
'        Conn.Open login.UfDbName
'End Sub
'Public Sub RemoveBaseObject(ByRef login As U8Login.clsLogin, ByRef Conn As ADODB.Connection)
'    Set Conn = Nothing
'    Set login = Nothing
'End Sub
Public Function MsgBox(ByVal sPrompt As String, Optional ByVal enumButtons As VbMsgBoxStyle = vbOKOnly, Optional ByVal cTitle As String = "", Optional ByVal cHelpFile As String = "", Optional ByVal context = "") As VbMsgBoxResult
    On Error Resume Next
    Call ReplyMessage(1)
    MsgBox = MuMsgBox(sPrompt, enumButtons, GetString("U8.SA.xsglsql.frmMain.00609"), cHelpFile, context)
    
End Function
Public Function writeLog(str As String)
    OutputDebugString str
'    Dim fs As Object
'    Dim oLogFile As Object
'    Set fs = CreateObject("Scripting.FileSystemObject")
'    If Dir(App.Path & "\sawfLogs") = "" Then
'        Set oLogFile = fs.CreateTextFile(App.Path & "\sawfLogs", True, True)
'    Else
'        Set oLogFile = fs.OpenTextFile(App.Path & "\sawfLogs", ForAppending, False, TristateTrue)
'    End If
'
'    Call oLogFile.WriteLine(CStr(DateTime.Now) & "    " & str)
'    oLogFile.Close
'    Set oLogFile = Nothing
End Function

'判断当前单据是否关闭
Public Function ErrVoucherClosed(vouchId As String, vouchType As String, Conn As ADODB.Connection) As String
    Dim sql As String
    ErrVoucherClosed = ""
    Dim rst As New ADODB.Recordset
    Select Case vouchType
    Case "01"
        sql = "select ccloser from dispatchlist where isnull(ccloser,N'')=N'' and dlid=" + vouchId
    Case "16"
        sql = "select ccloser from sa_quomain where isnull(ccloser,N'')=N'' and id=" + vouchId
    Case "17"
        sql = "select ccloser from so_somain where isnull(ccloser,N'')=N'' and id=" + vouchId
    Case Else
        GoTo exitsub:
    End Select
    rst.CursorLocation = adUseClient
    rst.Open sql, Conn, adOpenForwardOnly, adLockReadOnly
    If rst.BOF And rst.EOF Then
        ErrVoucherClosed = GetString("U8.SA.xsglsql_2.saworkflowsrv.019") '该单据已经关闭,不能审核或者弃审!
    End If
    rst.Close
exitsub:
    Set rst = Nothing
End Function
Public Function getAccinformation(CN As ADODB.Connection, strSysID As String, strName As String) As String
    Dim rst As New ADODB.Recordset
    
    rst.CursorLocation = adUseClient
    rst.Open "Select cValue from accinformation where cSysID=N'" & strSysID & "' and cName=N'" & strName & "'", CN, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rst.EOF Then
        getAccinformation = ""
    Else
        If IsNull(rst(0)) Then
            getAccinformation = ""
        Else
            getAccinformation = rst(0)
        End If
    End If
    rst.Close
    Set rst = Nothing
End Function

Public Function GetVoucherMoney(CN As ADODB.Connection, strBodyTable As String, strMainKeyName As String, VoucherId As String) As Double
    Dim rst As New ADODB.Recordset
    Dim sKeyRst As String
    
    rst.CursorLocation = adUseClient
    GetVoucherMoney = 0
    Select Case strSaVouchType
        Case "05", "06"
            sKeyRst = "isnull(sum(case when isnull(iQuantity,0)<>0  then (isnull((iQuantity-isnull(iSettleQuantity,0)-isnull(iRetQuantity,0))*((isnull(iNatSum,0))/iQuantity),0)) "
            sKeyRst = sKeyRst & "  else case when isnull(tbquantity,0)<>0 then (isnull((tbquantity-isnull(iSettleQuantity,0)-isnull(iRetQuantity,0))*((isnull(iNatSum,0))/tbquantity),0))"
            sKeyRst = sKeyRst & "  else inatsum-isnull(isettlenum,0)*inatsum/(case when isnull(isum,0)=0 then 1 else isum end)"
            sKeyRst = sKeyRst & "  End  end ),0)"
            rst.Open "select " & sKeyRst & " as inatSum,count(idlsid) as ReCount from " & strBodyTable & " where " & strMainKeyName & "= " & VoucherId, CN, adOpenForwardOnly, adLockReadOnly
        Case "97"
            sKeyRst = "( isnull(sum(case isnull(iQuantity,0) when 0 then (case  when isnull(iFhmoney,0)-isnull(iKPmoney,0)>=0 then (isum-isnull(ifhmoney,0))"
            'sKeyRst = sKeyRst & " else (isum-isnull(ikpmoney,0)) End * inatsum/isnull(isum,1)) "
            sKeyRst = sKeyRst & " else (isum-isnull(ikpmoney,0)) End * inatsum/(case isnull(isum,0) when 0 then 1 else isum end)) "
            sKeyRst = sKeyRst & " else (isnull(inatsum,0))/isnull(iQuantity,1) * (isnull(iQuantity,0)-(case when (isnull(ifhQuantity,0)-isnull(ikpQuantity,0))>0 then isnull(ifhQuantity,0) else isnull(ikpQuantity,0) end) )"
            sKeyRst = sKeyRst & "  End ),0)) "
            rst.Open "select " & sKeyRst & " as inatSum,count(autoid) as ReCount from " & strBodyTable & " where " & strMainKeyName & " = " & VoucherId, CN, adOpenForwardOnly, adLockReadOnly      '& " and isnull(isum,0) <>0"
        Case "26", "27", "28", "29"
            rst.Open "select isnull(sum(isnull(" & sKey & ",0)-isnull(iMoneySum,0)),0) as inatSum,count(autoid) as ReCount from " & strBodyTable & " where " & strMainKeyName & " = " & VoucherId, CN, adOpenForwardOnly, adLockReadOnly
        Case "07"
            rst.Open "select isnull(sum(isnull(inatSum,0)),0) as inatSum,count(inatSum) as ReCount from " & strBodyTable & " where " & strMainKeyName & " = " & VoucherId, CN, adOpenForwardOnly, adLockReadOnly
        Case Else
            rst.Open "select isnull(sum(isnull(inatmoney,0)),0) as inatSum,count(inatmoney) as ReCount from " & strBodyTable & " where " & strMainKeyName & " = " & VoucherId, CN, adOpenForwardOnly, adLockReadOnly
    End Select
    If Not rst.EOF Then
        GetVoucherMoney = rst(0)
    End If
    rst.Close
    Set rst = Nothing
End Function
Public Function CheckCredit(CheckType As String, login As U8Login.clsLogin, Conn As ADODB.Connection, VoucherType As String, VoucherId As String, strCreditInfo As String, errMsg As String) As Double
    Dim blnARStart As Boolean
    Dim bCredit As Boolean
    Dim dblTotal As Double
    Dim dom As New DOMDocument
    Dim rst As New ADODB.Recordset
    Dim strMainName As String
    Dim strMainKeyName As String
    Dim strCuscode As String
    Dim strVouCtrol As String
    Dim strBodyTable As String
    Dim strVoucherInfo As String
    Dim intVouchType As UFCreditSA.vouchType
    Dim objCheckCredit As New UFCreditSA.UsSaleCredit
    Dim strCreditType As String
    Dim strCodeName As String
    
    On Error GoTo ErrHandler
    rst.CursorLocation = adUseClient
    strVoucherInfo = GetVouchInfoEx(Conn, VoucherType, VoucherId)
    If strVoucherInfo = " </Data>" Then GoTo ErrHandler
    dom.LoadXML strVoucherInfo
    strBodyTable = dom.documentElement.Attributes.getNamedItem("detailtbl").nodeValue
    strMainKeyName = dom.documentElement.Attributes.getNamedItem("mainkey").nodeValue
    strMainName = dom.documentElement.Attributes.getNamedItem("maintbl").nodeValue
    strVouCtrol = dom.documentElement.Attributes.getNamedItem("bctrlcredit").nodeValue
    intVouchType = dom.documentElement.Attributes.getNamedItem("credittype").nodeValue
    writeLog ("账期自动服务:参数获得正常")
    Select Case CheckType
        Case "1"
            strCreditType = "bCredit"
            strCodeName = "ccuscode"
        Case "2"
            strCreditType = "bCrCheckPe"
            strCodeName = "cpersoncode"
        Case "3"
            strCreditType = "bCrCheckDe"
            strCodeName = "cdepcode"
    End Select
    rst.Open "select " + strCodeName + " from " & strMainName & " where " & strMainKeyName & "=" & VoucherId, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rst.BOF And rst.EOF Then
        errMsg = GetString("U8.SA.USSASERVER.clsvouchsave.00542") 'zh-CN：该单据已经不存在或已被其他人修改
        GoTo ErrHandler
        Exit Function
    End If
    strCuscode = rst(0)
    rst.Close
    blnARStart = IIf(LCase(getAccinformation(Conn, "SA", strVouCtrol)) = "true", True, False)
    If Not blnARStart Then
        Exit Function
    End If
    Set dom = Nothing
    
    blnARStart = IIf(Trim(getAccinformation(Conn, "AR", "dARStartDate")) <> "", True, False)
    Call objCheckCredit.init(Conn)
    bCredit = IIf(LCase(getAccinformation(Conn, "SA", strCreditType)) = "true", True, False)         '是否有客户信用额度的控制
    Dim cCrCheckFunction As String
    cCrCheckFunction = getAccinformation(Conn, "SA", "cCrCheckFunction")       '信用检查公式
    dblTotal = 0
    dblTotal = GetVoucherMoney(Conn, strBodyTable, strMainKeyName, VoucherId)
    errMsg = ""
    If dblTotal > 0 Then
        If bCredit Then
            CheckCredit = objCheckCredit.SACreCheck(Conn, Val(CheckType), strCuscode, cCrCheckFunction, dblTotal, intVouchType, VoucherId, False, strCreditInfo, errMsg)
            If errMsg <> "" Then
                GoTo ErrHandler
            End If
        End If
    End If
    Set objCheckCredit = Nothing
    Set dom = Nothing
    Set rst = Nothing
    Exit Function
ErrHandler:
    Set dom = Nothing
    Set rst = Nothing
    Set objCheckCredit = Nothing
    writeLog ("ErrHandler" & VBA.Err.Description)
    On Error GoTo 0
    Err.Raise vbObjectError + 1000, "", Err.Description & errMsg
End Function
Public Function CheckCreditDate(CheckType As String, login As U8Login.clsLogin, Conn As ADODB.Connection, VoucherType As String, VoucherId As String, strCreditInfo As String, errMsg As String) As Long
    Dim blnARStart As Boolean
    Dim bCredit As Boolean
    Dim dblTotal As Double
    Dim dom As New DOMDocument
    Dim rst As New ADODB.Recordset
    Dim strMainName As String
    Dim strMainKeyName As String
    Dim strCuscode As String
    Dim strVouCtrol As String
    Dim strBodyTable As String
    Dim strVoucherInfo As String
    Dim intVouchType As UFCreditSA.vouchType
    Dim objCheckCredit As New UFCreditSA.UsSaleCredit
    Dim strCreditType As String
    Dim strCodeName As String
    Dim strDate As String
    
    On Error GoTo ErrHandler
    rst.CursorLocation = adUseClient
    strVoucherInfo = GetVouchInfoEx(Conn, VoucherType, VoucherId)
    If strVoucherInfo = " </Data>" Then GoTo ErrHandler
    dom.LoadXML strVoucherInfo
    strBodyTable = dom.documentElement.Attributes.getNamedItem("detailtbl").nodeValue
    strMainKeyName = dom.documentElement.Attributes.getNamedItem("mainkey").nodeValue
    strMainName = dom.documentElement.Attributes.getNamedItem("maintbl").nodeValue
    strVouCtrol = dom.documentElement.Attributes.getNamedItem("bctrlcredit").nodeValue
    intVouchType = dom.documentElement.Attributes.getNamedItem("credittype").nodeValue
    writeLog ("账期自动服务:参数获得正常")
    Select Case CheckType
        Case "1"
            strCreditType = "bCredit"
            strCodeName = "ccuscode"
        Case "2"
            strCreditType = "bCrCheckPe"
            strCodeName = "cpersoncode"
        Case "3"
            strCreditType = "bCrCheckDe"
            strCodeName = "cdepcode"
    End Select
    rst.Open "select " + strCodeName + ",ddate from " & strMainName & " where " & strMainKeyName & "=" & VoucherId, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rst.BOF And rst.EOF Then
        errMsg = GetString("U8.SA.USSASERVER.clsvouchsave.00542") 'zh-CN：该单据已经不存在或已被其他人修改
        GoTo ErrHandler
        Exit Function
    End If
    strCuscode = rst(0)
    strDate = rst(1)
    rst.Close
    blnARStart = IIf(LCase(getAccinformation(Conn, "SA", strVouCtrol)) = "true", True, False)
    If Not blnARStart Then
        Exit Function
    End If
    Set dom = Nothing
    
    blnARStart = IIf(Trim(getAccinformation(Conn, "AR", "dARStartDate")) <> "", True, False)
    Call objCheckCredit.init(Conn)
    bCredit = IIf(LCase(getAccinformation(Conn, "SA", strCreditType)) = "true", True, False)         '是否有客户信用额度的控制
    Dim cCrCheckFunction As String
    Dim bCrCheckDateAcc As Boolean
    bCrCheckDateAcc = IIf(LCase(getAccinformation(Conn, "SA", "bCrCheckDateAcc")) = "true" Or getAccinformation(Conn, "SA", "bCrCheckDateAcc") = "1", True, False)       '是否有客户信用额度的控制
    If bCrCheckDateAcc Then
        cCrCheckFunction = getAccinformation(Conn, "SA", "cCrChAccDateFunction")       '信用检查公式
    Else
        cCrCheckFunction = getAccinformation(Conn, "SA", "cCrChDateFunction")       '信用检查公式
    End If
    dblTotal = 0
    dblTotal = GetVoucherMoney(Conn, strBodyTable, strMainKeyName, VoucherId)
    errMsg = ""
    If dblTotal > 0 Then
        If bCredit Then
            If bCrCheckDateAcc Then
                writeLog ("SACreDateCheckAcc")
                CheckCreditDate = objCheckCredit.SACreDateCheckAcc(Conn, Val(CheckType), strCuscode, cCrCheckFunction, DateValue(strDate), intVouchType, VoucherId, False, strCreditInfo, errMsg)
            Else
                writeLog ("SACreDateCheck")
                CheckCreditDate = objCheckCredit.SACreDateCheck(Conn, Val(CheckType), strCuscode, cCrCheckFunction, DateValue(strDate), intVouchType, VoucherId, False, strCreditInfo, errMsg)
            End If
            If errMsg <> "" Then
                GoTo ErrHandler
            End If
        End If
    End If
    Set objCheckCredit = Nothing
    Set dom = Nothing
    Set rst = Nothing
    Exit Function
ErrHandler:
    Set dom = Nothing
    Set rst = Nothing
    Set objCheckCredit = Nothing
    writeLog ("ErrHandler" & VBA.Err.Description)
    On Error GoTo 0
    Err.Raise vbObjectError + 1000, "", Err.Description & errMsg
End Function


