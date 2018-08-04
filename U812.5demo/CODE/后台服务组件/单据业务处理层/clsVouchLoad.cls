VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVouchLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
Public clsSys As New clsSystem
Dim m_VoucherID As Variant
Dim strReturn As String, strWhereVType As String
Dim sKeyHead As String     '表头关键字
Dim sKeyBody As String     '表体和表头关联关键字
Dim sBodyID As String      ''子表关键字
Dim sViewHead As String    '表头视图
Dim sViewBody As String    '表体视图
Public bHaveBodyGrid As Boolean
Private sMViewHead As String, sMViewBody As String
Private sLookHView As String, strLookWhere As String, strLookAuth As String, sLookHViewOld As String
Private sTableHead As String, sTableBody As String
Dim m_BOF As Boolean, m_EOF As Boolean
Dim bSTR As Boolean        ''判断关键字是否为文本型
Dim bFirst As Boolean
Private bAR As Boolean
Private LMaxID As Variant, LMinId As Variant     '取得最大最小id
Private strARAuthC As String  ''应收权限串

'Public Enum enuVouType
'    MMT01 = "87"
'    MMT02 = "88"
'    MMT03 = "89"
'    MMT04 = "90"
'    MMT05 = "91"
'    MMT06 = "92"
'    MMT07 = "93"
'    MMT08 = "94"
'    MMT09 = "95"
'    MMT11 = "96"
'    MMT12 = "97"
'    MMT66 = "98"
'End Enum
'Private MuvdVouchType As enuVouType
Public strVouchType   As String  ' 单据类型 CardNumbers
Public m_Login As U8Login.clsLogin

Public Function GetVoucherDataWEB(domHead As DOMDocument, domBody As DOMDocument, _
        ByVal m_Conn As Connection, ByVal CardNumbers As String, _
        ByVal bReturn As Boolean, ErrMsg As String, Optional ByVal varId As Variant, Optional blnAuth) As Boolean
    Dim rstTmp As New ADODB.Recordset, strMaxsql As String, strwhere As String
    Dim rstHeadTmp As New Recordset
    Dim rstBodyTmp As New Recordset
    Dim strHeadSql As String, strBodySql As String
    Dim strAuth As String
    Dim strAuth_B As String
    Dim clsAuth As New clsUserInterfaceInit  ''   U8RowAuthsvr.clsRowAuth
    
    On Error GoTo err
    
    ErrMsg = ""
    GetVoucherDataWEB = True
    
    'by ahzzd 20100926 增加记录权限功能
    clsAuth.Init m_Conn, clsSys.CurrentUserID, "voucher", CardNumbers
    strAuth = clsAuth.GetAuthStringByOperation("search", "T")
    strAuth_B = clsAuth.GetAuthStringByOperation("search", "B") '取表体数据权限
    
    
    bSTR = getkey(CardNumbers, bReturn, m_Conn)    '//''获取关键字,判断关键字是否为文本
    If Not IsMissing(varId) And Not IsNull(varId) Then
         m_VoucherID = varId
    Else
        If m_VoucherID = "" Or Val(m_VoucherID) = 0 Then
            MoveTo "last", m_Conn, domHead, domBody
            Exit Function
        End If
    End If
        
    If Not bSTR Then            '' 关键字是数字型的
        strHeadSql = "Select * from " & sViewHead & " where " & sKeyHead & " = " & Val(m_VoucherID)
        strBodySql = "Select * ,'' as editprop From " & sViewBody & " where " & sKeyBody & " = " & Val(m_VoucherID)
    Else                        '' 关键字是文本型的
        strHeadSql = "Select * from " & sViewHead & " where " & sKeyHead & " ='" & m_VoucherID & "'"
        strBodySql = "Select * ,'' as editprop From " & sViewBody & " where " & sKeyBody & " = '" & m_VoucherID & "'"
    End If
    
    If Trim(strAuth) <> "" Then strHeadSql = strHeadSql & " and " & strAuth
    If Trim(strAuth_B) <> "" Then strBodySql = strBodySql & " and " & strAuth_B
     
     
    rstHeadTmp.CursorLocation = adUseClient
    rstBodyTmp.CursorLocation = adUseClient
    rstHeadTmp.Open strHeadSql, m_Conn, 3, 4
    rstHeadTmp.Save domHead, adPersistXML
    If rstHeadTmp.EOF And Val(m_VoucherID) <> 0 Then
        ErrMsg = "单据不存在或者已经被删除或者没有权限！"
        GetVoucherDataWEB = False
        strBodySql = strBodySql & " and 1=2 "
    End If
    rstHeadTmp.Close
    
    strBodySql = strBodySql + " order by " + sBodyID
    If Trim(sViewBody) <> "" Then  '
        rstBodyTmp.Open strBodySql, m_Conn, 3, 4
        rstBodyTmp.Save domBody, adPersistXML
        rstBodyTmp.Close
    End If
    Set rstHeadTmp = Nothing
    Set rstBodyTmp = Nothing
    
    If m_VoucherID = LMaxID Then
        If m_VoucherID = LMinId Then
            m_EOF = True
            m_BOF = True
        Else
            m_EOF = True
            m_BOF = False
        End If
    Else
        If m_VoucherID = LMinId Then
            m_EOF = False
            m_BOF = True
        Else
            m_EOF = False
            m_BOF = False
        End If
    End If
    Call SetSum(domHead, domBody)
    Exit Function
err:
    If err.Number > 0 Then
        ErrMsg = err.Description
    End If
    GetVoucherDataWEB = False
    Set rstHeadTmp = Nothing
    Set rstBodyTmp = Nothing
End Function



''取得表头，表体参数
Private Function getkey(ByVal CardNumbers As String, ByVal bReturn As Boolean, CN As ADODB.Connection) As Boolean
    Dim sql As String
    Dim strSQL As String
    getkey = False
    bAR = False
'    MuvdVouchType = uvdVouchType
    strReturn = "cvouchtype ='" & CardNumbers & "'"
On Error GoTo err
    Select Case CardNumbers
    '/////////////////////////////////////////////////////////////////////
    '个性化设置处理
        Case "94"            '支出借款单
            strVouchType = "94"
            strReturn = strReturn
            sKeyHead = "id"
            sKeyBody = "id"
            strWhereVType = ""
            sMViewHead = "MT_budget"
            sMViewBody = "MT_budget"
            sViewHead = "V_MT_budget04"
            sViewBody = "V_MT_budget04"
            sKeyHead = "id"
            sKeyBody = "id"
            sBodyID = "id"
            
        Case Else
            Call Getvoucherinf(CN, CardNumbers, sKeyHead, sKeyBody, sMViewHead, sMViewBody, sViewHead, sViewBody, sBodyID)
    End Select
    sTableHead = sMViewHead
    sTableBody = sMViewBody
    ''优化翻页效率
    sLookHViewOld = sViewHead  '主表视图
    sLookHView = sLookHViewOld
    strLookWhere = strWhereVType
    
    
    'shanlan add 090811---------------------------------------------------------------
    Dim clsAuth As New U8RowAuthsvr.clsRowAuth
    Dim strAuth As String
    If clsSys.bAdmin Then Exit Function
    clsAuth.Init CN.ConnectionString, clsSys.CurrentUserID
    sLookHView = ""
    
    Dim Rs As New ADODB.Recordset
    Dim intCyc As Long
    Dim strInSql As String

    
    '项目
    If clsSys.bAuth_Inv Then
        strAuth = clsAuth.getAuthString("fitem", "", "R")
        strLookAuth = strAuth
        If strAuth = "" Then
            strAuth = " 1= 1 "
            strLookAuth = strAuth
        ElseIf Trim(strAuth) <> "1=2" Then
                strLookAuth = "(citemcode =N'' or citemcode is null or citemcode=authFitem.cACCode )"
                sLookHView = sLookHView + " left join (" + strAuth + ") as authFitem on  citemcode =authFitem.cACCode"
                strAuth = "(isnull(citemcode,N'')=''  or citemcode in (" + strAuth + "))"
        ElseIf Trim(strAuth) = "1=2" Then
            strLookAuth = "( isnull(citemcode,N'')=N''  )"
            sLookHView = sLookHView
            strAuth = "(isnull(citemcode,N'')=N''  )"
        End If
        If strAuth <> " 1= 1 " Then
            strWhereVType = strWhereVType + IIf(strWhereVType = "", " ", " and ") + strAuth
            strLookWhere = strLookWhere + IIf(strLookWhere = "", " ", " and ") + strLookAuth
        End If
    End If
    sLookHView = sLookHViewOld + sLookHView
    Set clsAuth = Nothing
    If sLookHView = "" Then
        sLookHView = sMViewHead
    End If
    '----------------------------------------------------------------------------------------------
Exit Function
err:
    If rds_M.State <> 0 Then rds_M.Close
    MsgBox "单据预制存在错误，请检查!", vbOKOnly, "提示" & err.Description
End Function

Public Sub Getvoucherinf(CN As ADODB.Connection, CardNumber As String, _
                            Optional sKeyHead As String, Optional sKeyBody As String, _
                            Optional BtTblname As String, Optional BwTblname As String, _
                            Optional sViewHead As String, Optional sViewBody As String, Optional sBodyID As String)
    Dim rd As New ADODB.Recordset
    rd.CursorLocation = adUseClient
    rd.Open "select * from Vouchers_base where CardNumber = '" & CardNumber & "' ", CN, 3, 4
    '主表关键字
    sKeyHead = "id"
    
    '子表和主表关联关键字
    sKeyBody = "id"
    
    '子表关键字
    sBodyID = "autoid"
    
    If rd.RecordCount > 0 Then
        '主表表名
        BtTblname = rd.Fields("BTTblName")
        
        '子表表名
        BwTblname = IIf(IsNull(rd.Fields("BWTblName")), "", rd.Fields("BWTblName"))
        
        '主表视图名
        sViewHead = rd.Fields("BTQName")
        
        '子表视图名
        sViewBody = IIf(IsNull(rd.Fields("BWQName")), "", rd.Fields("BWQName"))
        
        '是否有表体
        bHaveBodyGrid = IIf(LCase(rd("HaveBodyGrid") & "") = "true", True, False)
    End If
    Set rd = Nothing
End Sub

Public Sub GetPosition(m_Conn As Connection)
    
    Dim rstTmp As New ADODB.Recordset
    Dim vTmpID As Variant   '临时id
    Dim strTmpID As String
    Dim strwhere As String, strWhere1 As String, strWhere2 As String, _
    strOrder As String
    Dim strMaxID As String, strHeadSql As String, strBodySql As String
        If strVouchType <> "" Then strWhere1 = " cVouchType ='" & strVouchType & "'"
        If strReturn <> "" Then
            If strWhere1 = "" Then
                strWhere1 = strReturn
            Else
                strWhere1 = strWhere1 + " and " + strReturn
            End If
        End If
        
    vTmpID = m_VoucherID
    
    strMaxID = "select isnull(Max(" & sKeyHead & "),0) as MaxID,isnull(Min(" & sKeyHead & "),0) as MinID from " & sViewHead
    If strWhere1 <> "" Then strMaxID = strMaxID + " where " + strWhere1
    
    rstTmp.Open strMaxID, m_Conn, 3, 1
    If vTmpID = rstTmp(0) Or Val(vTmpID) = Val(rstTmp(0)) Then
        If vTmpID = rstTmp(1) Or Val(vTmpID) = Val(rstTmp(1)) Then
            m_EOF = True
            m_BOF = True
        Else
            m_EOF = True
            m_BOF = False
        End If
    Else
        If vTmpID = rstTmp(1) Or Val(vTmpID) = Val(rstTmp(1)) Then
            m_EOF = False
            m_BOF = True
        Else
            m_EOF = False
            m_BOF = False
        End If
    End If
End Sub


Public Function MoveTo(strMoveTo As String, ByVal m_Conn As ADODB.Connection, domHead As DOMDocument, domBody As DOMDocument)
    Dim rstHeadTmp As New Recordset
    Dim rstBodyTmp As New Recordset
    Dim rstTmp As New ADODB.Recordset
    Dim vTmpID As Variant   '临时id
    Dim strTmpID As String
    Dim strwhere As String, strWhere1 As String, strWhere2 As String, _
    strOrder As String, strWhereAuth As String
    Dim strMaxID As String, strHeadSql As String, strBodySql As String
'    Dim clsAuth As New U8RowAuthsvr.clsRowAuth
    Dim clsAuth As New clsUserInterfaceInit  ''   U8RowAuthsvr.clsRowAuth
    Dim strAuth As String
    Dim strAuth_B As String
    
    Dim ErrMsg As String
    On Error GoTo DoErr
    'by ahzzd 20100926 增加记录权限功能
    clsAuth.Init m_Conn, clsSys.CurrentUserID, "voucher", strVouchType
    strAuth = clsAuth.GetAuthStringByOperation("search", "T")
    strAuth_B = clsAuth.GetAuthStringByOperation("search", "B") '取表体数据权限
    strWhere1 = strAuth
    
    
    If strReturn <> "" Then
        If strWhere1 = "" Then
            strWhere1 = strReturn
        Else
            strWhere1 = strWhere1 + " and " + strReturn
        End If
    End If
    If strLookWhere <> "" Then
        If strWhere1 = "" Then
            strWhere1 = strLookWhere
        Else
            strWhere1 = strWhere1 + " and " + strLookWhere
        End If
    End If
    
    

    
    
    Select Case LCase(strMoveTo)
        Case "first"
            strWhere2 = ""
            If clsSys.bVouOrderByDate Then
                strOrder = " order by ddate asc ," & sKeyHead & " ASC "
            Else
                strOrder = " order by " & sKeyHead & " ASC "
            End If
        Case "last"
            strWhere2 = ""
            If clsSys.bVouOrderByDate Then
                strOrder = " order by ddate DESC," & sKeyHead & " DESC "
            Else
                strOrder = " order by " & sKeyHead & " DESC "
            End If
        Case "previous"
            If bSTR Then
                strWhere2 = sKeyHead & "< '" & m_VoucherID & "'"
            Else
                strWhere2 = sKeyHead & "< " & m_VoucherID
            End If
            If clsSys.bVouOrderByDate Then
                strOrder = " order by ddate DESC," & sKeyHead & " DESC "
            Else
                strOrder = " order by " & sKeyHead & " DESC "
            End If
            'strOrder = " order by " & sKeyHead & " DESC "
        Case "next"
            If bSTR Then
                strWhere2 = sKeyHead & "> '" & Val(m_VoucherID) & "'"
            Else
                strWhere2 = sKeyHead & "> " & Val(m_VoucherID)
            End If
            If clsSys.bVouOrderByDate Then
                strOrder = " order by ddate ASC," & sKeyHead & " ASC "
            Else
                strOrder = " order by " & sKeyHead & " ASC "
            End If
    End Select
    If strWhere1 <> "" Then
        If strWhere2 <> "" Then
            strwhere = "where " + strWhere1 + " and " + strWhere2
        Else
            strwhere = "where " + strWhere1
        End If
    Else
        If strWhere2 <> "" Then
            strwhere = "where " + strWhere2
        End If
    End If
    If LCase(strMoveTo) = "previous" Or LCase(strMoveTo) = "next" Then
        'strTmpID = "select top 1 " & sKeyHead & " from " & sMViewHead _
                        & " " & strwhere & " " & strOrder
        strTmpID = "select top 1 " & sKeyHead & " from " & sLookHView & " " & strwhere & " " & strOrder
        If rstTmp.State = 1 Then rstTmp.Close
        rstTmp.Open strTmpID, m_Conn, 3, 1
        If Not rstTmp.EOF Then
            vTmpID = rstTmp(0)
        Else
            vTmpID = m_VoucherID
            'strMoveTo = "last"   ''CLin 09-12-15 去掉向上翻页时可从第一张直接翻到最后一张单据的情况
        End If
        rstTmp.Close
    End If
    If clsSys.bVouOrderByDate Then
        strMaxID = "select isnull(Max(" & sKeyHead & "),0) as MaxID,isnull(Min(" & sKeyHead & "),0) as MinID from ( "
        If strWhere1 <> "" Then
            strMaxID = strMaxID + "select * from (select top 1 " & sKeyHead & ",ddate from  " & sLookHView & " where " + strWhere1 + " order by ddate asc ," & sKeyHead & " asc ) as a union all "
            strMaxID = strMaxID + "select * from (select top 1 " & sKeyHead & ",ddate from  " & sLookHView & " where " + strWhere1 + " order by ddate desc," & sKeyHead & " desc) as b  ) as  " & sLookHView
        Else
            strMaxID = strMaxID + "select * from (select top 1 " & sKeyHead & ",ddate from  " & sLookHView & " order by ddate asc," & sKeyHead & "  asc ) as a union all "
            strMaxID = strMaxID + "select * from (select top 1 " & sKeyHead & ",ddate from  " & sLookHView & " order by ddate desc," & sKeyHead & " desc) as b  ) as  " & sLookHView
        End If
    Else
        strMaxID = "select isnull(Max(" & sKeyHead & "),0) as MaxID,isnull(Min(" & sKeyHead & "),0) as MinID from " & sLookHView
        If strWhere1 <> "" Then
            strMaxID = strMaxID + "  WHERE " + sKeyHead + " >0 and (" + strWhere1 + ")"
        Else
            strMaxID = strMaxID + "  WHERE " + sKeyHead + " >0 "
        End If
    End If
    
'    If uvdVouchType = "96" Then
    
    rstTmp.Open strMaxID, m_Conn, 3, 1
    Select Case LCase(strMoveTo)
        Case "first"
            vTmpID = rstTmp(1)
        Case "last"
            vTmpID = rstTmp(0)
    End Select
    LMaxID = rstTmp(0)
    LMinId = rstTmp(1)
    If vTmpID = rstTmp(0) Then
        If vTmpID = rstTmp(1) Then
            m_EOF = True
            m_BOF = True
        Else
            m_EOF = True
            m_BOF = False
        End If
    Else
        If vTmpID = rstTmp(1) Then
            m_EOF = False
            m_BOF = True
        Else
            m_EOF = False
            m_BOF = False
        End If
    End If
    m_VoucherID = vTmpID
    
    If Not bSTR Then            '' 关键字是数字型的
        strHeadSql = "Select * from " & sViewHead & " where " & sKeyHead & "=" & m_VoucherID
        strBodySql = "Select * ,'' as editprop From " & sViewBody & " where " & sKeyBody & " = " & m_VoucherID
    Else                        '' 关键字是文本型的
        strHeadSql = "Select * from " & sViewHead & " where " & sKeyHead & "='" & m_VoucherID & "'"
        strBodySql = "Select * ,'' as editprop From " & sViewBody & " where " & sKeyBody & " = '" & m_VoucherID & "'"
    End If
    
    
    If Trim(strAuth_B) <> "" Then
        strBodySql = strBodySql & " and " & strAuth_B
    End If
    
    rstHeadTmp.CursorLocation = adUseClient
    rstBodyTmp.CursorLocation = adUseClient
    rstHeadTmp.Open strHeadSql, m_Conn, 3, 4
    rstHeadTmp.Save domHead, adPersistXML
    
'    If UCase(strVouchType) <> "EFBWGL020601" Then
'        If strVouchType <> "EFBWGL020301" Or strVouchType <> "EFBWGL020401" Or strVouchType <> "EFBWGL020501" Or strVouchType <> "EFBWGL020502" Or strVouchType <> "EFBWGL020601" Or strVouchType <> "EFYZGL030201" Or strVouchType <> "EFYZGL030303" Then
            If sViewBody <> "" And bHaveBodyGrid = True Then rstBodyTmp.Open strBodySql, m_Conn, 3, 4
            If sViewBody <> "" And bHaveBodyGrid = True Then rstBodyTmp.Save domBody, adPersistXML
'        End If
'    End If
    
    If Not rstHeadTmp.EOF Then
        m_VoucherID = rstHeadTmp.Fields(sKeyHead)
    End If
    
    Call SetSum(domHead, domBody)
    rstHeadTmp.Close
    If rstBodyTmp.State <> 0 Then
      rstBodyTmp.Close
    End If
    Set rstHeadTmp = Nothing
    Set rstBodyTmp = Nothing
    Exit Function
DoErr:
    MoveTo = err.Description
    Set rstHeadTmp = Nothing
    Set rstBodyTmp = Nothing
End Function

Private Sub SetSum(domHead As DOMDocument, domBody As DOMDocument)
    Dim ele As IXMLDOMElement, NdList As IXMLDOMNodeList
    Dim iSum As Variant ' Double
    Dim strSumDX As Variant
    If strVouchType = "97" Then Exit Sub
    strSumDX = ""
    iSum = 0
    Set NdList = domBody.selectNodes("//z:row")
    For Each ele In NdList
        iSum = iSum + CDec(Val(IIf(IsNull(ele.getAttribute("isum")), 0, ele.getAttribute("isum"))))
    Next
    Num2Chinese Format(iSum, "#.00"), strSumDX
    Set ele = domHead.selectSingleNode("//z:row")
    If Not ele Is Nothing Then
        ele.setAttribute "isumdx", strSumDX
        ele.setAttribute "isumx", iSum
        ele.setAttribute "zdsumdx", strSumDX
        ele.setAttribute "zdsum", iSum
    End If
    Set NdList = Nothing
    Set ele = Nothing
End Sub

'功 能: 保存单据
Public Function VouchSave(m_Conn As Connection, ByVal strCardNum As String, _
        ByVal domHead As DOMDocument, ByVal domBody As DOMDocument, _
        ByVal VoucherState As Integer, sErrMsg As String, Optional vNewID As Variant, Optional DomConfig As DOMDocument) As Boolean
    Dim domHeadOld As New DOMDocument
    Dim domBodyOld As New DOMDocument
    Dim clsCommSave As ClsVouchSave
    Dim strMTblName As String
    Dim strSTblName As String
    Dim strError As String
    Dim strSQL As String
    Dim bSuccess As Boolean
    Dim strError1 As String
    Dim strError2 As String
    Dim sNewID As String
    Dim ele As IXMLDOMElement
    Dim strBusType As String
    Set domHeadOld = domHead
    Set domBodyOld = domBody
    Dim intCyc As Integer
    Dim bTrans As Boolean
    Dim UserCheck As Object  'by ahzzd 20100121保存前检查处理类
    Dim sql As String
    Dim rds_temp As New Recordset
    Dim Id As String
    Dim UFTS As String
    
    
    
    
  
    VouchSave = True
    strError1 = ""
    strError2 = ""
    bTrans = False
    Dim lngID As Long
    Dim Userdll As New UserDefineDll_MOD
'    Set Userdll = New UserDefineDll_MOD


   ' --这里实现保存前的业务校验处理
    Select Case strCardNum
        
        Case ""
            Set UserCheck = New clsUsercheck_efyzgl
            Set UserCheck.clsSys = clsSys
            UserCheck.strVouchType = strCardNum
            bSuccess = UserCheck.VoucherCheckForSave(m_Conn, domHead, domBody, VoucherState, strError1, strError2)
            Set UserCheck = Nothing
            
            
            
        Case Else
        'ahzzd 增加保存前检查插件
            Userdll.Before_Save m_Login, m_Conn, strCardNum, domHead, domBody, sErrMsg, bSuccess
            
    End Select
    
    
    Getvoucherinf m_Conn, strCardNum, , , strMTblName, strSTblName
    Id = Val(GetHeadItemValue(domHead, "id"))
    If Id <> 0 Then
        UFTS = GetHeadItemValue(domHead, "ufts")
        strSQL = "select * from " & strMTblName & "  where id=" & Id & " and CONVERT(char, CONVERT(money, ufts), 2)='" & UFTS & "'"
        If rds_temp.State <> 0 Then rds_temp.Close
        rds_temp.Open strSQL, m_Conn, 3, 4
        If rds_temp.EOF Then
            sErrMsg = "当前业务单据已被他人修改，请刷新后再试！"
            bSuccess = False
        End If
    End If
    
    If Not bSuccess Then
        VouchSave = False
        sErrMsg = sErrMsg & strError1
        Exit Function
    End If
     
    Set clsCommSave = New ClsVouchSave
'    clsSys.bManualTrans = True
    Set clsCommSave.clsSys = clsSys
    If clsSys.bManualTrans = False And bTrans = False Then
        m_Conn.BeginTrans
        bTrans = True
        clsSys.bManualTrans = True
    End If
    Set clsCommSave.m_Login = m_Login
    strError = clsCommSave.Save(domHead, domBody, strMTblName, strSTblName, m_Conn, VoucherState, strCardNum, sNewID, , , DomConfig)
    
    
    If Len(strError) = 0 Then
        Userdll.After_Save m_Login, m_Conn, strCardNum, sNewID, domHead, domBody, strError, bSuccess
        If bSuccess Then
            If bTrans Then
                m_Conn.CommitTrans
                bTrans = False
                clsSys.bManualTrans = False
            End If
            
        Else
            If bTrans Then
                m_Conn.RollbackTrans
                bTrans = False
                clsSys.bManualTrans = False
            End If
        End If
        
    Else
    '保存失败后直接回滚
        If bTrans Then
            m_Conn.RollbackTrans
            bTrans = False
            clsSys.bManualTrans = False
        End If
    End If
    
    
    If strError <> "" Then
        VouchSave = False
        sErrMsg = strError
    Else
        vNewID = sNewID
    End If
'    clsSys.bManualTrans = False
    Set clsCommSave = Nothing
    Set Userdll = Nothing
End Function

'//------------取得相应外币汇率，用于项目报价单-------------
'add by Clin 2009-9-21
Public Function GetiExchRate(StrCexch_Name As String, Intiperiod As Integer, DBConn As ADODB.Connection) As String
    Dim rst As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "select nflat from exch where cexch_name = '" & StrCexch_Name & "' and iperiod = '" & Intiperiod & "'"
    rst.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rst.EOF Then
        GetiExchRate = ""
    Else
        If IsNull(rst(0)) Then
            GetiExchRate = ""
        Else
            GetiExchRate = rst(0)
        End If
    End If
    rst.Close
    Set rst = Nothing
End Function

'从后台Dom对象中验证必填项是否正确，字段内容长度限制是否正确
Public Function CheckVoucherItems(m_Conn As Connection, domHead As DOMDocument, domBody As DOMDocument, iState As Integer, CardNumber As String, ByRef errDec As String) As Boolean
    Dim strSQL As String
    Dim Rs As ADODB.Recordset
    Dim domItems As DOMDocument
    Dim eleMent As IXMLDOMElement
    Dim Attr As IXMLDOMAttribute
    Dim vtid As Long
    Dim Row As Long
    
    errDec = ""
    Set Attr = domHead.selectSingleNode("//z:row").Attributes.getNamedItem("vt_id")
    If Attr Is Nothing Then
        vtid = 0
    Else
        vtid = Val(Attr.value & "")
    End If
    If vtid = 0 Then
        strSQL = "select * from voucheritems where vt_id=(select min(vt_id) from vouchertemplates_base where vt_cardnumber='" & VBA.Replace(CardNumber, "'", "''") & "')"
    Else
        strSQL = "select * from voucheritems where vt_id=" & vtid
    End If
    strSQL = strSQL & " order by cardsection"
    Set Rs = m_Conn.Execute(strSQL)
    Set domItems = New DOMDocument: Rs.Save domItems, adPersistXML: Set Rs = Nothing
    '验证表头数据
    If CheckVoucherItem(m_Conn, domHead.selectSingleNode("//z:row"), 0, domItems.selectNodes("//z:row[@CardSection='T']"), 0, iState, CardNumber, errDec) = False Then
    End If
    Row = 1
    For Each eleMent In domBody.selectNodes("//z:row")
        If CheckVoucherItem(m_Conn, eleMent, Row, domItems.selectNodes("//z:row[@CardSection='B']"), 1, iState, CardNumber, errDec) = False Then
                    
        End If
        Row = Row + 1
    Next
    If errDec <> "" Then
        CheckVoucherItems = False
    Else
        CheckVoucherItems = True
    End If
End Function

'从后台Dom对象中验证必填项是否正确，字段内容长度限制是否正确
Public Function CheckVoucherItem(m_Conn As Connection, eleRow As IXMLDOMElement, Row As Long, eleItems As IXMLDOMNodeList, Section As Integer, iState As Integer, CardNumber As String, ByRef errDec As String) As Boolean
    Dim eleMent As IXMLDOMElement
    Dim Attr As IXMLDOMAttribute
    Dim sErrTop As String
    
    If Section = 0 Then
        sErrTop = "表头"
    Else
        sErrTop = "表体行[" & Row & "]"
    End If
    
    For Each eleMent In eleItems
        Select Case LCase(eleMent.getAttribute("FieldName"))
            Case "ufts", "id", "autoid"
            Case Else
                If LCase(eleMent.getAttribute("IsNull") & "") = "true" And LCase(eleMent.getAttribute("CanModify") & "") = "true" Then
                    Set Attr = eleRow.Attributes.getNamedItem(eleMent.getAttribute("FieldName"))
                    If Attr Is Nothing Then
                        errDec = errDec & sErrTop & "[" & eleMent.getAttribute("carditemname") & "]不能为空" & vbCrLf
                    ElseIf Trim(Attr.value & "") = "" Then
                        errDec = errDec & sErrTop & "[" & eleMent.getAttribute("carditemname") & "]不能为空" & vbCrLf
                    End If
                End If
                If LCase(eleMent.getAttribute("FieldType") & "") = "1" Then '长度当前只验证文本型
                    Set Attr = eleRow.Attributes.getNamedItem(eleMent.getAttribute("FieldName"))
                    If Not Attr Is Nothing Then
                        If Len(Attr.value) > Val(eleMent.getAttribute("MaxLength") & "") Then
                            errDec = errDec & sErrTop & "[" & eleMent.getAttribute("carditemname") & "]长度不能大于" & Val(eleMent.getAttribute("MaxLength") & "") & vbCrLf
                        End If
                    End If
                End If
        End Select
    Next
        
End Function

Private Sub CheckUfts(CN As Object, ByVal iState As Long, strMTblName As String, Id As String, UFTS As String, ByRef ErrMsg1 As String)
    Dim strSQL As String: Dim rec As ADODB.Recordset
    If iState = 1 Or iState = 2 Then   '1、2 表示变更或修改      0、表示增加
        strSQL = "select * from " & strMTblName & " where id=" & Id & IIf(UFTS = "", "", " AND convert(char, convert(money,ufts), 2)= N'" & UFTS & "'")
        Set rec = New ADODB.Recordset
        If rec.State = 1 Then rec.Close
        rec.Open strSQL, CN, 3, 1
        If rec.RecordCount <> 1 Then
            ErrMsg1 = "当前单据已经被其他操作员修改!"
        End If
    End If
End Sub

'单据审核、弃核处理
Public Function VouchVerify(CN As ADODB.Connection, sflag As Boolean, CardNumber As String, domHead As DOMDocument, TotalMsg As String) As Boolean
    Dim bVer As Boolean, bLock() As Boolean
    Dim iCount As Long, i As Long, iMode As Integer, sOpt As String
    Dim sTaskID As String, sVerMsg As String, strTable As String, ErrMsg As String
    Dim rst As New ADODB.Recordset
    Dim UFTS As String, strUserAuth As String, strUserName As String
    Dim tmpId As String 'Long,
    Dim UFTSFieldName As String '时间戳的字段名称
    Dim sqltemp As String
    Dim Trues As Boolean
    Dim Userdll As New UserDefineDll_MOD
    Dim BeginTrans As Boolean
    
    On Error GoTo ErrHandle
    If CN.State <> 1 Then
        ErrMsg = GetString("U8.SA.USSASERVER.clsfhfunc.01766") 'zh-CN：不能连接到数据库,可能是网络忙或者当前打开的连接太多，请稍后再试 'zh-CN：不能连接到数据库,可能是网络忙或者当前打开的连接太多，请稍后再试
        TotalMsg = ErrMsg
        VouchVerify = False
        Exit Function
    End If
    
    Getvoucherinf CN, CardNumber, , , strTable
    If sflag = True Then
        bVer = True
        iMode = 0   ''审核
        sOpt = GetString("U8.SA.USSASERVER.clsfhfunc.01767") 'zh-CN：审核
        'strTable = "MT_budget"
    ElseIf sflag = False Then
        bVer = False
        iMode = 1  ''弃审
        sOpt = GetString("U8.SA.USSASERVER.clsfhfunc.01768") 'zh-CN：弃审
        'strTable = "MT_budget"
    End If
    ErrMsg = ""
    
    tmpId = Val(GetHeadItemValue(domHead, "id"))
    UFTS = GetHeadItemValue(domHead, "ufts")

    If iMode = 0 Or iMode = 1 Then
'        ''先判断是否单据已经改变
'        ''弃审状态检测操作员
'        Select Case sType
        
'            Case "87"
'                sqltemp = "Select IsNULL(cverifier,'') As cChecker,ddate From v_mt_baseset01 where " & _
'                          "ID=" & tmpId & " and UFTS='" & UFTS & "'"
'            Case "88"
'                sqltemp = "Select IsNULL(cverifier,'') As cChecker,ddate From v_mt_baseset02 where " & _
'                          "ID=" & tmpId & " and UFTS='" & UFTS & "'"
'
'            Case "98"
'                sqltemp = "Select IsNULL(cverifier,'') As cChecker,ddate From V_MT_budget66 where " & _
'                          "ID=" & tmpId & " and UFTS='" & UFTS & "'"
'            Case Else
'                sqltemp = ""
'        End Select
        
        
        sqltemp = "Select IsNULL(cverifier,'') As cChecker,ddate From " & strTable & " where " & _
                 "ID=" & tmpId & IIf(UFTS = "", "", " AND convert(char, convert(money,ufts), 2)= N'" & UFTS & "'")
        
        If rst.State <> 0 Then rst.Close
        rst.CursorLocation = adUseClient
        rst.Open sqltemp, CN, adOpenForwardOnly, adLockReadOnly
        If rst.RecordCount = 0 Then
            ErrMsg = GetString("U8.SA.USSASERVER.clsfhfunc.01788") 'zh-CN：单据不存在或者已经被人修改!
            GoTo DoNext
        Else
            '如果不是跨年度单据，则控制审核日期必须大于等于单据日期
            If clsSys.getEndDate(12) >= rst.Fields("dDate") Then
                If clsSys.objlogin.CurDate < rst.Fields("dDate") And iMode = 0 Then
                    ErrMsg = GetString("U8.SA.ussaserver.VerifyFHD.861patch.001") '"审核日期必须大于等于单据日期。" '
                    GoTo DoNext
                End If
            End If
            If rst(0) <> "" And iMode = 0 Then
                ErrMsg = GetString("U8.SA.USSASERVER.clsfhfunc.01784") 'zh-CN：单据已经审核!
                GoTo DoNext
            ElseIf rst(0) = "" And iMode = 1 Then
                ErrMsg = "单据未审核!"
                GoTo DoNext
            End If
        
        End If
        If rst.State = 1 Then rst.Close
'        clsSys.bManualTrans = True

'MsgBox "VerifyVoucher clsSys.bManualTrans = " & clsSys.bManualTrans
            If sflag Then  '审核
                Userdll.Before_Verify m_Login, CN, CardNumber, domHead, ErrMsg, Trues
                If Not Trues Then GoTo DoNext
             
                If clsSys.bManualTrans = False And BeginTrans = False Then
                    CN.BeginTrans
                    BeginTrans = True
                    clsSys.bManualTrans = True
                End If
                ErrMsg = VerifyVoucher(CN, clsSys, strTable, str(tmpId), bVer, UFTS, CardNumber, domHead)
                If Len(ErrMsg) = 0 Then
                    Userdll.After_Verify m_Login, CN, CardNumber, domHead, ErrMsg, Trues
                    If BeginTrans Then
                        If Trues Then
                            CN.CommitTrans
                            BeginTrans = False
                            clsSys.bManualTrans = False
                        Else
                            CN.RollbackTrans
                            BeginTrans = False
                            clsSys.bManualTrans = False
                        End If
                    End If
                End If
            Else '弃审
                Userdll.Before_UnVerify m_Login, CN, CardNumber, domHead, ErrMsg, Trues
                If Not Trues Then GoTo DoNext
                
                If clsSys.bManualTrans = False And BeginTrans = False Then
                    CN.BeginTrans
                    BeginTrans = True
                    clsSys.bManualTrans = True
                End If
                ErrMsg = VerifyVoucher(CN, clsSys, strTable, str(tmpId), bVer, UFTS, CardNumber, domHead)
                If Len(ErrMsg) = 0 Then
                    Userdll.After_UnVerify m_Login, CN, CardNumber, domHead, ErrMsg, Trues
                    If BeginTrans Then
                        If Trues Then
                            CN.CommitTrans
                            BeginTrans = False
                            clsSys.bManualTrans = False
                        Else
                            CN.RollbackTrans
                            BeginTrans = False
                            clsSys.bManualTrans = False
                        End If
                    End If
                End If
            End If
            

        
    End If
DoNext:
    If ErrMsg <> "" Then
        TotalMsg = ErrMsg
        VouchVerify = False
        If rst.State = 1 Then rst.Close
        Set rst = Nothing
    Else
        VouchVerify = True
    End If
    
     If BeginTrans Then
        CN.RollbackTrans
        BeginTrans = False
    End If
'    clsSys.bManualTrans = False
    Exit Function
ErrHandle:
     If BeginTrans Then
        CN.RollbackTrans
        BeginTrans = False
        clsSys.bManualTrans = False
    End If
    VouchVerify = False
    TotalMsg = err.Description
End Function

'单据表体行关闭、打开处理
Public Function VouchBodyLineClose(CN As ADODB.Connection, bClose As Boolean, CardNumber As String, domHead As DOMDocument, autoids As String, TotalMsg As String) As Boolean
    'Dim bVer As Boolean, bLock() As Boolean
    'Dim iCount As Long, i As Long, iMode As Integer
    Dim sOpt As String
    'Dim sTaskID As String, sVerMsg As String,
    Dim strMainTable As String, ErrMsg As String
    Dim strBodyTable As String
    Dim rst As New ADODB.Recordset
    Dim AffectedLine As Long
    Dim UFTS As String ', strUserAuth As String, strUserName As String
    Dim tmpId As String 'Long,
    'Dim UFTSFieldName As String '时间戳的字段名称
    Dim sqltemp As String
    Dim bTrans As Boolean
    On Error GoTo ErrHandle
    If CN.State <> 1 Then
        ErrMsg = GetString("U8.SA.USSASERVER.clsfhfunc.01766") 'zh-CN：不能连接到数据库,可能是网络忙或者当前打开的连接太多，请稍后再试 'zh-CN：不能连接到数据库,可能是网络忙或者当前打开的连接太多，请稍后再试
        TotalMsg = ErrMsg
        VouchBodyLineClose = False
        Exit Function
    End If
    Getvoucherinf CN, CardNumber, , , strMainTable, strBodyTable
    
    sOpt = IIf(bClose, "关闭", "打开")
    ErrMsg = ""
    tmpId = Val(GetHeadItemValue(domHead, "id"))
    UFTS = GetHeadItemValue(domHead, "ufts")
    sqltemp = "Select IsNULL(ccloser,'') As ccloser,ddate From " & strMainTable & " where " & _
             "ID=" & tmpId & IIf(UFTS = "", "", " AND convert(char, convert(money,ufts), 2)= N'" & UFTS & "'")
    
    If rst.State <> 0 Then rst.Close
    rst.CursorLocation = adUseClient
    rst.Open sqltemp, CN, adOpenForwardOnly, adLockReadOnly
    If rst.RecordCount = 0 Then
        ErrMsg = GetString("U8.SA.USSASERVER.clsfhfunc.01788") 'zh-CN：单据不存在或者已经被人修改!
        GoTo DoNext
    Else
'        '如果不是跨年度单据，则控制审核日期必须大于等于单据日期
'        If clsSys.getEndDate(12) >= rst.Fields("dDate") Then
'            If clsSys.objlogin.CurDate < rst.Fields("dDate") And iMode = 0 Then
'                ErrMsg = GetString("U8.SA.ussaserver.VerifyFHD.861patch.001") '"审核日期必须大于等于单据日期。" '
'                GoTo DoNext
'            End If
'        End If
        If rst(0) <> "" And bClose Then
            ErrMsg = "单据已经关闭!"
            GoTo DoNext
        End If
    
    End If
    If rst.State = 1 Then rst.Close
    
    bTrans = False
    If clsSys.bManualTrans = False And bTrans = False Then
       CN.BeginTrans  ''每张单据一个事务
       bTrans = True
       clsSys.bManualTrans = True
    End If
    
    sqltemp = " Update " & strBodyTable & " SET cbCloser=" & IIf(bClose, "N'" & clsSys.CurrentUserName & "'", "NULL") & _
                " ,cbCloseDate =" & IIf(bClose, "N'" & clsSys.dWorkDate & "'", "NULL") & _
         " WHERE autoid in(" & autoids & ")"
    CN.Execute sqltemp, AffectedLine
            
    If AffectedLine = -1 Then
        If bTrans Then
            CN.RollbackTrans: bTrans = False
            clsSys.bManualTrans = False
        End If
        TotalMsg = "不成功，请检查数据是否正确!"
        Exit Function
    Else
        If bTrans Then CN.CommitTrans
        TotalMsg = ""
    End If
    
    'ErrMsg = VerifyVoucher(CN, clsSys, strMainTable, str(tmpId), bVer, UFTS, sType, domHead)
  
DoNext:
    If ErrMsg <> "" Then
        TotalMsg = ErrMsg
        VouchBodyLineClose = False
        If rst.State = 1 Then rst.Close
        Set rst = Nothing
    Else
        VouchBodyLineClose = True
    End If
    Exit Function
ErrHandle:
    If bTrans Then
        CN.RollbackTrans: bTrans = False
        clsSys.bManualTrans = False
    End If
    VouchBodyLineClose = False
    TotalMsg = err.Description
End Function
'单据关闭、打开处理
Public Function VouchClose(CN As ADODB.Connection, bClose As Boolean, CardNumber As String, domHead As DOMDocument, TotalMsg As String) As Boolean
    'Dim bVer As Boolean, bLock() As Boolean
    'Dim iCount As Long, i As Long, iMode As Integer
    Dim sOpt As String
    'Dim sTaskID As String, sVerMsg As String,
    Dim strTable As String, ErrMsg As String
    Dim rst As New ADODB.Recordset
    Dim AffectedLine As Long
    Dim UFTS As String ', strUserAuth As String, strUserName As String
    Dim tmpId As String 'Long,
    'Dim UFTSFieldName As String '时间戳的字段名称
    Dim sqltemp As String
    Dim bTrans As Boolean
    On Error GoTo ErrHandle
    If CN.State <> 1 Then
        ErrMsg = GetString("U8.SA.USSASERVER.clsfhfunc.01766") 'zh-CN：不能连接到数据库,可能是网络忙或者当前打开的连接太多，请稍后再试 'zh-CN：不能连接到数据库,可能是网络忙或者当前打开的连接太多，请稍后再试
        TotalMsg = ErrMsg
        VouchClose = False
        Exit Function
    End If
    Getvoucherinf CN, CardNumber, , , strTable
    
    sOpt = IIf(bClose, "关闭", "打开")
    ErrMsg = ""
    tmpId = Val(GetHeadItemValue(domHead, "id"))
    UFTS = GetHeadItemValue(domHead, "ufts")
    sqltemp = "Select IsNULL(ccloser,'') As ccloser,ddate From " & strTable & " where " & _
             "ID=" & tmpId & IIf(UFTS = "", "", " AND convert(char, convert(money,ufts), 2)= N'" & UFTS & "'")
    
    If rst.State <> 0 Then rst.Close
    rst.CursorLocation = adUseClient
    rst.Open sqltemp, CN, adOpenForwardOnly, adLockReadOnly
    If rst.RecordCount = 0 Then
        ErrMsg = GetString("U8.SA.USSASERVER.clsfhfunc.01788") 'zh-CN：单据不存在或者已经被人修改!
        GoTo DoNext
    Else
'        '如果不是跨年度单据，则控制审核日期必须大于等于单据日期
'        If clsSys.getEndDate(12) >= rst.Fields("dDate") Then
'            If clsSys.objlogin.CurDate < rst.Fields("dDate") And iMode = 0 Then
'                ErrMsg = GetString("U8.SA.ussaserver.VerifyFHD.861patch.001") '"审核日期必须大于等于单据日期。" '
'                GoTo DoNext
'            End If
'        End If
        If rst(0) <> "" And bClose Then
            ErrMsg = "单据已经关闭!"
            GoTo DoNext
        ElseIf rst(0) = "" And bClose = False Then
            ErrMsg = "单据未关闭!"
            GoTo DoNext
        End If
    
    End If
    If rst.State = 1 Then rst.Close
    
    bTrans = False
    If clsSys.bManualTrans = False And bTrans = False Then
       CN.BeginTrans  ''每张单据一个事务
       bTrans = True
       clsSys.bManualTrans = True
    End If
    
    sqltemp = " Update " & strTable & " SET ccloser=" & IIf(bClose, "N'" & clsSys.CurrentUserName & "'", "NULL") & _
                " ,dcloserdate =" & IIf(bClose, "N'" & clsSys.dWorkDate & "'", "NULL") & _
         " WHERE id=" & CLng(Val(tmpId)) & _
         IIf(UFTS = "", "", " AND Convert(char,Convert(Money,Ufts),2)= N'" & UFTS & "'")
    CN.Execute sqltemp, AffectedLine
            
    If AffectedLine = 0 Then
        If bTrans Then
            CN.RollbackTrans: bTrans = False
            clsSys.bManualTrans = False
        End If
        TotalMsg = GetStringPara("U8.SA.USSASERVER.modvoucher.01054", sOpt) 'zh-CN：弃审 'zh-CN：审核 'Para zh-CN：该单据已经被其他人改动，或{0}不成功
        Exit Function
    ElseIf AffectedLine = -1 Then
        If bTrans Then
            CN.RollbackTrans: bTrans = False
            clsSys.bManualTrans = False
        End If
        TotalMsg = "不成功，请检查数据是否正确!"
        Exit Function
    Else
        If bTrans Then CN.CommitTrans
        TotalMsg = ""
    End If
    
    'ErrMsg = VerifyVoucher(CN, clsSys, strTable, str(tmpId), bVer, UFTS, sType, domHead)
  
DoNext:
    If ErrMsg <> "" Then
        TotalMsg = ErrMsg
        VouchClose = False
        If rst.State = 1 Then rst.Close
        Set rst = Nothing
    Else
        VouchClose = True
    End If
    Exit Function
ErrHandle:
    If bTrans Then
        CN.RollbackTrans: bTrans = False
        clsSys.bManualTrans = False
    End If
    VouchClose = False
    TotalMsg = err.Description
End Function

'功 能: 返回当前单据是否在第一张。
Public Property Get BOF() As Boolean
    BOF = m_BOF
End Property

'功 能: 返回当前单据是否在最后一张。
Public Property Get EOF() As Boolean
    EOF = m_EOF
End Property
''是否期初单据
Public Property Let bFirstVouch(ByVal vNewValue As Boolean)
     bFirst = vNewValue
End Property
Public Property Let strARAuth(ByVal strNewValue As String)
    strARAuthC = strNewValue
End Property

Public Property Get m_VouchID() As Variant
    m_VouchID = m_VoucherID
End Property

Public Property Let m_VouchID(ByVal vNewValue As Variant)
    m_VoucherID = vNewValue
End Property


Public Sub GetsumDom(CN As ADODB.Connection, cDepCode As String, cItemCode As String, domHead As DOMDocument, domboby As DOMDocument)
    Dim rds As New ADODB.Recordset
    Dim sqlstr As String
    sqlstr = "select * from V_MT_budgetsum  where cDepCode='" & cDepCode & "' and cItemCode='" & cItemCode & "'"
    If rds.State <> 0 Then rds.Close
    rds.CursorLocation = adUseClient
    rds.Open sqlstr, CN, 3, 4
    rds.Save domHead, adPersistXML
    
    sqlstr = "select * from V_MT_budgetssum where cDepCode='" & cDepCode & "' and cItemCode='" & cItemCode & "'"
    If rds.State <> 0 Then rds.Close
    rds.CursorLocation = adUseClient
    rds.Open sqlstr, CN, 3, 4
    rds.Save domboby, adPersistXML

End Sub

Private Sub Class_Initialize()
    Set cls_Public = CreateObject("UF_Public_base.cls_log")
End Sub
