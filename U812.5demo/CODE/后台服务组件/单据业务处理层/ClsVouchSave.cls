VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsVouchSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public clsSys As clsSystem
Private bOnceGetTrueNO As Boolean
Public m_Login As U8Login.clsLogin
'Public uvdVouchTypeS As enuVouType

Public Function Save(ByVal domHead As DOMDocument, ByVal domBody As DOMDocument, ByVal sHeadTableName As String, ByVal sBodyTableName As String, ByVal DBConn As ADODB.Connection, ByVal iState As Integer, strCardNum As String, Optional VouchID As String, Optional bBatchMake As Boolean, Optional bBatchDispToBill As Boolean, Optional DomConfig As DOMDocument) As String
    Dim bReturnFlag As Boolean
    Dim bUpDateCurrentStock As Boolean
    Dim eleMent As IXMLDOMElement
    Dim eleMentBody As IXMLDOMElement
    Dim errDom As New DOMDocument
    Dim i As Long
    Dim HeadID As String
    Dim BodyID As String
    Dim sError As String
    Dim AffectedLine As Long
    Dim strSQL As String
    Dim lngMainID As String ' Long
    Dim lngSubID As String 'Long
    Dim lngSubIDCopy  As String ' As Long
    Dim lngCountSTable As Long        '//表体记录数
    Dim eleAttr As IXMLDOMAttribute
    Dim strVouchCode As String        '//单据编号
    Dim strSourceCode As String
    Dim ErrMsg1 As String
    Dim strBusType As String
    Dim strVouchType As String
    Dim tmpDomHead As DOMDocument, tmpDomBody As DOMDocument
    Dim bNeedCheckCredit As Boolean
    Dim strFldVouchCode As String
    Dim strRefType As String
    Dim iGetVouchCodeCount As Integer
    Dim RecTemp As New ADODB.Recordset
    Dim strTmp As String
    Dim strSeedName As String
    Dim bAR As Boolean
    Dim lngReGet As Long
    Dim rstGetBodySQL As New ADODB.Recordset
    Dim NdList As IXMLDOMNodeList
    Dim oAttachDom As DOMDocument, bSaveAttach As Boolean   '附件的dom
    
    Dim strErrMsg As String
    Dim DictCorUfts As New Dictionary        '对方单据时间戳集合 '860
    Dim bTrans As Boolean
    
    Dim blnEAI As Boolean
    Dim tempNode As IXMLDOMNode
    
'    If Not DomConfig Is Nothing Then
    If Not DomConfig.selectSingleNode("EAI") Is Nothing Then
        Set tempNode = DomConfig.selectSingleNode("EAI")
    End If
    
    If tempNode Is Nothing Then
        blnEAI = False
    Else
        blnEAI = tempNode.Text
    End If
    
    ''是否取过真号
    bOnceGetTrueNO = False
    RecTemp.CursorLocation = adUseClient

    On Error GoTo ErrSave
    
    
    cls_Public.WrtDBlog DBConn, m_Login.cUserId, "EFVoucherMo.ClsVouchSave.Save domHead=", domHead.xml
'    cls_Public.WrtDBlog DBConn, m_Login.cUserId, "EFVoucherMo.ClsVouchSave.Save domHead=", domHead.selectSingleNode("//rs:data").xml
'    domHead.Save App.Path & "\domHead.xml"
       
    cls_Public.WrtDBlog DBConn, m_Login.cUserId, "EFVoucherMo.ClsVouchSave.Save domBody=", domBody.xml
'    cls_Public.WrtDBlog DBConn, m_Login.cUserId, "EFVoucherMo.ClsVouchSave.Save domBody=", domBody.selectSingleNode("//rs:data").xml
'    domBody.Save App.Path & "\domBody.xml"
    
    If domHead Is Nothing Then
        Save = "前台传来的表头数据有误，不能保存"
        Set rstGetBodySQL = Nothing
        If RecTemp.State = 1 Then RecTemp.Close
        Set RecTemp = Nothing
        Exit Function
    End If
    If domBody Is Nothing Then
        Save = "前台传来的表体数据有误，不能保存"
        Set rstGetBodySQL = Nothing
        If RecTemp.State = 1 Then RecTemp.Close
        Set RecTemp = Nothing
        Exit Function
    End If
    
    '-------------------单据主表保存处理--------------------
    Dim strTableName As String
    Dim strID As String
    Dim strSubID As String
    Dim bRepeatReDo As Boolean
    Dim SBVID As Long
    Dim DomFormat As New DOMDocument
    Dim iverifystate As String  '判断是否单据已经提交
    
    Set eleMent = domHead.selectSingleNode("//z:row")
    lngReGet = 0
    If IsNull(eleMent.getAttribute("iverifystate")) Then
        iverifystate = "0"
    Else
        iverifystate = Trim(eleMent.getAttribute("iverifystate"))
    End If
    
    Select Case LCase(sHeadTableName)
       
       '没有特殊处理不要修改此处
        Case LCase("efbwgl_SelDeclare")  '选题论议审批
             strTableName = "efbwgl_SelDeclare"
            strVouchCode = GetAttrValue(eleMent, "ccode")       '/单据号
            strFldVouchCode = LCase("ccode")                    '/单据号字段名称
'            strVouchType = "EFBWGL020301"                        '/单据类型编码"
            strVouchType = strCardNum
            strID = "id"                                        '/主表关键字
            lngMainID = Val(GetAttrValue(eleMent, strID))       '/主表关键字的值
            
            '87X 工作流 added
            If (iState = 0 Or iState = 1) And iverifystate = "0" Then
                If getIsWfControl(clsSys, DBConn, strVouchType, ErrMsg1) Then
                    eleMent.setAttribute "iswfcontrolled", 1
                Else
                    eleMent.setAttribute "iswfcontrolled", 0
                End If
            End If
        Case Else
            strTableName = sHeadTableName
            strVouchCode = GetAttrValue(eleMent, "ccode")       '/单据号
            strFldVouchCode = LCase("ccode")                    '/单据号字段名称
            strVouchType = strCardNum                           '/单据类型编码
            strID = "id"                                        '/主表关键字
            lngMainID = Val(GetAttrValue(eleMent, strID))       '/主表关键字的值
            strSubID = "autoid"
            
            '处理审批流状态信息
            If (iState = 0 Or iState = 1) And iverifystate = "0" Then
                If getIsWfControl(clsSys, DBConn, strVouchType, ErrMsg1) Then
                    eleMent.setAttribute "iswfcontrolled", 1
                    eleMent.setAttribute "VoucherId", lngMainID
                    eleMent.setAttribute "VoucherType", strVouchType
                    eleMent.setAttribute "VoucherCode", strVouchCode
                Else
                    eleMent.setAttribute "iswfcontrolled", 0
                End If
            End If
            
    End Select
    
    '提取附件dom
    Call GetAttachDOM(strVouchType, domHead, oAttachDom, bSaveAttach)
    Call GetVouchNO(clsSys.Connectstr, strVouchType, domHead, strVouchCode, ErrMsg1, , True, , bRepeatReDo)

GetVouchCode:
    bTrans = False
    If Not bAR And Not clsSys.bManualTrans Then
        DBConn.BeginTrans
        bTrans = True
    End If
    
    If CheckVouchCode(strVouchType, sHeadTableName, strFldVouchCode, strVouchCode, bBatchMake, bRepeatReDo, ErrMsg1, domHead, DBConn, iState, Val(lngMainID), strID) = False Then    'bBatchDispToBill------   bBatchMake,
        Save = ErrMsg1
        If Not bAR Then
           If bTrans Then DBConn.RollbackTrans: bTrans = False
        End If
        Set rstGetBodySQL = Nothing
        If RecTemp.State = 1 Then RecTemp.Close
        Set RecTemp = Nothing
        Exit Function
    End If
    
    Set eleMent = domHead.selectSingleNode("//z:row")
    eleMent.setAttribute LCase(strFldVouchCode), strVouchCode
    eleMent.setAttribute "VoucherCode", strVouchCode
    If strVouchType = "94" Then
        lngCountSTable = 0 '//借款单没有子表
    Else
        lngCountSTable = domBody.selectNodes("//z:row[@editprop = 'A' or @editprop = 'a']").length
    End If
    
    If iState = 0 Then      '增加
        If Not blnEAI Then
            '//重新取得主表ID及子表IDS的起始值
            lngMainID = GetVouchID(strTableName, clsSys, lngSubID, lngCountSTable, ErrMsg1)
        End If
        
        If ErrMsg1 <> "" Then
            Save = ErrMsg1
            Set rstGetBodySQL = Nothing
            If RecTemp.State = 1 Then RecTemp.Close
            Set RecTemp = Nothing
            If bTrans Then DBConn.RollbackTrans: bTrans = False
            Exit Function
        End If
        eleMent.setAttribute LCase(strID), lngMainID
        eleMent.setAttribute "VoucherId", lngMainID
        strSQL = GetStrInsertFromDoc(eleMent, sHeadTableName, DBConn, True, lngMainID)
        eleMent.setAttribute LCase(strID), lngMainID
    Else       'iState=1    '修改操作
        If lngCountSTable <> 0 Then
            If Not blnEAI Then
                lngMainID = GetVouchID(strTableName, clsSys, lngSubID, lngCountSTable, ErrMsg1)
            End If
            If ErrMsg1 <> "" Then
                Save = ErrMsg1
                Set rstGetBodySQL = Nothing
                If RecTemp.State = 1 Then RecTemp.Close
                Set RecTemp = Nothing
                If bTrans Then DBConn.RollbackTrans: bTrans = False
                Exit Function
            End If
        End If
        lngMainID = GetHeadItemValue(domHead, strID)
        strSQL = GetStrUpdateFromDoc(eleMent, sHeadTableName, DBConn, strID, "", lngMainID)
    End If
    On Error Resume Next
    DBConn.Execute "set nocount off"
    AffectedLine = 0
    DBConn.Execute strSQL, AffectedLine
    On Error GoTo 0
    If AffectedLine = 0 Then
        lngReGet = lngReGet + 1
        If bTrans Then DBConn.RollbackTrans: bTrans = False
        If DBConn.Errors.Count > 0 Then
            If (DBConn.Errors(0).NativeError = 2627 Or DBConn.Errors(0).NativeError = 3604 Or DBConn.Errors(0).NativeError = 2601) And lngReGet < 10 Then GoTo GetVouchCode
        End If
        Save = GetString("U8.SA.USSASERVER.clsvouchsave.00510") 'zh-CN：本单据已被他人修改或修改不成功或网络冲突
        Set rstGetBodySQL = Nothing
        If RecTemp.State = 1 Then RecTemp.Close
        Set RecTemp = Nothing
        Exit Function
    ElseIf AffectedLine = -1 Then
        Save = GetString("U8.SA.USSASERVER.clsvouchsave.00511") 'zh-CN：保存表头数据时失败，不能保存
        If bTrans Then DBConn.RollbackTrans: bTrans = False
        Set rstGetBodySQL = Nothing
        If RecTemp.State = 1 Then RecTemp.Close
        Set RecTemp = Nothing
        Exit Function
    End If
    
        
    '-------------------单据子表保存处理,包装物没有子表，不进行如下操作---------------------
    If strVouchType <> "94" And strVouchType <> "EFBWGL020601" Then  '//借款单不进行子表的处理
        Dim strEditProp As String
        Dim lngStep As Long
        Set NdList = domBody.selectNodes(("//z:row[@editprop != '']"))
            
        If Not blnEAI Then   'by ahzzd 20060922 当不是EAI时才显示进度条
            Proc_ShowProcess (True)
        End If
        
        If NdList.length > 0 Then
            Setup_Frm.ProgressBar1.Max = NdList.length
        End If
        Setup_Frm.ProgressBar1.value = 0
        lngStep = 0
        For Each eleMent In NdList
            lngStep = lngStep + 1
            Setup_Frm.ProgressBar1.value = lngStep
            
            If iState = 0 Then
                strEditProp = Trim(eleMent.getAttribute("editprop"))
                If Not bBatchMake Then
                    eleMent.setAttribute LCase(strID), lngMainID
                End If
            Else
                strEditProp = Trim(eleMent.getAttribute("editprop"))
            End If
            
            Select Case LCase(sBodyTableName)
                Case LCase("MT_basesets")
                    strSubID = "autoid"
                    strID = "id"
                Case LCase("MT_budgets")
                    strSubID = "autoid"
                    strID = "id"
                Case Else
                    strSubID = "autoid"
                    strID = "id"
            End Select
            Select Case UCase(strEditProp)
                Case "A"
                    Select Case LCase(sHeadTableName)
                        Case "wwwww"
                            
                        Case Else
                            If Not blnEAI Then
                                strSQL = GetStrInsertFromDoc2(rstGetBodySQL, eleMent, sBodyTableName, DBConn, False, lngMainID, str(CDbl(Val(lngSubID)) + i))
                                i = i + 1
                            Else
                                lngSubID = eleMent.getAttribute(LCase(strSubID))
                                strSQL = GetStrInsertFromDoc2(rstGetBodySQL, eleMent, sBodyTableName, DBConn, False, lngMainID, CStr(lngSubID))
                            End If
                    End Select
                Case "M"
                    strSQL = GetStrUpdateFromDoc2(rstGetBodySQL, eleMent, sBodyTableName, DBConn, strSubID, , , strVouchCode)
                Case "D"
                    strSQL = "DELETE FROM " & sBodyTableName & " WHERE " & strSubID & "=" & CDbl(Val(GetAttrValue(eleMent, strSubID)))
                Case ""
                    strSQL = ""
            End Select
            If strSQL <> "" Then
                On Error Resume Next
                DBConn.Execute strSQL, AffectedLine
                On Error GoTo 0
                If AffectedLine = 0 Then
                    If bTrans Then DBConn.RollbackTrans: bTrans = False
                    lngReGet = lngReGet + 1
                    If DBConn.Errors.Count > 0 Then
                        If (DBConn.Errors(0).NativeError = 2627 Or DBConn.Errors(0).NativeError = 3604 Or DBConn.Errors(0).NativeError = 2601) And lngReGet < 10 Then GoTo GetVouchCode
                    End If
                    Save = GetString("U8.SA.USSASERVER.clsvouchsave.00513") 'zh-CN：本单据已被修改或网络冲突
                    If rstGetBodySQL.State = 1 Then rstGetBodySQL.Close
                    Set rstGetBodySQL = Nothing
                    If RecTemp.State = 1 Then RecTemp.Close
                    Set RecTemp = Nothing
                    Exit Function
                ElseIf AffectedLine = -1 Then
                    Save = GetString("U8.SA.USSASERVER.clsvouchsave.00514") 'zh-CN：保存表尾数据时失败，不能保存
                    If bTrans Then DBConn.RollbackTrans: bTrans = False
                    If rstGetBodySQL.State = 1 Then rstGetBodySQL.Close
                    Set rstGetBodySQL = Nothing
                    If RecTemp.State = 1 Then RecTemp.Close
                    Set RecTemp = Nothing
                    Exit Function
                ElseIf AffectedLine = 1 And DBConn.Errors.Count > 0 Then
                    If DBConn.Errors.Item(0).NativeError = 547 Then
                        Save = GetString("U8.SA.USSASERVER.clsvouchsave.00515") 'zh-CN：保存表尾数据时失败，不能保存!
                    Else
                        Save = GetString("U8.SA.USSASERVER.clsvouchsave.00516") '& DBConn.Errors(0).Description 'zh-CN：保存表尾数据时失败，不能保存!
                    End If
                    If bTrans Then DBConn.RollbackTrans: bTrans = False
                    If rstGetBodySQL.State = 1 Then rstGetBodySQL.Close
                    Set rstGetBodySQL = Nothing
                    If RecTemp.State = 1 Then RecTemp.Close
                    Set RecTemp = Nothing
                    Exit Function
                End If
            End If
        Next
        If Not blnEAI Then
            Proc_ShowProcess (False)
        End If
    End If
    
    
    If bSaveAttach Then
        If SaveAttach(DBConn, strVouchType, domHead, oAttachDom, strCardNum, ErrMsg1) = False Then
            Save = ErrMsg1
            If bTrans Then DBConn.RollbackTrans: bTrans = False
            If rstGetBodySQL.State = 1 Then rstGetBodySQL.Close
            Set rstGetBodySQL = Nothing
            If RecTemp.State = 1 Then RecTemp.Close
            Set RecTemp = Nothing
            Exit Function
        End If
    End If
    

    
    'UAP生单处理 Begin====================
    Dim ErrDesc As String
    Dim strRuleID As String
    Dim CardNumber As String
    
    strRuleID = GetAttrValue(domHead.documentElement, "makevoucherruleid")
 
    If ErrDesc <> "" Then
        Save = ErrDesc
        If bTrans Then DBConn.RollbackTrans: bTrans = False
        Exit Function
    End If
    'UAP生单处理 end========================
    
    If Not bAR Then
        If bTrans = True Then DBConn.CommitTrans
        bTrans = False
    End If
    
    VouchID = lngMainID
    If rstGetBodySQL.State = 1 Then rstGetBodySQL.Close
    Set rstGetBodySQL = Nothing
    If RecTemp.State = 1 Then RecTemp.Close
    Set RecTemp = Nothing
    Exit Function
ErrSave:
    If bTrans = True Or bAR = True Then
        '单据号重复
        If DBConn.Errors.Count > 0 Then
            If DBConn.Errors(0).NativeError = 2627 Or DBConn.Errors(0).NativeError = 3604 Or DBConn.Errors(0).NativeError = 2601 Then
               strVouchCode = ""
                If bTrans Then DBConn.RollbackTrans
                bTrans = False
               GoTo GetVouchCode
            Else
                If bTrans Then DBConn.RollbackTrans
                bTrans = False
            End If
        Else
            If bTrans Then DBConn.RollbackTrans
            bTrans = False
        End If
    End If
    Select Case err.Number
        Case Else
            Save = err.Description & vbCrLf & ErrMsg1
    End Select
    If rstGetBodySQL.State = 1 Then rstGetBodySQL.Close
    Set rstGetBodySQL = Nothing
End Function

'保存UAP的单据
'iState 0=表示新增 ,1=删除 ,  2=修改
'lngMainID=主键ID
'cardnumber=单据CardNum
'如果新增则传入规则ID,修改则不需要规则ID
Private Sub ProcessUAPMaker(iState As Integer, DBConn As ADODB.Connection, strRuleID As String, strVoucherID As String, ByVal CardNumber As String, ByRef ErrDesc As String)
    Dim maker As Object
    Set maker = CreateObject("UAPMaker.clsSave")
    If maker Is Nothing Then
        ErrDesc = "创建UAPMaker.clsSave失败!"
        Exit Sub
    End If
    maker.savetype = iState
    maker.CardNumer = CardNumber
    maker.LoginInfo = clsSys.objlogin
    maker.MakeVoucherRuleID = strRuleID
    maker.TransConnection = DBConn
    maker.VoucherId = strVoucherID
    maker.Check
    ErrDesc = maker.ErrDesc
End Sub
Private Function getCardNum(ByVal VoucherType As String, ByVal bReturnFlag As Boolean) As String
    Select Case VoucherType
        Case "97"
            getCardNum = "17"
        Case "05"
            If bReturnFlag = True Then
                getCardNum = "03"
            Else
                getCardNum = "01"
            End If
            
        Case "26"   '专用发票,红字专用发票
            getCardNum = "07"
        Case "27"     '普通发票,红字普通发票
            getCardNum = "13"
        Case "06"
            If bReturnFlag = True Then
                getCardNum = "06"   '委托代销退货单
            Else
                getCardNum = "05"   '委托代销发货单
            End If
        Case "16"                     '销售报价单
            getCardNum = "16"
    End Select
End Function
Private Function Proc_ShowProcess(bShow As Boolean)
    On Error Resume Next
    If bShow Then
        Setup_Frm.Show
    Else
        Setup_Frm.Hide
    End If
End Function

       
'' 通用的字段转换函数
Public Function ConvertFieldByType(ByVal objFld As ADODB.Field) As String
Dim vValue As Variant, sErrFrom As String
On Error GoTo ErrHandle
   sErrFrom = "Public:ConvertFieldByType"
   Select Case objFld.Type
          ''数值类型
          Case adBigInt, adTinyInt, adSmallInt, adSingle, adDouble, adNumeric, _
               adCurrency, adDecimal, adInteger, adVarNumeric, adUnsignedBigInt, _
               adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            
               If IsNull(objFld.value) Then
                  vValue = "NULL"
               Else
                  vValue = CDbl(objFld.value)
               End If

          ''字符串类型
          Case adBSTR, adChar, adLongVarChar, adLongVarWChar, adVarChar, adVarWChar, adWChar
              '' 2001.6
              '' 针对NULL和空串区分处理
              If IsNull(objFld.value) Then
                 vValue = "NULL"
              Else
                 If Trim(objFld.value) = "" Then
                    vValue = "NULL"
                 Else
                    vValue = "'" & objFld.value & "'"
                 End If
              End If
         
          ''布尔类型
          Case adBoolean
              vValue = IIf(objFld.value, 1, 0)
          
          ''日期类型,要根据情况转换
          Case adDate Or adDBTime Or adDBTimeStamp
              
              vValue = objFld.value
              If IsDate(vValue) Then
                 vValue = "'" & Format(vValue, "YYYY-MM-DD") & "'"
              Else
                 vValue = "NULL"
              End If
          Case Else
              vValue = NullToStr(objFld)
  End Select
  ConvertFieldByType = vValue
  Exit Function
ErrHandle:
  Select Case err.Number
    Case Else
         MsgBox sErrFrom & vbCrLf & CStr(err.Number) & err.Description, vbOKOnly + vbCritical
  End Select
  Exit Function
End Function

Private Function NullToStr(vValue As ADODB.Field) As String
  NullToStr = IIf(IsNull(vValue.value), "NULL", "'" & vValue.value & "'")
End Function


Public Function DeleteVouch(strVouchType As String, ByRef DBConn As ADODB.Connection, sTableName As String, domHead As DOMDocument) As String
    Dim errDom As New DOMDocument
    Dim domBody As New DOMDocument
    Dim lngID As String
    Dim lngSubID As String
    Dim dUfts As String
    Dim strSQL As String
    Dim strTblName As String
    Dim strSubID As String
    Dim bTrans As Boolean
    Dim strVouchName As String
    Dim bFromSO As Boolean
    Dim strErrMsg As String
    Dim sTmpTable As String
    Dim strID As String
    Dim ieleMent As IXMLDOMElement
    Dim bReturnFlag As Boolean
    Dim iAffCount As Long
    Dim strSubTblName As String
    Dim Userdll As New UserDefineDll_MOD
    Dim bSuc As Boolean
    
    Set ieleMent = domHead.selectSingleNode("//z:row")
    
    On Error GoTo err_Handle
    '工作流校验 87X 已提交的单据不能删除
    If GetHeadItemValue(domHead, "iswfcontrolled") = "1" And Not ( _
       (GetHeadItemValue(domHead, "iverifystate") = "0" Or GetHeadItemValue(domHead, "iverifystate") = "") Or _
       (GetHeadItemValue(domHead, "iverifystate") = "1" And Val(GetHeadItemValue(domHead, "ireturncount")) > 0)) Then
        DeleteVouch = GetString("U8.SA.xsglsql_2.saworkflowsrv.014") '"单据已经提交工作流,不能删除!"
        Exit Function
    End If
    
    dUfts = GetHeadItemValue(domHead, "ufts")
    
'-----------------add by asdf 2009-07-27-------------------
'设置主表，子表的关联字段
'------------------------Start-----------------------------
    Select Case UCase(strVouchType)

        Case "EFBWGL020601"
            strID = "id"
            lngID = GetAttrValue(ieleMent, strID)
            
        Case Else
            strID = "id"
            lngID = GetAttrValue(ieleMent, strID)
            strSubID = "autoid"
            lngSubID = GetAttrValue(ieleMent, strSubID)
    End Select
'------------------------END--------------------------------
    
    
    Dim rds_temp As New Recordset
    Dim sql As String
    Dim strMainTable As String
    Dim strSubTable As String
    

    sql = "select BTTblName,BTQName,BWTblName,BWQName,VchListQName from Vouchers " & _
        "where [CardNumber] ='" & strVouchType & "'"
    With rds_temp
        If .State <> 0 Then .Close
        .CursorLocation = adUseClient
        .Open sql, DBConn, 3, 4
        strMainTable = .Fields("BTTblName")
        strSubTable = IIf(IsNull(.Fields("BWTblName")), "", .Fields("BWTblName"))
    End With
    
    
    strErrMsg = ""
    CheckForDel strVouchType, domHead, DBConn, strErrMsg
    
    If strErrMsg <> "" Then
        DeleteVouch = strErrMsg
        Exit Function
    End If
    
    
    '删除前插件
    Userdll.Before_Delete m_Login, DBConn, strVouchType, domHead, strErrMsg, bSuc
    If bSuc = False Then
        DeleteVouch = strErrMsg
        Exit Function
    End If
    
    
    
    bTrans = False
    If clsSys.bManualTrans = False And bTrans = False Then
        DBConn.BeginTrans
        bTrans = True
        clsSys.bManualTrans = True
    End If
    
    
    If strSubTable <> "" Then          ''无表体跳过 Clin 09-12-03
        '判断表体是否存在数据
        strSQL = "select * from " & strSubTable & _
                " where " & strID & " = " & lngID
        With rds_temp
            If .State <> 0 Then .Close
            .CursorLocation = adUseClient
            .Open strSQL, DBConn, 3, 4
        End With
        
        If Not rds_temp.EOF Then
    
                strSQL = " DELETE  FROM " & strSubTable & _
                         " Where  " & strID & "= " & lngID
                         
                DBConn.Execute strSQL, iAffCount
                If iAffCount = 0 Then
                    DeleteVouch = GetString("U8.SA.USSASERVER.clsvouchsave.00542") 'zh-CN：该单据已经不存在或已被其他人修改
                    If bTrans Then
                         DBConn.RollbackTrans
                         clsSys.bManualTrans = False
                    End If
                    Exit Function
                End If
        End If
    End If
    
    
    '删除主表
    strSQL = " DELETE  FROM " & strMainTable & _
             " Where  " & strID & "= " & lngID
    DBConn.Execute strSQL, iAffCount
 
    

    
    
    If iAffCount = 0 Then
        DeleteVouch = GetString("U8.SA.USSASERVER.clsvouchsave.00543") 'zh-CN：该单据已经不存在或已被其他人修改
        If bTrans Then
             DBConn.RollbackTrans
             clsSys.bManualTrans = False
        End If
        Exit Function
    End If
    
     '删除附件
    If DelVouchAttach(DBConn, strVouchType, domHead, strErrMsg) = False Then
        If bTrans Then DBConn.RollbackTrans: bTrans = False
        DeleteVouch = strErrMsg
        Exit Function
    End If
    
   
    Userdll.After_Delete m_Login, DBConn, strVouchType, domHead, strErrMsg, bSuc
    If bSuc = False Then
        DeleteVouch = strErrMsg
    End If
    
    If DeleteVouch <> "" Then
        If bTrans Then
             DBConn.RollbackTrans
             clsSys.bManualTrans = False
        End If
        Exit Function
    End If


    
    'UAP生单删除 begin
    Dim ErrDesc As String
    Dim CardNumber As String
    CardNumber = getCardNum(strVouchType, bReturnFlag)
    If CardNumber <> "" Then
        'iState 0=表示新增 ,1=删除 ,  2=修改
        Call ProcessUAPMaker(1, DBConn, "", lngID, CardNumber, ErrDesc)
    End If
    If ErrDesc <> "" Then
        DeleteVouch = ErrDesc
        If bTrans Then DBConn.RollbackTrans: bTrans = False
        Exit Function
    End If
    'UAP生单删除 end
    
    
    
    
    
    
    If bTrans = True Then
        DBConn.CommitTrans: bTrans = False
    End If
    Exit Function
    
err_Handle:
    If DBConn.Errors.Count > 0 Then
        If DBConn.Errors(0).NativeError = 644 Then
           DeleteVouch = GetStringPara("U8.SA.USSASERVER.clsvouchsave.00546", sTableName)  'Para zh-CN：严重错误,请检查{0}主表或者子表!
        Else
           DeleteVouch = GetStringPara("U8.SA.USSASERVER.clsvouchsave.00548", strTblName, err.Number & err.Description) 'Para zh-CN：删除{0}失败,{1}
        End If
    End If
    If bTrans = True Then
        DBConn.RollbackTrans
    End If
    DeleteVouch = GetStringPara("U8.SA.USSASERVER.clsvouchsave.00550", strTblName, err.Number & err.Description) 'Para zh-CN：删除{0}失败,错误如下{1}
    
End Function


''删除单据前检查是否可以删除
Private Function CheckForDel(sVouchType As String, domHead As DOMDocument, CN As ADODB.Connection, ErrMsg1 As String) As Boolean
    Dim rec As New ADODB.Recordset
    Dim OriID As String, CurVouchName As String, CurVouchCode As String, sflag As String
    Dim strSQL As String, strUserName As String
    Dim bReturnFlag As Boolean, strUserAuth As String, UFTS As String
    
    On Error GoTo ErrHandle
    
    CheckForDel = False
    rec.CursorLocation = adUseClient
    UFTS = GetHeadItemValue(domHead, "ufts")
    
    If CN.State <> 1 Then
        ErrMsg1 = GetStringPara("U8.SA.USSASERVER.clsvouchsave.00576", clsSys.sDBName)  'Para zh-CN：不能连接到数据库[{0}]
        Exit Function
    End If
    
    bReturnFlag = CBool(Val(GetHeadItemValue(domHead, "breturnflag")))
    
    Select Case sVouchType
        Case "87", "88", "89", "90", "96", "97"
            CurVouchName = "设置"
            sflag = "SZ"
            OriID = GetHeadItemValue(domHead, "id")
            CurVouchCode = GetHeadItemValue(domHead, "ccode")
            strSQL = "select cMaker,isnull(cverifier,N'') as cverifier,cVouchType from MT_baseset  " _
                & " where MT_baseset.id=" & OriID & IIf(UFTS = "", "", " AND convert(char, convert(money,ufts), 2)= N'" & UFTS & "'")
        Case "91", "92", "93", "94", "95", "98"
            CurVouchName = "预算"
            sflag = "Budget"
            OriID = GetHeadItemValue(domHead, "id")
            CurVouchCode = GetHeadItemValue(domHead, "ccode")
            strSQL = "select cMaker,isnull(cverifier,N'') as cverifier,cVouchType from MT_budget  " _
                & " where MT_budget.id=" & OriID & IIf(UFTS = "", "", " AND convert(char, convert(money,ufts), 2)= N'" & UFTS & "'")
        Case "EFBWGL020601"
        
            OriID = GetHeadItemValue(domHead, "id")
            CurVouchCode = GetHeadItemValue(domHead, "ccode")
            strSQL = "select cMaker,isnull(cverifier,N'') as cverifier from EFBWGL_v_CipDataT  " _
                & " where EFBWGL_v_CipDataT.id=" & OriID & IIf(UFTS = "", "", " AND convert(char, convert(money,ufts), 2)= N'" & UFTS & "'")
    End Select
    
    
    OriID = GetHeadItemValue(domHead, "id")
    CurVouchCode = GetHeadItemValue(domHead, "ccode")
    
    Dim sql As String
    Dim strMainTable As String
    Dim Vouch_name As String
    
    sql = "select name, BTTblName,BTQName,BWTblName,BWQName,VchListQName from Vouchers " & _
        "where [CardNumber] ='" & sVouchType & "'"
    With rec
        If .State <> 0 Then .Close
        .CursorLocation = adUseClient
        .Open sql, CN, 3, 4
        strMainTable = .Fields("BTQName")
        Vouch_name = .Fields("name")
    End With
    
    strSQL = "select ccode ,cMaker,isnull(cverifier,N'') as cverifier from  " & strMainTable & " " _
            & " where id=" & OriID & IIf(UFTS = "", "", " AND convert(char, convert(money,ufts), 2)= N'" & UFTS & "'")
            
    If rec.State <> 0 Then rec.Close
    
    rec.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    If rec.EOF Then
        ErrMsg1 = GetStringPara("U8.SA.USSASERVER.clsvouchsave.00598", CurVouchName, CurVouchCode) 'Para zh-CN：{0}[{1}]不存在或者被其他操作员修改!
        rec.Close
        Set rec = Nothing
        GoTo DOExit
    End If
    Select Case sVouchType
        Case ""
        Case Else
            If rec("cverifier") <> "" Then
'                ErrMsg1 = GetStringPara("U8.SA.USSASERVER.clsvouchsave.00599", CurVouchName, CurVouchCode)  'Para zh-CN：{0}[{1}]已经审核!
                
                ErrMsg1 = "【" & Vouch_name & "】" & rec.Fields("cCode").value & "号单据已经审核，不能删除！"
                rec.Close
                Set rec = Nothing
                GoTo DOExit
            End If
    End Select
    
 
    If ErrMsg1 <> "" Then
        GoTo DOExit
    End If
    
    CheckForDel = True

    Exit Function

DOExit:
    CheckForDel = False
    Exit Function
    
ErrHandle:
     
    CheckForDel = False
    If err.Number <> 0 Then
        If ErrMsg1 = "" Then ErrMsg1 = "Public Function: clsSave.CheckforDel " & vbCrLf & "Error(" & err.Number & "):" & err.Description
    End If
    Set rec = Nothing
End Function

'================================================================================
'模块: PublicSub   Function: GetAttrValue
'--------------------------------------------------------------------------------
'说明: 获得某一节点的属性名称及值
'================================================================================
'Private Function bAttrExist(ByVal ele As IXMLDOMElement, ByVal sAttrName As String, Optional strName As String) As Boolean
Private Function GetAttrValue(ByVal ele As IXMLDOMElement, ByVal sAttrName As String, Optional strName As String) As String
    
    On Error GoTo Error_General_Handler:
    Const PROC_SIG As String = "PublicSub:bAttrExist"
    Dim Attr As IXMLDOMAttribute
    
    Set Attr = ele.getAttributeNode(LCase(sAttrName))
    If Attr Is Nothing Then
        For Each Attr In ele.Attributes
            If UCase(Attr.Name) = UCase(sAttrName) Then
               strName = Attr.Name
               GetAttrValue = Attr.value
               Exit For
            End If
        Next
    Else
        strName = Attr.Name
        GetAttrValue = Attr.value
    End If
    Exit Function
Error_General_Handler:
  Select Case err.Number
    Case Else
        GetAttrValue = ""
  End Select
End Function

'//组织Insert into语句
Public Function GetStrInsertFromDoc(ByVal eleMent As IXMLDOMElement, ByVal sTableName As String, ByVal DBConn As ADODB.Connection, Optional bHead As Boolean, Optional lngMainID As String, Optional lngSubID As String, Optional strSOCode As String) As String
    Dim strSQL As String
    Dim RecTemp As ADODB.Recordset
    Set RecTemp = New ADODB.Recordset
    Dim strValues As String
    Dim i As Long
    Dim strID As String
    Dim strSubID As String
    Dim ArrDomElement As IXMLDOMAttribute
    RecTemp.CursorLocation = adUseClient
    strSQL = "SELECT * FROM " & sTableName & " WHERE 1=0"
    RecTemp.Open strSQL, DBConn.ConnectionString, adOpenForwardOnly, adLockReadOnly
    Dim strValue As String
    Dim strTableName As String
    strSQL = "("
    strValues = "("
    Dim strFieldName As String
    Dim strItem As String
    For Each ArrDomElement In eleMent.Attributes
        strFieldName = ArrDomElement.Name
        Debug.Print strFieldName
        If FieldInFields(strFieldName, RecTemp.Fields) Then
            If LCase(strFieldName) <> "editprop" And LCase(strFieldName) <> "ufts" And LCase(strFieldName) <> "itimage" Then
               If RecTemp.Fields(strFieldName).Attributes <> 16 Then   '自动编号
                    If IsNull(RecTemp.Fields(strFieldName).Properties("BASETABLENAME")) Then
                        strSQL = strSQL & strFieldName & ","
                    ElseIf (LCase(RecTemp.Fields(strFieldName).Properties("BASETABLENAME")) = LCase(sTableName)) Then
                        strSQL = strSQL & strFieldName & ","
                    End If
                    strItem = Trim(ArrDomElement.value)
                    Select Case bHead
                        Case True
                            If UCase(strFieldName) = UCase(strID) Then
                                strItem = lngMainID
                            End If
                        Case False
                            If UCase(strFieldName) = UCase(strSubID) Then
                                strItem = lngSubID
                            Else
                                If UCase(strFieldName) = UCase(strID) Then
                                    strItem = lngMainID
                                End If
                            End If
                    End Select
                    Select Case RecTemp.Fields(strFieldName).Type
                        Case adVarChar, adVarWChar, adChar, adLongVarChar, adLongVarWChar, adWChar, 130  'string 10
                            strValues = strValues & IIf(strItem = "", "Null", "'" & strItem & "'") & ","
                        Case adDouble, adDecimal, adNumeric, adCurrency, adUnsignedTinyInt 'Doublue 7
                            If strItem = "" Then
                                strItem = "Null"
                            Else
                                strItem = CDbl(strItem)
                            End If
                            strValues = strValues & strItem & ","
                        Case adDate, adDBTime, adDBTimeStamp      'Date 8
                            strValues = strValues & IIf(strItem = "", "Null", "'" & Mid(strItem, 1, 10) & "'") & ","
                        Case adInteger, adSmallInt, adTinyInt, adSingle, adBigInt
                            If UCase(strFieldName) = "IDISP" Or UCase(strFieldName) = "ISALE" Then
                               strItem = IIf(strItem = "", 0, strItem)
                            Else
                               If strItem = "" Then
                                   strItem = "Null"
                               Else
                                   strItem = CDbl(strItem)
                               End If
                            End If
                            strValues = strValues & strItem & ","
                        Case adBoolean
                            If strItem = "" Then
                               strItem = 0
                            End If
                            If strItem = "是" Then strItem = "1"
                            If strItem = "否" Then strItem = "0"
                            strValues = strValues & IIf(CBool(strItem) = True, 1, 0) & ","
                        Case adGUID
                            If strItem = "" Then
                                strItem = "Null"
                            Else
                                strItem = "'" + Replace(Replace(strItem, "{", ""), "}", "") + "'"
                            End If
                            strValues = strValues & strItem & ","
                        Case Else
                            MsgBox strFieldName
                            MsgBox RecTemp.Fields(strFieldName).Type
                            MsgBox "Other"
                        End Select
                End If
            End If
        End If
    Next
    strItem = ""
    strSQL = " Insert Into " & sTableName & Left(strSQL, Len(strSQL) - 1) & ")"
    strValues = Left(strValues, Len(strValues) - 1) & ")"
    GetStrInsertFromDoc = strSQL & " Values " & strValues
    If RecTemp.State = 1 Then RecTemp.Close
    Set RecTemp = Nothing
End Function


Public Function FieldInFields(strFieldName As String, Flds As ADODB.Fields) As Boolean
Dim i As Long
FieldInFields = False
For i = 0 To Flds.Count - 1
    If UCase(Flds(i).Name) = UCase(strFieldName) Then
        FieldInFields = True
        Exit For
    End If
Next
End Function

'//组织Update语句
Public Function GetStrUpdateFromDoc(ByVal eleMent As IXMLDOMElement, ByVal sTableName As String, ByVal DBConn As ADODB.Connection, ByVal sKeyField As String, Optional ByVal dUfts As String, Optional lngMainID As String, Optional strSCOde As String) As String
    Dim strSQL As String
    Dim RecTemp As ADODB.Recordset
    Set RecTemp = New ADODB.Recordset
    Dim strValues As String
    Dim i As Long
    Dim ArrDomElement As IXMLDOMAttribute
    RecTemp.CursorLocation = adUseClient
    strSQL = "SELECT * FROM " & sTableName & " WHERE 1=0"
    RecTemp.Open strSQL, DBConn.ConnectionString, adOpenForwardOnly, adLockReadOnly
    Dim strValue As String
    Dim strFieldName As String
    Dim strItem As String
    strSQL = ""
    For Each ArrDomElement In eleMent.Attributes
        strFieldName = ArrDomElement.Name
        If FieldInFields(strFieldName, RecTemp.Fields) Then
            If LCase(strFieldName) <> "editprop" And LCase(strFieldName) <> "ufts" And LCase(strFieldName) <> "itimage" Then
                If UCase(sKeyField) = UCase(strFieldName) Then
                    lngMainID = Val(ArrDomElement.value)
                End If
                If Not (sTableName = "SO_SODetails" And LCase(strFieldName) = "autoid") And Not (sTableName = "SalePayVouchs" And LCase(strFieldName) = "autoid") And Not (sTableName = "ExpenseVouchs" And LCase(strFieldName) = "autoid") And Not (sTableName = "Sa_WrapLease" And LCase(strFieldName) = "autoid") And Not (LCase(sTableName) = LCase("DispatchLists") And LCase(strFieldName) = "autoid") Then          ' (Not recTemp.Fields(strFieldName).Properties("ISAUTOINCREMENT")) Then
                    If IsNull(RecTemp.Fields(strFieldName).Properties("BASETABLENAME")) Then
                        strSQL = strSQL & strFieldName & "="
                    ElseIf (LCase(RecTemp.Fields(strFieldName).Properties("BASETABLENAME")) = LCase(sTableName)) Then
                        strSQL = strSQL & strFieldName & "="
                    End If
                    strItem = Trim(ArrDomElement.value)
                    Select Case RecTemp.Fields(strFieldName).Type
                        Case adVarChar, adVarWChar, adChar, adLongVarChar, adLongVarWChar, adWChar, 130  'string 10
                            strSQL = strSQL & IIf(strItem = "", "Null", "'" & strItem & "'") & ","
                        Case adDouble, adDecimal, adNumeric, adCurrency, adUnsignedTinyInt 'Doublue 7
                            If strItem = "" Then
                                strItem = "Null"
                            Else
                                strItem = CDbl(strItem)
                            End If
                            strSQL = strSQL & IIf(strItem = "", "Null", strItem) & ","
                        Case adDate, adDBTime, adDBTimeStamp      'Date 8
                            If Right(strFieldName, 7) = "systime" Then
                                strSQL = strSQL & "convert(varchar,getdate(),120),"
                            Else
                                strSQL = strSQL & IIf(strItem = "", "Null", "'" & Mid(strItem, 1, 10) & "'") & ","
                            End If
                        Case adInteger, adSmallInt, adTinyInt, adSingle, adBigInt
                            If strItem = "" Then
                                 strItem = "Null"
                            Else
                                 strItem = CDbl(strItem)
                            End If
                            strSQL = strSQL & IIf(strItem = "", "Null", strItem) & "," 'Val(strItem)
                        Case adBoolean
                            If strItem = "" Then
                                  strItem = 0
                            End If
                            If strItem = "是" Then strItem = "1"
                            If strItem = "否" Then strItem = "0"
                            strSQL = strSQL & IIf(CBool(strItem) = True, "1", "0") & ","
                        Case adGUID
                            If strItem = "" Then
                                 strItem = "Null"
                            Else
                                     strItem = "'" + Replace(Replace(strItem, "{", ""), "}", "") + "'"
                            End If
                            strSQL = strSQL & IIf(strItem = "", "Null", strItem) & ","
                        Case Else
                             
                            MsgBox "Other"
                    End Select
                End If
            End If
        End If
        strItem = ""
    Next
    If LCase(sTableName) = "so_sodetails" Then
        strSQL = " Update " & sTableName & " SET " & strSQL & "cSOCode='" & strSCOde & "'"
    Else
        strSQL = " Update " & sTableName & " SET " & Left(strSQL, Len(strSQL) - 1)
    End If
    strSQL = strSQL & " Where  " & sKeyField & "= " & lngMainID
    GetStrUpdateFromDoc = strSQL
    If RecTemp.State = 1 Then RecTemp.Close
    Set RecTemp = Nothing
End Function


Public Function CheckVouchCode(strVouchType As String, sHeadTableName As String, strFldVouchCode As String, strVouchCode As String, bBatchDispToBill As Boolean, bRepeatReDo As Boolean, strMsg As String, domHead As DOMDocument, DBConn As ADODB.Connection, iState As Integer, lngMainID As String, strID As String) As Boolean
    Dim objBillNo As New UFBillComponent.clsBillComponent
    Dim strSeedValue As String
    Dim strSeedName As String
    Dim strSQL As String
    Dim iGetVouchCodeCount As Integer
    Dim RecTemp As New ADODB.Recordset
    Dim eleMent As IXMLDOMElement
    Dim strBVouchCode As String   '修改前的单据编号
    Dim strWVouch As String
    On Error GoTo err_Handle
    RecTemp.CursorLocation = adUseClient
GetVouchCode:
        Set eleMent = domHead.selectSingleNode("//z:row")
        If iState <> 0 Then '修改状态，获得修改前的单据编号
            If strVouchCode <> "" Then
                Select Case LCase(sHeadTableName)
                     Case "dispatchlist", "salebillvouch"
                          If iState <> 0 Then
                                strSQL = " SELECT " & strID & "," & strFldVouchCode & " AS cVouchCode FROM " & sHeadTableName & _
                                         " WHERE " & strID & "=" & lngMainID & " AND " & strFldVouchCode & " ='" & strVouchCode & "'"
                          Else
                                strSQL = "SELECT " & strFldVouchCode & "  AS cVouchCode FROM " & sHeadTableName & " WHERE " & strFldVouchCode & "='" & strVouchCode & "'  AND cVouchType='" & strVouchType & "'"
                          End If
                     Case Else
                          If iState <> 0 Then
                                strSQL = " SELECT " & strID & "," & strFldVouchCode & " AS cVouchCode FROM " & sHeadTableName & _
                                         " WHERE " & strID & "=" & lngMainID '& " AND " & strFldVouchCode & " ='" & strVouchCode & "'"
                          Else
                                strSQL = "SELECT " & strFldVouchCode & " AS cVouchCode FROM " & sHeadTableName & " WHERE " & strFldVouchCode & "='" & strVouchCode & "'"
                          End If
                End Select
                If RecTemp.State <> adStateClosed Then
                     RecTemp.Close
                End If
        
                RecTemp.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
                If RecTemp.RecordCount > 0 Then
                    strBVouchCode = RecTemp!cVouchCode
                End If
                If RecTemp.State <> adStateClosed Then
                     RecTemp.Close
                End If
            End If
        End If
        
        If bBatchDispToBill Then
           bRepeatReDo = True
        End If
        If strVouchCode = "" Then
            If bRepeatReDo = False Then  '(或虽然单据编号允许手工输入，但当前操作是批量生单)，" & vbCr & "
                strMsg = "该单据非自动编号，当前单据号重复或非法，请手工修改单据编号或修改单据编号允许自动重取!"
                If RecTemp.State <> adStateClosed Then
                     RecTemp.Close
                End If
                Set RecTemp = Nothing
                Set objBillNo = Nothing
                Exit Function
            End If
           
            If iGetVouchCodeCount < 2 And bRepeatReDo Then
                If bOnceGetTrueNO = False Then
                    If GetVouchNO(clsSys.Connectstr, strVouchType, domHead, strVouchCode, strMsg, , , , bRepeatReDo, clsSys.sRemoteID, strSeedName, strSeedValue, True) = False Then
                    
                        CheckVouchCode = False
                        If RecTemp.State <> adStateClosed Then
                             RecTemp.Close
                        End If
                        Set RecTemp = Nothing
                        Set objBillNo = Nothing
                        Exit Function
                    End If
                    ''保证只取1次真号
                    bOnceGetTrueNO = True
                End If
            Else
                strMsg = "无法取得单据编号，请稍后再试!"
                If RecTemp.State <> adStateClosed Then
                     RecTemp.Close
                End If
                Set RecTemp = Nothing
                Set objBillNo = Nothing
                Exit Function
            End If
            If bUniCode(sHeadTableName, strFldVouchCode, strVouchCode, strVouchType, DBConn) Then
                 iGetVouchCodeCount = iGetVouchCodeCount + 1
                 If iGetVouchCodeCount < 2 And bRepeatReDo And Not bOnceGetTrueNO Then
                    strVouchCode = ""
                 End If
                 GoTo GetVouchCode
            End If
        Else
            If iState <> 0 Then
               If strBVouchCode = strVouchCode Then
                   CheckVouchCode = True
                    If RecTemp.State <> adStateClosed Then
                         RecTemp.Close
                    End If
                    Set RecTemp = Nothing
                    Set objBillNo = Nothing
                    Exit Function
               End If
            
            End If
            
            If iGetVouchCodeCount > 2 Then
                If Trim(strSeedName) <> "" Then
                    strSeedValue = GetAttrValue(eleMent, strSeedName)
                End If
                objBillNo.InitBill DBConn.ConnectionString, GetstrCardNum(strVouchType)
                Select Case strVouchType
                    Case "05"
                        strWVouch = "cvouchtype='05'"
                    Case "06", "00"
                        strWVouch = "(cvouchtype='06' or cvouchtype='00' )"
                    Case "26", "27", "28", "29"
                        strWVouch = "(cvouchtype='" & strVouchType & "')"
                    Case "94"
                        strWVouch = "cvouchtype='04'"
                End Select
                If objBillNo.DataSynchronization(sHeadTableName, strFldVouchCode, strVouchCode, strSeedValue, strWVouch) = True Then
                    If GetVouchNO(clsSys.Connectstr, strVouchType, domHead, strVouchCode, strMsg, , , , bRepeatReDo, clsSys.sRemoteID, strSeedName, strSeedValue, True) = False Then
                        If RecTemp.State <> adStateClosed Then
                             RecTemp.Close
                        End If
                        Set objBillNo = Nothing
                        Set RecTemp = Nothing
                        Exit Function
                    End If
                    If bUniCode(sHeadTableName, strFldVouchCode, strVouchCode, strVouchType, DBConn) Then
                        strMsg = "单据控件无法自动获得一个正确的单据编号,请手工修改后再试!"
                        If RecTemp.State <> adStateClosed Then
                             RecTemp.Close
                        End If
                        Set RecTemp = Nothing
                        Set objBillNo = Nothing
                        Exit Function
                    End If
                    
                    
                Else
                    strMsg = "单据控件同步存取单据编号不成功,请稍后再试!"
                    If RecTemp.State <> adStateClosed Then
                         RecTemp.Close
                    End If
                    Set RecTemp = Nothing
                    Set objBillNo = Nothing
                    Exit Function
                End If
            Else
                If bUniCode(sHeadTableName, strFldVouchCode, strVouchCode, strVouchType, DBConn) Then
                         iGetVouchCodeCount = iGetVouchCodeCount + 1
                         If bRepeatReDo Then    '自动重取
                            If iGetVouchCodeCount < 2 And bOnceGetTrueNO = False Then
                               strVouchCode = ""
                            End If
                            GoTo GetVouchCode
                         Else
                            strMsg = "该单据非自动编号，单据编号重复,请手工修改或稍后再试!"
                            If RecTemp.State <> adStateClosed Then
                                 RecTemp.Close
                            End If
                            Set RecTemp = Nothing
                            Set objBillNo = Nothing
                            Exit Function
                         End If
                    Else
                        'by ahzzd 20060907 不重号的时候  bOnceGetTrueNO
                        Dim vouchcode As String
                        Call GetVouchNO(clsSys.Connectstr, strVouchType, domHead, vouchcode, strMsg, , , , bRepeatReDo, clsSys.sRemoteID, strSeedName, strSeedValue, False)
                        
                        If (bOnceGetTrueNO = False) And strVouchCode = vouchcode Then
                            If GetVouchNO(clsSys.Connectstr, strVouchType, domHead, strVouchCode, strMsg, , , , bRepeatReDo, clsSys.sRemoteID, strSeedName, strSeedValue, True) = False Then
                                CheckVouchCode = False
                                If RecTemp.State <> adStateClosed Then
                                     RecTemp.Close
                                End If
                                Set RecTemp = Nothing
                                Set objBillNo = Nothing
                                Exit Function
                            End If
                            '保证只取1次真号
                            bOnceGetTrueNO = True
                        End If
                      '//////////////////////////////////'by ahzzd 20060907 不重号的时候
                    
                    End If
                End If
        End If
        CheckVouchCode = True
        If RecTemp.State <> adStateClosed Then
            RecTemp.Close
        End If
        Set RecTemp = Nothing
        Set objBillNo = Nothing
        Exit Function
err_Handle:
        strMsg = "Function: CheckVouchCode" & strMsg & strMsg & err.Description & vbCr & IIf(iGetVouchCodeCount >= 5, "单据控件同步存取单据编号不成功", "")
        CheckVouchCode = False
        If RecTemp.State <> adStateClosed Then
            RecTemp.Close
        End If
        Set RecTemp = Nothing
        Set objBillNo = Nothing
        Exit Function
End Function

'是否唯一编码
Private Function bUniCode(sHeadTableName As String, strFldVouchCode As String, strVouchCode As String, strVouchType As String, DBConn As ADODB.Connection) As Boolean
    Dim strSQL As String
    Dim RecTemp As New ADODB.Recordset
    RecTemp.CursorLocation = adUseClient
    Select Case LCase(sHeadTableName)
         Case "dispatchlist", "salebillvouch"
              strSQL = "SELECT " & strFldVouchCode & " FROM " & sHeadTableName & " WHERE " & strFldVouchCode & "='" & strVouchCode & "'  AND cVouchType='" & strVouchType & "'"
           

         Case Else
              strSQL = "SELECT " & strFldVouchCode & " FROM " & sHeadTableName & " WHERE " & strFldVouchCode & "='" & strVouchCode & "'AND cVouchType ='" & strVouchType & "'"
    End Select
    If RecTemp.State <> adStateClosed Then
        RecTemp.Close
    End If
    RecTemp.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly
    If RecTemp.RecordCount > 0 Then
        bUniCode = True
    Else
        bUniCode = False
    End If
    If RecTemp.State <> adStateClosed Then
        RecTemp.Close
    End If
    Set RecTemp = Nothing
    Exit Function
End Function


Private Function GetStrInsertFromDoc2(RecTemp As ADODB.Recordset, ByVal eleMent As IXMLDOMElement, ByVal sTableName As String, ByVal DBConn As ADODB.Connection, Optional bHead As Boolean, Optional lngMainID As String, Optional lngSubID As String, Optional strSOCode As String) As String
    Dim strSQL As String
    Dim strValues As String
    Dim i As Long
    Dim strID As String
    Dim strSubID As String
    Dim ArrDomElement As IXMLDOMAttribute
    If RecTemp.State <> 1 Then
        RecTemp.CursorLocation = adUseClient
        strSQL = "SELECT * FROM " & sTableName & " WHERE 1=0 "
        RecTemp.Open strSQL, DBConn.ConnectionString, adOpenForwardOnly, adLockReadOnly
    End If
    Dim strValue As String
    Select Case LCase(sTableName)
        Case Else
                strSubID = "AutoID"
                strID = "id"
    End Select
    strSQL = "("
    strValues = "("
    Dim strFieldName As String
    Dim strItem As String
    If bHead = False And strSubID <> "" Then
        eleMent.setAttribute LCase(strSubID), lngSubID
    End If
    For Each ArrDomElement In eleMent.Attributes
        strFieldName = ArrDomElement.Name
        If FieldInFields(strFieldName, RecTemp.Fields) Then
            If LCase(strFieldName) <> "editprop" And LCase(strFieldName) <> "ufts" Then
               If RecTemp.Fields(strFieldName).Attributes <> 16 Then   '自动编号
                    If IsNull(RecTemp.Fields(strFieldName).Properties("BASETABLENAME")) Then
                        strSQL = strSQL & strFieldName & ","
                    ElseIf (LCase(RecTemp.Fields(strFieldName).Properties("BASETABLENAME")) = LCase(sTableName)) Then
                        strSQL = strSQL & strFieldName & ","
                    End If
                        strItem = Trim(ArrDomElement.value)
                        Select Case bHead
                            Case True
                                If UCase(strFieldName) = UCase(strID) Then
                                    strItem = lngMainID
                                End If
    
                            Case False
                                If UCase(strFieldName) = UCase(strSubID) Then
                                    strItem = lngSubID
                                Else
                                    If UCase(strFieldName) = UCase(strID) Then
                                        strItem = lngMainID
                                    End If
                                End If
                        End Select
                        Select Case RecTemp.Fields(strFieldName).Type
                             Case adVarChar, adVarWChar, adChar, adLongVarChar, adLongVarWChar, adWChar, 130  'string 10
                                  strValues = strValues & IIf(strItem = "", "Null", "'" & strItem & "'") & ","
                             Case adDouble, adDecimal, adNumeric, adCurrency, adUnsignedTinyInt 'Doublue 7
                                   If strItem = "" Then
                                        strItem = "Null"
                                   Else
                                            strItem = CDbl(strItem)
                                   End If
                                   strValues = strValues & strItem & ","
                             Case adDate, adDBTime, adDBTimeStamp      'Date 8
                                  strItem = Mid(strItem, 1, 10)
                                  If strItem <> "" Then
                                    If CDate(Format(strItem, "yyyy-mm-dd")) < CDate("1900-01-01") Then
                                        strItem = "1900-01-01"
                                    End If
                                  End If
                                  strValues = strValues & IIf(strItem = "", "Null", "'" & strItem & "'") & ","
                             Case adInteger, adSmallInt, adTinyInt, adSingle, adBigInt
                                     If UCase(strFieldName) = "IDISP" Or UCase(strFieldName) = "ISALE" Then
                                        strItem = IIf(strItem = "", 0, strItem)
                                     Else
                                        If strItem = "" Then
                                            strItem = "Null"
                                        Else
                                            strItem = CDbl(strItem)
                                        End If
                                     End If
                                     strValues = strValues & strItem & ","
                             Case adBoolean
                                  If strItem = "" Then
                                     strItem = 0
                                  End If
                                  If strItem = "是" Then strItem = "1"
                                  If strItem = "否" Then strItem = "0"
                                  strValues = strValues & IIf(CBool(strItem) = True, 1, 0) & ","
                             Case adGUID
                                   If strItem = "" Then
                                        strItem = "Null"
                                   Else
                                            strItem = "'" + Replace(Replace(strItem, "{", ""), "}", "") + "'"
                                   End If
                                   strValues = strValues & strItem & ","
                             Case Else
                                  MsgBox strFieldName
                                  MsgBox RecTemp.Fields(strFieldName).Type
                                  MsgBox "Other"
                        End Select
'                    End If
                End If
            End If
        End If
    Next
    strItem = ""
    strSQL = " Insert Into " & sTableName & Left(strSQL, Len(strSQL) - 1) & ")"
    strValues = Left(strValues, Len(strValues) - 1) & ")"
    GetStrInsertFromDoc2 = strSQL & " Values " & strValues

End Function

Public Function GetStrUpdateFromDoc2(RecTemp As ADODB.Recordset, ByVal eleMent As IXMLDOMElement, ByVal sTableName As String, ByVal DBConn As ADODB.Connection, ByVal sKeyField As String, Optional ByVal dUfts As String, Optional lngMainID As String, Optional strSCOde As String) As String
    Dim strSQL As String
    Dim strValues As String
    Dim i As Long
    Dim ArrDomElement As IXMLDOMAttribute
    
    If RecTemp.State <> 1 Then
        RecTemp.CursorLocation = adUseClient
        strSQL = "SELECT * FROM " & sTableName & " WHERE 1=0"
        RecTemp.Open strSQL, DBConn.ConnectionString, adOpenForwardOnly, adLockReadOnly
    End If
    Dim strValue As String
    Dim strFieldName As String
    Dim strItem As String
    strSQL = ""
    For Each ArrDomElement In eleMent.Attributes
        strFieldName = ArrDomElement.Name
        If FieldInFields(strFieldName, RecTemp.Fields) Then
                If LCase(strFieldName) <> "editprop" And LCase(strFieldName) <> "ufts" Then
                    If UCase(sKeyField) = UCase(strFieldName) Then
                        lngMainID = Val(ArrDomElement.value)
                    End If
                     If Not (sTableName = "SO_SODetails" And LCase(strFieldName) = "autoid") And Not (sTableName = "SalePayVouchs" And LCase(strFieldName) = "autoid") And Not (sTableName = "ExpenseVouchs" And LCase(strFieldName) = "autoid") And Not (sTableName = "Sa_WrapLease" And LCase(strFieldName) = "autoid") And Not (LCase(sTableName) = LCase("DispatchLists") And LCase(strFieldName) = "autoid") Then
                                If IsNull(RecTemp.Fields(strFieldName).Properties("BASETABLENAME")) Then
                                    strSQL = strSQL & strFieldName & "="
                                ElseIf (LCase(RecTemp.Fields(strFieldName).Properties("BASETABLENAME")) = LCase(sTableName)) Then
                                    strSQL = strSQL & strFieldName & "="
                                End If
                                strItem = Trim(ArrDomElement.value)
                                 Select Case RecTemp.Fields(strFieldName).Type
                                     Case adVarChar, adVarWChar, adChar, adLongVarChar, adLongVarWChar, adWChar, 130  'string 10
                                          strSQL = strSQL & IIf(strItem = "", "Null", "'" & strItem & "'") & ","
                                     Case adDouble, adDecimal, adNumeric, adCurrency, adUnsignedTinyInt 'Doublue 7
                                               If strItem = "" Then
                                                    strItem = "Null"
                                               Else
                                                    strItem = CDbl(strItem)
                                               End If
                                          strSQL = strSQL & IIf(strItem = "", "Null", strItem) & ","

                                          
                                     Case adDate, adDBTime, adDBTimeStamp      'Date 8
                                        strItem = Mid(strItem, 1, 10)
                                        If strItem <> "" Then
                                            If CDate(Format(strItem, "yyyy-mm-dd")) < "1900-01-01" Then
                                                strItem = "1900-01-01"
                                            End If
                                        End If
                                          strSQL = strSQL & IIf(strItem = "", "Null", "'" & strItem & "'") & ","
                                     Case adInteger, adSmallInt, adTinyInt, adSingle, adBigInt
                                               If strItem = "" Then
                                                    strItem = "Null"
                                               Else
                                                    strItem = CDbl(strItem)
                                               End If
                                             
                                             
                                              strSQL = strSQL & IIf(strItem = "", "Null", strItem) & "," 'Val(strItem)
                                     
                                     
                                     Case adBoolean
                                          If strItem = "" Then
                                                strItem = 0
                                          End If
                                          If strItem = "是" Then strItem = "1"
                                          If strItem = "否" Then strItem = "0"
                                          strSQL = strSQL & IIf(CBool(strItem) = True, "1", "0") & ","
                                     Case adGUID
                                            If strItem = "" Then
                                                 strItem = "Null"
                                            Else
                                                     strItem = "'" + Replace(Replace(strItem, "{", ""), "}", "") + "'"
                                            End If
                                            strSQL = strSQL & IIf(strItem = "", "Null", strItem) & ","
                                     Case Else
                                     
                                          MsgBox "Other update SQL组织 有问题!"
                                 End Select
                    End If
                End If
        End If
        strItem = ""
    Next
    If LCase(sTableName) = "so_sodetails" Then
        strSQL = " Update " & sTableName & " SET " & strSQL & "cSOCode='" & strSCOde & "'"
    Else
        strSQL = " Update " & sTableName & " SET " & Left(strSQL, Len(strSQL) - 1)
    End If
    strSQL = strSQL & " Where  " & sKeyField & "= " & lngMainID & _
             IIf(dUfts = "", "", " AND convert(char, convert(money,ufts), 2)= '" & dUfts & "'")
    GetStrUpdateFromDoc2 = strSQL

End Function

Public Function AttrExist(ByVal ele As IXMLDOMElement, ByVal sAttr As String) As Boolean
    AttrExist = False
    Dim i As Integer
    If ele Is Nothing Then
        Exit Function
    End If
    sAttr = LCase(sAttr)
    On Error GoTo Err_info
    Dim nd As IXMLDOMAttribute
    Set nd = ele.Attributes.getNamedItem(sAttr)
    If nd Is Nothing Then
        Exit Function
    End If
    AttrExist = True
    Exit Function
Err_info:
    AttrExist = False '没有该元素节点
End Function


Private Function GetAttachDOM(strVouchType As String, domHead As DOMDocument, oAttachDom As DOMDocument, bSaveAttach As Boolean) As Boolean
    Dim nd As IXMLDOMNode
    Dim strXMl As String
    Dim ndrs As IXMLDOMNode
    If domHead.selectNodes("//rs:data/voucherattached").length = 0 Then
        bSaveAttach = False
        GetAttachDOM = True
        Exit Function
    Else
        bSaveAttach = True
    End If
    Set oAttachDom = New DOMDocument
    Set ndrs = domHead.selectSingleNode("//rs:data")
    Set nd = domHead.selectSingleNode("//rs:data/voucherattached")
    strXMl = nd.Text
    ndrs.removeChild nd
    Set nd = Nothing
    oAttachDom.loadXML strXMl
End Function

'存附件保存
Private Function SaveAttach(CN As ADODB.Connection, strVouchType As String, oDomH As DOMDocument, oAttachDom As DOMDocument, sCardNumber As String, ErrMsg As String) As Boolean
    Dim oVouchServer As Object
    Dim strXMl As String
    Dim sMainID As String
    Dim bRed As Boolean, sReturnFlag As String
    Set oVouchServer = CreateObject("UFVoucherServer85.clsVoucherTemplate")
    sReturnFlag = GetHeadItemValue(oDomH, "breturnflag")
    If sReturnFlag <> "" Then
        bRed = CBool(sReturnFlag)
    End If
    sMainID = "id"
    Call oAttachDom.documentElement.setAttribute("VoucherTypeID", sCardNumber)
    Call oAttachDom.documentElement.setAttribute("VoucherID", GetHeadItemValue(oDomH, sMainID))
    strXMl = oAttachDom.xml
    ErrMsg = ""
    Call oVouchServer.SaveAccessories(strXMl, CN, ErrMsg)
    If ErrMsg = "" Then
        SaveAttach = True
    End If
End Function

'861存附件删除
Private Function DelVouchAttach(CN As ADODB.Connection, strVouchType As String, oDomH As DOMDocument, ErrMsg As String) As Boolean
    Dim oVouchServer As Object
    Dim strXMl As String
    Dim sMainID As String
    Dim sCardNumber As String
    Set oVouchServer = CreateObject("UFVoucherServer85.clsVoucherTemplate")
    sMainID = GetHeadItemValue(oDomH, "id") '得到单据关键字的值
    sCardNumber = GetstrCardNum(strVouchType)
    strXMl = "<Data AccessoriesChanged='1' Deleted='1' VoucherTypeID='" + sCardNumber + "' VoucherID='" + sMainID + "' />"
    Call oVouchServer.SaveAccessories(strXMl, CN, ErrMsg)
    DelVouchAttach = True
End Function






Private Sub Class_Initialize()
    Set cls_Public = CreateObject("UF_Public_base.cls_log")
End Sub
