Attribute VB_Name = "ModVoucher"
Option Explicit

Public Enum EState
    LVRead = 0
    lvSave = 1
    LVSaveDefault = 2
End Enum

Public Function bSetNoEANoCheck(domHead As DOMDocument, strErrMsg As String) As Boolean
    Dim ele As IXMLDOMElement
    
    On Error GoTo DoErr
    Set ele = domHead.selectSingleNode("//z:row")
    ele.setAttribute "bnochecker", "1"
    ele.setAttribute "cnochecker", strErrMsg
    Set ele = Nothing
    bSetNoEANoCheck = True
    Exit Function
DoErr:
    Set ele = Nothing
End Function

Public Sub Num2Chinese(strSumXX, strSumDX)
    Dim oNum2Chinese As Object
    Set oNum2Chinese = CreateObject("FormulaParse.Calculator")
    strSumDX = ""
    oNum2Chinese.Num2Chinese strSumXX, strSumDX
    If strSumDX = "圆整" Then
        strSumDX = "零圆零角零分"
    Else
        If Left(strSumDX, Len("圆")) = "圆" Then
            'strSumDX = "零" + strSumDX
            strSumDX = Mid(strSumDX, 2)
        End If
        If Left(strSumDX, Len("零")) = "零" Then
            strSumDX = Mid(strSumDX, 2)
        End If
        If Left(strSumDX, Len("角")) = "角" Then
            strSumDX = Mid(strSumDX, 2)
        End If
        If Right(strSumDX, Len("角")) = "角" Then
            strSumDX = strSumDX + "整"
        End If
        If Abs(strSumXX) < 0.1 Then
            strSumDX = strSumDX + "`"
        End If
    End If
    Set oNum2Chinese = Nothing
    Exit Sub
End Sub

Public Function GetHeadItemValue(ByVal domHead As DOMDocument, ByVal sKey As String) As String
    sKey = LCase(sKey)
    If Not domHead.selectSingleNode("//z:row").Attributes.getNamedItem(sKey) Is Nothing Then
        GetHeadItemValue = domHead.selectSingleNode("//z:row").Attributes.getNamedItem(sKey).nodeValue
    Else
        GetHeadItemValue = ""
    End If
End Function
' by ahzzd 2006/05/09 修改DomHead 里的值
Public Function SetHeadItemValue(ByVal domHead As DOMDocument, ByVal sKey As String, ByVal value As Variant) As Boolean
    sKey = LCase(sKey)
    If Not domHead.selectSingleNode("//z:row").Attributes.getNamedItem(sKey) Is Nothing Then
        domHead.selectSingleNode("//z:row").Attributes.getNamedItem(sKey).nodeValue = value
        SetHeadItemValue = True
    Else
        SetHeadItemValue = False
    End If
End Function
Public Function GetBodyItemValue(ByVal domBody As DOMDocument, ByVal sKey As String, ByVal R As Long) As String
    sKey = LCase(sKey)
    If Not domBody.selectNodes("//z:row").Item(R).Attributes.getNamedItem(sKey) Is Nothing Then
        GetBodyItemValue = domBody.selectNodes("//z:row").Item(R).Attributes.getNamedItem(sKey).nodeValue
    Else
        GetBodyItemValue = ""
    End If
End Function

Public Function SetBodyItemValue(ByVal domBody As DOMDocument, ByVal sKey As String, ByVal R As Long, ByVal value As Variant) As Boolean
    sKey = LCase(sKey)
    If Not domBody.selectNodes("//z:row").Item(R).Attributes.getNamedItem(sKey) Is Nothing Then
        domBody.selectNodes("//z:row").Item(R).Attributes.getNamedItem(sKey).nodeValue = value
        SetBodyItemValue = True
    Else
        SetBodyItemValue = False
    End If
End Function

Private Function FieldInFields(fld As ADODB.Field, Flds As ADODB.Fields) As Boolean
Dim fld_Tmp As Field
FieldInFields = False
For Each fld_Tmp In Flds
    If fld_Tmp.Name = fld.Name Then
        FieldInFields = True
        Exit For
    End If
Next
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
          Case adDate, adDBTime, adDBTimeStamp
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
    MsgBox sErrFrom & vbCrLf & CStr(err.Number) & err.Description, vbOKOnly + vbCritical
End Function

'' 检查指定单据是否可以审核/弃审
'' 参数: strTblName:表名, LngID:单据主表ID; bVer:是否审核; bNewCollection:是否新建连接
'  bVer=true 审核处理         bVer＝false 弃审处理
'' 返回: 如果不能审核/弃审 则返回相关信息
Public Function CanVerify(objsys As clsSystem, strVouchType As String, lngID As String, bVer As Boolean, CN As ADODB.Connection, Optional bSettle As Boolean = False) As String
    Dim rst As New ADODB.Recordset
    Dim strSQL As String, strFldID As String
    Dim strTblName As String
    Dim strVouchName As String
    Dim bFirst As Boolean
    Dim i As Long
    On Error GoTo ErrHandle
    CanVerify = ""
    rst.CursorLocation = adUseClient
    If CN.State <> 1 Then
       CanVerify = "不能连接到数据库[" & objsys.sDBName & "],可能是网络忙或者打开的数据库连接太多,请稍后再试."
       GoTo DOExit
    End If
    Select Case strVouchType
'        Case "97"   '资产卡片日常增加处理
'            strSQL = "select IsNULL(checkcode,'') as checkcode,IsNULL(checkname,'') as checkname,IsNULL(coutno_id,'') as coutno_id,ufts from wjbfa_Cards where id='" & lngID & "'"
'            rst.CursorLocation = adUseClient
'            rst.Open strSQL, CN, 3, 1
'
'            If rst.RecordCount = 0 Then
'                CanVerify = "单据不存在"
'                GoTo DOExit
'            End If
'            If bVer Then
'                If Trim(rst.Fields("checkname")) <> "" Then
'                    CanVerify = "单据已审核，不能审核。"
'                    GoTo DOExit
'                End If
'            Else 'by ahzzd 2006/05/30 弃审检查
'                If Trim(rst.Fields("checkname")) = "" Then
'                   CanVerify = "没有审核,不能弃审。"
'                   GoTo DOExit
'                End If
'                If Trim(rst.Fields("coutno_id")) <> "" Then
'                   CanVerify = "本张单据已生成凭证,不能弃审。"
'                   GoTo DOExit
'                End If
'            End If
    End Select
DOExit:
    If rst.State = 1 Then rst.Close
    Set rst = Nothing
    Exit Function
ErrHandle:
    If CanVerify = "" Then
       CanVerify = "网络异常请稍后造再试!"
    End If
End Function


''VouchID：主表关键字
''bVer：是否审核标志，TRUE审核，FALSE弃审
''dUfts：时间戳
Public Function VerifyVoucher(CN As ADODB.Connection, clsSys As clsSystem, strTblName As String, VouchID As String, bVer As Boolean, dUfts As String, Optional CardNumber As String, Optional oriDomHead As DOMDocument) As String
    Dim bTrans As Boolean
    Dim AffectedLine As Long
    Dim strSQL As String
    Dim ErrMsg As String
    Dim strID As String
    Dim rsttemp As New ADODB.Recordset
    
    
    On Error GoTo ErrHandle
    VerifyVoucher = ""
    If CN.State <> 1 Then
        VerifyVoucher = GetString("U8.SA.USSASERVER.modvoucher.01041") 'zh-CN：不能连接到数据库,可能是网络忙或者当前打开的数据库连接太多，请稍后再试
        Exit Function
    End If
     
    bTrans = False
    If Not clsSys.bManualTrans Then
       CN.BeginTrans  ''每张单据一个事务
       bTrans = True
    End If
    
     ' 单据审核和弃审前的校验------------------------------------------------------------------
    'Modify by Ktao 将删除时的校验也指向验证类UserCheck×××，以方便开发人员编写
    Dim UserCheck As Object
    Select Case UCase(CardNumber)
        Case "EFBWGL020201" '编务管理的所有CardNumber加在这里
 
        Case Else
            Set UserCheck = New clsUsercheck_efyzgl
    End Select
    Set UserCheck.clsSys = clsSys
    UserCheck.strVouchType = CardNumber
    VerifyVoucher = UserCheck.VoucherCheckForVerify(CN, clsSys, strTblName, VouchID, bVer, dUfts, CardNumber, oriDomHead)
    If VerifyVoucher <> "" Then
        If bTrans Then CN.RollbackTrans: bTrans = False
        Exit Function
    End If
     
    strID = GetHeadItemValue(oriDomHead, clsSys.getVouchMainIDName(CardNumber, CN))
     

            strSQL = " Update " & strTblName & " SET cverifier=" & IIf(bVer = True, "N'" & clsSys.CurrentUserName & "'", "NULL") & _
                     ",dverifydate=" & IIf(bVer, "case when ddate>'" & clsSys.objlogin.CurDate & "' then ddate else '" & clsSys.objlogin.CurDate & "' end", "null") & _
                     ",dnverifytime=" & IIf(bVer, "getdate()", "null") & _
                     " WHERE id=" & CLng(Val(VouchID)) & _
                     IIf(dUfts = "", "", " AND Convert(char,Convert(Money,Ufts),2)= N'" & dUfts & "'")
            
            CN.Execute strSQL, AffectedLine
            
            If AffectedLine = 0 Then
                If bTrans Then CN.RollbackTrans: bTrans = False
                VerifyVoucher = GetStringPara("U8.SA.USSASERVER.modvoucher.01054", IIf(bVer, GetString("U8.SA.USSASERVER.modvoucher.01052"), GetString("U8.SA.USSASERVER.modvoucher.01053"))) 'zh-CN：弃审 'zh-CN：审核 'Para zh-CN：该单据已经被其他人改动，或{0}不成功
                Exit Function
            ElseIf AffectedLine = -1 Then
                If bTrans Then CN.RollbackTrans: bTrans = False
                VerifyVoucher = "不成功，请检查数据是否正确!"
                Exit Function
            Else
                If bTrans Then CN.CommitTrans
                VerifyVoucher = ""
            End If
    
  
    Exit Function
ErrHandle:
    If VerifyVoucher = "" Then
       If CN.Errors.Count > 0 Then
            If CN.Errors(0).NativeError = 8153 Then
                VerifyVoucher = CN.Errors(1).Description
            Else
                VerifyVoucher = err.Description
            End If
       End If
       
    End If
    If bTrans Then CN.RollbackTrans: bTrans = False
End Function


Public Function GetVouchID(strTableName As String, clsSys As clsSystem, lngIDs As String, lngsTableCount As Long, ErrMsg As String) As String
    Dim AdoComm As ADODB.Command
    On Error GoTo DoErr
    Set AdoComm = New ADODB.Command
    With AdoComm
        .ActiveConnection = clsSys.dbSales
        .CommandText = "sp_GetID"
        .CommandType = adCmdStoredProc
        .Prepared = False
        .Parameters.Append .CreateParameter("RemoteId", adVarChar, adParamInput, 3, clsSys.objlogin.cacc_id)
        .Parameters.Append .CreateParameter("cAcc_Id", adVarChar, adParamInput, 3, clsSys.CurrentAccID)
        .Parameters.Append .CreateParameter("VouchType", adVarChar, adParamInput, 50, strTableName)
        .Parameters.Append .CreateParameter("iAmount", adInteger, adParamInput, 8, lngsTableCount)
        .Parameters.Append .CreateParameter("MaxID", adBigInt, adParamOutput)
        .Parameters.Append .CreateParameter("MaxIDs", adBigInt, adParamOutput)
        .Execute
        GetVouchID = CStr(.Parameters("MaxID"))
        lngIDs = .Parameters("MaxIDs") - lngsTableCount + 1
    End With
    Set AdoComm = Nothing
    Exit Function
DoErr:
    ErrMsg = "获取单据ID发生错误：" & err.Description
    Set AdoComm = Nothing
End Function


'    @MT_id nvarchar(50),             --当前单据ID
'    @cDepCode nvarchar(50),       --部门编码
'    @cItemCode nvarchar(50),      --项目编码
'    @cExpCode nvarchar(50)='',        --费用类别
'    @ErrStr nvarchar(500) OUTPUT  --出错信息

Public Sub GetMT_sum(conn As ADODB.Connection, MT_id As String, cDepCode As String, cItem_class As String, cItemCode As String, iPeriod As String, ErrStr As String, Optional cExpCode As String)
    Dim AdoComm As ADODB.Command
    On Error GoTo DoErr
    Set AdoComm = New ADODB.Command
    With AdoComm
        .ActiveConnection = conn
        .CommandTimeout = 120
        .CommandText = "MT_sum"
        .CommandType = adCmdStoredProc
        .Prepared = False
        .Parameters.Append .CreateParameter("MT_id", adVarChar, adParamInput, 50, MT_id)
        .Parameters.Append .CreateParameter("cDepCode", adVarChar, adParamInput, 50, cDepCode)
        .Parameters.Append .CreateParameter("cItem_Class", adVarChar, adParamInput, 50, cItem_class)
        .Parameters.Append .CreateParameter("cItemCode", adVarChar, adParamInput, 50, cItemCode)
        .Parameters.Append .CreateParameter("cExpCode", adVarChar, adParamInput, 50, cExpCode)
        .Parameters.Append .CreateParameter("iPeriod", adVarChar, adParamInput, 50, iPeriod)
        .Parameters.Append .CreateParameter("ErrStr", adVarChar, adParamOutput, 500)
        .Execute
        ErrStr = CStr(.Parameters("ErrStr"))
    End With
    Set AdoComm = Nothing
    'LDX 2009-05-17 修改 beg
    conn.Execute "update a set a.citemccode=b.citemccode,a.citemcname=c.citemcname from MT_budget a inner join fitemss" & cItem_class & " b on a.citemccode=b.citemccode inner join fitemss" & cItem_class & "class c on b.citemccode=c.citemccode  and a.citem_class='" & cItem_class & "' and a.id=" & MT_id
'    conn.Execute "update a set a.citemccode=b.citemccode,a.citemcname=c.citemcname from MT_budget a inner join fitemss00 b on a.citemccode=b.citemccode inner join fitemss00class c on b.citemccode=c.citemccode  and a.citem_class='" & cItem_class & "' and a.id=" & MT_id
    'LDX 2009-05-17 修改 end
    Exit Sub
DoErr:
    ErrStr = "发生错误：" & err.Description
    Set AdoComm = Nothing
End Sub


Public Function SupStr(ByVal sStr As String) As String
    Dim i As Integer
    Dim j As Integer
    
    Dim MaxCode As String
    Dim LCode As String
    Dim RCode As String
    
    If Left(sStr, 1) = "-" Then
        sStr = "0000001"
        Exit Function
    End If
    MaxCode = Right("0000000000" & sStr, 8)
    j = 0
    For i = 8 To 1 Step -1
        If Mid(MaxCode, i, 1) < "0" Or Mid(MaxCode, i, 1) > "9" Then
            j = i
            Exit For
        End If
    Next i
    
    LCode = Left(MaxCode, j)
    RCode = Right(MaxCode, 8 - j)
    If RCode <> "" Then
        If Len(CStr(Val(RCode) + 1)) > 8 - j Then
            RCode = Right("00000001", 8 - j)
            LCode = Left(LCode, Len(LCode) - 1) & Chr(Asc(Right(LCode, 1)) + 1)
        Else
            RCode = Right("00000000" & Trim(CStr(Val(RCode) + 1)), 8 - j)
        End If
    Else
        LCode = Left(LCode, Len(LCode) - 1) & Chr(Asc(Right(LCode, 1)) + 1)
    End If
    SupStr = LCode & RCode
End Function


Public Function LockVouch(cnn As ADODB.Connection, VouchType As String, Prop As String, User As String, ComputerName As String, ParamArray VouchID()) As Boolean
   LockVouch = True
End Function
Public Function UnLockVouch(cnn As ADODB.Connection, VouchType As String, ParamArray VouchID()) As Boolean
UnLockVouch = True
End Function

Private Function RetStr(SouStr As String, MaxLen As Long) As String
    Dim i As Integer
    Dim tmpStr As String
    If lstrlen(SouStr) <= MaxLen Then
        RetStr = SouStr
    Else
        tmpStr = SouStr
        Do Until 1 = 0
            tmpStr = Left(tmpStr, Len(tmpStr) - 1)
            If lstrlen(tmpStr) <= MaxLen Then
                RetStr = tmpStr
                Exit Function
            End If
        Loop
    End If
End Function


''通用检查，检查单据项目是否为空
Public Function ChkNull(CN As ADODB.Connection, rstHead As ADODB.Recordset, RstTail As ADODB.Recordset, Optional strErrMsg As String) As Boolean
    'Dim CN As New ADODB.Recordset
    Dim recH As New ADODB.Recordset
    Dim recB As New ADODB.Recordset
    Dim strSQL As String
    Dim fld As ADODB.Field
    On Error GoTo ErrNum
    Dim tmpfieldName As String
    ChkNull = True
    strSQL = "SELECT * FROM voucheritems WHERE vt_id=7 AND CardSection='T' AND IsNull=1 ORDER BY CardItemNum"
    recH.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    recH.MoveFirst
    Do While Not recH.EOF
        tmpfieldName = recH("Fieldname")
        If rstHead(tmpfieldName) = "" Or IsNull(rstHead(tmpfieldName)) Then
            strErrMsg = strErrMsg & Chr(13) & Chr(10) & recH("CardItemName") & "不能为空值"
            ChkNull = False
        End If
        recH.MoveNext
    Loop
    If recH.State = 1 Then
        recH.Close
    End If
    Set recH = Nothing
    strSQL = "SELECT * FROM voucheritems WHERE vt_id=7 AND CardSection='b' AND IsNull=1 ORDER BY CardItemNum"
    recB.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    recB.MoveFirst
    Do While Not recB.EOF
        RstTail.MoveFirst
        Do While Not RstTail.EOF
        tmpfieldName = recB("Fieldname")
            If RstTail(tmpfieldName) = "" Or IsNull(RstTail(tmpfieldName)) Then
                strErrMsg = strErrMsg & Chr(13) & Chr(10) & recB("CardItemName") & "不能为空值"
                ChkNull = False
            End If
            RstTail.MoveNext
        Loop
        recB.MoveNext
    Loop
    If recB.State = 1 Then
        recB.Close
    End If
    Set recB = Nothing
    Exit Function
ErrNum:
    strErrMsg = strErrMsg & Chr(13) & Chr(10) & err.Description
    ChkNull = False
End Function

'================================================================================
'模块: PublicSub   Function: bAttrExist
'--------------------------------------------------------------------------------
'说明: 用于判断结点的属性是否存在
'================================================================================

Public Function GetAttrValue(ByVal eleMent As IXMLDOMElement, ByVal sAttrName As String, Optional strName As String) As String
    On Error GoTo Error_Handler:
    Const PROC_SIG As String = "PublicSub:bAttrExist"
    Dim iAttr As IXMLDOMAttribute
    If Not eleMent.Attributes.getNamedItem(sAttrName) Is Nothing Then
        GetAttrValue = eleMent.getAttribute(sAttrName)
        Exit Function
    End If
    For Each iAttr In eleMent.Attributes
        If UCase(sAttrName) = UCase(iAttr.Name) Then
           GetAttrValue = iAttr.value
           strName = iAttr.Name
           Exit For
        End If
    Next
    Exit Function
Error_Handler:
  Select Case err.Number
    Case Else
        GetAttrValue = ""
  End Select
End Function

'取最大单据号
Public Function GetVouchNO(Connectstr As String, strCardNum As String, domHead As DOMDocument, strVouchNo As String, ErrMsg As String, Optional DomFormat As DOMDocument, Optional bGetFormatOnly As Boolean, Optional bUseSelfFormat As Boolean, Optional bRepeatReDo As Boolean, Optional sRemoteID As String, Optional strSeedName As String, Optional strSeedValue As String, Optional bTrueNO As Boolean = True, Optional bFlaseToTrueNO As Boolean = False) As Boolean
    Dim objBillNo As New UFBillComponent.clsBillComponent
    Dim ele As IXMLDOMElement
    Dim tmpVouchNO As String
    If IsMissing(bTrueNO) Then
        bTrueNO = True
    End If
    If IsMissing(DomFormat) = True Or DomFormat Is Nothing Then
        Set DomFormat = New DOMDocument
    End If
    On Error GoTo DoErr
    ErrMsg = ""
 
    If objBillNo.InitBill(Connectstr, strCardNum) = False Then
        ErrMsg = "初始化单据号码失败!"
        GetVouchNO = False
        Exit Function
    End If
    
    If IsMissing(bUseSelfFormat) = True Or bUseSelfFormat = False Then
        If DomFormat.loadXML(objBillNo.GetBillFormat) = False Then
            ErrMsg = "获得单据前缀格式失败!"
            GetVouchNO = False
            Exit Function
        End If
    End If
    Set ele = DomFormat.selectSingleNode("//单据编号")
    bRepeatReDo = ele.getAttribute("重号自动重取")
    If CBool(ele.getAttribute("允许手工修改")) = True And CBool(ele.getAttribute("重号自动重取")) = False Then
        bRepeatReDo = False
    Else
        bRepeatReDo = True
    End If
    If Len(Mid(ele.getAttribute("流水依据"), InStr(1, "单据类型 远程号", Space(1)) + 1)) > 0 Then '单据类型 仓库'
       Set ele = DomFormat.selectSingleNode("//单据编号/前缀")
       If ele.Attributes.getNamedItem("对象名称").nodeValue = "远程号" Then
          ele.setAttribute "种子", sRemoteID
       End If
       strSeedName = ele.getAttribute("源表字段名")
       strSeedValue = ele.getAttribute("种子")
    End If
    If bGetFormatOnly = True Then
        GoTo DOExit
    End If
    
    If IsMissing(bUseSelfFormat) = True Or bUseSelfFormat = False Then
        For Each ele In DomFormat.selectNodes("//单据编号/前缀[@对象类型!=5]")
            If ele.Attributes.getNamedItem("对象名称").nodeValue <> "远程号" Then
                'ele.setAttribute "种子", GetHeadItemValue(DOMHead, ele.Attributes.getNamedItem("档案表字段名").nodeValue)
                If GetHeadItemValue(domHead, ele.Attributes.getNamedItem("源表字段名").nodeValue) = "" Then
                    ErrMsg = "单据项目(" & ele.Attributes.getNamedItem("对象名称").nodeValue & ")为单据号流水依据，不能为空，请填写!"
                    GetVouchNO = False
                    Exit Function
                End If
                If ele.Attributes.getNamedItem("对象类型").nodeValue <> "2" Then
                    ele.setAttribute "种子", GetHeadItemValue(domHead, ele.Attributes.getNamedItem("源表字段名").nodeValue)
                    
                Else
                    ele.setAttribute "种子", Left(GetHeadItemValue(domHead, ele.Attributes.getNamedItem("源表字段名").nodeValue), 10)
                End If
            Else
                ele.setAttribute "种子", sRemoteID
            End If
        Next
    End If
'    Call objBillNo.InitBill("", "FA01")
'    objBillNo.GetBillFormat
    tmpVouchNO = objBillNo.GetNumber(DomFormat.xml, bTrueNO)
    If Trim(tmpVouchNO) = "" Then
        ErrMsg = "单据控件取单号失败!"
        GetVouchNO = False
        Exit Function
    ElseIf bFlaseToTrueNO Then
        
        If strVouchNo = tmpVouchNO Then
            tmpVouchNO = objBillNo.GetNumber(DomFormat.xml, True)
        End If
        strVouchNo = tmpVouchNO
    Else
        If Trim(strVouchNo) <> "" Then
            While Val(Right(tmpVouchNO, Len(tmpVouchNO) - 3)) < Val(Right(strVouchNo, Len(strVouchNo) - 3))
            tmpVouchNO = objBillNo.GetNumber(DomFormat.xml, True)
'            strVouchNo = tmpVouchNO
            Wend
        Else
            strVouchNo = tmpVouchNO
        End If
    End If
DOExit:
    GetVouchNO = True
    Exit Function
DoErr:
    ErrMsg = "错误位于Function GetVouchNO:" & err.Description & objBillNo.GetLastErrorA
    GetVouchNO = False
End Function

'某种单据是否自动编号
Public Function bAutoVouchCode(CN As ADODB.Connection, strVouchType As String) As Boolean
    Dim strSQL As String
    Dim strCardNum As String
    Dim RecTemp As New ADODB.Recordset
    RecTemp.CursorLocation = adUseClient
    strCardNum = GetstrCardNum(strVouchType)
If RecTemp.State <> 0 Then RecTemp.Close
strSQL = "SELECT bAllowHandWork,bRepeatReDo FROM VoucherNumber WHERE CardNumber='" & strCardNum & "'"
RecTemp.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
If RecTemp.RecordCount > 0 Then
    bAutoVouchCode = IIf(RecTemp!bAllowHandWork = False, True, False)
End If
End Function
''检查权限
Public Function CheckAuth(clsSys As clsSystem, objlogin As U8Login.clsLogin, domHead As DOMDocument, domBody As DOMDocument, ErrMsg As String, CN As ADODB.Connection) As Boolean
    Dim objAuth As U8RowAuthsvr.clsRowAuth
    Dim strAuth As String, strSQL As String
    Dim rstTmp As ADODB.Recordset
    Dim i As Integer
    Dim eleList As IXMLDOMNodeList
    Dim ndrs    As IXMLDOMNode
    Dim nd      As IXMLDOMNode
    
    CheckAuth = True
    Set objAuth = New U8RowAuthsvr.clsRowAuth
    Set rstTmp = New ADODB.Recordset
    If objlogin.isAdmin = True Then GoTo DOExit

    ''部门
    If clsSys.bAuth_Dep And GetHeadItemValue(domHead, "cdepcode") <> "" Then
        objAuth.Init objlogin.UfDbName, objlogin.cUserId, False, "SA"
        strAuth = objAuth.getAuthString("DEPARTMENT", , "W")
        If Not strAuth = "" Then
            If Trim(strAuth) = "1=2" Then
                ErrMsg = ErrMsg + "无部门权限!" & Chr(13) & Chr(10)
                CheckAuth = False
            Else
                rstTmp.Open "select cDepcode from department where cdepcode ='" & GetHeadItemValue(domHead, "cdepcode") & "' and cdepcode in (" & strAuth & ")", CN, adOpenForwardOnly, adLockReadOnly
                If rstTmp.EOF Then
                    ErrMsg = "无部门" & GetHeadItemValue(domHead, "cdepcode") & "权限!" & Chr(13) & Chr(10)
                    CheckAuth = False
                    rstTmp.Close
                    GoTo DOExit
                Else
                    rstTmp.Close
                End If

            End If
        End If
    End If
    
    If clsSys.bAuth_Per And GetHeadItemValue(domHead, "cpersoncode") <> "" Then
        objAuth.Init objlogin.UfDbName, objlogin.cUserId, False, "SA"
        strAuth = objAuth.getAuthString("PERSON", , "W")
        If Not strAuth = "" Then
            If Trim(strAuth) = "1=2" Then
                ErrMsg = ErrMsg + "无业务员权限!" & Chr(13) & Chr(10)
                CheckAuth = False
            Else
                rstTmp.Open "select cpersoncode from person where cpersoncode ='" & GetHeadItemValue(domHead, "cpersoncode") & "' and cpersoncode in (" & strAuth & ")", CN, adOpenForwardOnly, adLockReadOnly
                If rstTmp.EOF Then
                    ErrMsg = "无业务员" & GetHeadItemValue(domHead, "cpersoncode") & "权限!" & Chr(13) & Chr(10)
                    CheckAuth = False
                    GoTo DOExit
                Else
                    rstTmp.Close
                End If

            End If
        End If
    End If
DOExit:
    If rstTmp.State = 1 Then rstTmp.Close
    Set rstTmp = Nothing
    Set nd = Nothing
    Set eleList = Nothing
    Set ndrs = Nothing
End Function

'/////////////////////////////////////////////////////////////////////////////////////
'
'根据自定义类型得到单据的 CardNumber 号
'by 客户化开发中心 2006/03/01
'将860sp的GetstrCardNum函数的参数从新定义（由以前的一个变成三个）
'Public Function GetstrCardNum(strVouchType As String) As String
'Public Function GetstrCardNum(strVouchType As String, Optional bRed As Boolean = False, Optional bGetTrue As Boolean = False) As String
'//////////////////////////////////////////////////////////////////////////////////////
Public Function GetstrCardNum(strVouchType As String, Optional bRed As Boolean = False, Optional bGetTrue As Boolean = False) As String
    '//xzq
    Select Case strVouchType
        Case "98"
            GetstrCardNum = "MT66"
        Case Else
            GetstrCardNum = strVouchType
        
    End Select

End Function

''说明：转换xml中的保留字
Public Function FormatStrForDOM(strXMl As Variant) As String
    If Not IsNull(strXMl) Then
        FormatStrForDOM = Replace(strXMl, "&", "&amp;")
        FormatStrForDOM = Replace(FormatStrForDOM, "<", "&lt;")
        FormatStrForDOM = Replace(FormatStrForDOM, ">", "&gt;")
        FormatStrForDOM = Replace(FormatStrForDOM, """", "&quot;")
        FormatStrForDOM = Replace(FormatStrForDOM, "'", "&apos;")
    End If
    
End Function


'是否唯一编码
Public Function bUniCode(sHeadTableName As String, strFldVouchCode As String, strVouchCode As String, strVouchType As String, DBConn As ADODB.Connection) As Boolean
    Dim strSQL As String
    Dim RecTemp As New ADODB.Recordset
    RecTemp.CursorLocation = adUseClient
    Select Case LCase(sHeadTableName)
         Case "dispatchlist", "salebillvouch"
              strSQL = "SELECT " & strFldVouchCode & " FROM " & sHeadTableName & " WHERE " & strFldVouchCode & "='" & strVouchCode & "'  AND cVouchType='" & strVouchType & "'"
         Case "mt_baseset"
         Select Case strVouchType
           Case "97"
             strSQL = "SELECT " & strFldVouchCode & " FROM " & sHeadTableName & " WHERE " & strFldVouchCode & "='" & strVouchCode & "'  AND cVouchType ='12'"
           Case "96"
             strSQL = "SELECT " & strFldVouchCode & " FROM " & sHeadTableName & " WHERE " & strFldVouchCode & "='" & strVouchCode & "'  AND cVouchType ='11'"
           Case "87"
             strSQL = "SELECT " & strFldVouchCode & " FROM " & sHeadTableName & " WHERE " & strFldVouchCode & "='" & strVouchCode & "'  AND cVouchType ='01'"
           Case "88"
             strSQL = "SELECT " & strFldVouchCode & " FROM " & sHeadTableName & " WHERE " & strFldVouchCode & "='" & strVouchCode & "'  AND cVouchType ='02'"

         End Select
        Case "mt_budget"
          Select Case strVouchType
           Case "91"
            strSQL = "SELECT " & strFldVouchCode & " FROM " & sHeadTableName & " WHERE " & strFldVouchCode & "='" & strVouchCode & "'  AND cVouchType ='01'"
           Case "98"
            strSQL = "SELECT " & strFldVouchCode & " FROM " & sHeadTableName & " WHERE " & strFldVouchCode & "='" & strVouchCode & "'  AND cVouchType ='02' and iperiod='13 月'"
           Case "92"
            strSQL = "SELECT " & strFldVouchCode & " FROM " & sHeadTableName & " WHERE " & strFldVouchCode & "='" & strVouchCode & "'  AND cVouchType ='02' and iperiod<>'13 月'"
           Case "93"
            strSQL = "SELECT " & strFldVouchCode & " FROM " & sHeadTableName & " WHERE " & strFldVouchCode & "='" & strVouchCode & "'  AND cVouchType ='03'"
           Case "94"
            strSQL = "SELECT " & strFldVouchCode & " FROM " & sHeadTableName & " WHERE " & strFldVouchCode & "='" & strVouchCode & "'  AND cVouchType ='04'"
          End Select
         Case Else
              strSQL = "SELECT " & strFldVouchCode & " FROM " & sHeadTableName & " WHERE " & strFldVouchCode & "='" & strVouchCode & "'"
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

Public Function GetNodeAtrVal(IXNOde As IXMLDOMNode, sKey As String) As String
    sKey = LCase(sKey)
    If IXNOde.Attributes.getNamedItem(sKey) Is Nothing Then
        GetNodeAtrVal = ""
    Else
        GetNodeAtrVal = IXNOde.Attributes.getNamedItem(sKey).nodeValue
    End If
End Function

''j 无用，仅为了保留兼容性
Public Function GetElEAtrVal(IXEle As IXMLDOMElement, sKey As String, Optional j As Long) As String
    sKey = LCase(sKey)
    GetElEAtrVal = IIf(IsNull(IXEle.getAttribute(sKey)), "", IXEle.getAttribute(sKey))
End Function


 

''根据单据类型返回 数量金额合计
Public Function SumForEA(strVouchType As String, domHead As DOMDocument, domBody As DOMDocument, dSumQuantity As Double, dSumSum As Double, ErrMsg As String) As Boolean
    Dim strFldMoneyName As String, strFldQuantityName As String
    Dim o_Ele As IXMLDOMElement
    Dim o_Ndlist As IXMLDOMNodeList
    
    On Error GoTo DoErr
    Select Case strVouchType
        Case "97"
            strFldMoneyName = "inatsum"
            strFldQuantityName = "iquantity"
        Case Else
            ErrMsg = "错误的单据类型"
            SumForEA = False
            Exit Function
    End Select
    strFldQuantityName = LCase(strFldQuantityName)
    strFldMoneyName = LCase(strFldMoneyName)
    dSumQuantity = 0
    dSumSum = 0
    Set o_Ndlist = domBody.selectNodes("//z:row[@editprop!='D' and @editprop!='d']")  '不计算删除列
    If Not o_Ndlist Is Nothing Then
        For Each o_Ele In o_Ndlist
            dSumQuantity = dSumQuantity + CDbl(Val(GetElEAtrVal(o_Ele, strFldQuantityName)))
            dSumSum = dSumSum + CDbl(Val(GetElEAtrVal(o_Ele, strFldMoneyName)))
        Next
    End If
    dSumQuantity = Abs(dSumQuantity)
    dSumSum = Abs(dSumSum)
    Set o_Ele = Nothing
    Set o_Ndlist = Nothing
    SumForEA = True
    Exit Function
    
DoErr:
    SumForEA = False
    ErrMsg = err.Description
End Function


 

'函数功能：870 added 判断是否启用工作流
'参数说明：bizObjectID-业务对象标识，也即单据类型
'增加时间：2009-1-31
Public Function getIsWfControl(clsSys As Object, myConn As Connection, bizObjectID As String, ByRef ErrMsg As String) As Boolean
    Dim isWfCtl As Boolean
    Call GetIsWFControlled(myConn, bizObjectID, bizObjectID & ".Submit", clsSys.objlogin.cIYear, clsSys.objlogin.cacc_id, isWfCtl, ErrMsg)
    getIsWfControl = isWfCtl
End Function

Public Function AttrExists(ByVal ele As IXMLDOMElement, ByVal sAttr As String) As Boolean
    AttrExists = False
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
    AttrExists = True
    Exit Function
Err_info:
    AttrExists = False '没有该元素节点
End Function
Public Function FormatNum(NumValue, Dec As Integer) As Variant
    Dim tmpStr As String, tmpFString As String
    If Dec < 0 Then Dec = 0
    tmpFString = "####0" & IIf(Dec = 0, "", ".") & String(Val(Dec), "0")
    FormatNum = (Format(Val(NumValue), tmpFString))
    'FormatNum = tmpStr
End Function

Public Function GetiMassUnit(ByVal cMassUnit As String) As Integer
    Select Case cMassUnit
    Case "年", "1"
        GetiMassUnit = 1
    Case "月", "2"
        GetiMassUnit = 2
    Case "日", "3"
        GetiMassUnit = 3
    Case Else
        GetiMassUnit = 0
    End Select
End Function

Public Function GetiMassUnitName(ByVal cMassUnitCode As String) As String
    Select Case cMassUnitCode
    Case "1", "年"
        GetiMassUnitName = "年"
    Case "2", "月"
        GetiMassUnitName = "月"
    Case "3", "日"
        GetiMassUnitName = "日"
    Case Else
        GetiMassUnitName = ""
    End Select
End Function

''取得rst中的字段值，将null转换为0
Public Function GetRstVal(rst As ADODB.Recordset, FieldName As String, Optional bConverStrForDom As Boolean = True) As Variant
    If IsNull(rst(FieldName)) = True Then
        If rst(FieldName).Type = adChar Or rst(FieldName).Type = adVarChar Or rst(FieldName).Type = adDate _
            Or rst(FieldName).Type = adDBDate Or rst(FieldName).Type = adDBTime Or rst(FieldName).Type = adDBTimeStamp _
            Or rst(FieldName).Type = adVarWChar Or rst(FieldName).Type = adLongVarChar Or rst(FieldName).Type = adLongVarWChar _
            Or rst(FieldName).Type = adWChar Or rst(FieldName).Type = adBSTR Then
            GetRstVal = ""
        Else
            GetRstVal = 0
        End If
    Else
        If rst(FieldName).Type = adChar Or rst(FieldName).Type = adVarChar Or rst(FieldName).Type = adDate _
            Or rst(FieldName).Type = adDBDate Or rst(FieldName).Type = adDBTime Or rst(FieldName).Type = adDBTimeStamp _
            Or rst(FieldName).Type = adVarWChar Or rst(FieldName).Type = adLongVarChar Or rst(FieldName).Type = adLongVarWChar _
            Or rst(FieldName).Type = adWChar Or rst(FieldName).Type = adBSTR Then
            If bConverStrForDom Then
                GetRstVal = FormatStrForDOM(rst(FieldName))
            Else
                GetRstVal = rst(FieldName)
            End If
        ElseIf rst(FieldName).Type = adBoolean Then
            GetRstVal = IIf(rst(FieldName), "1", "0")
        Else
            GetRstVal = rst(FieldName)
        End If
    End If
    
End Function


''自定义项目是否必填
Public Function bNeedDefCheck(CN As ADODB.Connection, Id As String, Optional strDefType As String, Optional bFixLen As Boolean, Optional intLen As Integer) As Boolean
    Dim rstTmp As New ADODB.Recordset
    bNeedDefCheck = False
    rstTmp.Open "select isnull(binput,0),cType,isnull(ilength,0),bfixlength from userdef where cid=N'" & Id & "' and citemname is not null", CN, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTmp.EOF Then
        bNeedDefCheck = IIf(rstTmp(0) = 1 Or rstTmp(0) = True, True, False)
        strDefType = rstTmp(1)
        If rstTmp(1) <> "文本" Then
            bNeedDefCheck = False
        End If
        bFixLen = IIf(rstTmp(3) = "0" Or LCase(rstTmp(3)) = "false", False, True)
        intLen = rstTmp(2)
    Else
        bNeedDefCheck = False
        bFixLen = False
    End If
    rstTmp.Close
    Set rstTmp = Nothing
End Function



Public Function ConvDateForDOM(strDate As String) As Date
    Dim tmpPosition As Integer
    
    On Error Resume Next
    tmpPosition = InStr(strDate, "T")
    If tmpPosition > 0 Then
        ConvDateForDOM = CDate(Left(strDate, tmpPosition - 1))
    Else
        ConvDateForDOM = CDate(strDate)
    End If
    ConvDateForDOM = Format(ConvDateForDOM, "yyyy-mm-dd")
End Function


Public Function GetCardNumber(strVouchType As String, domHead As DOMDocument) As String
    Dim strCardNum As String
    Select Case strVouchType
        
        Case "26"
            strCardNum = "07"
        Case "27"
            strCardNum = "13"
            
       ' Case "05"
       '     strCardNum = "01"
        Case "05"
            If GetHeadItemValue(domHead, "breturnflag") = "0" Or LCase(GetHeadItemValue(domHead, "breturnflag")) = "false" Then
                strCardNum = "01"
            Else
                strCardNum = "03"
            End If
        Case "97"
            strCardNum = "17"
        Case "98"
            strCardNum = "08"
        Case "99"
            strCardNum = "09"
        Case "29"
            strCardNum = "14"
        Case "28"
            strCardNum = "15"
        Case "06"
            If GetHeadItemValue(domHead, "breturnflag") = "0" Or LCase(GetHeadItemValue(domHead, "breturnflag")) = "false" Then
                strCardNum = "05"
            Else
                strCardNum = "06"
            End If
        Case "07"
            If GetHeadItemValue(domHead, "breturnflag") = "0" Or LCase(GetHeadItemValue(domHead, "breturnflag")) = "false" Then
                strCardNum = "02"
            Else
                strCardNum = "04"
            End If
        Case "95"
            strCardNum = "10"
            
        Case "92"
            strCardNum = "11"
            
        Case "16"
            strCardNum = "16"
        Case "00"
            strCardNum = "28"
    End Select
    GetCardNumber = strCardNum
End Function

'函数功能：870 added 获取是否启用工作流
'参数说明：cBizObjectId-业务对象标识，也即单据类型;
'          cBizEventId-业务事件类型，应该是提交审批的业务事件类型标识 (如订单为：17.submit)
'          cAccId-账套标识
'          iYear-年度
'          bWFControlled-是否受流程的控制，1表示受流程控制、0表示不受流程控制
'增加时间：2009-1-31
Public Function GetIsWFControlled(conn As Connection, ByVal cBizObjectId As String, ByVal cBizEventId As String, ByVal iYear As Integer, ByVal cAccId As String, ByRef bWFControlled As Boolean, ByRef ErrMsg As String) As Boolean
    Dim cmd As Command
    
    On Error GoTo ErrHandler
    
    Set cmd = New Command
    cmd.ActiveConnection = conn
    cmd.CommandText = "usp_WF_IsFlowControlled"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("@cBizObjectId", adVarWChar, adParamInput, 40, cBizObjectId)
    cmd.Parameters.Append cmd.CreateParameter("@cBizEventId", adVarWChar, adParamInput, 40, cBizEventId)
    cmd.Parameters.Append cmd.CreateParameter("@iYear", adSmallInt, adParamInput, , iYear)
    cmd.Parameters.Append cmd.CreateParameter("@cAcc_id", adVarWChar, adParamInput, 3, cAccId)
    cmd.Parameters.Append cmd.CreateParameter("@bControlled", adBoolean, adParamOutput)
    
    cmd.Prepared = True
    
    cmd.Execute
    
    bWFControlled = CBool(cmd.Parameters("@bControlled").value)
    
    Set cmd = Nothing
    GetIsWFControlled = True
    Exit Function
    
ErrHandler:
    ErrMsg = VBA.err.Description
    Set cmd = Nothing
    GetIsWFControlled = False
    
End Function


'取得参数值
Public Function getAccinformation(CN As ADODB.Connection, strSysID As String, strName As String, Optional cID As String = "") As String
    Dim rst As New ADODB.Recordset
    Dim strSQL As String
    If cID = "" Then
        strSQL = "Select cValue from accinformation where cSysID='" & strSysID & "' and cName='" & strName & "'"
    Else
        strSQL = "select cvalue from accinformation where cSysid='" & strSysID & "' and cID='" & cID & "' and cname='" & strName & "'"
    End If
    rst.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly, adCmdText
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
