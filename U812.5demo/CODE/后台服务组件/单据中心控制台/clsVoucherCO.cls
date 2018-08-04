VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsVoucherCO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Enum ErrMode
    ErrorInformation
    ErrorQuestion
End Enum
 

Private strCardNum As String  '单据类型号
Private strVouchType As String, bReturnFlag As Boolean, sTemplateID As String, sTable As String
Private m_Login As U8Login.clsLogin
Public m_Conn As Connection
Private rstVoucher As Recordset
Private sUseMode As String, bAR As Boolean
Private rstID As Recordset
Private lngIDCount As Long
Event OnError(ByVal sErr As String, ByVal ErrMode As ErrMode, bContinue As Boolean)
Private bSelfclsSys As Boolean  ''是否需要释放
Private strSaleTypeCode As String, strSaleTypeName As String
Private bFirst As Boolean
Private cBCode As String, CBName As String, cBAccount As String
Private bSelfConn As Boolean   ''是否是内部打开的联接
Public bNeedLock As Boolean    ''是否需要进行功能申请
Private bLocked As Boolean     ''是否功能申请
Private strTaskID As String    ''功能任务

'////////////////////////////////以上为需要修改部分09-01-31//////////////////////////////////////////////
Private clsDatacheck As New EFVoucherMo.clsCommCheck   '单据在编辑状态下需要后台较验函数
Private clsVouchLoad As New EFVoucherMo.clsVouchLoad   '存取单据
Public clsSystemWeb As EFVoucherMo.clsSystem           '后台参数类需要


'Enum VoucherType
'    MT01    '0   费用类别与科目对照表
'    MT02    '0   借款科目对照表
'    MT03    '0   费用类别比例设置表
'    MT04    '0   费用分类比例设置表
'    MT05    '0   预算编制期初录入
'    MT06    '0   预算编制单
'    MT07    '0   预算编制调整单
'    MT08    '0   支票借款单
'    MT09    '0   节目制作经费报账单
'    MT11    '0    预算编制模板对照表
'    MT12    '0    预算机构费用项目对照
'    MT66    '0    年度预算编制单
'
'    MT67
'    MT68
'    MT69
'    MT70
'    MT71
'    MT72
'    MT73
'    MT74
'    MT75
'    MT76
'    MT77
'    MT78
'End Enum

Public Property Get HaveBodyGrid() As Boolean
    HaveBodyGrid = clsVouchLoad.bHaveBodyGrid
End Property


Public Sub Init(ByVal CardNumber As String, Optional oLogin As U8Login.clsLogin, Optional oConn As ADODB.Connection = Nothing, Optional strUserMode As String, Optional sysObj As Object)    'EFVoucherMo.clsSystem
    strVouchType = CardNumber
    strCardNum = CardNumber
    clsVouchLoad.strVouchType = CardNumber
    If Not oLogin Is Nothing Then
        Set m_Login = oLogin
        Set clsVouchLoad.m_Login = oLogin
    End If
    
    If Not oConn Is Nothing Then Set m_Conn = oConn
    If m_Conn Is Nothing Then
        Set m_Conn = New Connection
        m_Conn.CommandTimeout = 500
        m_Conn.ConnectionTimeout = 600
        m_Conn.Open m_Login.UfDbName
        bSelfConn = True
    Else
        bSelfConn = False
    End If
    If Not IsMissing(sysObj) And Not sysObj Is Nothing Then
        Set clsSystemWeb = sysObj
        bSelfclsSys = False
    ElseIf LCase(strUserMode) = "bs" Then
        Set clsSystemWeb = New EFVoucherMo.clsSystem   '需修改处09-01-31
        bSelfclsSys = True
    End If
    '判断是否进行功能申请
    
    If LCase(strUserMode) <> "cs" Then
        bNeedLock = True
    End If
    If LCase(strUserMode) = "bs" Then
        sUseMode = LCase(strUserMode)
        bAR = IIf(sUseMode = "ar", True, False)
        clsSystemWeb.bARUse = bAR
        clsSystemWeb.sUseMode = sUseMode
        If bSelfclsSys = True Then
            clsSystemWeb.Init oLogin
        End If
    Else
        If strUserMode <> "" Then
            clsSystemWeb.sUseMode = sUseMode
        End If
    End If
    clsSystemWeb.ReSetCuriMonth '判断当月的该模块的结帐状态
    Set clsVouchLoad.clsSys = clsSystemWeb
'    Set clsVouchDMO.clsSys = clsSystemWeb
    Exit Sub
End Sub


'Public Sub GetVouchInf(m_enmVoucherType As String, _
'        sCardNum As String, sVouchType As String, bRed As Boolean, bVouchFirst As Boolean, strTable As String)
'    strVouchType = sCardNum
'    bReturnFlag = bRed
'    strCardNum = sCardNum
'    sTable = strTable
'    bFirst = bVouchFirst
'End Sub

 

'功能说明：获取指定VouchID单据的表头domHead,表体domBody
Public Function GetVoucherData(domHead As DOMDocument, domBody As DOMDocument, Optional VouchID As Variant, Optional blnAuth As Boolean = True) As String
    Dim strHeadsql As String, strBodysql As String
    Dim tmpId As Variant
    Dim strErrMsg As String
    On Error GoTo err
    clsVouchLoad.strVouchType = strCardNum
    clsVouchLoad.GetVoucherDataWEB domHead, domBody, m_Conn, strCardNum, False, strErrMsg, VouchID, blnAuth
    If strErrMsg <> "" Then GetVoucherData = strErrMsg
    Exit Function
err:
    If err.Number > 0 Then
        GetVoucherData = err.Description
    End If
End Function

'功 能: 保存单据
Public Function Save(ByVal domHead As DOMDocument, ByVal domBody As DOMDocument, ByVal VoucherState As Integer, Optional vNewID As Variant, Optional DomConfig As DOMDocument) As String
    Dim ele As IXMLDOMNodeList
    Dim strError As String
    Dim iSum As Double, iQuantity As Double, iTax As Double, iMoney As Double, iTaxRate As Double
    Dim iUnitPrice As Double, iTaxUnitPrice As Double
    Dim iSum2 As Double, iMoney2 As Double, iQuantity2 As Double
    Dim i As Integer, j As Integer
    Dim ndrs    As IXMLDOMNode
    Dim nd      As IXMLDOMNode
    Dim strOperater As String
    Dim tmpVtid As Long
    Dim eleH As IXMLDOMElement
    Dim sDate As String
    If Val(GetHeadItemValue(domHead, "vt_id")) = 0 And VoucherState <> 2 Then
         
        tmpVtid = clsSystemWeb.GetVTID(m_Conn, strCardNum)
        Set eleH = domHead.selectSingleNode("//z:row")
        eleH.setAttribute "vt_id", CStr(tmpVtid)
    End If
    Set eleH = Nothing

    Select Case VoucherState
        Case 0
            strOperater = "增加"
        Case 1
            strOperater = "修改"
        Case 2
            strOperater = "变更"
    End Select
    clsSystemWeb.ReSetCuriMonth
    If VoucherState = 0 Then
        Set ndrs = domBody.selectSingleNode("//rs:data")
        Set ele = domBody.selectNodes("//z:row[@editprop != 'A']")
        If (Not ndrs Is Nothing) And ele.length <> 0 Then
            For Each nd In ele
                ndrs.removeChild nd
            Next
        End If
    End If
    strError = ""
    If VoucherTask(strOperater, strError) = False Then
        Save = strError
        Exit Function
    End If

     
    Set clsVouchLoad.clsSys = clsSystemWeb
    clsVouchLoad.VouchSave m_Conn, strCardNum, domHead, domBody, VoucherState, strError, vNewID, DomConfig
    If Len(strError) <> 0 Then
        If m_Conn.Errors.Count > 0 Then
            If m_Conn.Errors(0).NativeError = 1205 Or m_Conn.Errors(0).NativeError = 2627 Or m_Conn.Errors(0).NativeError = 3604 Or m_Conn.Errors(0).NativeError = 2601 Then
                strError = GetString("U8.SA.VoucherCO_Sa.clsvoucher_sac.00090") 'zh-CN：操作过程中发生资源共享冲突(可能单据号重复)，请稍后重试！ 'zh-CN：操作过程中发生资源共享冲突(可能单据号重复)，请稍后重试！
            End If
        End If
        If Trim(strError) = "" Then
            strError = GetString("U8.SA.VoucherCO_Sa.clsvoucher_sac.00090") 'zh-CN：操作过程中发生资源共享冲突(可能单据号重复)，请稍后重试！
        End If
    End If
    If strError <> "" Then
        Save = strError
    End If
    If m_Conn.State = 0 Then
        m_Conn.Open
    End If

DoExit:
    Call FreeTask
End Function

''审核、弃审操作
''参数说明：domHead ：表头, domBody：表体, bVerify：true审核    false弃审
Public Function VerifyVouch(domHead As DOMDocument, bVerify As Boolean) As String
    Dim lngId As Variant, UFTS As String, sSOKey As String
    Dim errMsg As New Collection, strMsg As String, i As Integer, bReturnFlag As Boolean
    Dim strOperater As String
    
    clsSystemWeb.ReSetCuriMonth
 
    Select Case strVouchType
        Case Else
            If bVerify Then
                strOperater = "审核"
            Else
                strOperater = "弃审"
            End If
    End Select
    
    If LCase(clsSystemWeb.sUseMode) = "bs" Then
        If VoucherTask(strOperater, strMsg) = False Then
               VerifyVouch = strMsg
               Exit Function
        End If
    End If

    Set clsVouchLoad.clsSys = clsSystemWeb
    Select Case strCardNum
'        Case MT01, MT02, MT03, MT04, MT11, MT12
'            lngId = GetHeadItemValue(domHead, "id")
'            If Not clsVouchLoad.VouchVerify(m_Conn, bVerify, strVouchType, domHead, strMsg) Then
'                VerifyVouch = strMsg
'            End If
'        Case MT05, MT06, MT07, MT08, MT09, MT66
'            lngId = GetHeadItemValue(domHead, "id")
'            If Not clsVouchLoad.VouchVerify(m_Conn, bVerify, strVouchType, domHead, strMsg) Then
'                VerifyVouch = strMsg
'            End If
            
        Case Else
        
            lngId = GetHeadItemValue(domHead, "id")
            If Not clsVouchLoad.VouchVerify(m_Conn, bVerify, strCardNum, domHead, strMsg) Then
                VerifyVouch = strMsg
            End If
        
    End Select
    
    If Len(VerifyVouch) <> 0 Then
        If m_Conn.Errors.Count > 0 Then
            If m_Conn.Errors(0).NativeError = 1205 Then
                VerifyVouch = GetString("U8.SA.VoucherCO_Sa.clsvoucher_sac.00054") 'zh-CN：所操作单据正被他人操作，请重试！
            End If
        End If
        If Trim(VerifyVouch) = "" Then
            VerifyVouch = GetString("U8.SA.VoucherCO_Sa.clsvoucher_sac.00054") 'zh-CN：所操作单据正被他人操作，请重试！
        End If
    End If
    Call FreeTask
End Function


''关闭、打开操作
''参数说明：domHead ：表头， bClose：true关闭    false打开
Public Function CloseVouch(domHead As DOMDocument, bClose As Boolean) As String
    Dim strMsg As String
    If Not clsVouchLoad.VouchClose(m_Conn, bClose, strVouchType, domHead, strMsg) Then
        CloseVouch = strMsg
    End If
    Exit Function
End Function

''关闭、打开操作
''参数说明：domHead ：表头， bClose：true关闭    false打开
Public Function CloseVouchBodyLine(domHead As DOMDocument, bClose As Boolean, autoids As String) As String
    Dim strMsg As String
    If Not clsVouchLoad.VouchBodyLineClose(m_Conn, bClose, strVouchType, domHead, autoids, strMsg) Then
        CloseVouchBodyLine = strMsg
    End If
    Exit Function
End Function


'//////////////////////////////////////////////////////////////////////////////////////////////////////
'功    能：新增单据，返回值或状态通过相关属性访问。
'根据您的业务需要,设置单据在新增状态时的默认值
'by 客户化开发中心  2006/03/01
'/////////////////////////////////////////////////////////////////////////////////////////////////////
Public Function AddNew(domHead As DOMDocument, Optional bCopy As Boolean, Optional domBody As DOMDocument)
    Dim iEleMent As IXMLDOMElement
    Dim iAttr As IXMLDOMAttribute
    Dim clsSystem As New EFVoucherMo.clsSystem
    Dim strXMLHead As String
    Dim rsTemp As New ADODB.Recordset
    Dim strBudgetType As String
    Dim ibudmode As Integer
    rsTemp.CursorLocation = adUseClient
    
    On Error Resume Next
    Set iEleMent = domHead.selectSingleNode("//z:row")
    If iEleMent Is Nothing Then GoTo setBody
    For Each iAttr In iEleMent.Attributes
    
        Select Case LCase(iAttr.Name)  '

  
            Case "ddate", "dinputdate"              '//盘点单制单日期
                iAttr.value = m_Login.CurDate
                '资产卡片

            Case "cmaker"         '//制单人编码
                iAttr.value = m_Login.cUserName
            
            Case "cdate"
                iAttr.value = m_Login.CurDate
 
                
            Case "chandler", "cverifier", "cinvalider", "caccounter", "ccloser"
                iAttr.value = ""
                
            '//新增通用字段的处理
            Case "dmoddate", "dverifydate", "cmodifier", "dmodifysystime", "dverifysystime", "dcreatesystime"
                iAttr.value = ""
                    

                
            Case "iverifystate"    '872added
                iAttr.value = 0
            Case "ireturncount"    '872added
                iAttr.value = ""
            Case "iswfcontrolled"  '872added
                iAttr.value = ""
        End Select
    Next
setBody:
    If bCopy = True Then
        For Each iEleMent In domBody.selectNodes("//z:row")
            iEleMent.setAttribute LCase("iSettleNum"), "0"
            iEleMent.setAttribute LCase("iSettleQuantity"), "0"
        Next
    End If
End Function

Private Function getVoucherCodeName() As String
    Dim KeyCode As String
    Select Case strVouchType
        Case "87"
            KeyCode = "ccode"
        Case Else
            KeyCode = "ccode"
    End Select
    getVoucherCodeName = KeyCode
End Function

'功    能: 表头检查方法
'参数说明: rstHead表头数据集。
'          RstBody表体数据集。
'          sKey当前编辑单元的关键字。
'返 回 值：True检查通过，通过访问相关属性得到相关数据。False为保存失败，并触发错误事件。
Public Function HeadCheck(sKey As String, domHead As DOMDocument, Optional domBody As DOMDocument = Nothing) As String
    Dim iType As Integer
    Dim skeyCode As String
    Dim errMsg As String
    Dim oriDomHead As DOMDocument
    Dim ele As IXMLDOMElement
    Dim eleN As IXMLDOMElement
    Dim bGetCusExch As Boolean
    
    HeadCheck = ""
    Set clsDatacheck.m_Conn = m_Conn
    clsDatacheck.strVouchType = strVouchType
    Select Case LCase(sKey)
        Case "iexchrate", "cbustype", "itaxrate"      '需要计算表体：(税率,汇率)
            If LCase(sKey) = "iexchrate" Then
                If GetHeadItemValue(domHead, "cexch_name") = clsSystemWeb.sCurrencyName Then
                    If Val(GetHeadItemValue(domHead, sKey)) <> 1 Then
                        HeadCheck = GetString("U8.SA.VoucherCO_Sa.clsvoucher_sac.00092") 'zh-CN：本位币的汇率必须是1！
                        Exit Function
                    End If
                End If
            End If
            
        Case "ddate", "dpredatebt", "dpremodatebt"     '单据日期、预发货日期、预完工日期
        
        Case "cdlcode"
        Case "csocode"  ''可能引起客户改变
            Set clsDatacheck.clsSys = clsSystemWeb
            HeadCheck = clsDatacheck.HeadCheck(sKey, domHead, domBody)
        Case "ccode"
        Case "ccuscode", "ccusabbname"
            
            
            If Not domBody Is Nothing Then
                If Not bAR Then bGetCusExch = True
            End If
            
            Select Case strVouchType '单据是否参照了其他信息
                Case "16"
                Case "97"
                    If Not domBody Is Nothing Then
                        For Each ele In domBody.selectNodes("//z:row")
                            If GetEleAtrVal(ele, "ccontractid") <> "" Then
                                bGetCusExch = False
                                Exit For
                            End If
                        Next
                    End If
                Case "05", "06", "00"
                Case "26", "27", "28", "29"
                Case "98", "99"
                Case "07"
            End Select
            If bGetCusExch Then
                Set oriDomHead = New DOMDocument
                oriDomHead.loadXML domHead.xml
                Set ele = domHead.selectSingleNode("//z:row")
                ele.setAttribute "bcheckcusexchname", "1"   '是否校验客户默认币种
            End If
            Set clsDatacheck.clsSys = clsSystemWeb
            
            HeadCheck = clsDatacheck.HeadCheck(sKey, domHead, domBody)
            If bGetCusExch Then
                Set ele = domHead.selectSingleNode("//R[@K='bCountExch']")
                If Not ele Is Nothing Then
                    If ele.getAttribute("V") = "1" Then '重新计算汇率
                        Set ele = domHead.selectSingleNode("//R[@K='iexchrate']")
                        Set eleN = oriDomHead.selectSingleNode("//z:row")
                        eleN.setAttribute "iexchrate", ele.getAttribute("V")
                        Set ele = domHead.selectSingleNode("//R[@K='cexch_name']")
                        eleN.setAttribute "cexch_name", ele.getAttribute("V")
                        Call HeadCheck("iexchrate", oriDomHead, domBody)
                    End If
                End If
            End If
            
            
        Case Else                       '不需要计算表体 (如供应商、部门等)
           
    End Select
End Function

'功    能: 表体检查方法
'参数说明: rstHead表头数据集。
'          RstCurRow当前行表体数据集。
'          sKey当前编辑单元的关键字。
'返 回 值：True检查通过，通过访问相关属性得到相关数据。False为保存失败，并触发错误事件。
Public Function BodyCheck(sKey As String, domBody As DOMDocument, Optional domHead As DOMDocument = Nothing, Optional R As Long) As String
    Dim strMessage As String
    Dim strComunitCode As String
    Dim iEleMent As IXMLDOMElement
    Dim iQuantity As Double
    Dim InvenExchRate As Double
    Dim bReturnFlag As Boolean
    Dim KL2 As Double
    Dim kl2Old As Double
    Dim strUnitName As String, strUnitCode As String
    Dim strXMLBody As String
    Dim igrouptype As Integer
    BodyCheck = ""
    If domBody.selectNodes("//z:row").length = 0 Then
        Exit Function
    End If
    Set iEleMent = domBody.selectSingleNode("//z:row")
    sKey = LCase(sKey)
    Select Case LCase(sKey)
        Case "citemcode", "citem_class"
        Case Else

    End Select
    
End Function


'//////////////////////////////////////以下部分一般不需要修改09-01-31////////////////////////////////////////////////////////

Private Function ChangeVouchNOForWeb(domHead As DOMDocument, errMsg As String) As Boolean
    Dim ele As IXMLDOMElement
    Dim DomFormat As New DOMDocument
    Dim strVouchNo As String
    If clsSystemWeb.GetVoucherNO(strVouchType, domHead, strVouchNo, errMsg, DomFormat, True, False) = False Then
        ChangeVouchNOForWeb = False
        Exit Function
    End If
    If LCase(DomFormat.selectSingleNode("//单据编号").Attributes.getNamedItem("允许手工修改").nodeValue) = "true" Then
        ChangeVouchNOForWeb = True
        Exit Function
    End If
    If clsSystemWeb.GetVoucherNO(strVouchType, domHead, strVouchNo, errMsg, DomFormat, False, False) = False Then
        ChangeVouchNOForWeb = False
        Exit Function
    Else
        Set ele = domHead.selectSingleNode("//z:row")
        ele.setAttribute LCase(getVoucherCodeName()), strVouchNo
        ChangeVouchNOForWeb = True
        Exit Function
    End If
End Function

Private Function GetNodeAtrVal(IXNOde As IXMLDOMNode, sKey As String) As String
    sKey = LCase(sKey)
    If IXNOde.Attributes.getNamedItem(sKey) Is Nothing Then
        GetNodeAtrVal = ""
    Else
        GetNodeAtrVal = IXNOde.Attributes.getNamedItem(sKey).nodeValue
    End If
End Function

Private Function GetEleAtrVal(ele As IXMLDOMElement, sKey As String) As String
    sKey = LCase(sKey)
    If IsNull(ele.getAttribute(sKey)) Then
        GetEleAtrVal = ""
    Else
        GetEleAtrVal = ele.getAttribute(sKey)
    End If
End Function

'功    能： 获取单据格式记录集
'参数说明：strCardNum-单据模版号 sTemplateID-单据模板VTID
Public Function GetVoucherFormat(ByVal sTemplateID As String, ByVal strCardNum As String) As Recordset
    Dim TmpUFTemplate As UFVoucherServer85.clsVoucherTemplate
    Set TmpUFTemplate = New UFVoucherServer85.clsVoucherTemplate
    Set GetVoucherFormat = TmpUFTemplate.GetTemplateData2(m_Conn, strCardNum, sTemplateID)
    Exit Function
err_log:
    MsgBox CStr(err.Number) + " " + err.Description
End Function

'获得固定标签记录集
Public Function GetRecVchFixed(ByVal sTemplateID As String) As ADODB.Recordset
    Dim TmpUFTemplate As UFVoucherServer85.clsVoucherTemplate
    Set TmpUFTemplate = New UFVoucherServer85.clsVoucherTemplate
    Set GetRecVchFixed = TmpUFTemplate.GetFixedData(m_Conn, sTemplateID)
    Exit Function
err_log:
    MsgBox CStr(err.Number) + " " + err.Description

End Function

'功 能: 移到第一页
'参数说明: rstHead表头数据集。
'          RstBody表体数据集。
Public Function MoveFirst(domHead As DOMDocument, domBody As DOMDocument)
    MoveFirst = clsVouchLoad.MoveTo("first", m_Conn, domHead, domBody)
End Function

'功 能: 移到上一页
'参数说明: rstHead表头数据集。
'          RstBody表体数据集。
Public Function MovePrevious(domHead As DOMDocument, domBody As DOMDocument)
    MovePrevious = clsVouchLoad.MoveTo("Previous", m_Conn, domHead, domBody)
End Function

'功 能: 移到下一页
'参数说明: rstHead表头数据集。
'          RstBody表体数据集。
Public Function MoveNext(domHead As DOMDocument, domBody As DOMDocument)
    MoveNext = clsVouchLoad.MoveTo("Next", m_Conn, domHead, domBody)
End Function

'功 能: 移到最后一页
'参数说明: rstHead表头数据集。
'          RstBody表体数据集。
Public Function MoveLast(domHead As DOMDocument, domBody As DOMDocument)
    MoveLast = clsVouchLoad.MoveTo("Last", m_Conn, domHead, domBody)
End Function

'功    能: 移动到指定位置。
'参数说明: rstHead表头数据集。
'          RstBody表体数据集。
Public Function MoveTo(ByVal varID As Variant, ByRef rsthead As Recordset, ByRef rstbody As Recordset)

End Function

''取得单据号(假号码)
''bGetFormatOnly true:只取得单据号码前缀格式串
Public Function GetVoucherNO(domHead As DOMDocument, strVouchNo As String, errMsg As String, Optional DomFormat As DOMDocument, Optional bGetFormatOnly As Boolean, Optional bSelfFormat As Boolean) As Boolean
    GetVoucherNO = clsSystemWeb.GetVoucherNO(strCardNum, domHead, strVouchNo, errMsg, DomFormat, bGetFormatOnly, bSelfFormat)
End Function

''取得单据号(真号码)
''bGetFormatOnly true:只取得单据号码前缀格式串
Public Function GetVoucherTrueNO(domHead As DOMDocument, strVouchNo As String, errMsg As String, Optional DomFormat As DOMDocument, Optional bGetFormatOnly As Boolean, Optional bSelfFormat As Boolean) As Boolean
    GetVoucherTrueNO = clsSystemWeb.GetVoucherTrueNO(strVouchType, domHead, strVouchNo, errMsg, DomFormat, bGetFormatOnly, bSelfFormat)
End Function

'功 能: 修改单据
'返 回 值：True为可修改，False为不可修改，并触发错误事件。
Public Function Edit() As Boolean

End Function

'功能:删除单据
Public Function Delete(domHead As DOMDocument) As String
    Dim strError As String
    Dim sBusType As String
    Dim sDate As String
    
    Dim clsVouchDel As New EFVoucherMo.ClsVouchSave
    Set clsVouchDel.clsSys = clsSystemWeb
    Set clsVouchDel.m_Login = m_Login
    
    strError = ""
    If clsSystemWeb.CheckDJMBAuth(GetHeadItemValue(domHead, "vt_id"), "W") = False Then
        Delete = GetString("U8.SA.VoucherCO_Sa.clsvoucher_sac.00051") 'zh-CN：没有单据模版使用权限,不能删除此单据!
        Exit Function
    End If
    
    '//CS程序权限校验放在View层，BS程序权限校验放在CO层。
    If LCase(clsSystemWeb.sUseMode) = "bs" Then
        If VoucherTask("删除", strError) = False Then
            Delete = strError
        End If
    End If
    
    strError = clsVouchDel.DeleteVouch(strCardNum, m_Conn, sTable, domHead)
    If strError <> "" Then
        Delete = strError
    End If
    If Len(Delete) <> 0 Then
        If m_Conn.Errors.Count > 0 Then
            If m_Conn.Errors(0).NativeError = 1205 Then
                Delete = GetString("U8.SA.VoucherCO_Sa.clsvoucher_sac.00054") 'zh-CN：所操作单据正被他人操作，请重试！
            End If
        End If
        If Trim(Delete) = "" Then
            Delete = GetString("U8.SA.VoucherCO_Sa.clsvoucher_sac.00054") 'zh-CN：所操作单据正被他人操作，请重试！
        End If
    End If
    Set clsVouchDel = Nothing
    
    '//CS程序功能释放放在View层，BS程序功能释放放在CO层。
    If LCase(clsSystemWeb.sUseMode) = "bs" Then
        Call FreeTask
    End If
    
End Function

''获得字符串长度
Private Function GetStrTrueLenth(sUnicode As String) As Integer
    Dim strAnsi As String
    strAnsi = StrConv(sUnicode, vbFromUnicode)
    GetStrTrueLenth = LenB(strAnsi)
End Function

Property Get BOF() As Boolean
    BOF = clsVouchLoad.BOF
End Property

'功 能: 返回当前单据是否在最后一张。
Property Get EOF() As Boolean
    EOF = clsVouchLoad.EOF
End Property

Private Function GetHeadItemValue(ByVal domHead As DOMDocument, ByVal sKey As String) As String
    sKey = LCase(sKey)
    If Not domHead.selectSingleNode("//z:row").Attributes.getNamedItem(sKey) Is Nothing Then
        GetHeadItemValue = domHead.selectSingleNode("//z:row").Attributes.getNamedItem(sKey).nodeValue
    Else
        GetHeadItemValue = ""
    End If
End Function

Private Function GetBodyItemValue(ByVal domBody As DOMDocument, ByVal sKey As String, ByVal R As Long) As String
    sKey = LCase(sKey)
    If Not domBody.selectNodes("//z:row").Item(R).Attributes.getNamedItem(sKey) Is Nothing Then
        GetBodyItemValue = domBody.selectNodes("//z:row").Item(R).Attributes.getNamedItem(sKey).nodeValue
    Else
        GetBodyItemValue = ""
    End If
End Function

Private Sub Class_Terminate()
    On Error Resume Next
    Set clsVouchLoad = Nothing   '存取单据
    If bSelfConn = True Then
        If m_Conn.State = 1 Then m_Conn.Close
        Set m_Conn = Nothing
    End If
    If bSelfclsSys = True Then
        Set clsSystemWeb = Nothing
    End If
End Sub

Public Function InitSys(obj As U8Login.clsLogin, Optional ByVal enmVoucherType As String, Optional sysObj As Object) As Boolean
    If Not IsMissing(sysObj) Then
        Set clsSystemWeb = sysObj
    Else
        clsSystemWeb.Init obj
    End If
    Set m_Login = obj
 
    clsSystemWeb.ReSetCuriMonth
    bNeedLock = False
End Function

Public Function GetVoucherTaskID(sKey As String, sVouchType As String, Optional bRed As Boolean) As String
    Dim strID As String
    If IsMissing(bRed) Then bRed = bReturnFlag
    strID = clsSystemWeb.GetVoucherTaskID(sKey, sVouchType, bRed)
    GetVoucherTaskID = strID
End Function

''功能申请
Private Function VoucherTask(sKey As String, errMsg As String) As Boolean
    Dim strID As String
    
    If bNeedLock = False Then
        bLocked = False
        VoucherTask = True
        Exit Function
    End If
    strID = GetVoucherTaskID(sKey, strVouchType, bReturnFlag)
    If strID <> "" Then
        strTaskID = strID
        VoucherTask = clsSystemWeb.SAApplyTask(strTaskID, errMsg)
        If VoucherTask = False Then
            If Trim(errMsg) = "" Then
                errMsg = "功能申请失败,错误信息无!"
            End If
        End If
        bLocked = VoucherTask
            
    Else
        VoucherTask = True
        bLocked = False
    End If
End Function

''释放任务
Private Function FreeTask(Optional errMsg As String) As Boolean
    If bLocked = True Then
        FreeTask = clsSystemWeb.SAFreeTask(strTaskID, errMsg)
    End If
    bLocked = False
End Function

''格式化数字
Private Function FormatNum(NumValue, Dec As Integer) As Variant
    Dim tmpStr As String, tmpFString As String
    
    If Dec < 0 Then Dec = 0
    tmpFString = "####0" & IIf(Dec = 0, "", ".") & String(Val(Dec), "0")
    FormatNum = (Format(Val(NumValue), tmpFString))
End Function

Public Function CopyRow(domBody As DOMDocument) As Boolean
    Dim iEleMent As IXMLDOMElement
    Dim iAttr As IXMLDOMAttribute
    Dim clsSystem As New EFVoucherMo.clsSystem
    Dim strXMLHead As String
    CopyRow = True
    On Error Resume Next
    For Each iEleMent In domBody.selectNodes("//z:row")
        iEleMent.setAttribute "editprop", "A"
    Next
End Function

' 修改DomHead 里的值
Public Function SetHeadItemValue(ByVal domHead As DOMDocument, ByVal sKey As String, ByVal value As Variant) As Boolean
    sKey = LCase(sKey)
    If Not domHead.selectSingleNode("//z:row").Attributes.getNamedItem(sKey) Is Nothing Then
        domHead.selectSingleNode("//z:row").Attributes.getNamedItem(sKey).nodeValue = value
        SetHeadItemValue = True
    Else
        SetHeadItemValue = False
    End If
End Function


'取得参数值
Public Function getAccinformation(strSysID As String, strName As String, Optional cID As String = "") As String
    Dim rst As New ADODB.Recordset
    Dim strSql As String
    If cID = "" Then
        strSql = "Select cValue from accinformation where cSysID='" & strSysID & "' and cName='" & strName & "'"
    Else
        strSql = "select cvalue from accinformation where cSysid='" & strSysID & "' and cID='" & cID & "' and cname='" & strName & "'"
    End If
    rst.Open strSql, m_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
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



