VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUsercheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
''为了web调用不出现冲突
Public clsSys As New clsSystem
Public strVouchType   As String  'by ahzzd 2006/06/02单据类型
Private RetMsg As New Collection
Private NewSOID As New Collection      ''允许删除的订单编号
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'
' 实现业务逻辑层的基本校验，如保存前的检查,审核、弃审 等'
' by 客户化开发中心 2006/03/01
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'' 保存前的合法性检查
Public Function VoucherCheckForSave(CN As ADODB.Connection, domHead As DOMDocument, domBody As DOMDocument, _
     ByVal iState As Integer, ErrMsg1 As String, errmsg2 As String, Optional ByVal bNewCollection As Boolean = False) As Boolean
    Dim bCheckBefore As Boolean, bSavePer As Boolean
    Dim Rec As New ADODB.Recordset, rec2 As New ADODB.Recordset
    Dim iHeadCount As Long
    Dim I As Long, j As Long, lRow As Long, bChang As Boolean
    Dim strSql As String, OriSOCode As String
    Dim OriID As String 'OriID As Long,
    Dim bApply As Boolean
    Dim ErrMsgPer As String, ErrMsgAll As String
    Dim sFilter As String
    Dim UFTS As String, iAffected As Integer
    Dim strUserAuth As String
    Dim ele As IXMLDOMElement
    Dim NdList As IXMLDOMNodeList
    On Error GoTo ErrSaveSO
    VoucherCheckForSave = False
    I = 1
    Rec.CursorLocation = adUseClient
    Set rec2 = New ADODB.Recordset
    rec2.CursorLocation = adUseClient
    If CN.State = 1 Then
        iHeadCount = domHead.selectNodes("//z:row").length
        OriID = Val(GetHeadItemValue(domHead, "id"))
        OriSOCode = GetHeadItemValue(domHead, "scardnum")
        Select Case strVouchType
            Case "96", "97" '资产期初和日常增加
                If iState = 1 Or iState = 2 Then   '1、2 表示变更或修改      0、表示增加
                    strSql = "select * from wjbfa_Cards where id=" & OriID
                    If Rec.State = 1 Then Rec.Close
                    Rec.Open strSql, CN, 3, 1
                    If Rec.RecordCount <> 1 Then
                        ErrMsg1 = "当前卡片[" & OriSOCode & "]已经被其他操作员修改！"
                        GoTo DOExit
                    Else
                        Set ele = domHead.selectSingleNode("//z:row")
                        ele.setAttribute "dinputdate", Rec("dinputdate")
                    End If

                    If Not clsSys.IsHoldRWAuth("97", str(OriID), ErrMsg1, "W") Then
                        GoTo DOExit
                    End If
                    If Not IsBlank(Rec("checkname")) Then
                        If iState = 1 Then
                            ErrMsg1 = "卡片[" & OriSOCode & "]已经审核,不能修改"
                            Rec.Close
                            Set Rec = Nothing
                            GoTo DOExit
                        End If
                    End If
                End If
        End Select
        bCheckBefore = CheckBeforeSave(domHead, domBody, ErrMsg1, errmsg2, iState, CN)
        VoucherCheckForSave = bCheckBefore
        If Trim(ErrMsg1) = "" Then ErrMsg1 = ErrMsgAll
        
     Else
        ErrMsg1 = "不能访问数据库 " & clsSys.sDBName & ",可能是网络忙或打开的数据库连接太多,请稍后再试"
     End If
     
DOExit:
   On Error Resume Next
   If bApply Then clsSys.SAFreeTask "SA03010101", ErrMsg1
   If Rec.State = 1 Then Rec.Close
   If rec2.State = 1 Then rec2.Close
   Set Rec = Nothing
   Set rec2 = Nothing
   Exit Function
ErrSaveSO:
    On Error Resume Next
    VoucherCheckForSave = False
    If err.Number <> 0 Then
        ErrMsg1 = "Public Function: VoucherCheckForSave " & vbCrLf & " Error(" & err.Number & "):" & err.Description
    Else
        ErrMsg1 = "Public Function: VoucherCheckForSave 发生错误，请重试!"
    End If
    If Rec.State = 1 Then Rec.Close
    If rec2.State = 1 Then rec2.Close
    Set Rec = Nothing
    Set rec2 = Nothing
    VoucherCheckForSave = False
   
End Function


Private Function CheckBeforeSave(domHead As DOMDocument, domBody As DOMDocument, _
    ErrMsg1 As String, errmsg2 As String, iMode As Integer, CN As ADODB.Connection) As Boolean
    Dim iTailCount As Long, j As Long
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim Par As New ADODB.Parameter
    Dim strCurSOCode As String, strCurInvCode As String, strCurCusCode As String, strSql As String
    Dim ErrRet As Variant
    Dim I As Integer
    Dim strErrMsg As String
    Dim ddate As String
    Dim bContinue  As Boolean
    Dim strStockChkMsg As String
    Dim bUpDateCurrentStock As Boolean
    Dim bPerInvType As Boolean, bPerService As Boolean
    Dim ele As IXMLDOMElement
    Dim eleList As IXMLDOMNodeList
    Dim ndrs    As IXMLDOMNode
    Dim nd      As IXMLDOMNode
    Dim iExchRate As Double
    Dim bCal As Boolean
    Dim lnRow As Long
    Dim strEditProp As String
    Dim sOrDate As String
    Dim bNeedReCheck As Boolean
    Dim sEditProp As String
    Dim ErrMsg As String
    Dim bCusChange As Boolean
    Dim sErr As String, iPartId As Long '调用制造结构自由项用
    Dim strOppCode As String ''商机编码
    Dim o_CRM As Object  'UFCRMSRVSALE.clsCheckOppCus
    On Error GoTo DoErr
    CheckBeforeSave = False
    If CN.State = 1 Then
        Select Case strVouchType
            Case "97"  '资产期初 和日常增加
                If Trim(GetHeadItemValue(domHead, "caretakercode")) = "" Or Trim(GetHeadItemValue(domHead, "caretakername")) = "" Then
                    ErrMsg1 = "使馆编码不能为空或使馆名称不能为空！"
                    GoTo DOExit
                End If
                
                If rs.State <> 0 Then
                    rs.Close
                End If
                strSql = "select * from fa_Origins where iType=1 and lchilds=0 and saddid='" & Trim(GetHeadItemValue(domHead, "sorgaddid")) & "'  and sname='" & Trim(GetHeadItemValue(domHead, "sorgaddname")) & "'"
                rs.Open strSql, CN, adOpenStatic, adLockReadOnly
                If rs.RecordCount = 0 Then
                   ErrMsg1 = "增减方式不合法"
                   GoTo DOExit
                End If
                
                If rs.State <> 0 Then
                    rs.Close
                End If
                strSql = " SELECT cPersonCode, cPersonName FROM Person where cPersonCode ='" & Trim(GetHeadItemValue(domHead, "managercode")) & "' and cPersonName='" & Trim(GetHeadItemValue(domHead, "managername")) & "' "
                rs.Open strSql, CN, adOpenStatic, adLockReadOnly
                If rs.RecordCount = 0 Then
                   ErrMsg1 = "管理人不合法"
                   GoTo DOExit
                End If
                
                If Trim(GetHeadItemValue(domHead, "stypenum")) = "" Then '批量导入时根据资产类型选取单据模板
                    ErrMsg1 = "国标分类代码不能为空."
                    GoTo DOExit
                Else
                        If rs.State <> 0 Then rs.Close
                        rs.Open "SELECT isnull(showid,''),isnull(sname,'')  FROM fa_AssetTypes where sid<>1 and  sNum='" & Trim(GetHeadItemValue(domHead, "stypenum")) & "'", CN, adOpenStatic, adLockReadOnly
                        If rs.RecordCount > 0 Then
                        SetHeadItemValue domHead, "ivtid", rs.Fields(0)
                        SetHeadItemValue domHead, "sassetname", rs.Fields(1)
                        If rs.State <> 0 Then rs.Close
                        Else
                            ErrMsg1 = "【" & GetHeadItemValue(domHead, "stypenum") & "】国标分类代码不合法！"
                            GoTo DOExit
                        End If
                End If
                If Trim(GetHeadItemValue(domHead, "direction")) <> "" Then
                    If Trim(GetHeadItemValue(domHead, "direction")) <> "左" Then
                        If Trim(GetHeadItemValue(domHead, "direction")) <> "右" Then
                            ErrMsg1 = "舵向只能是左 或 右"
                            GoTo DOExit
                        End If
                   End If
                End If
                
                If Trim(GetHeadItemValue(domHead, "usestate")) = "" Then
                   ErrMsg1 = "状态不能为空"
                   GoTo DOExit
                End If
                
                If Trim(GetHeadItemValue(domHead, "usestate")) <> "在用" Then
                    If Trim(GetHeadItemValue(domHead, "usestate")) <> "在建" Then
                        ErrMsg1 = "状态只能是在用或在建"
                        GoTo DOExit
                    End If
                End If
                
                If Trim(GetHeadItemValue(domHead, "usestate")) = "在建" Then
                    If Trim(GetHeadItemValue(domHead, "citemcode")) = "" Or Trim(GetHeadItemValue(domHead, "citemcode")) = "" Then
                       ErrMsg1 = "状态是“在建”时项目不能为空"
                       GoTo DOExit
                    Else '项目不能为空
                        If rs.State <> 0 Then
                            rs.Close
                        End If
                        strSql = " SELECT citemcode , citemname FROM V_MT_Items   where bclose=0 and citemcode ='" & Trim(GetHeadItemValue(domHead, "citemcode")) & "' and citemname='" & Trim(GetHeadItemValue(domHead, "citemname")) & "' "
                        rs.Open strSql, CN, adOpenStatic, adLockReadOnly
                        If rs.RecordCount = 0 Then
                           ErrMsg1 = "项目不合法"
                           GoTo DOExit
                        End If
                    End If
                End If
                
                        
                SetHeadItemValue domHead, "scardnum", GetHeadItemValue(domHead, "sassetnum")
                strCurSOCode = GetHeadItemValue(domHead, "scardnum")
                If Not bAutoVouchCode(CN, "97") Then
                    If strCurSOCode = "" Then
                       ErrMsg1 = "卡片编号不能为空."
                       GoTo DOExit
                    End If
                End If
                
                ddate = GetHeadItemValue(domHead, "dinputdate")
                If ddate = "" Then
                   ErrMsg1 = "卡片[" & strCurSOCode & "]:录入日期不能为空."
                   GoTo DOExit
                End If
                
                ddate = GetHeadItemValue(domHead, "purchasedate")
                If ddate = "" Then
                   ErrMsg1 = "卡片[" & strCurSOCode & "]:购置时间不能为空."
                   GoTo DOExit
                End If
                
                If Trim(GetHeadItemValue(domHead, "sassetbrand")) = "" Then
                   ErrMsg1 = "品牌不能为空"
                   GoTo DOExit
                End If
                
        
                If Trim(GetHeadItemValue(domHead, "sstyle")) = "" Then
                   ErrMsg1 = "规格型号不能为空"
                   GoTo DOExit
                End If
        
                If Trim(GetHeadItemValue(domHead, "ssite")) = "" Then
                   ErrMsg1 = "存放地点不能为空"
                   GoTo DOExit
                End If
                
                If Trim(GetHeadItemValue(domHead, "sdeptnum")) = "" Then
                   ErrMsg1 = "部门不能为空"
                   GoTo DOExit
                End If
                
                If rs.State <> 0 Then
                    rs.Close
                End If
                strSql = "SELECT cDepCode , cDepName FROM Department  WHERE (bDepEnd = 1)  and  cDepCode='" & Trim(GetHeadItemValue(domHead, "sdeptnum")) & "' and cDepName='" & Trim(GetHeadItemValue(domHead, "sdeptname")) & "' "
                rs.Open strSql, CN, adOpenStatic, adLockReadOnly
                If rs.RecordCount = 0 Then
                   ErrMsg1 = "部门不合法"
                   GoTo DOExit
                End If
                
                If rs.State <> 0 Then
                    rs.Close
                End If
                strSql = " SELECT cPersonCode, cPersonName FROM Person where cPersonCode ='" & Trim(GetHeadItemValue(domHead, "usercode")) & "' and cPersonName='" & Trim(GetHeadItemValue(domHead, "username")) & "' "
                rs.Open strSql, CN, adOpenStatic, adLockReadOnly
                If rs.RecordCount = 0 Then
                   ErrMsg1 = "使用人不合法"
                   GoTo DOExit
                End If

                If rs.State <> 0 Then
                    rs.Close
                End If
                strSql = "select cpersoncode,cdepcode,* from Person where  cdepcode='" & Trim(GetHeadItemValue(domHead, "sdeptnum")) & "' and cpersoncode='" & Trim(GetHeadItemValue(domHead, "usercode")) & "'"
                rs.Open strSql, CN, adOpenStatic, adLockReadOnly
                If rs.RecordCount = 0 Then
                   ErrMsg1 = "部门和人员不匹配不合法"
                   GoTo DOExit
                End If
              
                If Val(Trim(GetHeadItemValue(domHead, "sassetnums"))) = 0 Then
                   ErrMsg1 = "资产数量不能为空."
                   GoTo DOExit
                End If
                
                If Val(Trim(GetHeadItemValue(domHead, "dblvalue"))) = 0 Then
                   ErrMsg1 = "本币金额不能为空 "
                   GoTo DOExit
                End If
                
                If Val(Trim(GetHeadItemValue(domHead, "usdollar"))) = 0 Then
                   ErrMsg1 = "美元金额不能为空 "
                   GoTo DOExit
                End If
                
                If Val(Trim(GetHeadItemValue(domHead, "usexchangerate"))) = 0 Then
                    SetHeadItemValue domHead, "usexchangerate", Format(Val(Trim(GetHeadItemValue(domHead, "dblvalue"))) / Val(Trim(GetHeadItemValue(domHead, "usdollar"))), "0.#######")
                End If
                
                If Val(Trim(GetHeadItemValue(domHead, "dblfcvalue"))) = 0 Then
                   ErrMsg1 = "驻在国货币金额不能为空 "
                   GoTo DOExit
                End If
                
                If Val(Trim(GetHeadItemValue(domHead, "dblexchangerate"))) = 0 Then
                    SetHeadItemValue domHead, "dblexchangerate", Format(Val(Trim(GetHeadItemValue(domHead, "dblvalue"))) / Val(Trim(GetHeadItemValue(domHead, "dblfcvalue"))), "0.#######")
                End If
        End Select
    Else
        ErrMsg1 = "不能访问数据库 " & clsSys.sDBName & ",请稍后再试."
        GoTo DOExit
    End If
    If ErrMsg1 <> "" Then
        CheckBeforeSave = False
    Else
        CheckBeforeSave = True
    End If
    
    Set ele = Nothing
    Set eleList = Nothing
    Set ndrs = Nothing
    Set nd = Nothing
    Exit Function
DOExit:
    If ErrMsg1 <> "" Then
        CheckBeforeSave = False
    Else
        CheckBeforeSave = True
    End If
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
    Set ele = Nothing
    Set eleList = Nothing
    Set ndrs = Nothing
    Set nd = Nothing
    Exit Function
DoErr:
    ErrMsg1 = err.Description
    If CN.Errors.Count > 0 Then
        ErrMsg1 = err.Description
    End If
    CheckBeforeSave = False
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
    Set ele = Nothing
    Set eleList = Nothing
    Set ndrs = Nothing
    Set nd = Nothing
End Function

Private Sub Class_Terminate()
 On Error Resume Next
    ClearCol NewSOID
    Set NewSOID = Nothing
    ClearCol RetMsg
    Set RetMsg = Nothing
End Sub


           




