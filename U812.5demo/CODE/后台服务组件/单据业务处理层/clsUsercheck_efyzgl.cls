VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUsercheck_efyzgl"
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

'' 保存前的合法性检查
Public Function VoucherCheckForSave(CN As ADODB.Connection, domHead As DOMDocument, domBody As DOMDocument, _
     ByVal iState As Integer, ErrMsg1 As String, errmsg2 As String, Optional ByVal bNewCollection As Boolean = False) As Boolean
    Dim bCheckBefore As Boolean, bSavePer As Boolean
    Dim rec As New ADODB.Recordset, rec2 As New ADODB.Recordset
    Dim iHeadCount As Long
    Dim i As Long, j As Long, lRow As Long, bChang As Boolean
    Dim strSQL As String, OriSOCode As String
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
    i = 1
    rec.CursorLocation = adUseClient
    Set rec2 = New ADODB.Recordset
    rec2.CursorLocation = adUseClient
    If CN.State = 1 Then
        iHeadCount = domHead.selectNodes("//z:row").length
        OriID = Val(GetHeadItemValue(domHead, "id"))
        UFTS = GetHeadItemValue(domHead, "ufts")
        OriSOCode = GetHeadItemValue(domHead, "ccode")
        'convert(nchar,convert(money,dbo.MT_baseset.ufts),2) as ufts
'        Select Case strVouchType
'        End Select
        bCheckBefore = CheckBeforeSave(domHead, domBody, ErrMsg1, errmsg2, iState, CN)
        VoucherCheckForSave = bCheckBefore
        If Trim(ErrMsg1) = "" Then ErrMsg1 = ErrMsgAll
        
     Else
        ErrMsg1 = "不能访问数据库 " & clsSys.sDBName & ",可能是网络忙或打开的数据库连接太多,请稍后再试"
     End If
     
DOExit:
   On Error Resume Next
   If bApply Then clsSys.SAFreeTask "SA03010101", ErrMsg1
   If rec.State = 1 Then rec.Close
   If rec2.State = 1 Then rec2.Close
   Set rec = Nothing
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
    If rec.State = 1 Then rec.Close
    If rec2.State = 1 Then rec2.Close
    Set rec = Nothing
    Set rec2 = Nothing
    VoucherCheckForSave = False
   
End Function


Private Function CheckBeforeSave(domHead As DOMDocument, domBody As DOMDocument, ErrMsg1 As String, errmsg2 As String, iMode As Integer, CN As ADODB.Connection) As Boolean
    Dim Rs As New ADODB.Recordset
    Dim strSQL As String
    Dim Row As Long
    Dim SUM_rate As Double
    On Error GoTo DoErr
    Rs.CursorLocation = adUseClient
    If CN.State = 1 Then
        Select Case strVouchType
            Case "EFYZGL030201"  ''制版单
                ''校验参照填充字段
                If UCase(getAccinformation(CN, "PB", "bMustPB_checknote", "BWGL")) = "TRUE" Then ''制版单必有发稿记录
                    If GetHeadItemValue(domHead, "presscode") = "" Then
                        ErrMsg1 = "收稿记录号不能为空！"
                        GoTo DOExit
                    End If
'                    If GetHeadItemValue(domHead, "bsccode") = "" Then
'                        ErrMsg1 = "稿源单据号不能为空！"
'                        GoTo DOExit
'                    End If
                End If
            Case "EFYZGL030303"  ''印刷委托书
                ''校验参照填充字段
                If UCase(getAccinformation(CN, "PB", "bMustPB_preconsign", "YZGL")) = "TRUE" Then ''印刷委托书必有付印通知单
                    If GetHeadItemValue(domHead, "pressid") = "" Then
                        ErrMsg1 = "付印通知单号不能为空！"
                        GoTo DOExit
                    End If
'                    If GetHeadItemValue(domHead, "bsccode") = "" Then
'                        ErrMsg1 = "稿源单据号不能为空！"
'                        GoTo DOExit
'                    End If
                End If
            Case "EFYZGL030203" ''纸样（软片）入库单
                For Row = 0 To domBody.selectNodes("//z:row").length - 1
                    
                    If Trim(GetBodyItemValue(domBody, "iquantity", Row)) <> "" Then  '' 数量不能小于0
                        If Trim(GetBodyItemValue(domBody, "iquantity", Row)) < 0 Then
                            ErrMsg1 = "表体第 " & Row + 1 & " 行 数量不能小于0!"
                            GoTo DOExit
                        End If
                    End If
                    
                    If Trim(GetBodyItemValue(domBody, "iunitcost", Row)) <> "" Then  '' 单价不能小于0
                        If Trim(GetBodyItemValue(domBody, "iunitcost", Row)) < 0 Then
                            ErrMsg1 = "表体第 " & Row + 1 & " 行 单价不能小于0!"
                            GoTo DOExit
                        End If
                    End If
                    
                    If Trim(GetBodyItemValue(domBody, "iprice", Row)) <> "" Then  '' 金额不能小于0
                        If Trim(GetBodyItemValue(domBody, "iprice", Row)) < 0 Then
                            ErrMsg1 = "表体第 " & Row + 1 & " 行 金额不能小于0!"
                            GoTo DOExit
                        End If
                    End If
                Next Row
            Case "EFYZGL030205" ''纸样（软片）出库单
                For Row = 0 To domBody.selectNodes("//z:row").length - 1
                    
                    If Trim(GetBodyItemValue(domBody, "iquantity", Row)) <> "" Then  '' 数量不能小于0
                        If Trim(GetBodyItemValue(domBody, "iquantity", Row)) < 0 Then
                            ErrMsg1 = "表体第 " & Row + 1 & " 行 数量不能小于0!"
                            GoTo DOExit
                        End If
                    End If
                    
                    If Trim(GetBodyItemValue(domBody, "iunitcost", Row)) <> "" Then  '' 单价不能小于0
                        If Trim(GetBodyItemValue(domBody, "iunitcost", Row)) < 0 Then
                            ErrMsg1 = "表体第 " & Row + 1 & " 行 单价不能小于0!"
                            GoTo DOExit
                        End If
                    End If
                    
                    If Trim(GetBodyItemValue(domBody, "iprice", Row)) <> "" Then  '' 金额不能小于0
                        If Trim(GetBodyItemValue(domBody, "iprice", Row)) < 0 Then
                            ErrMsg1 = "表体第 " & Row + 1 & " 行 金额不能小于0!"
                            GoTo DOExit
                        End If
                    End If
                Next Row
            Case "EFFYGL040201", "EFFYGL040301" ''费用预估单
                For Row = 0 To domBody.selectNodes("//z:row").length - 1
                    
                    If Trim(GetBodyItemValue(domBody, "inumber", Row)) <> "" Then  '' 数量不能小于0
                        If Trim(GetBodyItemValue(domBody, "inumber", Row)) < 0 Then
                            ErrMsg1 = "表体第 " & Row + 1 & " 行 数量不能小于0!"
                            GoTo DOExit
                        End If
                    End If
                    
                    If Trim(GetBodyItemValue(domBody, "iunitprice", Row)) <> "" Then  '' 单价不能小于0
                        If Trim(GetBodyItemValue(domBody, "iunitprice", Row)) < 0 Then
                            ErrMsg1 = "表体第 " & Row + 1 & " 行 单价不能小于0!"
                            GoTo DOExit
                        End If
                    End If
                    If Val(GetHeadItemValue(domHead, "iperiod")) > 12 Or Val(GetHeadItemValue(domHead, "iperiod")) < 1 Then
                        ErrMsg1 = "期间输入不合法，值不能小于1大于12。"
                    End If
                    If Trim(GetBodyItemValue(domBody, "imoney", Row)) <> "" Then  '' 金额不能小于0
                        If Trim(GetBodyItemValue(domBody, "imoney", Row)) < 0 Then
                            ErrMsg1 = "表体第 " & Row + 1 & " 行 金额不能小于0!"
                            GoTo DOExit
                        End If
                    End If
                Next Row
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
    If Rs.State = 1 Then Rs.Close
    Set Rs = Nothing
    
    Exit Function
DOExit:
    If ErrMsg1 <> "" Then
        CheckBeforeSave = False
    Else
        CheckBeforeSave = True
    End If
    If Rs.State = 1 Then Rs.Close
    Set Rs = Nothing
    Exit Function
DoErr:
    ErrMsg1 = err.Description
    If CN.Errors.Count > 0 Then
        ErrMsg1 = err.Description
    End If
    CheckBeforeSave = False
    If Rs.State = 1 Then Rs.Close
    Set Rs = Nothing
End Function

Private Sub Class_Terminate()
 On Error Resume Next
    ClearCol NewSOID
    Set NewSOID = Nothing
    ClearCol RetMsg
    Set RetMsg = Nothing
End Sub

''校验单据是否被引用过
Public Function IsQuote(CN As ADODB.Connection, StrCcode As String, Optional ByVal cToTable As String = "") As Boolean
    Dim rst As New ADODB.Recordset
    
    Select Case UCase(strVouchType)
            Case "EFYZGL030301"  ''付印通知单
                If cToTable = "" Then
                    rst.Open "Select * from EFYZGL_pressconsign where pressid = '" & StrCcode & "'", CN, adOpenForwardOnly, adLockReadOnly, adCmdText
                ElseIf LCase(cToTable) = "rdrecord" Then
                    rst.Open "select cdefine1 from rdrecord where cdefine1 = '" & StrCcode & "'", CN, adOpenForwardOnly, adLockReadOnly, adCmdText
                End If
'            Case "EFFYGL040201"  ''费用预估单
'                rst.Open "Select * from EFYZGL_pressconsign where pressid = '" & StrCcode & "'", CN, adOpenForwardOnly, adLockReadOnly, adCmdText
    End Select
    
    If rst.EOF Then
        IsQuote = False
    Else
        IsQuote = True
    End If
    rst.Close
    Set rst = Nothing
End Function

'是否已经制单


'删除前的校验
'm_Conn为数据连接，当前校验之前已经启动事务，本过程内容所有数据库操作请使用m_Conn连接，验证成功或失败都不要提交或回滚事务，只赋值错误信息和返回False
Public Function VoucherCheckForDelete(strCardNum As String, domHead As DOMDocument, m_Conn As ADODB.Connection, ErrMsg1 As String) As Boolean
    VoucherCheckForDelete = True
'    Dim strCcode As String
'    strCcode = GetHeadItemValue(domHead, "ccode")
'
'    If IsQuote(m_Conn, strCcode, 1) Then
'        VoucherCheckForDelete = False
'        ErrMsg1 = "删除失败，该报价单已被项目立项引用"
'    End If
End Function

'审核前的校验
'm_Conn为数据连接，当前校验之前已经启动事务，本过程内容所有数据库操作请使用m_Conn连接，验证成功或失败都不要提交或回滚事务，只返回错误信息
''bVer：是否审核标志，TRUE审核，FALSE弃审
''dUfts：时间戳
Public Function VoucherCheckForVerify(m_Conn As ADODB.Connection, clsSys As clsSystem, strTblName As String, VouchID As String, bVer As Boolean, dUfts As String, Optional strVouchType As String, Optional domHead As DOMDocument) As String
    Dim StrCcode As String
    StrCcode = GetHeadItemValue(domHead, "ccode")
    
    If bVer = False Then
        Select Case UCase(strVouchType)
            Case "EFYZGL030301"  ''付印通知单
                If IsQuote(m_Conn, StrCcode) Then
                    VoucherCheckForVerify = "无法弃审，已经被印刷委托书引用。": Exit Function
                End If
                If IsQuote(m_Conn, StrCcode, "rdrecord") Then
                    VoucherCheckForVerify = "无法弃审，已经被库存模块引用。": Exit Function
                End If
            Case "EFFYGL040201", "EFFYGL040301" ''费用预估单
                If Not domHead.selectSingleNode("//z:row").Attributes.getNamedItem("bbuild") Is Nothing Then
                    If CBool(domHead.selectSingleNode("//z:row").Attributes.getNamedItem("bbuild").Text) Then
                        VoucherCheckForVerify = "无法弃审，该单据已经制单。": Exit Function
                    End If
                End If
'                If IsQuote(m_Conn, StrCcode) Then
'                    VoucherCheckForVerify = "无法弃审，已经被费用结算单引用。"
'                End If
            
        End Select
    End If
End Function

