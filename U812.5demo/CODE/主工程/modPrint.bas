Attribute VB_Name = "modPrint"

Option Explicit
Public bPrtCancel As Boolean  '是否打印取消
Public domPrint As DOMDocument
Public bPrtPreview As Boolean
Public domPrintStyle As DOMDocument
Public m_strPrintType As String

Public DLID() As Long



Public Sub Num2Chinese(strSumXX, strSumDX)
    Dim oNum2Chinese As Object
    strSumDX = strSumXX
    Exit Sub
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
            'strSumDX = "零" + strSumDX
            strSumDX = Mid(strSumDX, 2)
        End If
        If Left(strSumDX, Len("角")) = "角" Then
            strSumDX = Mid(strSumDX, 2)
        End If
        If Right(strSumDX, Len("角")) = "角" Then
            'strSumDX = strSumDX + "整"
        End If
    End If
End Sub
 


'单据打印

'考虑共打印几页，当前页码
Public Function VoucherPrn(strVouchtype As String, Voucher As UAPVoucherControl85.ctlVoucher, strCardNum As String, Optional VTID As Long, Optional PrnType As String = "Print", Optional bPm As Boolean = False, Optional lngVouchID As Long, Optional iPrtCount As Long = 1, Optional iRecordCount As Integer = 1, Optional bBatchPrint As Boolean) As Boolean
    Dim striSumDX As String
    Dim striSumX As String
    Dim striSum As String
    Dim striMoney As String
    Dim striTax As String
    Dim striDiscount As String
    Dim striQuantity As String
    Dim striNum As String
    Dim VoucherTD As UAPVoucherControl85.ctlVoucher
    Dim i As Long
    Dim j As Long
    Dim iGrade As Integer
    Dim blnBatch As Boolean
    Dim strTblName As String
    Dim strVouchID As String
'    Dim lngVouchID As Long
    Dim iGradeRuleLen As Integer
    Dim recTemp As New ADODB.Recordset
    Dim recTempPrint As New ADODB.Recordset
    Dim recFixed As New ADODB.Recordset
    Dim domHead As New DOMDocument
    Dim domBody As New DOMDocument
    Dim strSumDX  As Variant '(必须用variant类型)
    Dim iSumHJ As Double
    Dim strsql As String
    bPrtCancel = False
    Set VoucherTD = Voucher
    If iPrtCount < 2 Then
        Set domPrint = New DOMDocument
        Set domPrintStyle = New DOMDocument
    End If
    recTemp.CursorLocation = adUseClient
    VoucherPrn = True
    Select Case strVouchtype
        Case "www"

        Case Else
            strVouchID = "ID"
    End Select
    
    
    On Error Resume Next
    With VoucherTD
        iGradeRuleLen = Left(GetGradeRule("inventoryclass"), 1)

        If lngVouchID = 0 Then
            lngVouchID = CLng(val(.headerText(strVouchID)))
        End If

        strSumDX = ""     '要求先初始化
        If bBatchPrint Then
            strsql = "SELECT ISNULL(SUM(ISNULL(iSum,0)),0) AS iSum FROM " & strTblName & " WITH  (NOLOCK) WHERE " & strVouchID & "=" & lngVouchID
            If recTemp.State <> adStateClosed Then
                recTemp.Close
            End If
            
            recTemp.Open strsql, DBconn, adOpenForwardOnly, adLockReadOnly
            Num2Chinese Format((recTemp!iSum), "#.00"), strSumDX
            iSumHJ = recTemp!iSum
            recTemp.Close
        Else
            Num2Chinese Format((val(.TotalText("iSum"))), "#.00"), strSumDX
            iSumHJ = val(.TotalText("iSum"))
        End If
        .headerText("zdsumdx") = strSumDX
        .headerText("zdsum") = iSumHJ
        Select Case strVouchtype
            Case Else
                Call SetSumItems(VoucherTD)
        End Select
        .headerText("zdsumdx") = strSumDX
        .headerText("zdsum") = iSumHJ
        Call SetPrintOtherItems(VoucherTD, VTID, strCardNum, PrnType, , iPrtCount, iRecordCount, bBatchPrint)
        .VoucherStatus = VSNormalMode
    End With
    Exit Function
End Function

'获得编码规则
Public Function GetGradeRule(strKeyWord As String) As String
    Dim recTemp As ADODB.Recordset
    Set recTemp = New ADODB.Recordset
    On Error GoTo Err_Handle
    recTemp.CursorLocation = adUseClient
    recTemp.Open "SELECT * FROM GradeDef WHERE KEYWORD='" & strKeyWord & "'", DBconn, adOpenForwardOnly, adLockReadOnly
    If recTemp.RecordCount = 0 Then Exit Function
    GetGradeRule = recTemp!CodingRule
    Exit Function
Err_Handle:
    MsgBox "GetGradeRule" & Err.Description
End Function


'设置打印项，折扣横打，元、角、分
'iRecordCount  '分单张数

Public Sub SetPrintOtherItems(Voucher As ctlVoucher, VTID As Long, strCardNum As String, Optional PrnType As String = "Print", Optional bSetupPrintDialog As Boolean = False, Optional iCount As Long = 1, Optional iRecordCount As Integer = 1, Optional bBatchPrint As Boolean, Optional iPrtCount As Long = 1)
    On Error Resume Next
    Dim iRowQuantity As Double
    Dim i As Long
    
    Dim sConnString As String
    Dim oServer As Object
    Set oServer = CreateObject("UFVoucherServer85.clsVoucherTemplate")
    sConnString = DBconn.ConnectionString
    
    With Voucher
            '加载公司logo
            .LogoPictureDataString = oServer.GetCompanyLogoImageData(sConnString)
            Set oServer = Nothing
'            .VoucherStatus = VSeEditMode
            If BillPrnSet <> 11 Then
                For i = 1 To .BodyRows
                    If val(.bodyText(i, "iQuantity")) = 0 Then
                        .bodyText(i, "iQuantity") = " "
                        .bodyText(i, "iUnitPrice") = " "
                        .bodyText(i, "iTaxUnitPrice") = " "
                        .bodyText(i, "iNatUnitPrice") = " "
                    Else

                    End If
                Next i
                Dim clsVoucherCO As New clsVoucherCO
                Set clsVoucherCO.m_Conn = DBconn
                Dim recTempPrint As ADODB.Recordset
                Dim recFixed As ADODB.Recordset
                Dim m_sData As String
                Dim sStyle As String
                Dim PControl As PrintControl
                Dim nRet As Long
                Dim ndRoot  As IXMLDOMNode
                Set recTempPrint = clsVoucherCO.GetVoucherFormat(VTID, strCardNum)
                Set recFixed = clsVoucherCO.GetRecVchFixed(VTID)
                
'                recTempPrint!VT_Header = "(" & iRecordCount & "-" & iCount & ")" & recTempPrint!VT_Header
                If PrnType = "Print" Or PrnType = "BatchPrint" Then
                     If iCount = 1 Then
                        .PreParePrint recTempPrint, recFixed, sStyle, m_sData
                        domPrint.loadXML m_sData
                        domPrintStyle.loadXML sStyle
                     Else
                        Call PreParePrint(Voucher, recTempPrint, recFixed)
                     End If
                     If iRecordCount = iCount Then
                        If .PrintVoucher(recTempPrint, recFixed, False) <> 0 Then
                            bPrtCancel = True
                            Exit Sub
                        End If
                     End If
                     
                ElseIf PrnType = "Preview" Then
                    If iCount = 1 And iPrtCount = 1 Then
                         .PreParePrint recTempPrint, recFixed, sStyle, m_sData
                         domPrint.loadXML m_sData
                         domPrintStyle.loadXML sStyle
                    Else
                        Call PreParePrint(Voucher, recTempPrint, recFixed)
                    End If
                    If iRecordCount = iCount And bBatchPrint = False Then   '非批打
                       .VoucherStatus = VSNormalMode
                       .AutoAggregate
                       .PrintVoucher recTempPrint, recFixed, True
                    End If
                End If
            End If
    
    End With
    Exit Sub
Err_Handle:
    MsgBox "SetPrintOtherItems" & Err.Description
'    Resume
End Sub


'全部汇总打印
Public Function PrintByCollect(Voucher As ctlVoucher, lngVouchID As Long, strVouchtype As String)
   Dim striSumDX As String
   Dim striSumX As String
   Dim striSum As String
   Dim striMoney As String
   Dim striTax As String
   Dim striDiscount As String
   
   Dim strNatMoney As String
   Dim strNatTax  As String
   Dim strNatSum As String
   Dim strNatDisCount As String
                    
   Dim striQuantity As String
   Dim striNum As String
   Dim domHead As New DOMDocument
   
   Dim domBody As New DOMDocument
   Dim striNumB As String
   Dim i As Long
   
'   On Error GoTo Err_Handle
   On Error Resume Next
With Voucher
        Set domHead = .GetHeadDom
        Set domBody = .GetLineDom(1)

        FillVouchTD lngVouchID, Voucher, strVouchtype
        striSumDX = gcSales.FourFive(Abs(val(.TotalText("iSum"))), 2) ' SumtoChiness(gcSales.FourFive(Abs(Val(.TotalText("iSum"))), 2))
        striSumX = Format(CDbl(.TotalText("iSum")), "#.00") 'cXxSumT & Format(CDbl(.TotalText("iSum")), "#.00")
        striSum = Format(CDbl(.TotalText("iSum")), "#.00")  'Format(CDbl(.TotalText("iSum")), cXxSumT & "#.00")
        striMoney = Format(CDbl(.TotalText("iMoney")), "#.00")  'Format(CDbl(.TotalText("iMoney")), cXxSumT & "#.00")
        striTax = Format(CDbl(.TotalText("iTax")), "#.00")  ' Format(CDbl(.TotalText("iTax")), cXxSumT & "#.00")
        striDiscount = Format(CDbl(.TotalText("iDiscount")), "#.00")  ' Format(CDbl(.TotalText("iDiscount")), cXxSumT & "#.00")
        striQuantity = Format(CDbl(.TotalText("iQuantity")), "#.00")
        striNumB = Format(CDbl(.TotalText("iNum")), "#.00")
        strNatMoney = Format(CDbl(.TotalText("iNatMoney")), "#.00")
        strNatTax = Format(CDbl(.TotalText("iNatTax")), "#.00")
        strNatSum = Format(CDbl(.TotalText("iNatSum")), "#.00")
        strNatDisCount = Format(CDbl(.TotalText("iNatDisCount")), "#.00")
         '清空表体
        .setVoucherDataXML domHead, domBody
         For i = 0 To .BodyInfoCount - 1
             .bodyText(1, .ItemState(i + 1, sibody).sFieldName) = ""
         Next
        .bodyText(1, "cInvName") = BillInvName
        .headerText("iSumDX") = striSumDX
        .headerText("iSumX") = striSumX
        .bodyText(1, "iSum") = striSum
        .bodyText(1, "iMoney") = striMoney
        .bodyText(1, "iTax") = striTax
        .bodyText(1, "iNatMoney") = strNatMoney
        .bodyText(1, "iNatTax") = strNatTax
        .bodyText(1, "iNatSum") = strNatSum
        .bodyText(1, "iNatDisCount") = strNatDisCount
        .bodyText(1, "iDiscount") = striDiscount
        .bodyText(1, "iQuantity") = striQuantity
        .bodyText(1, "iNum") = striNumB
         Call SetSumItems(Voucher)
End With
Exit Function
Err_Handle:

MsgBox "PrintByCollect" & Err.Description

End Function


Public Sub FillVouchTD(lngVouchID As Long, Voucher As ctlVoucher, strVouchtype As String, Optional sWhere As String = "")
    Dim recTemp As New ADODB.Recordset
    Dim i As Long, j As Long
    Dim iID As Long, iRows As Long
    Dim strSqlT As String, strsql As String
    Dim domHead As New DOMDocument
    Dim domBody As New DOMDocument
    Dim strBTQName As String
    Dim sOrderBy As String
    
    On Error GoTo Err_Handle
    
    With Voucher
        If lngVouchID = 0 Then Exit Sub
        Select Case strVouchtype
        Case "05", "06", "00"
            strsql = "Select Sales_FHD_W.*,' ' AS editprop  From Sales_FHD_W Where DLID = " & lngVouchID & IIf(sWhere = "", "", " And " & sWhere)
            If strVouchtype = "05" Then
                strBTQName = "Sales_FHD_T"
            Else
                strBTQName = "Sales_DXFH_T"
            End If
            strSqlT = "Select " & strBTQName & ".*  From " & strBTQName & " Where DLID = " & lngVouchID '& IIf(sWhere = "", "", " And " & sWhere)
            sOrderBy = " Order by autoid "
        End Select
        strsql = strsql & IIf(gbInvSort = True, " ORDER BY cInvCode ", sOrderBy)
        
        recTemp.CursorLocation = adUseClient
        If recTemp.State <> adStateClosed Then
            recTemp.Close
        End If
        recTemp.Open strsql, DBconn, adOpenForwardOnly, adLockReadOnly
        If recTemp.RecordCount = 0 Then Exit Sub
        recTemp.Save domBody, adPersistXML
        If recTemp.State <> adStateClosed Then
            recTemp.Close
        End If
        recTemp.Open strSqlT, DBconn, adOpenForwardOnly, adLockReadOnly
        If recTemp.RecordCount = 0 Then Exit Sub
        recTemp.Save domHead, adPersistXML
'        .setVoucherDataXML Domhead, Dombody
'        .ExamineFlowAuditInfo = GetEAStream(strVouchType, Domhead, voucher)
        ResetVoucherData Voucher, domHead, domBody, strVouchtype
        .BodyTotal
        If recTemp.State <> adStateClosed Then
            recTemp.Close
        End If
        Set recTemp = Nothing
    End With
    Exit Sub
Err_Handle:
    MsgBox "FillVouchTD" & Err.Description
End Sub

Private Function ResetVoucherData(Voucher As Object, domHead As DOMDocument, domBody As DOMDocument, strVouchtype As String)
    With Voucher
        .setVoucherDataXML domHead, domBody
        If .headerText("cexch_name") <> "" Then
           .ItemState("iexchrate", siHeader).nNumPoint = clsSAWeb.GetExchRateDec(GetHeadItemValue(domHead, "cexch_name"))
           .headerText("iexchrate") = GetHeadItemValue(domHead, "iexchrate")
        End If
        .ExamineFlowAuditInfo = GetEAStream(Voucher, strVouchtype)
    End With
End Function


Public Sub VouchOutPut(Voucher As ctlVoucher, VTID As Long, strCardNum As String)
Dim recTempPrint As New ADODB.Recordset
Dim recFixed As New ADODB.Recordset

Dim clsVoucherCO As New clsVoucherCO
Set clsVoucherCO.m_Conn = DBconn
Set recTempPrint = clsVoucherCO.GetVoucherFormat(VTID, strCardNum)
Set recFixed = clsVoucherCO.GetRecVchFixed(VTID)
If Voucher.ExportToFile(recTempPrint, recFixed) = False Then
'   MsgBox "输出文件失败"  'shanlan090803
End If
Set recTempPrint = Nothing
Set recFixed = Nothing
End Sub

'增加读或写权限,以前为固定的R权限,现增加参数RorW
Public Function getAuthString(ByVal sSysID As String, ByVal strBusObId As String, conn As ADODB.Connection, Login As Object, strVouchtype As String, Optional RorW As String = "R") As String
    Dim objRowAuthsrv As New U8RowAuthsvr.clsRowAuth
    Dim strTmp As String
        objRowAuthsrv.Init Login.UfDbName, Login.cUserId '"UFSOFT"
        
            If clsSAWeb.bAuth_Inv Then
                strTmp = objRowAuthsrv.getAuthString("fitem", , RorW)
                If strTmp <> "" Then
                    If strTmp = "1=2" Then
                        getAuthString = strTmp
                        Exit Function
                    Else
                     Select Case strVouchtype
                       Case "MT012"
                        getAuthString = IIf(getAuthString = "", "", getAuthString & " ") & "( v_mt_baseset11.citemcode " & "In (" & strTmp & ") or v_mt_baseset11.citemcode is null)"
                       Case "MT013"
                        getAuthString = IIf(getAuthString = "", "", getAuthString & " ") & "( v_mt_baseset12.citemcode " & "In (" & strTmp & ") or v_mt_baseset12.citemcode is null)"
                       Case "MT005"
                        getAuthString = IIf(getAuthString = "", "", getAuthString & " ") & "( v_mt_budget01.citemcode " & "In (" & strTmp & ") or v_mt_budget01.citemcode is null)"
                       Case "MT011"
                        getAuthString = IIf(getAuthString = "", "", getAuthString & " ") & "( v_mt_budget66.citemcode " & "In (" & strTmp & ") or v_mt_budget66.citemcode is null)"
                       Case "MT006"
                        getAuthString = IIf(getAuthString = "", "", getAuthString & " ") & "( v_mt_budget02.citemcode " & "In (" & strTmp & ") or v_mt_budget02.citemcode is null)"
                       Case "MT007"
                        getAuthString = IIf(getAuthString = "", "", getAuthString & " ") & "( v_mt_budget03.citemcode " & "In (" & strTmp & ") or v_mt_budget03.citemcode is null)"
                       Case "MT001"
                        getAuthString = IIf(getAuthString = "", "", getAuthString & " ") & "( v_mt_baseset01.citemcode " & "In (" & strTmp & ") or v_mt_baseset01.citemcode is null)"
                       Case "MT021"
                        getAuthString = IIf(getAuthString = "", "", getAuthString & " ") & "( v_mt_budget66.citemcode " & "In (" & strTmp & ") or v_mt_budget66.citemcode is null)"
                       Case "MT020"
                        getAuthString = IIf(getAuthString = "", "", getAuthString & " ") & "( v_mt_budget02.citemcode " & "In (" & strTmp & ") or v_mt_budget02.citemcode is null)"
                       Case "MT008"
                        getAuthString = IIf(getAuthString = "", "", getAuthString & " ") & "( v_mt_budget04.citemcode " & "In (" & strTmp & ") or v_mt_budget04.citemcode is null)"
                     End Select
                    End If
                End If
            End If
            If clsSAWeb.bAuth_dep Then
                strTmp = objRowAuthsrv.getAuthString("DEPARTMENT", , RorW)
                If strTmp <> "" Then
                    If strTmp = "1=2" Then
                        getAuthString = strTmp
                        Exit Function
                    Else
                        getAuthString = IIf(getAuthString = "", "", getAuthString & " and ") & " (cdepcode " & "In (" & strTmp & ") or cdepcode is null)"
                    End If
                End If
            End If
End Function
Public Function PreParePrint(Voucher As ctlVoucher, rstFormat As ADODB.Recordset, rstFixed As ADODB.Recordset, Optional sData As String)
        Dim domTmp  As New DOMDocument
        Dim ndRoot  As IXMLDOMNode
        Dim nd      As IXMLDOMNode
        Dim strData As String
        
        Dim rs As Object
        Dim rsFix As Object
        Set ndRoot = domPrint.selectSingleNode("//数据")
        Voucher.PreParePrint rstFormat, rstFixed, "", strData
        domTmp.loadXML strData
        Set nd = domTmp.selectSingleNode("//任务")
        ndRoot.appendChild nd


End Function


'设置合计行，并进行格式化
Public Sub SetSumItems(Voucher As ctlVoucher)
Dim bReturnVouch As Boolean
On Error Resume Next
    With Voucher
        bReturnVouch = CBool(IIf(val(.headerText("bReturnFlag")) <> 0, True, False))
        .headerText("iSumDX") = val(.TotalText("iSum")) 'Format(Abs(Val(.TotalText("iSum"))), "#.00") ' SumtoChiness(Format(Abs(Val(.TotalText("iSum"))), "#.00"))
        .headerText("iSumX") = Format(CDbl(.TotalText("iSum")), "#.00") 'cXxSumT & Format(CDbl(.TotalText("iSum")), "#.00")
        '.headerText("iSum") = Format(CDbl(.TotalText("iSum")), "#.00")  'Format(CDbl(.TotalText("iSum")), cXxSumT & "#.00")
        '.headerText("iMoney") = Format(CDbl(.TotalText("iMoney")), "#.00")  'Format(CDbl(.TotalText("iMoney")), cXxSumT & "#.00")
        '.headerText("iTax") = Format(CDbl(.TotalText("iTax")), "#.00")  'Format(CDbl(.TotalText("iTax")), cXxSumT & "#.00")
        .headerText("iDiscount") = Format(CDbl(.TotalText("iDiscount")), "#.00")  'Format(CDbl(.TotalText("iDiscount")), cXxSumT & "#.00")
        
        If BillPrnSet = 3 Or BillPrnSet = 4 Or BillPrnSet = 2 Then
            .TotalText("iQuantity") = " "
            .TotalText("iNum") = " "
        Else
            .TotalText("iQuantity") = Format(CDbl(.TotalText("iQuantity")), "#.00")
            .TotalText("iNum") = Format(CDbl(.TotalText("iNum")), "#.00")
        End If
    End With
    Exit Sub
Err_Handle:
    MsgBox "SetSumItems" & Err.Description
End Sub

'活动文本审批流处理
Public Function GetEAStream(Voucher As Object, strVouchtype As String) As String
    Dim Mid As String
    Dim VoucherType As String
    With Voucher
        Select Case strVouchtype
        Case "26", "27", "28", "29" '发票
            VoucherType = "07": Mid = .headerText("sbvid")
        Case "05", "06", "00" '发货
            VoucherType = "01": Mid = .headerText("dlid")
        Case "97" '订单
            VoucherType = "17": Mid = .headerText("id")
        Case "16" '报价单
            VoucherType = "16": Mid = .headerText("id")
        Case "98" '代垫
            VoucherType = "08": Mid = .headerText("id")
        Case "99" '费用支出
            VoucherType = "09": Mid = .headerText("id")
        Case "07" '结算
            VoucherType = "02": Mid = .headerText("id")
        Case "00"
            VoucherType = "28": Mid = .headerText("dlid")
    '    Case "95", "92" '包装物
    '        vouchertype = "10": Mid = "autoid"
        Case Else
            VoucherType = strVouchtype
            Mid = .headerText("id")
        End Select
    
        If .headerText("iswfcontrolled") = "1" Then 'wxyadded for 工作流
            GetEAStream = "GET" & "," & VoucherType & "," & Mid
        Else
            GetEAStream = ""
        End If
    End With
End Function
'批量打印
Public Sub PrintBatchVouch(strCardNum As String, sTemplateID As Long, strVouchtype As String, Voucher As ctlVoucher)

    Dim strsql As String
    Dim i As Long
    Dim strViewName As String
    Dim strAuth As String
    Dim strPrintType As String
    Dim strwhere As String
    Dim iRecordCount As Integer
    Dim recTempPrint As ADODB.Recordset
    Dim recFixed As ADODB.Recordset
    Dim clsVoucherCO As New clsVoucherCO
    Dim strPrintVtid As String
    
    bPrtPreview = True
        If UBound(DLID) > 0 Then
            strPrintType = m_strPrintType
            For i = 1 To UBound(DLID)
                If DLID(i) > 0 Then
                    iRecordCount = iRecordCount + 1
                End If
            Next
            For i = 1 To 1
                If bPrtCancel Or Not bPrtPreview Then
                    Voucher.VoucherStatus = VSNormalMode

                   Exit Sub
                End If
                If DLID(i) > 0 Then
                    If i = 1 Then
                        VoucherPrn strVouchtype, Voucher, strCardNum, sTemplateID, strPrintType, True, 0, i, 1, True   ', iRecordCount
                        strPrintVtid = sTemplateID
                    Else
                        VoucherPrn strVouchtype, Voucher, strCardNum, CLng(strPrintVtid), strPrintType, False, 0, i, 1, True ', iRecordCount
                    
                    End If
                End If
            Next i
            If bPrtCancel Or Not bPrtPreview Then
                Voucher.VoucherStatus = VSNormalMode

               Exit Sub
            End If
            Set clsVoucherCO.m_Conn = DBconn
            Set recTempPrint = clsVoucherCO.GetVoucherFormat(sTemplateID, strCardNum)
            If Not recTempPrint Is Nothing Then
                Set recFixed = clsVoucherCO.GetRecVchFixed(sTemplateID)
                If strPrintType = "Print" Or strPrintType = "BatchPrint" Then
                    bPrtCancel = False
                    Voucher.VoucherStatus = VSNormalMode

                ElseIf strPrintType = "Preview" Then
                    Voucher.VoucherStatus = VSNormalMode
                    On Error Resume Next
                    Voucher.AutoAggregate
                    Voucher.PrintVoucher recTempPrint, recFixed, True
                End If
            End If
        End If



End Sub


