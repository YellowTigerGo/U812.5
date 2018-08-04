VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAutoVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public domVouchers As New DOMDocument
Public domFieldConfig As New DOMDocument
Public domCellCheck As New DOMDocument
Public domRowColChange As New DOMDocument
Public domCheckBeforeSave As New DOMDocument
Public strCardNums As String

Public Sub SubInit(strCardNum As String)
    Dim Rst As New ADODB.Recordset
    Dim strsql As String
On Error Resume Next
    strCardNums = strCardNum
    Rst.CursorLocation = adUseClient
    Rst.Open "select * from vouchers_base where cardnumber='" + strCardNum + "'", DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Rst.Save domVouchers, adPersistXML
    Rst.Close
    Rst.Open "select * from sa_voucherfieldconfig where cardnumber=N'" + strCardNum + "' order by cardsection", DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Rst.Save domFieldConfig, adPersistXML
    Rst.Close
    Rst.Open "select * from sa_vouchercellcheck where cardnum=N'" + strCardNum + "' order by fldcardsection,fieldname,checkindex", DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Rst.Save domCellCheck, adPersistXML
    Rst.Close
    Rst.Open "select *  from SA_RowColChange where cardnumber=N'" + strCardNum + "' order by cardsection,fieldname ", DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Rst.Save domRowColChange, adPersistXML
    Rst.Close
    Rst.Open "select *  from SA_CheckBeforeSave where cardnumber=N'" + strCardNum + "'order by cardsection desc ", DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Rst.Save domCheckBeforeSave, adPersistXML
    Rst.Close
    Set Rst = Nothing
    
End Sub


 
Public Function CopyLine(Voucher As ctlVoucher, clsRefer As clsAutoRefer) As Boolean
    Dim Dombody As New DOMDocument
    Dim lst As IXMLDOMNodeList
    Dim nodData As IXMLDOMElement
    Dim nod As IXMLDOMNode
    Dim blnOk As Boolean
    
    Set Dombody = Voucher.GetLineDom
    blnOk = Voucher.AddLine()
    If blnOk Then
        Set lst = domFieldConfig.selectNodes("//z:row[@cardsection='B' and @copyclear='True']")
        Set nodData = Dombody.selectSingleNode("//z:row")
        For Each nod In lst
            nodData.setAttribute nod.Attributes.getNamedItem("fieldname").Text, ""
        Next
        Voucher.UpdateLineData Dombody, Voucher.BodyRows
        Set lst = domFieldConfig.selectNodes("//z:row[@cardsection='B' and @copyclear='True' and @addnewdefaultvalue!='']")
        For Each nod In lst
            SetFldDefaultValue Voucher, nod, Voucher.BodyRows, clsRefer
        Next
    End If
    Set Dombody = Nothing
End Function

Private Sub SetFldDefaultValue(Voucher As ctlVoucher, nod As IXMLDOMElement, lngRow As Long, clsRefer As clsAutoRefer)
    Dim varPar As Variant
    Dim Rst As New ADODB.Recordset
    Dim strValue As String
    Dim i As Long
    
    Rst.CursorLocation = adUseClient
'    MsgBox nod.Attributes.getNamedItem("addnewdefaultvalue").Text
    varPar = Split(nod.Attributes.getNamedItem("addnewdefaultvalue").Text, ":")
    Select Case varPar(0)
        Case "0"
            strValue = ReplaceVoucherItems(CStr(varPar(1)), Voucher, lngRow)
            SetVoucherItemValue Voucher, nod.Attributes.getNamedItem("cardsection").Text, nod.Attributes.getNamedItem("fieldname").Text, strValue, lngRow
        Case "1"        '全局变量
            SetVoucherItemValue Voucher, nod.Attributes.getNamedItem("cardsection").Text, nod.Attributes.getNamedItem("fieldname").Text, GetGlobalVariant(CStr(varPar(1))), lngRow
        Case "2"        'sql
            Rst.Open varPar(1), DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not Rst.EOF Then
                For i = 0 To Rst.Fields.Count - 1
                    SetVoucherItemValue Voucher, nod.Attributes.getNamedItem("cardsection").Text, Rst(i).Name, Rst(i).value, lngRow
                Next
                clsRefer.ReferCellCheck Voucher, nod.Attributes.getNamedItem("cardsection").Text, nod.Attributes.getNamedItem("fieldname").Text, lngRow
            End If
            Rst.Close
    End Select
    Set Rst = Nothing
End Sub

Private Sub SetVoucherItemValue(Voucher As ctlVoucher, strSection As String, strFieldName As String, strValue As String, Optional lngRow As Long)
    Dim intNumPoint As Integer
    If strSection = "B" Then
        If Voucher.ItemState(strFieldName, sibody).nFieldType = 4 Then
            intNumPoint = Voucher.ItemState(strFieldName, sibody).nNumPoint
            strValue = Format(strValue, IIf(intNumPoint = 0, "###0", "###0." & String(intNumPoint, "0")))
        End If
        If Voucher.ItemState(strFieldName, sibody).nReferType = 3 Then
            strValue = Format(strValue, "YYYY-MM-DD")
        End If
        Voucher.bodyText(lngRow, strFieldName) = strValue
    End If
    If strSection = "T" Then
        If Voucher.ItemState(strFieldName, siheader).nFieldType = 4 Then
            intNumPoint = Voucher.ItemState(strFieldName, siheader).nNumPoint
            strValue = Format(strValue, IIf(intNumPoint = 0, "###0", "###0." & String(intNumPoint, "0")))
        End If
        If Voucher.ItemState(strFieldName, siheader).nReferType = 3 Then
            strValue = Format(strValue, "YYYY-MM-DD")
        End If
        Voucher.headerText(strFieldName) = strValue
    End If
End Sub

Public Function GetGlobalVariant(strName As String) As String
    Select Case LCase(strName)
        Case "@username"
            GetGlobalVariant = m_Login.cUserName
        Case "@curdate"
            GetGlobalVariant = m_Login.CurDate
'        Case "@ifloatraterule"
'            GetGlobalVariant = clsSAWeb.iFloatRateRule
'        Case "@bsaleprice"
'            GetGlobalVariant = IIf(clsSAWeb.bSalePrice, 1, 0)
'        Case "@bmostart"
'            GetGlobalVariant = IIf(clsSAWeb.bMOStart, 1, 0)
'        Case "@bbostart"
'            GetGlobalVariant = IIf(IsDate(clsSAWeb.GetSysDicOption("BO", "dBOFirstDate")), 1, 0)
'        Case "@bmpstart"
'            GetGlobalVariant = IIf(IsDate(clsSAWeb.GetSysDicOption("MP", "dMPFirstDate")), 1, 0)
'        Case "@bmqstart"
'            GetGlobalVariant = IIf(IsDate(clsSAWeb.GetSysDicOption("MQ", "dMQFirstDate")), 1, 0)
'        Case "@sastartdate"
'            GetGlobalVariant = clsSAWeb.SAStDate
'        Case "@bcusinvlimited"
'            GetGlobalVariant = IIf(clsSAWeb.bCusInvLimited, 1, 0)
        Case Else
            GetGlobalVariant = strName
    End Select
End Function

Public Function CellCheck(Voucher As ctlVoucher, strCardSection As String, strFieldName As String, bChange As Long, clsRefer As clsAutoRefer, Optional lngRow As Long, Optional blnShowMsg As Boolean = True)
    Dim lst As IXMLDOMNodeList
    Dim nod As IXMLDOMNode
    Dim strCellCheck As String
    Dim clsComp As New UsSaCompStr.clsCompStr
    Dim strValue As String
    Dim strError As String
    Dim i As Long
    
    Set lst = domCellCheck.selectNodes("//z:row[@fldcardsection='" + strCardSection + "' and @fieldname='" + strFieldName + "']")
    
    For Each nod In lst
'        If Not nod.Attributes.getNamedItem("checkcondition") Is Nothing Then
'            strCellCheck = nod.Attributes.getNamedItem("checkcondition").nodeValue
'            strCellCheck = ReplaceSysPara(strCellCheck)
'            strCellCheck = ReplaceVoucherItems(strCellCheck, Voucher, lngRow)
'
'            strValue = clsComp.GetExp(strCellCheck)
'        Else
'            strValue = "True"
'        End If
'        If CBool(strValue) Then
        If Not nod.Attributes.getNamedItem("cellcheck") Is Nothing Then
            strCellCheck = ReplaceVoucherItems(nod.Attributes.getNamedItem("cellcheck").Text, Voucher, lngRow)
        End If
        If Not nod.Attributes.getNamedItem("checktype") Is Nothing Then
            Select Case nod.Attributes.getNamedItem("checktype").Text
                Case "checkother"
                    strValue = GetConditionValue(Voucher, nod, Voucher.row)
                    If CBool(strValue) Then
                        CellCheck Voucher, nod.Attributes.getNamedItem("cardsection").nodeValue, nod.Attributes.getNamedItem("desfldname").nodeValue, bChange, clsRefer, lngRow, False
                    End If
                Case "check"
                    strValue = GetConditionValue(Voucher, nod, Voucher.row)
                    If CBool(strValue) Then
                        strCellCheck = ReplaceSysPara(strCellCheck)
                        strValue = clsComp.GetExp(strCellCheck)
                        If Not CBool(strValue) Then
                            If Not nod.Attributes.getNamedItem("errresid") Is Nothing Then
                                strError = nod.Attributes.getNamedItem("errresid").Text
                                SetVoucherItemValue Voucher, nod.Attributes.getNamedItem("cardsection").nodeValue, nod.Attributes.getNamedItem("fieldname").Text, "", lngRow
                                If blnShowMsg Then
                                    bChange = UAPVoucherControl85.CheckRet.ReTry
                                    MsgBox GetString(strError)
                                End If
                                Exit For
                            End If
                        End If
                    End If
                Case "calc"
                    If nod.Attributes.getNamedItem("cardsection").nodeValue = "T" Or (nod.Attributes.getNamedItem("cardsection").nodeValue = "B" And strCardSection = "B") Then
                        strValue = GetConditionValue(Voucher, nod, Voucher.row)
                        If CBool(strValue) Then
                            strValue = clsComp.GetExp(strCellCheck)
                            SetVoucherItemValue Voucher, nod.Attributes.getNamedItem("cardsection").nodeValue, nod.Attributes.getNamedItem("desfldname").Text, strValue, lngRow
                        End If
                    Else
                        For i = 1 To Voucher.BodyRows
                            strValue = GetConditionValue(Voucher, nod, i)
                            If CBool(strValue) Then
                                strCellCheck = ReplaceVoucherItems(nod.Attributes.getNamedItem("cellcheck").Text, Voucher, i)
                                strValue = clsComp.GetExp(strCellCheck)
                                SetVoucherItemValue Voucher, nod.Attributes.getNamedItem("cardsection").nodeValue, nod.Attributes.getNamedItem("desfldname").Text, strValue, i
                            End If
                        Next
                    End If
                Case "refervoucher"
                    strValue = GetConditionValue(Voucher, nod, Voucher.row)
                    If CBool(strValue) Then
                        CellCheck = nod.xml
                    End If
                Case "refercellcheck"
                    If nod.Attributes.getNamedItem("cardsection").nodeValue = "T" Or (nod.Attributes.getNamedItem("cardsection").nodeValue = "B" And strCardSection = "B") Then
                        strValue = GetConditionValue(Voucher, nod, lngRow)
                        If CBool(strValue) Then
                            clsRefer.ReferCellCheck Voucher, nod.Attributes.getNamedItem("cardsection").nodeValue, nod.Attributes.getNamedItem("desfldname").Text, lngRow
                            If nod.Attributes.getNamedItem("cardsection").nodeValue = "B" Then
                                If Not Voucher.ItemState(nod.Attributes.getNamedItem("desfldname").Text, sibody) Is Nothing Then
                                    Voucher.CallAutoFillBackEvent nod.Attributes.getNamedItem("desfldname").Text, lngRow
                                End If
                            End If
                        End If
                    Else
                        For i = 1 To Voucher.BodyRows
                            strValue = GetConditionValue(Voucher, nod, i)
                            If CBool(strValue) Then
                                clsRefer.ReferCellCheck Voucher, nod.Attributes.getNamedItem("cardsection").nodeValue, nod.Attributes.getNamedItem("desfldname").Text, i
                                If nod.Attributes.getNamedItem("cardsection").nodeValue = "B" Then
                                    If Not Voucher.ItemState(nod.Attributes.getNamedItem("desfldname").Text, sibody) Is Nothing Then
                                        Voucher.CallAutoFillBackEvent nod.Attributes.getNamedItem("desfldname").Text, i
                                    End If
                                End If
                            End If
                        Next
                    End If
                Case "referclear"
                    strValue = GetConditionValue(Voucher, nod, Voucher.row)
                    If CBool(strValue) Then
                        If nod.Attributes.getNamedItem("cardsection").nodeValue = "B" Then
                            If Voucher.bodyText(Voucher.row, nod.Attributes.getNamedItem("desfldname").Text) = "" Then
                                clsRefer.ClearReferItems Voucher, nod.Attributes.getNamedItem("cardsection").nodeValue, nod.Attributes.getNamedItem("desfldname").Text, strCellCheck, Voucher.row
                            End If
                            Voucher.CallAutoFillBackEvent nod.Attributes.getNamedItem("desfldname").Text, Voucher.row
                        Else
                            If Voucher.headerText(nod.Attributes.getNamedItem("desfldname").Text) = "" Then
                                clsRefer.ClearReferItems Voucher, nod.Attributes.getNamedItem("cardsection").nodeValue, nod.Attributes.getNamedItem("desfldname").Text, strCellCheck, Voucher.row
                            End If
                        End If
                    End If
            End Select
        End If
'        End If
    Next
    Set clsComp = Nothing
    
End Function

Private Function GetConditionValue(Voucher As ctlVoucher, nod As IXMLDOMElement, lngRow As Long) As String
    Dim strCellCheck As String
    Dim clsComp As New UsSaCompStr.clsCompStr
'    Dim strValue As String
    
    If Not nod.Attributes.getNamedItem("checkcondition") Is Nothing Then
        strCellCheck = nod.Attributes.getNamedItem("checkcondition").nodeValue
        strCellCheck = ReplaceSysPara(strCellCheck)
        strCellCheck = ReplaceVoucherItems(strCellCheck, Voucher, lngRow)
        
        GetConditionValue = clsComp.GetExp(strCellCheck)
    Else
        GetConditionValue = "True"
    End If
End Function


Public Function CheckBeforeSave(strCheckPoint As String, Domhead As DOMDocument, Dombody As DOMDocument, domErrorMsg As DOMDocument) As Boolean
    Dim domCheck As New DOMDocument
    Dim nod As IXMLDOMNode
    Dim eleRow As IXMLDOMElement
    Dim strError As String
    Dim Rst As New ADODB.Recordset
    Dim strCheckName As String
    Dim strErrorType As String
    Dim nodHead As IXMLDOMNode
    Dim nodBody As IXMLDOMNode
    Dim nodList As IXMLDOMNodeList
    Dim lngRow As Long
    Dim nodVoucher As IXMLDOMNode
    Dim strNodValue As String
    Set nodVoucher = domVouchers.selectSingleNode("//z:row")
'    rst.CursorLocation = adUseClient
'    rst.Open "select * from sa_checkbeforesave where cardnumber=N'" + strCardNum + "' and cardsection='T' and ccheckpoint=N'" + strCheckPoint + "'", g_conn, adOpenForwardOnly, adLockReadOnly, adCmdText
'    rst.Save domCheck, adPersistXML
'    rst.Close
'    Set domCheck = domCheckBeforeSave.selectNodes("//z:row[@cardsection='T']")
    
    Set nodHead = Domhead.selectSingleNode("//z:row")
    For Each nod In domCheckBeforeSave.selectNodes("//z:row[@cardsection='T']")
        strCheckName = nod.Attributes.getNamedItem("checkname").Text
        strErrorType = nod.Attributes.getNamedItem("checktype").Text
        Set eleRow = domErrorMsg.documentElement.selectSingleNode("row[@name='" + strCheckName + "']")
        If eleRow Is Nothing Then
            strError = CallCheckFunction(nod, DBConn, strCardNums, nodVoucher, nodHead, nodHead, domErrorMsg)
            If strError <> "" Then
                Set eleRow = domErrorMsg.createElement("row")
                eleRow.setAttribute "name", strCheckName
                eleRow.setAttribute "type", strErrorType
                eleRow.setAttribute "ignore", "false"
                eleRow.setAttribute "message", strError
                domErrorMsg.documentElement.appendChild eleRow
                If strErrorType = "error" Then
                    CheckBeforeSave = False
                    Exit For
                End If
            End If
        Else
            If Not CBool(eleRow.Attributes.getNamedItem("ignore").Text) Then
                strError = CallCheckFunction(nod, DBConn, strCardNums, nodVoucher, nodHead, nodHead, domErrorMsg)
                If strError <> "" Then
                    eleRow.setAttribute "message", strError
                    If strErrorType = "error" Then
                        CheckBeforeSave = False
                        Exit For
                    End If
                End If
            End If
        End If
    Next
    
'    rst.Open "select * from sa_checkbeforesave where cardnumber=N'" + strCardNum + "' and cardsection='B' and ccheckpoint=N'" + strCheckPoint + "'", g_conn, adOpenForwardOnly, adLockReadOnly, adCmdText
'    rst.Save domCheck, adPersistXML
'    rst.Close
'    lngRow = 0

    Set nodList = Dombody.selectNodes("//z:row[@editprop!='D']")
    For Each nodBody In nodList
        lngRow = lngRow + 1
        For Each nod In domCheckBeforeSave.selectNodes("//z:row[@cardsection='B']") ' domCheck.selectNodes("//z:row")
            strCheckName = nod.Attributes.getNamedItem("checkname").Text
            strErrorType = nod.Attributes.getNamedItem("checktype").Text
            Set eleRow = domErrorMsg.documentElement.selectSingleNode("row[@name='" + strCheckName + "']")
            If eleRow Is Nothing Then
                strError = CallCheckFunction(nod, DBConn, strCardNums, nodVoucher, nodHead, nodBody, domErrorMsg, lngRow)
                If strError <> "" Then
                    Set eleRow = domErrorMsg.createElement("row")
                    eleRow.setAttribute "name", strCheckName
                    eleRow.setAttribute "type", strErrorType
                    eleRow.setAttribute "ignore", "false"
                    eleRow.setAttribute "message", strError
                    domErrorMsg.documentElement.appendChild eleRow
                    If strErrorType = "error" Then
                        CheckBeforeSave = False
'                        Exit For
                    End If
                End If
            Else
                If Not CBool(eleRow.Attributes.getNamedItem("ignore").Text) Then
                    strNodValue = eleRow.Attributes.getNamedItem("message").nodeValue
                    strError = CallCheckFunction(nod, DBConn, strCardNums, nodVoucher, nodHead, nodBody, domErrorMsg, lngRow)
                    If strError <> "" Then
                        strNodValue = strNodValue & IIf(strNodValue = "", "", Chr(13)) & strError
                        eleRow.setAttribute "message", strNodValue
                        If strErrorType = "error" Then
                            CheckBeforeSave = False
'                            Exit For
                        End If
                    End If
                End If
            End If
        Next
    Next
    Set domCheck = Nothing
    Set Rst = Nothing
End Function

Private Function CallCheckFunction(nod As IXMLDOMNode, cn As ADODB.Connection, strCardNum As String, nodVoucher As IXMLDOMNode, nodHead As IXMLDOMNode, nodBody As IXMLDOMNode, domErrorMsg As DOMDocument, Optional lngRow As Long) As String
    Dim strCheckName As String
    
    strCheckName = nod.Attributes.getNamedItem("checkname").Text
    Select Case strCheckName
        Case "pricefree"
            CallCheckFunction = CheckPriceFree(cn, strCardNum, nodVoucher, nodHead, nodBody, lngRow)
        Case "invpriceunique"
            CallCheckFunction = CheckInvPriceUnique(cn, strCardNum, nodBody, lngRow)
        Case "cuspriceunique"
            CallCheckFunction = CheckCusPriceUnique(cn, strCardNum, nodBody, lngRow)
        Case "checksignquantity"
            CallCheckFunction = CheckSignQuantity(cn, nodBody, lngRow)
        Case Else
            CallCheckFunction = CheckFormular(cn, strCardNum, nod, nodHead, nodBody, lngRow)
    End Select
End Function
'检查调价单表体自由项
Private Function CheckPriceFree(cn As ADODB.Connection, strCardNum As String, nodVoucher As IXMLDOMNode, nodHead As IXMLDOMNode, nodData As IXMLDOMElement, lngRow As Long) As String
    Dim lst As IXMLDOMNodeList
    Dim nod As IXMLDOMNode
'    Dim lstData As IXMLDOMNodeList
'    Dim nodData As IXMLDOMElement
    Dim strVoucherName As String
    Dim strVoucherCode As String
    Dim Rst As New ADODB.Recordset
    Dim domFieldConfig As New DOMDocument
    
    strVoucherName = nodVoucher.Attributes.getNamedItem("Name").Text
    strVoucherCode = nodHead.Attributes.getNamedItem(LCase(nodVoucher.Attributes.getNamedItem("ReceiptNOFieldName").Text)).Text
    Rst.CursorLocation = adUseClient
    Rst.Open "select fieldname,mustinputtype from sa_voucherfieldconfig where cardnumber=N'" + strCardNum + "' and isnull(mustinputtype,N'')<>N''", cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Rst.Save domFieldConfig, adPersistXML
    Rst.Close
'    Set lstData = domBody.selectNodes("//z:row[@editprop!='D']")
    Set lst = domFieldConfig.selectNodes("//z:row[@mustinputtype='pricefree']")
    If lst.length > 0 Then
        Dim strFreePriceType As String
        Dim strInvCode As String
        Dim strFieldName As String
        Dim strMark As String
        Dim strValue As String
'        Dim lngRow As Long
        
'        lngRow = 0
        strFreePriceType = getAccinformation("SA", "FreePriceType")
'        For Each nodData In lstData
        strInvCode = nodData.Attributes.getNamedItem("cinvcode").Text
        Rst.Open "select bfree1,bfree2,bfree3,bfree4,bfree5,bfree6,bfree7,bfree8,bfree9,bfree10 from inventory where cinvcode=N'" + strInvCode + "'", cn, adOpenForwardOnly, adLockReadOnly, adCmdText
        For Each nod In lst
            strFieldName = nod.Attributes.getNamedItem("fieldname").Text
            strMark = Right(strFieldName, Len(strFieldName) - 5)
            If Mid(strFreePriceType, val(strMark), 1) = "1" Then
                If CBool(Rst.Fields("bfree" + strMark)) Then
                    strValue = nodData.Attributes.getNamedItem(strFieldName).Text
                    If strValue = "" Then
                        CheckPriceFree = CheckPriceFree & Chr(13) & Chr(10) & GetStringPara("U8.SA.USSASERVER.clsfhfunc.01866", strVoucherName, strVoucherCode, lngRow, strInvCode, getDefineName(cn, strFieldName))  'Para zh-CN：{0}[{1}]第{2}行: 存货[{3}]自由项[{4}]不可以为空
                    End If
                Else
                    nodData.setAttribute strFieldName, ""
                End If
            Else
                nodData.setAttribute strFieldName, ""
            End If
        Next
        Rst.Close
'        Next
    End If
    Set Rst = Nothing
    Set domFieldConfig = Nothing
End Function


Private Function CheckInvPriceUnique(cn As ADODB.Connection, strCardNum As String, nodBody As IXMLDOMElement, lngRow As Long) As String
'    Dim nod As IXMLDOMElement
'    Dim lst As IXMLDOMNodeList
'    Dim I As Long
'    Dim strWhere As String
    Dim strError As String
    Dim strAutoID As String
'    Dim strValue As String
    
    strError = ""
'    Set lst = domBody.selectNodes("//z:row[@operationtype='新增']")
'    For Each nod In lst
    If nodBody.Attributes.getNamedItem("operationtype").nodeValue = "新增" Then
        strAutoID = GetInvUniqueAutoid(cn, nodBody)
        If strAutoID <> "" Then
'            strValue = nod.Attributes.getNamedItem("cinvcode").Text
            strError = IIf(strError = "", "", strError + Chr(13) + Chr(10)) & GetStringPara("U8.SA.xsglsql_2.invpricemulline", CStr(lngRow))         '"第" & CStr(lngRow) + "行在存货价格表中重复,是否更新原记录"
            nodBody.setAttribute "operationtype", "修改"
            nodBody.setAttribute "ipriceid", strAutoID
        End If
    End If
'    Next
    CheckInvPriceUnique = strError
End Function

Private Function CheckCusPriceUnique(cn As ADODB.Connection, strCardNum As String, nodBody As IXMLDOMElement, lngRow As Long) As String
'    Dim nod As IXMLDOMElement
'    Dim lst As IXMLDOMNodeList
'    Dim I As Long
'    Dim strWhere As String
    Dim strError As String
    Dim strAutoID As String
'    Dim strValue As String
    
'    Set lst = domBody.selectNodes("//z:row[@operationtype='新增']")
'    For Each nod In lst
    If nodBody.Attributes.getNamedItem("operationtype").nodeValue = "新增" Then
        strAutoID = GetCusUniqueAutoid(cn, nodBody)
        If strAutoID <> "" Then
'            strValue = nod.Attributes.getNamedItem("cinvcode").Text
            strError = IIf(strError = "", "", strError + Chr(13) + Chr(10)) & GetStringPara("U8.SA.xsglsql_2.cuspricemulline", CStr(lngRow))     '"第" & CStr(lngRow) + "行在客户价格表中重复,是否更新原行"
            nodBody.setAttribute "operationtype", "修改"
            nodBody.setAttribute "ipriceid", strAutoID
        End If
    End If
'    Next
    CheckCusPriceUnique = strError
End Function

Private Function CheckSignQuantity(CNN As ADODB.Connection, nodBody As IXMLDOMNode, lngRow As Long) As String
'    Dim nod As IXMLDOMElement
'    Dim lst As IXMLDOMNodeList
    Dim lngDLsid As Long
    Dim Rst As New ADODB.Recordset
    
    Rst.CursorLocation = adUseClient
    CheckSignQuantity = ""
'    Set lst = domBody.selectNodes("//z:row[@editprop!='D']")
'    For Each nod In lst
    lngDLsid = nodBody.Attributes.getNamedItem("idlsid").Text
    Rst.Open "select isnull(foutquantity,0) as foutquantity,isnull(fsumsignquantity,0) as fsumsignquantity from dispatchlists where idlsid=" & CStr(lngDLsid), CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not Rst.EOF Then
        If Rst.Fields("foutquantity") > 0 Then
            If Rst(1) > Rst(0) Then
                CheckSignQuantity = CheckSignQuantity & IIf(CheckSignQuantity = "", "", Chr(13)) & GetStringPara("U8.SA.xsglsql_2.toomuchquantity", CStr(lngRow))  '"第" + CStr(lngRow) + "行签回数量大于累计出库数量"
            End If
        Else
            If Rst(1) < Rst(0) Then
                CheckSignQuantity = CheckSignQuantity & IIf(CheckSignQuantity = "", "", Chr(13)) & GetStringPara("U8.SA.xsglsql_2.toomuchquantity", CStr(lngRow))    '"第" + CStr(lngRow) + "行签回数量大于累计出库数量"
            End If
        End If
    End If
    Rst.Close
'    Next
    Set Rst = Nothing
End Function


Private Function CheckFormular(cn As ADODB.Connection, strCardNum As String, nodCheck As IXMLDOMNode, nodHead As IXMLDOMNode, nodBody As IXMLDOMNode, lngRow As Long) As String
    Dim nod As IXMLDOMElement
    Dim lst As IXMLDOMNodeList
    Dim lngDLsid As Long
    Dim strValue As String
    Dim clsComStr As New UsSaCompStr.clsCompStr
    Dim strFormular As String
    Dim strErrorRes As String
    Dim strCardSection As String
    
    strFormular = nodCheck.Attributes.getNamedItem("checkname").Text
    strErrorRes = ReplaceItemName(cn, strCardNum, GetString(nodCheck.Attributes.getNamedItem("errresid").Text))
    strCardSection = GetString(nodCheck.Attributes.getNamedItem("cardsection").Text)
    CheckFormular = ""
    strValue = ReplaceItemValue(strFormular, nodHead, nodBody)
        
    strValue = clsComStr.GetExp(strValue)
    If Not CBool(strValue) Then
        CheckFormular = IIf(strCardSection = "T", "", GetStringPara("U8.SA.xsglsql_2.rowhaserror", CStr(lngRow))) & strErrorRes '"单据体第" + CStr(lngRow) + "行有错误"
    End If
End Function


''取自定义项目的名称
Public Function getDefineName(cn As ADODB.Connection, strFldName As String, Optional cItemName As String) As String
    Dim strID As String
    Dim Rst As New ADODB.Recordset
'    Dim tmpCn As New adodb.Connection
'    tmpCn.Open clsSys.Connectstr, , clsSys.SysPassword
    getDefineName = ""
    Rst.CursorLocation = adUseClient
    strID = GetDefineID(strFldName)
    Rst.Open "Select isnull(cItemName,cItemLang),ISNULL(cItemLang,'') AS cItemName from userdef where cID='" & strID & "'", cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not Rst.EOF Then
        getDefineName = Rst(0)
        cItemName = Rst(1)
    End If
    Rst.Close
    Set Rst = Nothing
    
End Function

Private Function GetInvUniqueAutoid(cn As ADODB.Connection, nod As IXMLDOMNode) As String
    Dim Rst As New ADODB.Recordset
    Dim strInvCode As String
    Dim strFree1 As String
    Dim strFree2 As String
    Dim strStartDate As String
    Dim strEndDate As String
    Dim strQuantity As String
    Dim strSales As String
    
    Rst.CursorLocation = adUseClient
    strInvCode = nod.Attributes.getNamedItem("cinvcode").Text
    strStartDate = GetItemValueStr("5", nod.Attributes.getNamedItem("dstartdate").Text)
    strQuantity = nod.Attributes.getNamedItem("fminquantity").Text
    If Not nod.Attributes.getNamedItem("bsales") Is Nothing Then
        strSales = GetItemValueStr("4", nod.Attributes.getNamedItem("bsales").Text)
    End If
    If Not nod.Attributes.getNamedItem("cfree1") Is Nothing Then
        strFree1 = nod.Attributes.getNamedItem("cfree1").Text
    End If
    If Not nod.Attributes.getNamedItem("cfree2") Is Nothing Then
        strFree2 = nod.Attributes.getNamedItem("cfree2").Text
    End If
    If Not nod.Attributes.getNamedItem("cfree2") Is Nothing Then
        strFree2 = nod.Attributes.getNamedItem("cfree2").Text
    End If
    strEndDate = "''"
    If strEndDate = "''" Then strEndDate = ""
    Rst.Open "select autoid from sa_invuprice where cinvcode=N'" + strInvCode + "' and isnull(cfree1,N'')=N'" + strFree1 + "' and isnull(cfree2,N'')=N'" + strFree2 + "' and bsales=" + strSales _
            & " and isnull(fminquantity,0)=" & strQuantity & " and ((isnull(sa_invuprice.dstartdate,N'1900-1-1')>=" + strStartDate + "" & IIf(strEndDate = "", ")", " and isnull(sa_invuprice.dstartdate,N'1900-1-1')<=" + strEndDate + ")") _
            & " or (isnull(sa_invuprice.dstartdate,N'1900-1-1')<=" + strStartDate + " and isnull(sa_invuprice.denddate,N'9999-12-31')>=" + strStartDate + "))", cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not Rst.EOF Then
        GetInvUniqueAutoid = Rst(0)
    End If
    Rst.Close
    Set Rst = Nothing
End Function



Private Function GetCusUniqueAutoid(cn As ADODB.Connection, nod As IXMLDOMNode) As String
    Dim Rst As New ADODB.Recordset
    Dim strInvCode As String
    Dim strCusCode As String
    Dim strCusCCode As String
    Dim strFree1 As String
    Dim strFree2 As String
    Dim strStartDate As String
    Dim strEndDate As String
    Dim strQuantity As String
    Dim strSales As String
    Dim strInvSCost As String
    Dim strsql As String
    
    If Not nod.Attributes.getNamedItem("ccuscode") Is Nothing Then
        strCusCode = nod.Attributes.getNamedItem("ccuscode").Text
    End If
    If Not nod.Attributes.getNamedItem("ccusccode") Is Nothing Then
        strCusCCode = nod.Attributes.getNamedItem("ccusccode").Text
    End If
    Rst.CursorLocation = adUseClient
    strInvCode = nod.Attributes.getNamedItem("cinvcode").Text
    strStartDate = GetItemValueStr("5", nod.Attributes.getNamedItem("dstartdate").Text)
    strQuantity = nod.Attributes.getNamedItem("fminquantity").Text
    If Not nod.Attributes.getNamedItem("bsales") Is Nothing Then
        strSales = GetItemValueStr("4", nod.Attributes.getNamedItem("bsales").Text)
    End If
    If Not nod.Attributes.getNamedItem("cfree1") Is Nothing Then
        strFree1 = nod.Attributes.getNamedItem("cfree1").Text
    End If
    If Not nod.Attributes.getNamedItem("cfree2") Is Nothing Then
        strFree2 = nod.Attributes.getNamedItem("cfree2").Text
    End If
    strEndDate = "''"
    If Not nod.Attributes.getNamedItem("denddate") Is Nothing Then
        strEndDate = GetItemValueStr("5", nod.Attributes.getNamedItem("denddate").Text)
    End If
    If strEndDate = "''" Then strEndDate = ""
    If Not nod.Attributes.getNamedItem("iinvscost") Is Nothing Then
        strInvSCost = GetItemValueStr("4", nod.Attributes.getNamedItem("iinvscost").Text)
    End If
    If strInvSCost = "" Then
        strsql = "select autoid from sa_cusuprice where isnull(ccusccode,'')='" + strCusCCode + "' and isnull(ccuscode,'')='" + strCusCode + "' and isnull(iinvscost,'')='' " _
            & " and cinvcode='" + strInvCode + "' and isnull(cfree1,'')='" + strFree1 + "' and isnull(cfree2,'')='" + strFree2 + "' and bsales=" + strSales _
            & " and isnull(fminquantity,0)=" & strQuantity & " and ((isnull(sa_cusuprice.dstartdate,'1900-1-1')>=" + strStartDate + "" & IIf(strEndDate = "", ")", " and isnull(sa_cusuprice.dstartdate,'1900-1-1')<=" + strEndDate + ")") _
            & " or (isnull(sa_cusuprice.dstartdate,'1900-1-1')<=" + strStartDate + " and isnull(sa_cusuprice.denddate,'9999-12-31')>=" + strStartDate + "))"
    Else
        strsql = "select autoid from sa_cusuprice where isnull(ccusccode,'')='" + strCusCCode + "' and isnull(ccuscode,'')='" + strCusCode + "' and isnull(iinvscost,'')<>'' " _
            & " and cinvcode='" + strInvCode + "' and isnull(cfree1,'')='" + strFree1 + "' and isnull(cfree2,'')='" + strFree2 + "' and bsales=" + strSales _
            & " and isnull(fminquantity,0)=" & strQuantity & " and ((isnull(sa_cusuprice.dstartdate,'1900-1-1')>=" + strStartDate + "" & IIf(strEndDate = "", ")", " and isnull(sa_cusuprice.dstartdate,'1900-1-1')<=" + strEndDate + ")") _
            & " or (isnull(sa_cusuprice.dstartdate,'1900-1-1')<=" + strStartDate + " and isnull(sa_cusuprice.denddate,'9999-12-31')>=" + strStartDate + "))"
    End If
    Rst.Open ConvertSQLString(strsql), cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not Rst.EOF Then
        GetCusUniqueAutoid = Rst(0)
    End If
    Rst.Close
    Set Rst = Nothing
End Function


Private Function ReplaceItemName(cn As ADODB.Connection, strCardNum As String, strFormular As String) As String
    Dim lngPos1 As Long
    Dim lngPos2 As Long
    Dim strTmp As String
    Dim varStr As Variant
    Dim strFieldName As String
    Dim strValue As String
    Dim strCardSection As String
    
    strTmp = strFormular
    lngPos1 = InStr(1, strTmp, "(")
    Do While lngPos1 > 0
        lngPos2 = InStr(lngPos1, strTmp, ")")
        strFieldName = Mid(strTmp, lngPos1 + 1, lngPos2 - lngPos1 - 1)
        varStr = Split(strFieldName, ",")
        strCardSection = varStr(0)
        strFieldName = varStr(1)
        strValue = ""
        strValue = GetCardItemName(cn, strCardNum, strCardSection, strFieldName)
'        If strCardSection = "T" Then
'            If Not nodHead.Attributes.getNamedItem(strFieldName) Is Nothing Then
'                strValue = nodHead.Attributes.getNamedItem(strFieldName).nodeValue
'            End If
'        Else
'            If Not nodBody.Attributes.getNamedItem(strFieldName) Is Nothing Then
'                strValue = nodBody.Attributes.getNamedItem(strFieldName).nodeValue
'            End If
'        End If
        strTmp = Replace(strTmp, "(" & strCardSection & "," & strFieldName & ")", strValue)
        lngPos1 = InStr(1, strTmp, "(")
    Loop
    ReplaceItemName = strTmp
End Function



Private Function ReplaceItemValue(strFormular As String, nodHead As IXMLDOMNode, nodBody As IXMLDOMNode) As String
    Dim lngPos1 As Long
    Dim lngPos2 As Long
    Dim strTmp As String
    Dim varStr As Variant
    Dim strFieldName As String
    Dim strValue As String
    Dim strCardSection As String
    
    strTmp = strFormular
    lngPos1 = InStr(1, strTmp, "[")
    Do While lngPos1 > 0
        lngPos2 = InStr(lngPos1, strTmp, "]")
        strFieldName = Mid(strTmp, lngPos1 + 1, lngPos2 - lngPos1 - 1)
        varStr = Split(strFieldName, ",")
        strCardSection = varStr(0)
        strFieldName = varStr(1)
        strValue = ""
        If strCardSection = "T" Then
            If Not nodHead.Attributes.getNamedItem(strFieldName) Is Nothing Then
                strValue = nodHead.Attributes.getNamedItem(strFieldName).nodeValue
            End If
        Else
            If Not nodBody.Attributes.getNamedItem(strFieldName) Is Nothing Then
                strValue = nodBody.Attributes.getNamedItem(strFieldName).nodeValue
            End If
        End If
        strTmp = Replace(strTmp, "[" & strCardSection & "," & strFieldName & "]", strValue)
        lngPos1 = InStr(1, strTmp, "[")
    Loop
    ReplaceItemValue = strTmp
End Function


Public Function GetDefineID(strFldName As String) As String
    Select Case LCase(strFldName)
        Case "cdefine1"
            GetDefineID = "01"  '单据头自定义项1
        Case "cdefine2"
            GetDefineID = "02"  '单据头自定义项2
        Case "cdefine3"
            GetDefineID = "03"  '单据头自定义项3
        Case "cdefine4"
            GetDefineID = "04"  '单据头自定义项4
        Case "cdefine5"
            GetDefineID = "05"  '单据头自定义项5
        Case "cdefine6"
            GetDefineID = "06"  '单据头自定义项6
        Case "cdefine7"
            GetDefineID = "07"  '单据头自定义项7
        Case "cdefine8"
            GetDefineID = "08"  '单据头自定义项8
        Case "cdefine9"
            GetDefineID = "09"  '单据头自定义项9
        Case "cdefine10"
            GetDefineID = "10"  '单据头自定义项10
        Case "cdefine11"
            GetDefineID = "36"  '单据头自定义项11
        Case "cdefine12"
            GetDefineID = "37"  '单据头自定义项12
        Case "cdefine13"
            GetDefineID = "38"  '单据头自定义项13
        Case "cdefine14"
            GetDefineID = "39"  '单据头自定义项14
        Case "cdefine15"
            GetDefineID = "40"  '单据头自定义项15
        Case "cdefine16"
            GetDefineID = "41"  '单据头自定义项16
            
        Case "cdefine22"
            GetDefineID = "22"  '单据体自定义项1
        Case "cdefine23"
            GetDefineID = "23"  '单据体自定义项2
        Case "cdefine24"
            GetDefineID = "24"  '单据体自定义项3
        Case "cdefine25"
            GetDefineID = "25"  '单据体自定义项4
        Case "cdefine26"
            GetDefineID = "26"  '单据体自定义项5
        Case "cdefine27"
            GetDefineID = "27"  '单据体自定义项6
        Case "cdefine28"
            GetDefineID = "42"  '单据体自定义项7
        Case "cdefine29"
            GetDefineID = "43"  '单据体自定义项8
        Case "cdefine30"
            GetDefineID = "44"  '单据体自定义项9
        Case "cdefine31"
            GetDefineID = "45"  '单据体自定义项10
        Case "cdefine32"
            GetDefineID = "46"  '单据体自定义项11
        Case "cdefine33"
            GetDefineID = "47"  '单据体自定义项12
        Case "cdefine34"
            GetDefineID = "48"  '单据体自定义项13
        Case "cdefine35"
            GetDefineID = "49"  '单据体自定义项14
        Case "cdefine36"
            GetDefineID = "50"  '单据体自定义项15
        Case "cdefine37"
            GetDefineID = "51"  '单据体自定义项16
            
        Case "cinvdefine1"
            GetDefineID = "17"
        Case "cinvdefine2"
            GetDefineID = "18"
        Case "cinvdefine3"
            GetDefineID = "19"
        Case "cfree1"
            GetDefineID = "20"
        Case "cfree2"
            GetDefineID = "21"
        Case "cfree3"
            GetDefineID = "28"
        Case "cfree4"
            GetDefineID = "29"
        Case "cfree5"
            GetDefineID = "30"
        Case "cfree6"
            GetDefineID = "31"
        Case "cfree7"
            GetDefineID = "32"
        Case "cfree8"
            GetDefineID = "33"
        Case "cfree9"
            GetDefineID = "34"
        Case "cfree10"
            GetDefineID = "35"
            
        Case "cinvdefine4"       ''自定义项4
            GetDefineID = "52"
        Case "cinvdefine5"       ''自定义项5
            GetDefineID = "53"
        Case "cinvdefine6"       ''自定义项6
            GetDefineID = "54"
        Case "cinvdefine7"       ''自定义项7
            GetDefineID = "55"
        Case "cinvdefine8"       ''自定义项8
            GetDefineID = "56"
        Case "cinvdefine9"       ''自定义项9
            GetDefineID = "57"
        Case "cinvdefine10"       ''自定义项10
            GetDefineID = "58"
        Case "cinvdefine11"       ''自定义项11
            GetDefineID = "59"
        Case "cinvdefine12"       ''自定义项12
            GetDefineID = "60"
        Case "cinvdefine13"       ''自定义项13
            GetDefineID = "61"
        Case "cinvdefine14"       ''自定义项14
            GetDefineID = "62"
        Case "cinvdefine15"       ''自定义项15
            GetDefineID = "63"
        Case "cinvdefine16"       ''自定义项16
            GetDefineID = "64"
            
'   客户
        Case "ccusdefine1"       ''自定义项7
            GetDefineID = "11"
        Case "ccusdefine2"       ''自定义项8
            GetDefineID = "12"
        Case "ccusdefine3"       ''自定义项9
            GetDefineID = "13"
        Case "ccusdefine4"       ''自定义项10
            GetDefineID = "65"
        Case "ccusdefine5"       ''自定义项11
            GetDefineID = "66"
        Case "ccusdefine6"       ''自定义项12
            GetDefineID = "67"
        Case "ccusdefine7"       ''自定义项13
            GetDefineID = "68"
        Case "ccusdefine8"       ''自定义项14
            GetDefineID = "69"
        Case "ccusdefine9"       ''自定义项15
            GetDefineID = "70"
        Case "ccusdefine10"       ''自定义项16
            GetDefineID = "71"
            
        Case "ccusdefine11"       ''自定义项11
            GetDefineID = "72"
        Case "ccusdefine12"       ''自定义项12
            GetDefineID = "73"
        Case "ccusdefine13"       ''自定义项13
            GetDefineID = "74"
        Case "ccusdefine14"       ''自定义项14
            GetDefineID = "75"
        Case "ccusdefine15"       ''自定义项15
            GetDefineID = "76"
        Case "ccusdefine16"       ''自定义项16
            GetDefineID = "77"
            
    End Select
End Function



Private Function GetItemValueStr(strFieldType As String, strValue As String) As String
    Select Case strFieldType
        Case "3", "4"
            If LCase(strValue) = "false" Or LCase(strValue) = "true" Then
                GetItemValueStr = IIf(CBool(strValue), "1", "0")
            Else
                GetItemValueStr = strValue
            End If
        Case "5"
            Dim lngPos1 As String
            lngPos1 = InStr(1, strValue, "T")
            If lngPos1 > 0 Then
                GetItemValueStr = "N'" + Left(strValue, lngPos1 - 1) + "'"
            Else
                GetItemValueStr = "N'" + strValue + "'"
            End If
        Case Else
            GetItemValueStr = "N'" + strValue + "'"
    End Select
End Function


Private Function GetCardItemName(CNN As ADODB.Connection, strCardNum As String, strCardSection As String, strFieldName As String) As String
    Dim Rst As New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    
    Rst.Open "select carditemname from voucheritems where cardnum=N'" + strCardNum + "' and cardsection='" + strCardSection + "' and fieldname=N'" + strFieldName + "'", CNN, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not Rst.EOF Then
        GetCardItemName = Rst(0)
    Else
        GetCardItemName = ""
    End If
    Rst.Close
    Set Rst = Nothing
End Function
