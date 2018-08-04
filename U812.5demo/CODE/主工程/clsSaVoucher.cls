VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSaVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public domVouchers As New DOMDocument
Public domDispVtid As New DOMDocument
Public domPrntVtid As New DOMDocument
Public domVouTemplate As New DOMDocument
Public domFieldConfig As New DOMDocument
Public domCodeXml As New DOMDocument
Private domFieldState As New DOMDocument
Dim domCellCheck As New DOMDocument
Dim domPopMnuConfig As New DOMDocument
Dim domRowColChange As New DOMDocument
'Dim m_lngClr1 As Long
'Dim m_lngClr2 As Long
Dim clsAuth As New SaVoucherService.clsSAAuth
Dim m_BodyMaxRows As Integer
Dim m_strCardNum As String
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




Public Function Init(strCardNum As String, strErrorResId As String) As Boolean
    Dim rst As New ADODB.Recordset
    Dim strAuth As String
    Dim strSql As String
    m_strCardNum = strCardNum
    rst.CursorLocation = adUseClient
    rst.Open "select * from vouchers_base where cardnumber=N'" + strCardNum + "'", DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rst.Save domVouchers, adPersistXML
    rst.Close
    Set clsAuth.clsAuth = m_login.RowAuth
    clsAuth.Init DBConn, m_login.cUserId, "voucher", strCardNum
    strAuth = clsAuth.getAuthString("DJMB")
    If strAuth = "1=2" Then strAuth = "''"
    strSql = "SELECT VT_Name,VT_ID,isnull(VT_PrintTemplID,DEF_ID_PRN) as DEF_ID_PRN From vouchertemplates inner join vouchers on vouchertemplates.VT_CardNumber=vouchers.cardnumber WHERE VT_CardNumber=N'" + strCardNum + "' AND (VT_TemplateMode = 0) " & IIf(strAuth <> "", "and vt_id in (" & strAuth & ")", "") + " order by case when vt_id=def_id then vt_id else null end desc"
    rst.Open strSql
    If rst.RecordCount <= 0 Then
        strErrorResId = GetString("U8.SA.xsglsql.01.frmbillvouch.00171") '"没有显示模版权限"
        rst.Close
        Set rst = Nothing
        Init = False
        Exit Function
    End If
    rst.Save domDispVtid, adPersistXML
    rst.Close
    strSql = "SELECT VT_Name,VT_ID  From vouchertemplates WHERE VT_CardNumber=N'" + strCardNum + "' AND (VT_TemplateMode = 1) " & IIf(strAuth <> "", "and vt_id in (" & strAuth & ")", "")
    rst.Open strSql
    rst.Save domPrntVtid, adPersistXML
    rst.Close
    rst.Open "select * from sa_voucherfieldconfig where cardnumber=N'" + strCardNum + "' order by cardsection", DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rst.Save domFieldConfig, adPersistXML
    rst.Close
    rst.Open "select * from sa_vouchercellcheck where cardnum=N'" + strCardNum + "' order by fldcardsection,fieldname,checkindex", DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rst.Save domCellCheck, adPersistXML
    rst.Close
    rst.Open "select * from sa_popmenuconfig where cardnumber=N'" + strCardNum + "'", DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rst.Save domPopMnuConfig, adPersistXML
    rst.Close
    rst.Open "select * from sa_rowcolchange where cardnumber=N'" + strCardNum + "'", DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rst.Save domRowColChange, adPersistXML
    rst.Close
    
    Set rst = Nothing
    On Error Resume Next
    domFieldState.loadXML "<Data />"
    strErrorResId = GetVoucherCodeXml(strCardNum)
    Init = True
End Function
Public Function CellCheck(Voucher As ctlVoucher, strCardSection As String, strFieldName As String, bChange As Long, clsRefer As clsSaRefer, Optional lngRow As Long, Optional blnShowMsg As Boolean = True)
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
                                    bChange = UAPVoucherControl85.CheckRet.retry
                                    MsgBox GetString(strError)
                                End If
                                Exit For
                            End If
                        End If
                    End If
                Case "calc"
                    If nod.Attributes.getNamedItem("cardsection").nodeValue = "T" Or (nod.Attributes.getNamedItem("cardsection").nodeValue = "B" And strCardSection = "B") Then
                        strValue = GetConditionValue(Voucher, nod, Voucher.row)
                        If CBool(IIf(strValue = "", 0, 1)) Then
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

Public Function CopyLine(Voucher As ctlVoucher, clsRefer As clsSaRefer) As Boolean
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


Private Sub SetFldDefaultValue(Voucher As ctlVoucher, nod As IXMLDOMElement, lngRow As Long, clsRefer As clsSaRefer)
    Dim varPar As Variant
    Dim rst As New ADODB.Recordset
    Dim strValue As String
    Dim i As Long
    
    rst.CursorLocation = adUseClient
'    MsgBox nod.Attributes.getNamedItem("addnewdefaultvalue").Text
    varPar = Split(nod.Attributes.getNamedItem("addnewdefaultvalue").Text, ":")
    Select Case varPar(0)
        Case "0"
            strValue = ReplaceVoucherItems(CStr(varPar(1)), Voucher, lngRow)
            SetVoucherItemValue Voucher, nod.Attributes.getNamedItem("cardsection").Text, nod.Attributes.getNamedItem("fieldname").Text, strValue, lngRow
        Case "1"        '全局变量
            SetVoucherItemValue Voucher, nod.Attributes.getNamedItem("cardsection").Text, nod.Attributes.getNamedItem("fieldname").Text, GetGlobalVariant(CStr(varPar(1))), lngRow
        Case "2"        'sql
            rst.Open varPar(1), DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not rst.EOF Then
                For i = 0 To rst.Fields.Count - 1
                    SetVoucherItemValue Voucher, nod.Attributes.getNamedItem("cardsection").Text, rst(i).Name, rst(i).value, lngRow
                Next
                clsRefer.ReferCellCheck Voucher, nod.Attributes.getNamedItem("cardsection").Text, nod.Attributes.getNamedItem("fieldname").Text, lngRow
            End If
            rst.Close
    End Select
    Set rst = Nothing
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





Public Function GetVoucherItemValue(Voucher As ctlVoucher, strSection As String, strFieldName As String, Optional lngRow As Long) As String
    If strSection = "B" Then
        GetVoucherItemValue = Voucher.bodyText(lngRow, strFieldName)
    End If
    If strSection = "T" Then
        GetVoucherItemValue = Voucher.headerText(strFieldName)
    End If
End Function

Private Function GetVoucherCodeXml(strCardNum As String) As String
    Dim objBillNo As New UFBillComponent.clsBillComponent
    Dim strXml As String
    
    If objBillNo.InitBill(m_login.UfDbName, strCardNum) = False Then
        GetVoucherCodeXml = "U8.SA.USSASERVER.modvoucher.01186"     'zh-CN：初始化单据号码失败！
        Set objBillNo = Nothing
        Exit Function
    End If
    strXml = objBillNo.GetBillFormat()
    domCodeXml.loadXML strXml
End Function

