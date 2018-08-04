VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSaRefer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim clsCtlRefer As New U8RefService.IService  'New UFReferC.UFReferClient
Dim domFieldConfig As New DOMDocument
Dim domReferFillConfig As New DOMDocument
Dim clsAuth As New SaVoucherService.clsSAAuth
Dim domEnumValue As New DOMDocument
'Dim clsRefInfo As clsPublicInfo

Private Sub Class_Terminate()
    Set clsCtlRefer = Nothing
    Set domFieldConfig = Nothing
    Set domReferFillConfig = Nothing
    Set clsAuth = Nothing
    Set domEnumValue = Nothing
End Sub
Public Function Init(strCardNumber As String, strErrorResId) As Boolean
'    Set clsRefInfo = clsInfo
    Dim rst As New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open "select * from sa_voucherfieldconfig where cardnumber=N'" + strCardNumber + "' order by cardsection", DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rst.Save domFieldConfig, adPersistXML
    rst.Close
    rst.Open "select cardnumber, fieldname, cardsection, refername, sourcefieldname, breferfill,bcellcheckfill, destfieldname, bclear,isnull(bcheckchange,0) as bcheckchange,changemsg  from SA_ReferFillConfig where cardnumber=N'" + strCardNumber + "' order by cardsection", DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rst.Save domReferFillConfig, adPersistXML
    rst.Close
    Set rst = Nothing
'    clsCtlRefer.SetLogin m_Login       ',isnull(ifillother,0)
    Set clsAuth.clsAuth = m_Login.RowAuth
    clsAuth.Init DBConn, m_Login.cUserId, "refer", strCardNumber
    domEnumValue.loadXML "<Data />"
End Function
Private Function FillVoucherItems(voucher As ctlVoucher, strReferFieldName As String, strReferName As String, strsql As String, strCardSection As String, Optional lngRow As Long) As Boolean
    Dim rst As New ADODB.Recordset
    rst.CursorLocation = adUseClient
    Dim lst As IXMLDOMNodeList
    Dim nod As IXMLDOMNode
    Dim strSelect As String
    Dim strSourceFieldName As String
    Dim strDesFldName As String
    Dim strAuth As String
    Dim varFilter As Variant
    Dim strFilter As String
    Dim strValue1 As String
    Dim strValue2 As String
    Dim i As Long
    Dim lstChange As IXMLDOMNodeList
    Dim eleChange As IXMLDOMElement
    Dim strMsg As String
    Dim blnCheck As Boolean
    
    strAuth = clsAuth.GetAuthStringByOperation(strReferName)
    If strAuth <> "" Then
        strsql = strsql & " and (" & strAuth & ")"
    End If
    strSelect = strsql
    strSelect = ReplaceVoucherItems(strSelect, voucher, lngRow)
    rst.Open ConvertSQLString(strSelect), DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    FillVoucherItems = True
    If Not rst.EOF Then
        If rst.RecordCount > 1 Then
            Set nod = domFieldConfig.selectSingleNode("//z:row[@cardsection='" + UCase(strCardSection) + "' and @fieldname='" + LCase(strReferFieldName) + "']")
            If Not nod.Attributes.getNamedItem("filterfields") Is Nothing Then
                strFilter = nod.Attributes.getNamedItem("filterfields").nodeValue
                varFilter = Split(strFilter, ",")
                For i = 0 To UBound(varFilter)
                    If strCardSection = "T" Then
                        rst.Filter = varFilter(i) & "='" + voucher.headerText(strReferFieldName) + "'"
                    Else
                        rst.Filter = varFilter(i) & "='" + voucher.bodyText(lngRow, strReferFieldName) + "'"
                    End If
                    If rst.RecordCount = 0 Then rst.Filter = ""
                    If rst.RecordCount = 1 Then Exit For
                Next
            End If
        End If
        Set lstChange = domReferFillConfig.selectNodes("//z:row[@cardsection='" + UCase(strCardSection) + "' and @fieldname='" + LCase(strReferFieldName) + "' and @refername='" + LCase(strReferName) + "' and @bcellcheckfill='True' and @bcheckchange='1']")
        Set lst = domReferFillConfig.selectNodes("//z:row[@cardsection='" + UCase(strCardSection) + "' and @fieldname='" + LCase(strReferFieldName) + "' and @refername='" + LCase(strReferName) + "' and @bcellcheckfill='True']")
        Do While Not rst.EOF
            strMsg = ""
            For Each eleChange In lstChange
                strSourceFieldName = eleChange.Attributes.getNamedItem("sourcefieldname").Text
                strDesFldName = eleChange.Attributes.getNamedItem("destfieldname").Text
                If IsNull(rst.Fields(strSourceFieldName)) Then
                    strValue1 = ""
                Else
                    strValue1 = rst.Fields(strSourceFieldName)
                End If
                If strCardSection = "T" Then
                    strValue2 = voucher.headerText(strDesFldName)
                Else
                    strValue2 = voucher.bodyText(lngRow, strDesFldName)
                End If
                If strValue2 <> "" And strValue1 <> strValue2 Then
                    If Not eleChange.Attributes.getNamedItem("changemsg") Is Nothing Then
                        strMsg = strMsg & IIf(strMsg = "", "", Chr(13)) & GetString(eleChange.Attributes.getNamedItem("changemsg").Text)
                    End If
                End If
            Next
            If strMsg <> "" Then
                If MsgBox(strMsg, vbYesNo + vbQuestion) = vbNo Then 'zh-CN：录入的业务员不属于录入的部门，是否继续？
                    ClearReferItems voucher, strCardSection, strReferFieldName, strReferName, lngRow
                    Exit Function
                End If
            End If
            For Each nod In lst
                strSourceFieldName = nod.Attributes.getNamedItem("sourcefieldname").Text
                If strSourceFieldName = "''" Then
                    SetVoucherItemValue voucher, strCardSection, nod.Attributes.getNamedItem("destfieldname").Text, "", lngRow
                Else
                    blnCheck = IIf(nod.Attributes.getNamedItem("bcheckchange").Text = "0", False, True)
                    If Not blnCheck Then
                        SetVoucherItemValue voucher, strCardSection, nod.Attributes.getNamedItem("destfieldname").Text, IIf(IsNull(rst.Fields(strSourceFieldName)), "", rst.Fields(strSourceFieldName)), lngRow
                    Else
                        strDesFldName = nod.Attributes.getNamedItem("destfieldname").Text
                        If strCardSection = "T" Then
                            strValue2 = voucher.headerText(strDesFldName)
                        Else
                            strValue2 = voucher.bodyText(lngRow, strDesFldName)
                        End If
                        If strValue2 = "" Then
                            SetVoucherItemValue voucher, strCardSection, nod.Attributes.getNamedItem("destfieldname").Text, IIf(IsNull(rst.Fields(strSourceFieldName)), "", rst.Fields(strSourceFieldName)), lngRow
                        End If
                    End If
                End If
            Next
            rst.MoveNext
            If Not rst.EOF Then
                voucher.AddLine
                lngRow = voucher.BodyRows
            End If
        Loop
    Else
        ClearReferItems voucher, strCardSection, strReferFieldName, strReferName, lngRow
        FillVoucherItems = False
    End If
    rst.Close
    Set rst = Nothing
End Function
Public Function ClearReferItems(voucher As ctlVoucher, strCardSection As String, strReferFieldName As String, strReferName As String, Optional lngRow As Long) As Boolean
    Dim lst As IXMLDOMNodeList
    Dim nod As IXMLDOMNode
    
    Set lst = domReferFillConfig.selectNodes("//z:row[@cardsection='" + UCase(strCardSection) + "' and @fieldname='" + LCase(strReferFieldName) + "' and @refername='" + LCase(strReferName) + "' and @bclear='True']")
    For Each nod In lst
        SetVoucherItemValue voucher, strCardSection, nod.Attributes.getNamedItem("destfieldname").Text, "", lngRow
    Next
'    Set lst = domReferFillConfig.selectNodes("//z:row[@cardsection='" + UCase(strCardSection) + "' and @fieldname='" + LCase(strReferFieldName) + "' and @refername='" + LCase(strReferName) + "' and @bclear='True' and @ifillother!='0']")
'    For Each nod In lst
'        ClearCellCheck Voucher, nod.Attributes.getNamedItem("cardnumber").Text, strCardSection, nod.Attributes.getNamedItem("destfieldname").Text, lngRow
'    Next
End Function
Public Function CellCheck(strReferString As String, voucher As ctlVoucher, strCardSection As String, strFieldName As String, Optional lngRow As Long, Optional strwhere As String) As String
    Dim varRefer  As Variant
    Dim lngRowOld As Long
    Dim lngRowMax As Long
    Dim i As Long
    Dim nod As IXMLDOMNode
    Dim strsql As String
    Dim strValue As String
    Dim blnOk As Boolean
    Dim strError As String
    Dim strReferName As String
    Dim lst As IXMLDOMNodeList
    Dim clsComp As New UsSaCompStr.clsCompStr
    If strReferString <> "" Then
        varRefer = Split(strReferString, ",")
        strReferString = ""
        If varRefer(0) = strCardSection And varRefer(1) = strFieldName Then Exit Function
    End If
    Dim strReferType As String
    
    Set nod = domFieldConfig.selectSingleNode("//z:row[@cardsection='" + UCase(strCardSection) + "' and @fieldname='" + LCase(strFieldName) + "']")
    If nod Is Nothing Then Exit Function
    If nod.Attributes.getNamedItem("refertype") Is Nothing Then Exit Function
    strReferType = nod.Attributes.getNamedItem("refertype").Text
    Select Case strReferType
        Case "4"
            strValue = GetVoucherItemValue(voucher, strCardSection, strFieldName, lngRow)
            Set nod = domEnumValue.documentElement.selectSingleNode("row[@name='" + LCase(strFieldName) + "']")
            If nod Is Nothing Then
                strError = "false"
            Else
                Dim ele As IXMLDOMElement
                Set ele = nod.selectSingleNode("value[@code='" + ReplaceSpecialCode(strValue) + "']")
                If ele Is Nothing Then
                    strError = "false"
                End If
            End If
        Case "1", "itemclass"
            Set nod = domFieldConfig.selectSingleNode("//z:row[@cardsection='" + UCase(strCardSection) + "' and @fieldname='" + LCase(strFieldName) + "']")
            If Not nod Is Nothing Then
                strReferName = nod.Attributes.getNamedItem("refername").Text
                strValue = GetVoucherItemValue(voucher, strCardSection, strFieldName, lngRow)
                If strValue = "" Then
                    ClearReferItems voucher, strCardSection, strFieldName, strReferName, lngRow
                Else
                    strsql = nod.Attributes.getNamedItem("cellchecksql").Text
                    blnOk = FillVoucherItems(voucher, strFieldName, strReferName, strsql, strCardSection, lngRow)
                    If Not blnOk Then
                        If Not nod.Attributes.getNamedItem("errresid") Is Nothing Then
                            CellCheck = nod.Attributes.getNamedItem("errresid").Text
                            strError = nod.Attributes.getNamedItem("errresid").Text
                            MsgBox GetErrorString(strError, voucher, strFieldName, strCardSection)
                        End If
'                    Else
'                        Set lst = domReferFillConfig.selectNodes("//z:row[@cardsection='" + UCase(strCardSection) + "' and @fieldname='" + LCase(strFieldName) + "' and @refername='" + LCase(strReferName) + "' and @ifillother!='0']")
'                        For Each nod In lst
'                            ReferCellCheck Voucher, nod.Attributes.getNamedItem("cardnumber").Text, strCardSection, nod.Attributes.getNamedItem("destfieldname").Text, lngRow
'                        Next
                    End If
                End If
            End If
        Case "free", "define"
            CellCheck = CellCheckDefine(voucher, voucher.LookUpArray(strFieldName, IIf(strCardSection = "T", siheader, sibody)), IIf(strCardSection = "T", siheader, sibody), GetVoucherItemValue(voucher, strCardSection, strFieldName, lngRow))
            If strCardSection = "T" Then
                voucher.headerText(strFieldName) = CellCheck
            Else
                voucher.bodyText(lngRow, strFieldName) = CellCheck
            End If
        Case "item"
            Dim rst As New ADODB.Recordset
            rst.CursorLocation = adUseClient
            strValue = GetVoucherItemValue(voucher, strCardSection, strFieldName, lngRow)
            If strValue = "" Then
                If strCardSection = "T" Then
                    voucher.headerText("citemcode") = ""
                    voucher.headerText("citemname") = ""
                Else
                    voucher.bodyText(lngRow, "citemcode") = ""
                    voucher.bodyText(lngRow, "citemname") = ""
                End If
            Else
                If strCardSection = "T" Then
                Else
                    If voucher.bodyText(lngRow, "citem_class") = "" Then
                        MsgBox GetString("U8.SA.USSASERVER.clscommcheck.00450") 'zh-CN：请先填写项目大类
                        CellCheck = "cancel"
                        Exit Function
                    End If
                End If
                '//////朱培碧
                strsql = "select ctable,citem_name from fitem where citem_class=N'" & voucher.bodyText(lngRow, "citem_class") & "'"
                rst.Open strsql, DBConn
                If Not rst.EOF Then
                    If rst(0) = "inventory" Then
                        strsql = "select cInvCode as citemcode,cInvName as citemname from inventory where (cInvCode=N'" & strValue & "' or cInvName=N'" & strValue & "' ) "
                    Else
                        voucher.bodyText(lngRow, "citem_name") = rst.Fields("citem_name")
                        strsql = "select citemcode,citemname,bclose from " & rst(0) & " where (citemcode=N'" & strValue & "' or citemname=N'" & strValue & "' ) and bclose =0"
                    End If
                    rst.Close
                    rst.Open strsql, DBConn
                    If Not rst.EOF Then
                        voucher.bodyText(lngRow, "citemcode") = rst.Fields("citemcode")
                        voucher.bodyText(lngRow, "citemname") = rst.Fields("citemname")
                    Else
                        voucher.bodyText(lngRow, "citemcode") = ""
                        voucher.bodyText(lngRow, "citemname") = ""
                        strError = GetString("U8.SA.USSASERVER.clscommcheck.00452") 'zh-CN：项目大类有错误,请检查!
                        MsgBox strError
                    End If
                Else
                    strError = GetString("U8.SA.USSASERVER.clscommcheck.00451") 'zh-CN：项目大类有错误,请检查!
                    voucher.bodyText(lngRow, "citemcode") = ""
                    voucher.bodyText(lngRow, "citemname") = ""
                    MsgBox strError
                End If
                rst.Close
            End If
            Set rst = Nothing
    End Select
    CellCheck = strError
End Function
Private Function GetFieldName(voucher As ctlVoucher, strFieldName As String, strCardSection As String) As String
    If strCardSection = "T" Then
        GetFieldName = voucher.ItemCaption(strFieldName, siheader)
    Else
        GetFieldName = voucher.ItemCaption(strFieldName, sibody)
    End If
End Function
Private Function GetErrorString(strSource As String, voucher As ctlVoucher, strFieldName As String, strCardSection As String) As String
    Dim varError As Variant
    If InStr(strSource, ",") = 0 Then
        GetErrorString = GetString(strSource)
    Else
        varError = Split(strSource, ",")
        GetErrorString = GetStringPara(varError(0), GetFieldName(voucher, strFieldName, strCardSection))
    End If
End Function
''自定义项目检查，调用u8defpro
Private Function CellCheckDefine(voucher As ctlVoucher, Index As Variant, iVoucherSec As Integer, KeyValue As String) As String
    Dim clsDef As U8DefPro.clsDefPro
    Dim nDataSource As Long         '数据来源
    Dim nEnterType As Long         '输入方式
    Dim sDataRule As String       '数据公式
    Dim bValidityCheck As Boolean      '是否合法性检测
    Dim bBuildArchives As Boolean      '是否建档
    Dim sVouchType As String
    Dim sTableName As String, sFieldName As String, sCardNumber As String
    Dim sDefWhere As String
    Dim intReturn As Integer
    Dim bFixlenth As Boolean
    Dim lngFixLenth As Long
    Dim strTmp As String
    Dim strDataType As String
    Dim sKeyName As String      ''cellcheck的项目名称
    Dim bFree As Boolean        ''是否自由项
    Dim bBatch As Boolean
    If KeyValue = "" Then
        CellCheckDefine = ""
        Exit Function
    End If
    Set clsDef = New U8DefPro.clsDefPro
    If Not clsDef.Init(False, DBConn.ConnectionString, m_Login.cUserId) Then
        CellCheckDefine = ""
        MsgBox GetString("U8.SA.xsglsql.01.frmbillvouch.00147") 'zh-CN：初始化自定义项组件失败！
        Set clsDef = Nothing
        Exit Function
    End If
    With voucher
        sKeyName = LCase(.ItemState(Index, iVoucherSec).sFieldName)
        If Left(sKeyName, 5) = "cfree" Then
            bFree = True
        Else
            bFree = False
        End If
        If Left(sKeyName, 14) = "cbatchproperty" Then
            bBatch = True
        Else
            bBatch = False
        End If
        If bFree = False Then
            bValidityCheck = .ItemState(Index, iVoucherSec).bValidityCheck
            nDataSource = .ItemState(Index, iVoucherSec).nDataSource
            nEnterType = .ItemState(Index, iVoucherSec).nEnterType
            sDataRule = .ItemState(Index, iVoucherSec).sDataRule
            strTmp = .ItemState(Index, iVoucherSec).sDefaultValue
            bFixlenth = IIf(Left(strTmp, 1) = "1", True, False)
            lngFixLenth = val(Mid(strTmp, 3))
            strDataType = .ItemState(Index, iVoucherSec).nFieldType
            ''是否建立档
            bBuildArchives = .ItemState(Index, iVoucherSec).bBuildArchives
            If bValidityCheck = False Then
                CellCheckDefine = KeyValue
                If strDataType = "3" Then         '
                    If Abs(val(CellCheckDefine)) > 2147483647 Then
                        MsgBox GetStringPara("U8.SA.xsglsql.01.frmbillvouch.00148", CellCheckDefine) 'Para zh-CN：当前项目设置为整型，输入的值{0}的输入超出可取值范围（-2,147,483,647到2,147,483,647）!
                        CellCheckDefine = ""
                    End If
                End If
                If CellCheckDefine <> "" And bFixlenth Then
                    If GetStrTrueLenth(CellCheckDefine) <> lngFixLenth Then
                        CellCheckDefine = ""
                        MsgBox GetStringPara("U8.SA.xsglsql.01.frmbillvouch.00150", lngFixLenth)  'Para zh-CN：项目为定长{0},长度不符合！
                    End If
                End If
                If .ItemState(Index, iVoucherSec).nDataSource = 0 Then
                    Call clsDef.UsrDefCodeToValue(iVoucherSec, .ItemState(Index, iVoucherSec).sFieldName, KeyValue)
                    CellCheckDefine = KeyValue
                End If
                Set clsDef = Nothing
                Exit Function
            End If
            ''定长时如果长度不等于定长不建档
            If bFixlenth Then
                If GetStrTrueLenth(KeyValue) <> lngFixLenth Then
                    bBuildArchives = False
                End If
            End If

            Select Case nDataSource  '0表示手工输入；1表示档案；2表示单据
                Case 0
                    sTableName = "UserDefine"
                    sFieldName = "cValue"
                    sVouchType = ""
                Case 1
                    sTableName = Left(sDataRule, InStr(1, sDataRule, ",") - 1)
                    sFieldName = Mid(sDataRule, InStr(1, sDataRule, ",") + 1)
                    sVouchType = ""
                Case 2
                    sCardNumber = Left(sDataRule, InStr(1, sDataRule, ",") - 1)
                    sFieldName = Mid(sDataRule, InStr(1, sDataRule, ",") + 1)
            End Select
        Else
            ''自由项
            bBuildArchives = .ItemState(Index, iVoucherSec).bBuildArchives
        End If


        If bFree = False Then
            intReturn = clsDef.ValidateAr(nDataSource, iVoucherSec, .ItemState(Index, iVoucherSec).sFieldName, sTableName, sFieldName, KeyValue, sCardNumber, "", bBuildArchives)
        Else
            intReturn = clsDef.ValidateFreeArEx(.bodyText(.row, "cinvcode"), "<Data />", .ItemState(Index, iVoucherSec).sFieldName, KeyValue, bBuildArchives)
        End If
        Select Case intReturn
            Case 0  '合法性检查成功
                CellCheckDefine = KeyValue
            Case 1  '合法性检查失败，建档成功；
                CellCheckDefine = KeyValue
                'MsgBox "合法性检查失败，建档成功"
            Case -2 '-2表示合法性检查失败，建档失败
                If bValidityCheck = True Then
                    CellCheckDefine = ""
                    MsgBox GetString("U8.SA.xsglsql.01.frmbillvouch.00152") 'zh-CN：合法性检查失败，建档失败
                Else
                    If KeyValue <> "" Then CellCheckDefine = KeyValue
                End If
            Case -1, -3 '-1表示合法性检查失败；
                If bValidityCheck = True Or bFree Or bBatch Then
                    CellCheckDefine = ""
                    'MsgBox "合法性检查失败"
                    MsgBox GetString("U8.SA.xsglsql.01.frmbillvouch.00153") 'zh-CN：录入不合法，请检查
                Else
                    If KeyValue <> "" Then CellCheckDefine = KeyValue
                End If
        End Select
        If CellCheckDefine <> "" And bFixlenth Then
            If GetStrTrueLenth(CellCheckDefine) <> lngFixLenth Then
                CellCheckDefine = ""
                MsgBox GetStringPara("U8.SA.xsglsql.01.frmbillvouch.00150", lngFixLenth)  'Para zh-CN：项目为定长{0},长度不符合！
            End If
        End If
        If strDataType = "3" Then
                If Abs(val(CellCheckDefine)) > 2147483647 Then
                    MsgBox GetStringPara("U8.SA.xsglsql.01.frmbillvouch.00148", CellCheckDefine) 'Para zh-CN：当前项目设置为整型，输入的值{0}的输入超出可取值范围（-2,147,483,647到2,147,483,647）!
                    CellCheckDefine = ""
                End If
            End If
    End With
    Set clsDef = Nothing

End Function

Public Function ShowReferCtl(clsVoucher As clsSaVoucher, voucher As ctlVoucher, iVoucherSec As Integer, iIndex As Long, referPara As UapVoucherControl85.ReferParameter, Optional lngRow As Long, Optional strwhere As String) As String
    Dim nod As IXMLDOMNode
    Dim strReferType As String
    Dim strCardSection As String
    Dim strFieldName As String
    Dim skey As String
    Dim clsRefer As New UFReferC.UFReferClient
    
    strCardSection = IIf(iVoucherSec = siheader, "T", "B")
    strFieldName = voucher.ItemState(iIndex, iVoucherSec).sFieldName
    Set nod = domFieldConfig.selectSingleNode("//z:row[@cardsection='" + UCase(strCardSection) + "' and @fieldname='" + LCase(strFieldName) + "']")
    If nod Is Nothing Then
        referPara.Cancel = True
        Exit Function
    End If
    If nod.Attributes.getNamedItem("refertype") Is Nothing Then
        referPara.Cancel = True
        Exit Function
    End If
    strReferType = nod.Attributes.getNamedItem("refertype").Text
    Select Case strReferType
        Case "1"        '枚举
            ShowReferCtl = ShowEnumReferCtrl(clsVoucher, voucher, strCardSection, strFieldName, referPara, lngRow, strwhere)
        Case "free"        '自由项
            If voucher.headerText("cinvcode") = "" Then
               MsgBox "请先选定图书", vbOKOnly, "提示"
            Else
            referPara.Cancel = True
            Dim clsDef As U8DefPro.clsDefPro
            Set clsDef = New U8DefPro.clsDefPro
            If Not clsDef.Init(False, DBConn, m_Login.cUserId) Then Exit Function
            If Not clsDef.InitNew(m_Login, False) Then Exit Function
            Dim strValue As String
            
            strValue = clsDef.GetStruFreeRefVal(voucher.headerText("cinvcode"), strFieldName, voucher.headerText(strFieldName), False, 100000, 1)
            If strValue <> "" Then
                voucher.headerText(strFieldName) = strValue
            End If
            Set clsDef = Nothing
            If strCardSection = "B" Then
                ShowReferCtl = IIf(ShowReferCtl = "", "", ",") + CStr(lngRow)
            End If
            ShowReferCtl = strCardSection + "," + strFieldName + IIf(ShowReferCtl = "", "", ",") + ShowReferCtl
            End If
      Case "editionnum"
            skey = "cfree1"
            If voucher.headerText("cinvcode") = "" Then
               MsgBox "请先选定图书", vbOKOnly, "提示"
            Else
            referPara.Cancel = True
            Set clsDef = New U8DefPro.clsDefPro
            If Not clsDef.Init(False, DBConn, m_Login.cUserId) Then Exit Function
            strValue = clsDef.GetStruFreeRefVal(voucher.headerText("cinvcode"), skey, voucher.headerText(strFieldName), False, 100000, 1)
            If strValue <> "" Then
                voucher.headerText(strFieldName) = strValue
            End If
            Set clsDef = Nothing
            If strCardSection = "B" Then
                ShowReferCtl = IIf(ShowReferCtl = "", "", ",") + CStr(lngRow)
            End If
            ShowReferCtl = strCardSection + "," + strFieldName + IIf(ShowReferCtl = "", "", ",") + ShowReferCtl
            End If
       Case "printnum"
           skey = "cfree2"
            If voucher.headerText("cinvcode") = "" Then
               MsgBox "请先选定图书", vbOKOnly, "提示"
            Else
            referPara.Cancel = True
            Set clsDef = New U8DefPro.clsDefPro
            If Not clsDef.Init(False, DBConn, m_Login.cUserId) Then Exit Function
            strValue = clsDef.GetStruFreeRefVal(voucher.headerText("cinvcode"), skey, voucher.headerText(strFieldName), False, 100000, 1)
            If strValue <> "" Then
                voucher.headerText(strFieldName) = strValue
            End If
            Set clsDef = Nothing
            If strCardSection = "B" Then
                ShowReferCtl = IIf(ShowReferCtl = "", "", ",") + CStr(lngRow)
            End If
            ShowReferCtl = strCardSection + "," + strFieldName + IIf(ShowReferCtl = "", "", ",") + ShowReferCtl
            End If
        Case "define"
            referPara.Cancel = True
            RefDefine voucher, iIndex, iVoucherSec
        Case "itemclass"
            Dim strsql As String
            referPara.Cancel = True
            strsql = "select citem_class ,citem_name from fitem"
            If clsRefer.StrRefInit(m_Login, False, "", strsql, GetString("U8.SA.xsglsql.01.frmbillvouch.00125") + "," + GetString("U8.SA.xsglsql.01.frmbillvouch.00125.01"), "") = False Then Exit Function 'zh-CN：项目大类编码 'zh-CN：项目大类名称
            clsRefer.Show
            
            If Not clsRefer.recmx Is Nothing Then
                If Not clsRefer.recmx.EOF Then
                    voucher.bodyText(voucher.row, "citem_class") = clsRefer.recmx("citem_class")
                    voucher.bodyText(voucher.row, "citem_cname") = clsRefer.recmx("citem_name")
    '                sRet = clsRefer.recmx("citem_class")
                End If
            End If
        Case "item"
            referPara.Cancel = True
            strValue = voucher.bodyText(voucher.row, strFieldName)
            If voucher.bodyText(voucher.row, "citem_class") = "" Then
                clsRefer.EnumRefInit m_Login, enuTreeViewAndGrid, False, enuItem
                If voucher.bodyText(voucher.row, "citem_class") <> "ch" Then
                    clsRefer.SetReferFilterString "bclose=0"
                End If
            Else
                clsRefer.ItemRefInit m_Login, False, voucher.bodyText(voucher.row, "citem_class"), IIf(voucher.bodyText(voucher.row, "citem_class") <> "ch", "bclose=0", "")
            End If
            clsRefer.FillText = strValue
            'clsRefer.SetReferFilterString getReferString(sKey, sKeyValue)
            clsRefer.Show
            If Not clsRefer.recmx Is Nothing Then
                If voucher.bodyText(lngRow, "citem_class") = "" Then
                    If Not clsRefer.RstSelClass.EOF Then
                        voucher.bodyText(voucher.row, "citem_class") = clsRefer.RstSelClass("citem_class")
                        voucher.bodyText(voucher.row, "citem_cname") = clsRefer.RstSelClass("citem_name")
                    End If
                End If
                If LCase(voucher.bodyText(voucher.row, "citem_class")) <> "ch" Then
                    voucher.bodyText(lngRow, "citemcode") = clsRefer.recmx("citemcode")
                    voucher.bodyText(lngRow, "citemname") = clsRefer.recmx("citemname")
    '                sRet = clsRefer.recmx("citemcode")
                Else
                    voucher.bodyText(lngRow, "citemcode") = clsRefer.recmx("cinvcode")
                    voucher.bodyText(lngRow, "citemname") = clsRefer.recmx("cinvname")
    '                sRet = clsRefer.recmx("cinvcode")
                End If
            End If
    End Select
    Set clsRefer = Nothing
End Function
Private Function RefDefine(voucher As ctlVoucher, Index As Variant, iVoucherSec As Integer) As String
    Dim clsDef As U8DefPro.clsDefPro
    Dim nDataSource As Long         '数据来源
    Dim nEnterType As Long         '输入方式
    Dim sDataRule As String       '数据公式
    Dim bValidityCheck As Boolean      '是否合法性检测
    Dim bBuildArchives As Boolean      '是否建档
    Dim sVouchType As String
    Dim sTableName As String, sFieldName As String, sCardNumber As String
    Dim sDefWhere As String
    Dim strKeyValue As String

    Set clsDef = New U8DefPro.clsDefPro

    With voucher
        If iVoucherSec = siheader Then
            strKeyValue = .headerText(Index)
        Else
            strKeyValue = .bodyText(.row, Index)
        End If
        nDataSource = .ItemState(Index, iVoucherSec).nDataSource
        nEnterType = .ItemState(Index, iVoucherSec).nEnterType
        sDataRule = .ItemState(Index, iVoucherSec).sDataRule
        bValidityCheck = .ItemState(Index, iVoucherSec).bValidityCheck
        bBuildArchives = .ItemState(Index, iVoucherSec).bBuildArchives
        Select Case nDataSource  '0表示手工输入；1表示档案；2表示单据
            Case 0
                sTableName = "UserDefine"
                sFieldName = "cValue"
                sVouchType = ""
            Case 1
                sTableName = Left(sDataRule, InStr(1, sDataRule, ",") - 1)
                sFieldName = Mid(sDataRule, InStr(1, sDataRule, ",") + 1)
                sVouchType = ""
            Case 2
                sCardNumber = Left(sDataRule, InStr(1, sDataRule, ",") - 1)
                sFieldName = Mid(sDataRule, InStr(1, sDataRule, ",") + 1)
        End Select
        If Not clsDef.Init(False, DBConn.ConnectionString, m_Login.cUserId) Then
            RefDefine = ""
            MsgBox GetString("U8.SA.xsglsql.01.frmbillvouch.00146") 'zh-CN：初始化自定义项组件失败！
            Exit Function
        End If
        RefDefine = clsDef.GetRefVal(nDataSource, iVoucherSec, .ItemState(Index, iVoucherSec).sFieldName, sTableName, sFieldName, sCardNumber, strKeyValue, False, 40, 1)
        If iVoucherSec = siheader Then
            .headerText(Index) = RefDefine
        Else
            .bodyText(.row, Index) = RefDefine
        End If
    End With
    Set clsDef = Nothing
End Function

Public Sub FillList(voucher As ctlVoucher, lngRow As Long, lngCol As Long, pCom As Object)
    Dim strFieldName As String
    Dim nod As IXMLDOMNode
    Dim strsql As String
    Dim strReferName As String
    Dim strFilter As String
    Dim rst As New ADODB.Recordset
    Dim ele As IXMLDOMElement
    Dim eleValue As IXMLDOMElement
    rst.CursorLocation = adUseClient
    strFieldName = LCase(voucher.ItemState(lngCol, sibody).sFieldName)
    Set nod = domFieldConfig.selectSingleNode("//z:row[@cardsection='B' and @fieldname='" + LCase(strFieldName) + "']")
    If Not nod.Attributes.getNamedItem("refertype") Is Nothing Then
        If nod.Attributes.getNamedItem("refertype").Text = "4" Then
            strReferName = nod.Attributes.getNamedItem("refername").Text
            If Not nod.Attributes.getNamedItem("referdefaultfilter") Is Nothing Then
                strFilter = nod.Attributes.getNamedItem("referdefaultfilter").Text
            End If
            strsql = "select enumcode from aa_enum where enumtype=N'" + strReferName + "' and localeid=N'" + m_Login.LanguageRegion + "'" + IIf(strFilter = "", "", " and (" + strFilter + ")")
            rst.Open ConvertSQLString(strsql), DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
            Set ele = domEnumValue.documentElement.selectSingleNode("row[@name='" + strFieldName + "']")
            If ele Is Nothing Then
                Set ele = domEnumValue.createElement("row")
                ele.setAttribute "name", strFieldName
                domEnumValue.documentElement.appendChild ele
            End If
            Do While Not rst.EOF
                Set eleValue = ele.selectSingleNode("value[@code='" + rst(0) + "']")
                If eleValue Is Nothing Then
                    Set eleValue = domEnumValue.createElement("value")
                    eleValue.setAttribute "code", rst(0)
                    ele.appendChild eleValue
                End If
                pCom.AddItem rst(0)
                rst.MoveNext
            Loop
            rst.Close
        End If
    End If
    Set rst = Nothing
End Sub
Public Function GetEnumDom(strFieldName As String) As DOMDocument
'    Dim strFieldName As String
    Dim nod As IXMLDOMNode
    Dim strsql As String
    Dim strReferName As String
    Dim strFilter As String
    Dim rst As New ADODB.Recordset
    Dim dom As New DOMDocument
    rst.CursorLocation = adUseClient
'    strFieldName = LCase(voucher.ItemState(lngCol, sibody).sFieldName)
    Set nod = domFieldConfig.selectSingleNode("//z:row[@cardsection='B' and @fieldname='" + LCase(strFieldName) + "']")
    If Not nod.Attributes.getNamedItem("refertype") Is Nothing Then
        If nod.Attributes.getNamedItem("refertype").Text = "4" Then
            strReferName = nod.Attributes.getNamedItem("refername").Text
            If Not nod.Attributes.getNamedItem("referdefaultfilter") Is Nothing Then
                strFilter = nod.Attributes.getNamedItem("referdefaultfilter").Text
            End If
            strsql = "select enumcode,enumname,enumindex from aa_enum where enumtype='" + strReferName + "' and localeid='" + m_Login.LanguageRegion + "'" + IIf(strFilter = "", "", " and (" + strFilter + ")")
            rst.Open ConvertSQLString(strsql), DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
            rst.Save dom, adPersistXML
            rst.Close
        End If
    End If
    Set GetEnumDom = dom
    Set dom = Nothing
    Set rst = Nothing
End Function
'Private Sub SetReferAuth(strReferName As String, strCardSection As String)
'    Dim strAuthName As String
'    Dim strAuthType As String
'
'    clsAuth.GetEnumReferAuth strReferName, strCardSection, strAuthName, strAuthType
'    If strAuthName = "" Then
'        clsCtlRefer.SetRWAuth "customer", "R", False
'    Else
'        clsCtlRefer.SetRWAuth strAuthName, strAuthType, True
'    End If
'End Sub
Private Function ShowEnumReferCtrl(clsVoucher As clsSaVoucher, voucher As ctlVoucher, strCardSection As String, strFieldName As String, referPara As UapVoucherControl85.ReferParameter, Optional lngRow As Long, Optional strwhere As String) As String
    Dim strFilter As String
    Dim strDefFilter As String
    Dim nod As IXMLDOMNode
    Dim strReferName As String
    Dim strAuth As String
    Dim rstClass As ADODB.Recordset
    Dim rstGrid As ADODB.Recordset
    Dim strError As String
    
    ShowEnumReferCtrl = ""
    Set nod = domFieldConfig.selectSingleNode("//z:row[@cardsection='" + UCase(strCardSection) + "' and @fieldname='" + LCase(strFieldName) + "']")
    If Not nod Is Nothing Then
        If Not nod.Attributes.getNamedItem("referdefaultfilter") Is Nothing Then
            strDefFilter = ReplaceVoucherItems(nod.Attributes.getNamedItem("referdefaultfilter").Text, voucher, lngRow)
        End If
        If Not nod.Attributes.getNamedItem("referfilterfields") Is Nothing Then
            strFilter = GetFilterString(voucher, Trim(nod.Attributes.getNamedItem("referfilterfields").Text), lngRow)
        End If
        If strDefFilter = "" Then
            strFilter = strFilter
        Else
            strFilter = "(" + strDefFilter + ")" + IIf(strFilter = "", "", " and (" + strFilter + ")")
        End If
        If strFilter <> "" Then
            strwhere = strFilter + IIf(strwhere = "", "", " and (" + strwhere + ")")
        End If
        strReferName = nod.Attributes.getNamedItem("refername").Text
        
'        clsCtlRefer.SetRWAuth "customer", "R", False
        strAuth = clsAuth.GetAuthStringByOperation(strReferName)
        If strAuth <> "" Then
            strwhere = strAuth + IIf(strwhere = "", "", " and (" + strwhere + ")")
        End If
        referPara.ID = strReferName
        Dim strMuti As String
        Dim strShowStyle As String
        
        strMuti = nod.Attributes.getNamedItem("refermultilines").Text
        strShowStyle = "1"
        If Not nod.Attributes.getNamedItem("refershowstyle") Is Nothing Then
            strShowStyle = nod.Attributes.getNamedItem("refershowstyle").Text
        End If
        referPara.sSQL = ConvertSQLString(strwhere)
        referPara.AutoDisplayText = False
        referPara.RetField = ""
        referPara.ReferMetaXML = "<Ref><RefSet bControlledInvalidData='1' bAuth='0' bMultiSel= '" + IIf(CBool(strMuti), "1", "0") + "' iShowMode='" + nod.Attributes.getNamedItem("referdispmode").Text + "' iShowStyle='" + strShowStyle + "' /></Ref>" 'iShowMode
        If referPara.Cancel = True Then
            clsCtlRefer.RefID = strReferName
            clsCtlRefer.FillText = GetVoucherItemValue(voucher, strCardSection, strFieldName, lngRow)
            clsCtlRefer.FilterSQL = ConvertSQLString(strwhere)
            clsCtlRefer.MetaXML = "<Ref><RefSet bControlledInvalidData='1' bAuth='0' bMultiSel= '" + IIf(CBool(strMuti), "1", "0") + "' iShowMode='" + nod.Attributes.getNamedItem("referdispmode").Text + "' iShowStyle='" + strShowStyle + "' /></Ref>" 'iShowMode
'            strReferName = Replace(strReferName, "_aa", "")
'            If clsCtlRefer.EnumRefInit(m_Login, nod.Attributes.getNamedItem("referdispmode").Text, nod.Attributes.getNamedItem("refermultilines").Text, strReferName, ConvertSQLString(strWhere)) Then
            If clsCtlRefer.ShowRef(m_Login, rstClass, rstGrid, strError) Then
'                clsCtlRefer.Show
                If Not rstGrid Is Nothing Then
                    Set referPara.rstGrid = rstGrid
                    ShowEnumReferCtrl = FillItemsAfterBrowse(clsVoucher, voucher, strCardSection, strFieldName, rstGrid, lngRow)
                End If
            End If
        End If
    End If
End Function
Public Function FillItemsAfterBrowse(clsVoucher As clsSaVoucher, voucher As ctlVoucher, strCardSection As String, strFieldName As String, recResult As ADODB.Recordset, Optional lngRow As Long) As String
    Dim strsql As String
    Dim nod As IXMLDOMNode
    Dim strReferName As String
    Dim lst As IXMLDOMNodeList
    Dim subNod As IXMLDOMNode
    Dim lngOldRow As Long
    
    lngOldRow = lngRow
    Set nod = domFieldConfig.selectSingleNode("//z:row[@cardsection='" + UCase(strCardSection) + "' and @fieldname='" + LCase(strFieldName) + "']")
    If nod Is Nothing Then Exit Function
    If Not nod.Attributes.getNamedItem("fillselectsql") Is Nothing Then
        strsql = nod.Attributes.getNamedItem("fillselectsql").Text
    End If
    strReferName = nod.Attributes.getNamedItem("refername").Text
    Dim strValue As String
'    On Error Resume Next
'    strReferName = Replace(strReferName, "_aa", "")
    Do While Not recResult.EOF
        Set lst = domReferFillConfig.selectNodes("//z:row[@cardsection='" + UCase(strCardSection) + "' and @fieldname='" + LCase(strFieldName) + "' and @refername='" + LCase(strReferName) + "' and @breferfill='True']")
        For Each subNod In lst
            If bExistField(subNod.Attributes.getNamedItem("sourcefieldname").Text, recResult) Then
                If IsNull(recResult.Fields(subNod.Attributes.getNamedItem("sourcefieldname").Text)) Then
                    strValue = ""
                Else
                    strValue = recResult.Fields(subNod.Attributes.getNamedItem("sourcefieldname").Text)
                End If
            End If
            SetVoucherItemValue voucher, strCardSection, subNod.Attributes.getNamedItem("destfieldname").Text, strValue, lngRow
'            If strCardSection = "B" Then
'                Voucher.CallAutoFillBackEvent strFieldName, lngRow
'            End If
        Next
        If strsql <> "" Then
            FillVoucherItems voucher, strFieldName, strReferName, strsql, strCardSection, lngRow
        End If
        If strCardSection = "B" Then
            FillItemsAfterBrowse = IIf(FillItemsAfterBrowse = "", "", ",") + CStr(lngRow)
        End If
'        Set lst = domReferFillConfig.selectNodes("//z:row[@cardsection='" + UCase(strCardSection) + "' and @fieldname='" + LCase(strFieldName) + "' and @refername='" + LCase(strReferName) + "' and @ifillother!='0']")
'        For Each nod In lst
'            ReferCellCheck Voucher, nod.Attributes.getNamedItem("cardnumber").Text, strCardSection, nod.Attributes.getNamedItem("destfieldname").Text, lngRow
'        Next
        If strCardSection = "B" Then
            voucher.CallAutoFillBackEvent strFieldName, lngRow
        End If
        recResult.MoveNext
        If Not recResult.EOF Then
            voucher.row = lngOldRow
'            Dim strValue As String
            strValue = voucher.bodyText(lngRow, strFieldName)
            clsVoucher.CopyLine voucher, Me
            voucher.bodyText(lngRow, strFieldName) = strValue
            lngRow = voucher.BodyRows
        End If
    Loop
    FillItemsAfterBrowse = strCardSection + "," + strFieldName + IIf(FillItemsAfterBrowse = "", "", ",") + FillItemsAfterBrowse
End Function
'Private Sub ClearCellCheck(Voucher As ctlVoucher, strCardNum As String, strCardSection As String, strFieldName As String, Optional lngRow As Long)
'    Dim rst As New ADODB.Recordset
'    Dim strsql As String
'    Dim strReferName As String
'    rst.CursorLocation = adUseClient
'
'    rst.Open "select * from sa_referfillconfig where cardnumber='" + strCardNum + "' and cardsection='" + strCardSection + "' and fieldname='" + strFieldName + "' and bclear=1", DBConn, adOpenForwardOnly, adLockReadOnly
'    Do While Not rst.EOF
'        SetVoucherItemValue Voucher, strCardSection, rst.Fields("destfieldname"), "", lngRow
'        If strCardSection = "B" Then
'            Voucher.CallAutoFillBackEvent strFieldName, lngRow
'        End If
'        rst.MoveNext
'    Loop
'    rst.Close
'    Set rst = Nothing
'End Sub
Public Sub ReferCellCheck(voucher As ctlVoucher, strCardSection As String, strFieldName As String, Optional lngRow As Long)
'    Dim rst As New ADODB.Recordset
    Dim strsql As String
    Dim strReferName As String
    Dim nod As IXMLDOMElement
'    rst.CursorLocation = adUseClient
    
    Set nod = domFieldConfig.selectSingleNode("//z:row[@cardsection='" + UCase(strCardSection) + "' and @fieldname='" + LCase(strFieldName) + "']")
    If nod Is Nothing Then Exit Sub
    If Not nod.Attributes.getNamedItem("fillselectsql") Is Nothing Then
        strsql = nod.Attributes.getNamedItem("fillselectsql").Text
    End If
    strReferName = nod.Attributes.getNamedItem("refername").Text
'    rst.Open strsql, DBConn, adOpenForwardOnly, adLockReadOnly
'    If Not rst.EOF Then
    FillVoucherItems voucher, strFieldName, strReferName, strsql, strCardSection, lngRow
'    End If
'    rst.Close
'    Set rst = Nothing
End Sub
Private Function GetFilterString(voucher As ctlVoucher, strFilterFields As String, Optional lngRow As Long) As String
    Dim varFields As Variant
    Dim varMarks As Variant
    Dim varFirst As Variant
    Dim clsComp As New UsSaCompStr.clsCompStr
    Dim i As Integer
    Dim j As Integer
    Dim strSqlWhere As String
    Dim strValue As String
    strSqlWhere = ""
    If strFilterFields <> "" Then
        varFirst = Split(strFilterFields, "/")
        For j = 0 To UBound(varFirst)
            varFields = Split(varFirst(j), ";")
            strSqlWhere = ""
            varFields(0) = ReplaceSysPara(CStr(varFields(0)))
            strValue = ReplaceVoucherItems(CStr(varFields(0)), voucher, lngRow)
            strValue = clsComp.GetExp(strValue)
            If CBool(IIf(strValue = "", 0, 1)) Then
                strSqlWhere = ReplaceVoucherItems(CStr(varFields(1)), voucher, lngRow)
            End If
            If GetFilterString = "" Then
                GetFilterString = strSqlWhere
            Else
                If strSqlWhere <> "" Then
                    GetFilterString = "(" & GetFilterString & ") and (" & strSqlWhere & ")"
                End If
            End If
        Next
    End If
End Function
Public Function GetReferName(strFieldName As String) As String
    Dim nod As IXMLDOMElement
    Set nod = domFieldConfig.selectSingleNode("//z:row[@cardsection='B' and @fieldname='" + LCase(strFieldName) + "']")
    If Not nod Is Nothing Then
        If Not nod.Attributes.getNamedItem("refername") Is Nothing Then
            If Not nod.Attributes.getNamedItem("refertype") Is Nothing Then
                If nod.Attributes.getNamedItem("refertype").nodeValue = "1" Then
                    GetReferName = nod.Attributes.getNamedItem("refername").Text
                End If
            End If
        End If
    End If
End Function
Public Function ReplaceSpecialCode(strSourceCode As String) As String
    Dim strTmp As String
    strTmp = Replace(strSourceCode, "'", "&apos;")
    strTmp = Replace(strTmp, """", "&quot;")
    strTmp = Replace(strTmp, ">", "&gt;")
    strTmp = Replace(strTmp, "<", "&lt;")
    strTmp = Replace(strTmp, "&", "&amp;")
    strTmp = Replace(strTmp, "\", "\\")
    ReplaceSpecialCode = strTmp
End Function


Public Function bExistField(strFieldName As String, recTemp As ADODB.Recordset) As Boolean

Dim i As Integer
On Error GoTo Err_Handle
For i = 0 To recTemp.Fields.Count - 1

    If LCase(recTemp.Fields(i).Name) = strFieldName Then
        bExistField = True
        Exit For
        
    End If
Next
Exit Function
Err_Handle:

End Function
