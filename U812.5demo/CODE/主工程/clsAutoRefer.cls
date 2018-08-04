VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAutoRefer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim clsCtlRefer As New U8RefService.IService  'New UFReferC.UFReferClient
Dim domFieldConfig As New DOMDocument
Dim domReferFillConfig As New DOMDocument
Dim clsAuth As New clsAutoAuth
Dim claauthys As New U8RowAuthsvr.clsRowAuth
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
    Dim strsql As String
    Dim Rst As New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    'rst.Open "select * from sa_voucherfieldconfig where cardnumber=N'" + strCardNumber + "' order by cardsection", DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    strsql = "set nocount on" & vbCrLf & "declare @t table([id] [int] IDENTITY(1,1) NOT NULL,[ctmp] nvarchar(20))"
    strsql = strsql & vbCrLf & "insert into @t select top 100 'abc' as ctmp from syscolumns"
    strsql = strsql & vbCrLf & "select substring(m.fieldname,t.id,case charindex(',',m.fieldname,t.id) when 0 then len(m.fieldname)+1 else charindex(',',m.fieldname,t.id) end - t.id) as fieldname,  m.* from sa_voucherfieldconfig as m,@t as t where cardnumber=N'" + strCardNumber + "' and (right(left(m.fieldname,t.id-1),1)=',' or t.id=1) order by cardsection"

    Rst.Open strsql, DBconn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Rst.Save domFieldConfig, adPersistXML
    Rst.Close
    'rst.Open "select cardnumber, fieldname, cardsection, refername, sourcefieldname, breferfill,bcellcheckfill, destfieldname, bclear,isnull(bcheckchange,0) as bcheckchange,changemsg  from SA_ReferFillConfig where cardnumber=N'" + strCardNumber + "' order by cardsection", DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    strsql = "set nocount on" & vbCrLf & "declare @t table([id] [int] IDENTITY(1,1) NOT NULL,[ctmp] nvarchar(20))"
    strsql = strsql & vbCrLf & "insert into @t select top 100 'abc' as ctmp from syscolumns"
    strsql = strsql & vbCrLf & "select substring(m.fieldname,t.id,case charindex(',',m.fieldname,t.id) when 0 then len(m.fieldname)+1 else charindex(',',m.fieldname,t.id) end - t.id) as fieldname,  m.* from SA_ReferFillConfig as m,@t as t where cardnumber=N'" + strCardNumber + "' and (right(left(m.fieldname,t.id-1),1)=',' or t.id=1) order by cardsection"

    Rst.Open strsql, DBconn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Rst.Save domReferFillConfig, adPersistXML
    Rst.Close
    DBconn.Execute "set nocount off"
    Set Rst = Nothing
'    clsCtlRefer.SetLogin m_Login       ',isnull(ifillother,0)
    Set clsAuth.clsAuth = m_login.RowAuth
    clsAuth.Init DBconn, m_login.cUserId, "refer", strCardNumber
    domEnumValue.loadXML "<Data />"
End Function
Private Function FillVoucherItems(Voucher As ctlVoucher, strReferFieldName As String, strReferName As String, strsql As String, strCardSection As String, Optional lngRow As Long) As Boolean
    Dim Rst As New ADODB.Recordset
    Rst.CursorLocation = adUseClient
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
    strSelect = ReplaceVoucherItems(strSelect, Voucher, lngRow)
    Rst.Open ConvertSQLString(strSelect), DBconn, adOpenForwardOnly, adLockReadOnly, adCmdText
    FillVoucherItems = True
    If Not Rst.EOF Then
        If Rst.RecordCount > 1 Then
            Set nod = domFieldConfig.selectSingleNode("//z:row[@cardsection='" + UCase(strCardSection) + "' and @fieldname='" + LCase(strReferFieldName) + "']")
            If Not nod.Attributes.getNamedItem("filterfields") Is Nothing Then
                strFilter = nod.Attributes.getNamedItem("filterfields").nodeValue
                varFilter = Split(strFilter, ",")
                For i = 0 To UBound(varFilter)
                    If strCardSection = "T" Then
                        Rst.Filter = varFilter(i) & "='" + Voucher.headerText(strReferFieldName) + "'"
                    Else
                        Rst.Filter = varFilter(i) & "='" + Voucher.bodyText(lngRow, strReferFieldName) + "'"
                    End If
                    If Rst.RecordCount = 0 Then Rst.Filter = ""
                    If Rst.RecordCount = 1 Then Exit For
                Next
            End If
        End If
        Set lstChange = domReferFillConfig.selectNodes("//z:row[@cardsection='" + UCase(strCardSection) + "' and @fieldname='" + LCase(strReferFieldName) + "' and @refername='" + LCase(strReferName) + "' and @bcellcheckfill='True' and @bcheckchange='1']")
        Set lst = domReferFillConfig.selectNodes("//z:row[@cardsection='" + UCase(strCardSection) + "' and @fieldname='" + LCase(strReferFieldName) + "' and @refername='" + LCase(strReferName) + "' and @bcellcheckfill='True']")
        Do While Not Rst.EOF
            strMsg = ""
            For Each eleChange In lstChange
                strSourceFieldName = eleChange.Attributes.getNamedItem("sourcefieldname").Text
                strDesFldName = eleChange.Attributes.getNamedItem("destfieldname").Text
                If IsNull(Rst.Fields(strSourceFieldName)) Then
                    strValue1 = ""
                Else
                    strValue1 = Rst.Fields(strSourceFieldName)
                End If
                If strCardSection = "T" Then
                    strValue2 = Voucher.headerText(strDesFldName)
                Else
                    strValue2 = Voucher.bodyText(lngRow, strDesFldName)
                End If
                If strValue2 <> "" And strValue1 <> strValue2 Then
                    If Not eleChange.Attributes.getNamedItem("changemsg") Is Nothing Then
                        strMsg = strMsg & IIf(strMsg = "", "", Chr(13)) & GetString(eleChange.Attributes.getNamedItem("changemsg").Text)
                    End If
                End If
            Next
            If strMsg <> "" Then
                If MsgBox(strMsg, vbYesNo + vbQuestion) = vbNo Then 'zh-CN：录入的业务员不属于录入的部门，是否继续？
                    ClearReferItems Voucher, strCardSection, strReferFieldName, strReferName, lngRow
                    Exit Function
                End If
            End If
            For Each nod In lst
                strSourceFieldName = nod.Attributes.getNamedItem("sourcefieldname").Text
                If strSourceFieldName = "''" Then
                    SetVoucherItemValue Voucher, strCardSection, nod.Attributes.getNamedItem("destfieldname").Text, "", lngRow
                Else
                    If nod.Attributes.getNamedItem("bcheckchange") Is Nothing Then
                        blnCheck = False
                    Else
                        blnCheck = IIf(nod.Attributes.getNamedItem("bcheckchange").Text = "0", False, True)
                    End If
                    If Not blnCheck Then
                        SetVoucherItemValue Voucher, strCardSection, nod.Attributes.getNamedItem("destfieldname").Text, IIf(IsNull(Rst.Fields(strSourceFieldName)), "", Rst.Fields(strSourceFieldName)), lngRow
                    Else
                        strDesFldName = nod.Attributes.getNamedItem("destfieldname").Text
                        If strCardSection = "T" Then
                            strValue2 = Voucher.headerText(strDesFldName)
                        Else
                            strValue2 = Voucher.bodyText(lngRow, strDesFldName)
                        End If
                        If strValue2 = "" Then
                            SetVoucherItemValue Voucher, strCardSection, nod.Attributes.getNamedItem("destfieldname").Text, IIf(IsNull(Rst.Fields(strSourceFieldName)), "", Rst.Fields(strSourceFieldName)), lngRow
                        End If
                    End If
                End If
            Next
            Rst.MoveNext
            If Not Rst.EOF Then
                Voucher.AddLine
                lngRow = Voucher.BodyRows
            End If
        Loop
    Else
        ClearReferItems Voucher, strCardSection, strReferFieldName, strReferName, lngRow
        FillVoucherItems = False
    End If
    Rst.Close
    Set Rst = Nothing
End Function

Public Function ClearReferItems(Voucher As ctlVoucher, strCardSection As String, strReferFieldName As String, strReferName As String, Optional lngRow As Long) As Boolean
    Dim lst As IXMLDOMNodeList
    Dim nod As IXMLDOMNode
    
    Set lst = domReferFillConfig.selectNodes("//z:row[@cardsection='" + UCase(strCardSection) + "' and @fieldname='" + LCase(strReferFieldName) + "' and @refername='" + LCase(strReferName) + "' and @bclear='True']")
    For Each nod In lst
        SetVoucherItemValue Voucher, strCardSection, nod.Attributes.getNamedItem("destfieldname").Text, "", lngRow
    Next
End Function

Public Function CellCheck(strReferString As String, Voucher As ctlVoucher, strCardSection As String, strFieldName As String, ByRef RetValue As Variant, Optional lngRow As Long, Optional strwhere As String) As String
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
    Dim strAuth As String
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
            strValue = GetVoucherItemValue(Voucher, strCardSection, strFieldName, lngRow)
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
                strValue = GetVoucherItemValue(Voucher, strCardSection, strFieldName, lngRow)
                If strValue = "" Then
                    ClearReferItems Voucher, strCardSection, strFieldName, strReferName, lngRow
                Else
                    If nod.Attributes.getNamedItem("cellchecksql") Is Nothing Then
                        Exit Function
                    End If
                    strsql = nod.Attributes.getNamedItem("cellchecksql").Text
                    strsql = VBA.Replace(strsql, "[retvalue]", RetValue)
                    blnOk = FillVoucherItems(Voucher, strFieldName, strReferName, strsql, strCardSection, lngRow)
                    If Not blnOk Then
                        If Not nod.Attributes.getNamedItem("errresid") Is Nothing Then
                            CellCheck = nod.Attributes.getNamedItem("errresid").Text
                            strError = nod.Attributes.getNamedItem("errresid").Text
                            MsgBox GetErrorString(strError, Voucher, strFieldName, strCardSection)
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
            CellCheck = CellCheckDefine(Voucher, Voucher.LookUpArray(strFieldName, IIf(strCardSection = "T", siHeader, sibody)), IIf(strCardSection = "T", siHeader, sibody), GetVoucherItemValue(Voucher, strCardSection, strFieldName, lngRow))
            If strCardSection = "T" Then
                Voucher.headerText(strFieldName) = CellCheck
            Else
                Voucher.bodyText(lngRow, strFieldName) = CellCheck
            End If
        Case "item"
            Dim Rst As New ADODB.Recordset
            Rst.CursorLocation = adUseClient
            strValue = GetVoucherItemValue(Voucher, strCardSection, strFieldName, lngRow)
            If strValue = "" Then
                If strCardSection = "T" Then
                    Voucher.headerText("citemcode") = ""
                    Voucher.headerText("citemname") = ""
                Else
                    Voucher.bodyText(lngRow, "citemcode") = ""
                    Voucher.bodyText(lngRow, "citemname") = ""
                End If
            Else
                If strCardSection = "T" Then
                    If Voucher.headerText("citem_class") = "" Then
                        MsgBox GetString("U8.SA.USSASERVER.clscommcheck.00450") 'zh-CN：请先填写项目大类
                        CellCheck = "cancel"
                        Exit Function
                    End If
                    'LDX 2009-05-18  modify beg
                    strsql = "select ctable from fitem where citem_class=N'" & Voucher.headerText("citem_class") & "'"
'                    strSql = "select ctable from fitem where citem_class='00'"
                    'LDX 2009-05-18  modify beg
                    Rst.Open strsql, DBconn
                    
                    If Not Rst.EOF Then
                        If Rst(0) = "inventory" Then
                            strsql = "select cInvCode as citemcode,cInvName as citemname from inventory where (cInvCode=N'" & strValue & "' or cInvName=N'" & strValue & "' ) "
                        Else
                            strsql = "select citemcode,citemname,bclose from " & Rst(0) & " where (citemcode=N'" & strValue & "' or citemname=N'" & strValue & "' ) and bclose =0"
                        End If
                        Rst.Close
                        
                        'shanlan add 090811----------------------------------------------------
                        If Not clsSAWeb.bAdmin Then
                            If clsSAWeb.bAuth_Inv Then
                                       Set claauthys = New U8RowAuthsvr.clsRowAuth
                                        claauthys.Init m_login.UfDbName, m_login.cUserId '"UFSOFT"
                                strAuth = claauthys.getAuthString("Fitem", Voucher.headerText("citem_class"), "W")
                                If strAuth = "1=2" Then
                                    MsgBox "没有项目权限"
                                    CellCheck = "cancel"
                                ElseIf strAuth <> "" Then
                                    strsql = strsql + " and (isnull(citemcode,'')='' or citemcode in  (" & strAuth & "))"
                                End If
                            End If
                         End If
                        '-------------------------------------------------------------------------
                        Rst.Open strsql, DBconn
                        If Not Rst.EOF Then
                            Voucher.headerText("citemcode") = Rst.Fields("citemcode")
                            Voucher.headerText("citemname") = Rst.Fields("citemname")
                        Else
                            Voucher.bodyText(lngRow, "citemcode") = ""
                            Voucher.bodyText(lngRow, "citemname") = ""
                            strError = GetString("U8.SA.USSASERVER.clscommcheck.00452") 'zh-CN：项目大类有错误,请检查!
                            MsgBox strError
                        End If
                    Else
                        strError = GetString("U8.SA.USSASERVER.clscommcheck.00451") 'zh-CN：项目大类有错误,请检查!
                        Voucher.headerText("citemcode") = ""
                        Voucher.headerText("citemname") = ""
                        MsgBox strError
                    End If
                    Rst.Close
                    
                Else
                    If Voucher.bodyText(lngRow, "citem_class") = "" Then
                        MsgBox GetString("U8.SA.USSASERVER.clscommcheck.00450") 'zh-CN：请先填写项目大类
                        CellCheck = "cancel"
                        Exit Function
                    End If
                    strsql = "select ctable from fitem where citem_class=N'" & Voucher.bodyText(lngRow, "citem_class") & "'"
                    Rst.Open strsql, DBconn
                    If Not Rst.EOF Then
                        If Rst(0) = "inventory" Then
                            strsql = "select cInvCode as citemcode,cInvName as citemname from inventory where (cInvCode=N'" & strValue & "' or cInvName=N'" & strValue & "' ) "
                        Else
                            strsql = "select citemcode,citemname,bclose from " & Rst(0) & " where (citemcode=N'" & strValue & "' or citemname=N'" & strValue & "' ) and bclose =0"
                        End If
                        Rst.Close
                        Rst.Open strsql, DBconn
                        If Not Rst.EOF Then
                            Voucher.bodyText(lngRow, "citemcode") = Rst.Fields("citemcode")
                            Voucher.bodyText(lngRow, "citemname") = Rst.Fields("citemname")
                        Else
                            Voucher.bodyText(lngRow, "citemcode") = ""
                            Voucher.bodyText(lngRow, "citemname") = ""
                            strError = GetString("U8.SA.USSASERVER.clscommcheck.00452") 'zh-CN：项目大类有错误,请检查!
                            MsgBox strError
                        End If
                    Else
                        strError = GetString("U8.SA.USSASERVER.clscommcheck.00451") 'zh-CN：项目大类有错误,请检查!
                        Voucher.bodyText(lngRow, "citemcode") = ""
                        Voucher.bodyText(lngRow, "citemname") = ""
                        MsgBox strError
                    End If
                    Rst.Close
                End If
            End If
            Set Rst = Nothing
    End Select
    CellCheck = strError
End Function
Private Function GetFieldName(Voucher As ctlVoucher, strFieldName As String, strCardSection As String) As String
    If strCardSection = "T" Then
        GetFieldName = Voucher.ItemCaption(strFieldName, siHeader)
    Else
        GetFieldName = Voucher.ItemCaption(strFieldName, sibody)
    End If
End Function
Private Function GetErrorString(strSource As String, Voucher As ctlVoucher, strFieldName As String, strCardSection As String) As String
    Dim varError As Variant
    If InStr(strSource, ",") = 0 Then
        GetErrorString = GetString(strSource)
    Else
        varError = Split(strSource, ",")
        GetErrorString = GetStringPara(varError(0), GetFieldName(Voucher, strFieldName, strCardSection))
    End If
End Function
''自定义项目检查，调用u8defpro
Private Function CellCheckDefine(Voucher As ctlVoucher, Index As Variant, iVoucherSec As Integer, KeyValue As String) As String
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
    If Not clsDef.Init(False, DBconn.ConnectionString, m_login.cUserId) Then
        CellCheckDefine = ""
        MsgBox GetString("U8.SA.xsglsql.01.frmbillvouch.00147") 'zh-CN：初始化自定义项组件失败！
        Set clsDef = Nothing
        Exit Function
    End If
    With Voucher
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

Public Function ShowReferCtl(clsVoucher As clsAutoVoucher, Voucher As ctlVoucher, iVoucherSec As Integer, iIndex As Long, referPara As UAPVoucherControl85.ReferParameter, Optional lngRow As Long, Optional strwhere As String) As String
    Dim nod As IXMLDOMNode
    Dim strReferType As String
    Dim strCardSection As String
    Dim strFieldName As String
    Dim clsRefer As New UFReferC.UFReferClient
    
    strCardSection = IIf(iVoucherSec = siHeader, "T", "B")
    strFieldName = Voucher.ItemState(iIndex, iVoucherSec).sFieldName
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
            ShowReferCtl = ShowEnumReferCtrl(clsVoucher, Voucher, strCardSection, strFieldName, referPara, lngRow, strwhere)
        Case "free"        '自由项
            referPara.Cancel = True
            Dim clsDef As U8DefPro.clsDefPro
            Set clsDef = New U8DefPro.clsDefPro
            If Not clsDef.Init(False, DBconn, m_login.cUserId) Then Exit Function
            If Not clsDef.InitNew(m_login, False) Then Exit Function
            Dim strValue As String
            
            strValue = clsDef.GetStruFreeRefVal(IIf(UCase(strCardSection) = "T", Voucher.headerText("cinvcode"), Voucher.bodyText(Voucher.row, "cinvcode")), strFieldName, IIf(UCase(strCardSection) = "T", Voucher.headerText(strFieldName), Voucher.bodyText(Voucher.row, strFieldName)), False, 100000, 1)
            If strValue <> "" Then
                If UCase(strCardSection) = "T" Then
                    Voucher.headerText(strFieldName) = strValue
                Else
                    Voucher.bodyText(Voucher.row, strFieldName) = strValue
                End If
            End If
            Set clsDef = Nothing
            If strCardSection = "B" Then
                ShowReferCtl = IIf(ShowReferCtl = "", "", ",") + CStr(lngRow)
            End If
            ShowReferCtl = strCardSection + "," + strFieldName + IIf(ShowReferCtl = "", "", ",") + ShowReferCtl
        Case "define"
            referPara.Cancel = True
            RefDefine Voucher, iIndex, iVoucherSec
        Case "itemclass"
            Dim strsql As String
            referPara.Cancel = True
            strsql = "select citem_class ,citem_name from fitem"
            If clsRefer.StrRefInit(m_login, False, "", strsql, GetString("U8.SA.xsglsql.01.frmbillvouch.00125") + "," + GetString("U8.SA.xsglsql.01.frmbillvouch.00125.01"), "") = False Then Exit Function 'zh-CN：项目大类编码 'zh-CN：项目大类名称
            clsRefer.Show
            
            If Not clsRefer.recmx Is Nothing Then
                If Not clsRefer.recmx.EOF Then
                    If iVoucherSec = sibody Then
                        Voucher.bodyText(Voucher.row, "citem_class") = clsRefer.recmx("citem_class")
                        Voucher.bodyText(Voucher.row, "citem_cname") = clsRefer.recmx("citem_name")
                    Else
                        Voucher.headerText("citem_class") = clsRefer.recmx("citem_class")
                        Voucher.headerText("citem_cname") = clsRefer.recmx("citem_name")
                    End If
                End If
            End If
        Case "item"
            referPara.Cancel = True
            strValue = Voucher.bodyText(Voucher.row, strFieldName)
            If iVoucherSec = sibody Then
                If Voucher.bodyText(Voucher.row, "citem_class") = "" Then
                    clsRefer.EnumRefInit m_login, enuTreeViewAndGrid, False, enuItem
                    If Voucher.bodyText(Voucher.row, "citem_class") <> "ch" Then
                        clsRefer.SetReferFilterString "bclose=0"
                    End If
                Else
                    clsRefer.ItemRefInit m_login, False, Voucher.bodyText(Voucher.row, "citem_class"), IIf(Voucher.bodyText(Voucher.row, "citem_class") <> "ch", "bclose=0", "")
                End If
            Else
                If Voucher.headerText("citem_class") = "" Then
                    clsRefer.EnumRefInit m_login, enuTreeViewAndGrid, False, enuItem
                    If Voucher.headerText("citem_class") <> "ch" Then
                        clsRefer.SetReferFilterString "bclose=0"
                    End If
                    'shanlan add 090810--------------------------------------------------------------
                    If clsSAWeb.bAuth_Inv Then
                      clsRefer.SetRWAuth "FITEM", "W", True
                    Else '
                      clsRefer.SetRWAuth "FITEM", "R", False
                    End If
                    '-------------- -----------------------------------------------------------------
                Else
                    'LDX    2009-05-18 modify beg
                    
                    clsRefer.ItemRefInit m_login, False, Voucher.headerText("citem_class"), IIf(Voucher.headerText("citem_class") <> "ch", "bclose=0", "")
'                    clsRefer.ItemRefInit m_Login, False, "00", IIf(Voucher.headerText("citem_class") <> "ch", "bclose=0", "")
                    'LDX    2009-05-18 modify beg
                                        'shanlan add 090810--------------------------------------------------------------
                    If clsSAWeb.bAuth_Inv Then
                      clsRefer.SetRWAuth "FITEM", "W", True
                    Else '
                      clsRefer.SetRWAuth "FITEM", "R", False
                    End If
                    '-------------- -----------------------------------------------------------------
                End If
            End If
            clsRefer.FillText = strValue
            clsRefer.Show
            If Not clsRefer.recmx Is Nothing Then
                If iVoucherSec = sibody Then
                    If Voucher.bodyText(lngRow, "citem_class") = "" Then
                        If Not clsRefer.RstSelClass.EOF Then
                            Voucher.bodyText(Voucher.row, "citem_class") = clsRefer.RstSelClass("citem_class")
                            Voucher.bodyText(Voucher.row, "citem_cname") = clsRefer.RstSelClass("citem_name")
                        End If
                    End If
                    If LCase(Voucher.bodyText(Voucher.row, "citem_class")) <> "ch" Then
                        Voucher.bodyText(lngRow, "citemcode") = clsRefer.recmx("citemcode")
                        Voucher.bodyText(lngRow, "citemname") = clsRefer.recmx("citemname")
                    Else
                        Voucher.bodyText(lngRow, "citemcode") = clsRefer.recmx("cinvcode")
                        Voucher.bodyText(lngRow, "citemname") = clsRefer.recmx("cinvname")
                    End If
                Else
                    If Voucher.headerText("citem_class") = "" Then
                        If Not clsRefer.RstSelClass.EOF Then
                            Voucher.headerText("citem_class") = clsRefer.RstSelClass("citem_class")
                            Voucher.headerText("citem_cname") = clsRefer.RstSelClass("citem_name")
                        End If
                    End If
                    If LCase(Voucher.headerText("citem_class")) <> "ch" Then
                        Voucher.headerText("citemcode") = clsRefer.recmx("citemcode")
                        Voucher.headerText("citemname") = clsRefer.recmx("citemname")
                    Else
                        Voucher.headerText("citemcode") = clsRefer.recmx("cinvcode")
                        Voucher.headerText("citemname") = clsRefer.recmx("cinvname")
                    End If
                End If
            End If
    End Select
    Set clsRefer = Nothing
End Function

Private Function RefDefine(Voucher As ctlVoucher, Index As Variant, iVoucherSec As Integer) As String
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

    With Voucher
        If iVoucherSec = siHeader Then
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
                sCardNumber = "01"
            Case 1
                sTableName = Left(sDataRule, InStr(1, sDataRule, ",") - 1)
                sFieldName = Mid(sDataRule, InStr(1, sDataRule, ",") + 1)
                sVouchType = ""
            Case 2
                sCardNumber = Left(sDataRule, InStr(1, sDataRule, ",") - 1)
                sFieldName = Mid(sDataRule, InStr(1, sDataRule, ",") + 1)
        End Select
'        If Not clsDef.Init(False, DBconn.ConnectionString, m_login.cUserId) Then
'            RefDefine = ""
'            MsgBox GetString("U8.SA.xsglsql.01.frmbillvouch.00146") 'zh-CN：初始化自定义项组件失败！
'            Exit Function
'        End If
        If Not clsDef.InitNew(m_login, False) Then
            RefDefine = ""
            MsgBox GetString("U8.SA.xsglsql.01.frmbillvouch.00146") 'zh-CN：初始化自定义项组件失败！
            Exit Function
        End If
'        RefDefine = clsDef.GetRefVal(nDataSource, iVoucherSec, .ItemState(Index, iVoucherSec).sFieldName, sTableName, sFieldName, sCardNumber, strKeyValue, False, 40, 1)
        RefDefine = clsDef.GetRefVal(nDataSource, iVoucherSec, "c" & .ItemState(Index, iVoucherSec).sFieldName, sTableName, sFieldName, sCardNumber, strKeyValue, False, 40, 1)
        If iVoucherSec = siHeader Then
            .headerText(Index) = RefDefine
        Else
            .bodyText(.row, Index) = RefDefine
        End If
    End With
    Set clsDef = Nothing
End Function

Public Sub FillList(Voucher As ctlVoucher, lngRow As Long, lngCol As Long, pCom As Object)
    Dim strFieldName As String
    Dim nod As IXMLDOMNode
    Dim strsql As String
    Dim strReferName As String
    Dim strFilter As String
    Dim Rst As New ADODB.Recordset
    Dim ele As IXMLDOMElement
    Dim eleValue As IXMLDOMElement
    Rst.CursorLocation = adUseClient
    strFieldName = LCase(Voucher.ItemState(lngCol, sibody).sFieldName)
    Set nod = domFieldConfig.selectSingleNode("//z:row[@cardsection='B' and @fieldname='" + LCase(strFieldName) + "']")
    If Not nod.Attributes.getNamedItem("refertype") Is Nothing Then
        If nod.Attributes.getNamedItem("refertype").Text = "4" Then
            strReferName = nod.Attributes.getNamedItem("refername").Text
            If Not nod.Attributes.getNamedItem("referdefaultfilter") Is Nothing Then
                strFilter = nod.Attributes.getNamedItem("referdefaultfilter").Text
            End If
            strsql = "select enumcode from aa_enum where enumtype=N'" + strReferName + "' and localeid=N'" + m_login.LanguageRegion + "'" + IIf(strFilter = "", "", " and (" + strFilter + ")") + " order by enumindex"
            Rst.Open ConvertSQLString(strsql), DBconn, adOpenForwardOnly, adLockReadOnly, adCmdText
            Set ele = domEnumValue.documentElement.selectSingleNode("row[@name='" + strFieldName + "']")
            If ele Is Nothing Then
                Set ele = domEnumValue.createElement("row")
                ele.setAttribute "name", strFieldName
                domEnumValue.documentElement.appendChild ele
            End If
            Do While Not Rst.EOF
                Set eleValue = ele.selectSingleNode("value[@code='" + Rst(0) + "']")
                If eleValue Is Nothing Then
                    Set eleValue = domEnumValue.createElement("value")
                    eleValue.setAttribute "code", Rst(0)
                    ele.appendChild eleValue
                End If
                pCom.AddItem Rst(0)
                Rst.MoveNext
            Loop
            Rst.Close
        End If
    End If
    Set Rst = Nothing
End Sub

Public Sub FillHeadComboBox(Voucher As ctlVoucher, Index As Long, pCom As Object)
    Dim strFieldName As String
    Dim nod As IXMLDOMNode
    Dim strsql As String
    Dim strReferName As String
    Dim strFilter As String
    Dim Rst As New ADODB.Recordset
    Dim ele As IXMLDOMElement
    Dim eleValue As IXMLDOMElement
    Rst.CursorLocation = adUseClient
    strFieldName = LCase(Voucher.ItemState(Index, siHeader).sFieldName)
    Set nod = domFieldConfig.selectSingleNode("//z:row[@cardsection='T' and @fieldname='" + LCase(strFieldName) + "']")
    If Not nod.Attributes.getNamedItem("refertype") Is Nothing Then
        If nod.Attributes.getNamedItem("refertype").Text = "4" Then
            strReferName = nod.Attributes.getNamedItem("refername").Text
            If Not nod.Attributes.getNamedItem("referdefaultfilter") Is Nothing Then
                strFilter = nod.Attributes.getNamedItem("referdefaultfilter").Text
            End If
            strsql = "select enumcode from aa_enum where enumtype=N'" + strReferName + "' and localeid=N'" + m_login.LanguageRegion + "'" + IIf(strFilter = "", "", " and (" + strFilter + ")") + " order by enumindex"
            Rst.Open ConvertSQLString(strsql), DBconn, adOpenForwardOnly, adLockReadOnly, adCmdText
            Set ele = domEnumValue.documentElement.selectSingleNode("row[@name='" + strFieldName + "']")
            If ele Is Nothing Then
                Set ele = domEnumValue.createElement("row")
                ele.setAttribute "name", strFieldName
                domEnumValue.documentElement.appendChild ele
            End If
            Do While Not Rst.EOF
                Set eleValue = ele.selectSingleNode("value[@code='" + Rst(0) + "']")
                If eleValue Is Nothing Then
                    Set eleValue = domEnumValue.createElement("value")
                    eleValue.setAttribute "code", Rst(0)
                    ele.appendChild eleValue
                End If
                pCom.AddItem Rst(0)
                Rst.MoveNext
            Loop
            Rst.Close
        End If
    End If
    Set Rst = Nothing
End Sub
Public Function GetEnumDom(strFieldName As String) As DOMDocument
'    Dim strFieldName As String
    Dim nod As IXMLDOMNode
    Dim strsql As String
    Dim strReferName As String
    Dim strFilter As String
    Dim Rst As New ADODB.Recordset
    Dim dom As New DOMDocument
    Rst.CursorLocation = adUseClient
'    strFieldName = LCase(voucher.ItemState(lngCol, sibody).sFieldName)
    Set nod = domFieldConfig.selectSingleNode("//z:row[@cardsection='B' and @fieldname='" + LCase(strFieldName) + "']")
    If Not nod.Attributes.getNamedItem("refertype") Is Nothing Then
        If nod.Attributes.getNamedItem("refertype").Text = "4" Then
            strReferName = nod.Attributes.getNamedItem("refername").Text
            If Not nod.Attributes.getNamedItem("referdefaultfilter") Is Nothing Then
                strFilter = nod.Attributes.getNamedItem("referdefaultfilter").Text
            End If
            strsql = "select enumcode,enumname,enumindex from aa_enum where enumtype='" + strReferName + "' and localeid='" + m_login.LanguageRegion + "'" + IIf(strFilter = "", "", " and (" + strFilter + ")")
            Rst.Open ConvertSQLString(strsql), DBconn, adOpenForwardOnly, adLockReadOnly, adCmdText
            Rst.Save dom, adPersistXML
            Rst.Close
        End If
    End If
    Set GetEnumDom = dom
    Set dom = Nothing
    Set Rst = Nothing
End Function

Private Function ShowEnumReferCtrl(clsVoucher As clsAutoVoucher, Voucher As ctlVoucher, strCardSection As String, strFieldName As String, referPara As UAPVoucherControl85.ReferParameter, Optional lngRow As Long, Optional strwhere As String) As String
    Dim strFilter As String
    Dim strDefFilter As String
    Dim nod As IXMLDOMNode
    Dim strReferName As String
    Dim strAuth As String
    Dim rstClass As ADODB.Recordset
    Dim rstGrid As ADODB.Recordset
    Dim strError As String
    Dim bAuth As String
    
    Dim referdispmode As String
    
    
    ShowEnumReferCtrl = ""
    Set nod = domFieldConfig.selectSingleNode("//z:row[@cardsection='" + UCase(strCardSection) + "' and @fieldname='" + LCase(strFieldName) + "']")
    If Not nod Is Nothing Then
        If Not nod.Attributes.getNamedItem("referdefaultfilter") Is Nothing Then
            strDefFilter = ReplaceVoucherItems(nod.Attributes.getNamedItem("referdefaultfilter").Text, Voucher, lngRow)
        End If
        If Not nod.Attributes.getNamedItem("referfilterfields") Is Nothing Then
            strFilter = GetFilterString(Voucher, Trim(nod.Attributes.getNamedItem("referfilterfields").Text), lngRow)
        End If
        
        If Not nod.Attributes.getNamedItem("referdispmode") Is Nothing Then
            referdispmode = Trim(nod.Attributes.getNamedItem("referdispmode").Text)
        Else
            referdispmode = "1"
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
       
        strAuth = clsAuth.GetAuthStringByOperation(strReferName)
        If strAuth <> "" Then
            strwhere = strAuth + IIf(strwhere = "", "", " and (" + strwhere + ")")
        End If
        
        referPara.id = strReferName
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
        If referdispmode = "1" Then
            bAuth = 0
        Else
            bAuth = 1
        End If
        
        
        ' bAuth='1'  表示 进行权限控制    bAuth='0'  表示 不进行权限控制
        referPara.ReferMetaXML = "<Ref><RefSet bControlledInvalidData='1' bAuth='" & bAuth & "' bMultiSel= '" + IIf(CBool(strMuti), "1", "0") + "' iShowMode='" + referdispmode + "' iShowStyle='" + strShowStyle + "' /></Ref>" 'iShowMode
        If referPara.Cancel = True Then
            clsCtlRefer.RefID = strReferName
            clsCtlRefer.FillText = GetVoucherItemValue(Voucher, strCardSection, strFieldName, lngRow)
            clsCtlRefer.FilterSQL = ConvertSQLString(strwhere)
            clsCtlRefer.MetaXML = "<Ref><RefSet bControlledInvalidData='1' bAuth='0' bMultiSel= '" + IIf(CBool(strMuti), "1", "0") + "' iShowMode='" + nod.Attributes.getNamedItem("referdispmode").Text + "' iShowStyle='" + strShowStyle + "' /></Ref>" 'iShowMode
'            strReferName = Replace(strReferName, "_aa", "")
'            If clsCtlRefer.EnumRefInit(m_Login, nod.Attributes.getNamedItem("referdispmode").Text, nod.Attributes.getNamedItem("refermultilines").Text, strReferName, ConvertSQLString(strWhere)) Then
            

            If clsCtlRefer.ShowRef(m_login, rstClass, rstGrid, strError) Then
'                clsCtlRefer.Show
                If Not rstGrid Is Nothing Then
                    Set referPara.rstGrid = rstGrid
                    ShowEnumReferCtrl = FillItemsAfterBrowse(clsVoucher, Voucher, strCardSection, strFieldName, rstGrid, lngRow)
                End If
            End If
        End If
    End If
End Function
Public Function FillItemsAfterBrowse(clsVoucher As clsAutoVoucher, Voucher As ctlVoucher, strCardSection As String, strFieldName As String, recResult As ADODB.Recordset, ByRef RetValue As Variant, Optional lngRow As Long) As String
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
        strsql = VBA.Replace(strsql, "[retvalue]", RetValue)
    End If
    strReferName = nod.Attributes.getNamedItem("refername").Text
    Dim strValue As String
    
'    strReferName = Replace(strReferName, "_aa", "")
    Do While Not recResult.EOF
        Set lst = domReferFillConfig.selectNodes("//z:row[@cardsection='" + UCase(strCardSection) + "' and @fieldname='" + LCase(strFieldName) + "' and @refername='" + LCase(strReferName) + "' and @breferfill='True']")
        For Each subNod In lst
            Dim intCyc As Integer
            Dim blnExsitFied As Boolean
            For intCyc = 1 To recResult.Fields.Count
                If LCase(recResult.Fields(intCyc - 1).Name) = LCase(subNod.Attributes.getNamedItem("sourcefieldname").Text) Then
                    blnExsitFied = True
                    Exit For
                Else
                    blnExsitFied = False
                End If
            Next
            If blnExsitFied = True Then
                If IsNull(recResult.Fields(subNod.Attributes.getNamedItem("sourcefieldname").Text)) Then
                    strValue = ""
                Else
                    strValue = recResult.Fields(subNod.Attributes.getNamedItem("sourcefieldname").Text)
                End If
                SetVoucherItemValue Voucher, strCardSection, subNod.Attributes.getNamedItem("destfieldname").Text, strValue, lngRow
            End If
'            If strCardSection = "B" Then
'                Voucher.CallAutoFillBackEvent strFieldName, lngRow
'            End If
        Next
        If strsql <> "" Then
            FillVoucherItems Voucher, strFieldName, strReferName, strsql, strCardSection, lngRow
        End If
        If strCardSection = "B" Then
            FillItemsAfterBrowse = IIf(FillItemsAfterBrowse = "", "", ",") + CStr(lngRow)
        End If
'        Set lst = domReferFillConfig.selectNodes("//z:row[@cardsection='" + UCase(strCardSection) + "' and @fieldname='" + LCase(strFieldName) + "' and @refername='" + LCase(strReferName) + "' and @ifillother!='0']")
'        For Each nod In lst
'            ReferCellCheck Voucher, nod.Attributes.getNamedItem("cardnumber").Text, strCardSection, nod.Attributes.getNamedItem("destfieldname").Text, lngRow
'        Next
        If strCardSection = "B" Then
            Voucher.CallAutoFillBackEvent strFieldName, lngRow
        End If
        recResult.MoveNext
        If Not recResult.EOF Then
            Voucher.row = lngOldRow
'            Dim strValue As String
            strValue = Voucher.bodyText(lngRow, strFieldName)
            clsVoucher.CopyLine Voucher, Me
            Voucher.bodyText(lngRow, strFieldName) = strValue
            lngRow = Voucher.BodyRows
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
Public Sub ReferCellCheck(Voucher As ctlVoucher, strCardSection As String, strFieldName As String, Optional lngRow As Long)
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
    FillVoucherItems Voucher, strFieldName, strReferName, strsql, strCardSection, lngRow
'    End If
'    rst.Close
'    Set rst = Nothing
End Sub
Private Function GetFilterString(Voucher As ctlVoucher, strFilterFields As String, Optional lngRow As Long) As String
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
            strValue = ReplaceVoucherItems(CStr(varFields(0)), Voucher, lngRow)
            strValue = clsComp.GetExp(strValue)
            If CBool(strValue) Then
                strSqlWhere = ReplaceVoucherItems(CStr(varFields(1)), Voucher, lngRow)
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


