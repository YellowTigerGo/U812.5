VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAutoVoucherList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_strDefaultFilter As String
Dim m_strFrom As String
Dim m_strFilterName As String
Dim m_strSubSysID As String
Dim m_strFilter As String
Dim m_strColumnSetKey As String
Dim domListConfig As New DOMDocument
Dim domFieldConfig As New DOMDocument
Dim domGoToVoucher As New DOMDocument
Public m_oColSet As New U8ColumnSet.clsColSet   'U8colset.clsColSet           '项目设置对象
Dim m_strColFormatXml    As String
Dim clsAuth As New clsAutoAuth
Dim m_strKey As String
Dim m_bshowSumType As Boolean
Dim domDefine As New DOMDocument

Public Function Init(strLstKey As String, strErrorResId As String) As Boolean
    
    Dim rsVouchs As New ADODB.Recordset
    Dim strColumnKey As String
    Dim strAuth As String
    Dim sAuthWhere As String
    rsVouchs.CursorLocation = adUseClient
    strKey = strLstKey
    rsVouchs.Open "Select subsysid,filtername,listfrom,isnull(defaultfilter,N'') as defaultfilter,menuid,columnkey,maintbl,mainkey,detailtbl,detailkey,isnull(showsumoption,0) as showsumoption from SA_VoucherListConfig Where ckey=N'" & strLstKey & "'", DBconn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (rsVouchs.BOF And rsVouchs.EOF) Then
        If strDefaultFilter = "" Then
            strDefaultFilter = rsVouchs("defaultfilter")
        Else
            If rsVouchs("defaultfilter") <> "" Then
                strDefaultFilter = "(" & strDefaultFilter & ") and (" & rsVouchs("defaultfilter") & ")"
            End If
        End If
        If Not IsNull(rsVouchs("listfrom")) Then
            strFrom = rsVouchs("listfrom")
        End If
        If Not IsNull(rsVouchs("filtername")) Then
            StrFilterName = rsVouchs("filtername")
        End If
        strSubSysID = rsVouchs("subsysid")
    End If
    rsVouchs.Save domListConfig, adPersistXML
    rsVouchs.Close
    strColumnKey = domListConfig.selectSingleNode("//z:row").Attributes.getNamedItem("columnkey").Text
    strColumnSetKey = strColumnKey
    rsVouchs.Open "select * from sa_voucherlistfieldconfig where ckey=N'" + strColumnSetKey + "' and isnull(formatstr,N'')<>N''", DBconn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rsVouchs.Save domFieldConfig, adPersistXML
    rsVouchs.Close
    
    rsVouchs.Open "select * from sa_gotovoucherconfig where ckey=N'" + strColumnKey + "'", DBconn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rsVouchs.Save domGoToVoucher, adPersistXML
    rsVouchs.Close
    rsVouchs.Open "Select isnull(idecimaldigits,0) idecimaldigits,cid From UserDef Where cID In  (61,62,26,27,07,41,87,88,74,75)", DBconn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rsVouchs.Save domDefine, adPersistXML
    rsVouchs.Close
   
    Set rsVouchs = Nothing
    Set clsAuth.clsAuth = m_login.RowAuth
    clsAuth.Init DBconn, m_login.cUserId, "voucherlist", strColumnKey
    strAuth = clsAuth.GetAuthStringByOperation("search")
    If strAuth <> "" Then
        If strDefaultFilter = "" Then
            strDefaultFilter = strAuth
        Else
            strDefaultFilter = "(" & strDefaultFilter & ") and (" & strAuth & ")"
        End If
    End If
    'shanlan add 090810---------------------------------------------------------------
'    sAuthWhere = getAuthString("KJ", "", DBConn, m_Login, strLstKey, "R")
'    strDefaultFilter = IIf(sAuthWhere = "", "", sAuthWhere)
    
    InitColumnSet
End Function
Public Sub FormatVouchList(Voulist As Object)
    Dim lst As IXMLDOMNodeList
    Dim nod As IXMLDOMNode
    Dim strType As String
    
    Set lst = domFieldConfig.selectNodes("//z:row")
    For Each nod In lst
        strType = ""
        If Not nod.Attributes.getNamedItem("precisiontype") Is Nothing Then
            strType = nod.Attributes.getNamedItem("precisiontype").Text
        End If
        If strType = "define" Then
            Voulist.SetFormatString nod.Attributes.getNamedItem("fieldname").Text, GetDefinFormat(nod.Attributes.getNamedItem("formatstr").Text)
        Else
            Voulist.SetFormatString nod.Attributes.getNamedItem("fieldname").Text, nod.Attributes.getNamedItem("formatstr").Text
        End If
    Next
End Sub

Public Property Get strDefaultFilter() As String
    strDefaultFilter = m_strDefaultFilter
End Property

Public Property Let strDefaultFilter(ByVal vNewValue As String)
    m_strDefaultFilter = vNewValue
End Property

Public Property Get strFrom() As String
    strFrom = m_strFrom
End Property

Public Property Let strFrom(ByVal vNewValue As String)
    m_strFrom = vNewValue
End Property

Public Property Get StrFilterName() As String
    StrFilterName = m_strFilterName
End Property

Public Property Let StrFilterName(ByVal vNewValue As String)
    m_strFilterName = vNewValue
End Property
Public Function ShowFilter() As Boolean
    Dim clsFilter As New clsFilter
    Dim strFilterStr As String
    Dim strShowSumOption As String
    
    strShowSumOption = GetVoucherListSet("showsumoption")
    clsFilter.bHideSum = Not CBool(strShowSumOption)
    ShowFilter = clsFilter.ShowFilter(StrFilterName, strSubSysID, strFilterStr)
    If ShowFilter Then
        strFilter = strFilterStr
        m_bshowSumType = clsFilter.bShowSumType
    End If
    Set clsFilter = Nothing
End Function

Public Sub SetFilter()
    Dim clsFilter As New clsFilter
    clsFilter.SetFilter StrFilterName, strSubSysID
    Set clsFilter = Nothing
End Sub

Public Property Get strSubSysID() As String
    strSubSysID = m_strSubSysID
End Property

Public Property Let strSubSysID(ByVal vNewValue As String)
    m_strSubSysID = vNewValue
End Property

Public Property Get strFilter() As String
    strFilter = m_strFilter
End Property

Public Property Let strFilter(ByVal vNewValue As String)
    m_strFilter = vNewValue
End Property

Private Sub Class_Terminate()
    Set m_oColSet = Nothing
    Set domListConfig = Nothing
    Set domFieldConfig = Nothing
    Set domGoToVoucher = Nothing
    Set clsAuth = Nothing
    Set domDefine = Nothing
End Sub
Private Sub InitColumnSet()
'    Set m_oColSet = New U8colset.clsColSet                          '初始化项目设置对象
    m_oColSet.Init m_login.UfDbName, m_login.cUserId               'ColSet.setColMode (cKey) '设置项目返回模式
    Call m_oColSet.setColMode(strColumnSetKey, 0)
End Sub
Public Function ColumnSet() As Boolean
    '开始项目设置
    m_oColSet.isShowTitle = True
    m_oColSet.AllowMerge = False
    ColumnSet = m_oColSet.setCol
    If ColumnSet Then
        strColFormatXml = m_oColSet.getColInfo()
    End If
End Function


Public Property Get strColumnSetKey() As String
    strColumnSetKey = m_strColumnSetKey
End Property
Public Property Get bShowSumType() As Boolean
    bShowSumType = m_bshowSumType
End Property

Public Property Let strColumnSetKey(ByVal vNewValue As String)
    m_strColumnSetKey = vNewValue
End Property
'Public Sub ShowAddNewVoucher()
'
'End Sub
Public Sub CreateNewVoucher(vchList As VouchList, strReferKey As String)
    Dim domHead As DOMDocument
    Dim domBody As DOMDocument
    Dim nod As IXMLDOMElement
    Dim nodlst As IXMLDOMNode
    Dim strMenuName As String
    Dim strAuthID As String
    Dim strErrorResId As String
    
    Set domHead = vchList.GetListDom(True)
    If strReferKey <> "新增" Then
        If domHead.selectNodes("//z:row").length <= 0 Then
            MsgBox GetString("U8.SA.xsglsql_2.unselected")      '"没有选择记录"
            Set domHead = Nothing
            Set domBody = Nothing
            Exit Sub
        End If
    End If
    Set domBody = domHead
    Set nodlst = domListConfig.selectSingleNode("//z:row")
    strMenuName = nodlst.Attributes.getNamedItem("menuid").Text
    Set nod = domMenu.selectSingleNode("//z:row[@menuid='" + strMenuName + "']")
    nod.setAttribute "referkey", strReferKey
    nod.setAttribute "voucherid", ""
    nod.setAttribute "voucherautoid", ""
    If Not nod Is Nothing Then
        If Not nod.Attributes.getNamedItem("authid") Is Nothing Then
            strAuthID = nod.Attributes.getNamedItem("authid").Text
        End If
        If Not domHead Is Nothing Then
            nod.setAttribute "sourcedomhead", domHead.xml
        End If
        If Not domBody Is Nothing Then
            nod.setAttribute "sourcedombody", domBody.xml
        End If
        'If Not frmMain.ShowVoucher(nod, strErrorResId, strAuthID) Then
        '    MsgBox strErrorResId
        'End If
    End If
End Sub

Public Sub ShowVoucher(Voulist As Object)
    Dim nod As IXMLDOMElement
    Dim lst As IXMLDOMNodeList
    Dim nodlst As IXMLDOMNode
    Dim strCondition As String
    Dim Rst As New ADODB.Recordset
    Dim strFieldName As String
    Dim strFieldValue As String
    Dim strMenuName As String
    Dim strAuthID As String
    Dim strErrorResId As String
    Dim strMainName As String
    Dim strDetailName As String
    Dim strCurrentRow As String
    Dim strVouchID As String
    Dim strSourceDetailKey As String
    
    Rst.CursorLocation = adUseClient
    Set lst = domGoToVoucher.selectNodes("//z:row")
    For Each nodlst In lst
        If nodlst.Attributes.getNamedItem("condition") Is Nothing Then
            strMenuName = nodlst.Attributes.getNamedItem("menuid").Text
            Exit For
        Else
            strCondition = nodlst.Attributes.getNamedItem("condition").Text
            If strCondition = "" Then
                strMenuName = nodlst.Attributes.getNamedItem("menuid").Text
                Exit For
            Else
                If Not nodlst.Attributes.getNamedItem("mainfieldname") Is Nothing Then
                    strMainName = nodlst.Attributes.getNamedItem("mainfieldname").Text
                    Rst.Open "select * " & strCondition & " and " & strMainName & "=" & Voulist.TextMatrix(Voulist.row, Voulist.GridColIndex(strMainName)), DBconn, adOpenForwardOnly, adLockReadOnly, adCmdText
                    If Not Rst.EOF Then
                        strMenuName = nodlst.Attributes.getNamedItem("menuid").Text
                        Exit For
                    End If
                    Rst.Close
                End If
            End If
        End If
    Next
    Set Rst = Nothing
    If strMenuName = "" Then Exit Sub
    Set nod = domMenu.selectSingleNode("//z:row[@menuid='" + strMenuName + "']")
    If Not nod Is Nothing Then
        nod.setAttribute "voucherid", ""
        nod.setAttribute "voucherautoid", ""
        If Not nod.Attributes.getNamedItem("authid") Is Nothing Then
            strAuthID = nod.Attributes.getNamedItem("authid").Text
        End If
        nod.setAttribute "voucherid", ""
        If Not nodlst.Attributes.getNamedItem("mainfieldname") Is Nothing Then
            strMainName = nodlst.Attributes.getNamedItem("mainfieldname").Text
            If strMainName <> "" Then
                strVouchID = Voulist.TextMatrix(Voulist.row, Voulist.GridColIndex(strMainName))
                nod.setAttribute "voucherid", strVouchID
            End If
        End If
        nod.setAttribute "voucherautoid", ""
        nod.setAttribute "setcurrentrow", ""
        If Not nodlst.Attributes.getNamedItem("detailfieldname") Is Nothing Then
            strDetailName = nodlst.Attributes.getNamedItem("detailfieldname").Text
            If strDetailName <> "" Then
                nod.setAttribute "voucherautoid", Voulist.TextMatrix(Voulist.row, Voulist.GridColIndex(strDetailName))
                If Voulist.TextMatrix(Voulist.row, Voulist.GridColIndex(strDetailName)) <> "" Then
                    If Not nodlst.Attributes.getNamedItem("sourcedetailkey") Is Nothing Then
                        strSourceDetailKey = nodlst.Attributes.getNamedItem("sourcedetailkey").Text
                    Else
                        strSourceDetailKey = strDetailName
                    End If
                    strCurrentRow = "@" + strSourceDetailKey + "='" & Voulist.TextMatrix(Voulist.row, Voulist.GridColIndex(strDetailName)) & "'"
                    nod.setAttribute "setcurrentrow", strCurrentRow
                End If
            End If
        End If
        If nod.Attributes.getNamedItem("voucherid").Text <> "" Or nod.Attributes.getNamedItem("voucherautoid").Text <> "" Then
            frmMain.MenuClick strMenuName, strAuthID, strVouchID, strCurrentRow
        End If
    Else
        
        If Not nodlst.Attributes.getNamedItem("mainfieldname") Is Nothing Then
            strMainName = nodlst.Attributes.getNamedItem("mainfieldname").Text
            If strMainName <> "" Then
                strVouchID = Voulist.TextMatrix(Voulist.row, Voulist.GridColIndex(strMainName))
            End If
        End If
        If Not nodlst.Attributes.getNamedItem("detailfieldname") Is Nothing Then
            strDetailName = nodlst.Attributes.getNamedItem("detailfieldname").Text
            If strDetailName <> "" Then
                If Voulist.TextMatrix(Voulist.row, Voulist.GridColIndex(strDetailName)) <> "" Then
                    If Not nodlst.Attributes.getNamedItem("sourcedetailkey") Is Nothing Then
                        strSourceDetailKey = nodlst.Attributes.getNamedItem("sourcedetailkey").Text
                    Else
                        strSourceDetailKey = strDetailName
                    End If
                    strCurrentRow = "@" + strSourceDetailKey + "='" & Voulist.TextMatrix(Voulist.row, Voulist.GridColIndex(strDetailName)) & "'"
                End If
            End If
        End If
        If strVouchID <> "" Then
            frmMain.MenuClick strMenuName, strAuthID, strVouchID, strCurrentRow
        End If
    End If
End Sub

Public Function PrintVoucherList(VchLst As Object, strType As String, strColumnSetKey As String) As Boolean
    Dim strPrnXml As String
    Select Case strType
        Case "print"
            strPrnXml = getPrnSet(strColumnSetKey)
            If strPrnXml <> "" Then VchLst.InitPrintSetup strPrnXml
            VchLst.VchLstPrint
        Case "preview"
            strPrnXml = getPrnSet(strColumnSetKey)
            If strPrnXml <> "" Then VchLst.InitPrintSetup strPrnXml
            VchLst.VchLstPreview
        Case "output"
            VchLst.VchLstPrintToFile
    End Select
End Function

Private Function getPrnSet(strKey As String) As String
    Dim Rst As New ADODB.Recordset
    
    Rst.CursorLocation = adUseClient
    
    Rst.Open "select * from prn_format where moduleid=N'" & strKey & "_" & "Print_" & m_login.cUserId & "'", DBconn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Rst.EOF Then
        getPrnSet = ""
    Else
        getPrnSet = Rst.Fields("formatxml")
    End If
    Rst.Close
    Set Rst = Nothing
End Function

Public Sub SavePrnSet(strPrnXml As String, strKey As String)
    Dim Rst As New ADODB.Recordset
    
    Rst.CursorLocation = adUseClient
    
    Rst.Open "select * from prn_format where moduleid=N'" & strKey & "_" & "Print_" & m_login.cUserId & "'", DBconn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Rst.EOF Then
        DBconn.Execute "insert into prn_format (moduleid,formatxml) values (N'" & strKey & "_" & "Print_" & m_login.cUserId & "',N'" & strPrnXml & "')"
    Else
        DBconn.Execute "update prn_format set formatxml=N'" & strPrnXml & "' where moduleid='" & strKey & "_" & "Print_" & m_login.cUserId & "'"
    End If
    Rst.Close
    Set Rst = Nothing
End Sub

Public Property Get strKey() As String
    strKey = m_strKey
End Property

Public Property Let strKey(ByVal vNewValue As String)
    m_strKey = vNewValue
End Property

Public Property Get strColFormatXml() As String
    strColFormatXml = m_oColSet.getColInfo()
End Property

Public Property Let strColFormatXml(ByVal vNewValue As String)
    m_strColFormatXml = vNewValue
End Property

Private Function GetNumFormat(Dec As Integer) As Variant
    Dim tmpStr As String, tmpFString As String
    If Dec < 0 Then Dec = 0
    GetNumFormat = "####0" & IIf(Dec = 0, "", ".") & String(val(Dec), "0")
End Function

Private Function GetDefinFormat(strID As String) As String
    Dim ele As IXMLDOMElement
    Set ele = domDefine.selectSingleNode("//z:row[@cid='" + strID + "']")
    If Not ele Is Nothing Then
        GetDefinFormat = GetNumFormat(val(ele.Attributes.getNamedItem("idecimaldigits").Text))
    Else
        GetDefinFormat = GetNumFormat(6)
    End If
End Function

Public Function GetVoucherListSet(strName As String) As String
    Dim ele As IXMLDOMElement
    Set ele = domListConfig.selectSingleNode("//z:row")
    If Not ele.Attributes.getNamedItem(strName) Is Nothing Then
        GetVoucherListSet = ele.Attributes.getNamedItem(strName).nodeValue
    Else
        GetVoucherListSet = ""
    End If
End Function

Public Function GetSumRecord(strwhere As String) As ADODB.Recordset
    Dim Rst As New ADODB.Recordset
    Dim strSelect As String
    Dim strsql As String
    Dim clsAuth As New EFVoucherMo.clsUserInterfaceInit  ''   U8RowAuthsvr.clsRowAuth
    Dim strAuth As String
    
    On Error Resume Next
    Rst.CursorLocation = adUseClient
    strSelect = m_oColSet.GetSumStringKCGL
    
    
    
    'by ahzzd 20100926 增加记录权限功能
    clsAuth.Init DBconn, m_login.cUserId, "voucherlist", m_strKey
    strAuth = clsAuth.GetAuthStringByOperation("search")
    If Trim(strAuth) <> "" Then
        strwhere = strwhere & "and " & strAuth
    End If
    
    strsql = "select " & strSelect & " from " & strFrom & " Where " & strwhere
    cls_Public.WrtDBlog DBconn, m_login.cUserId, "EFmain", "EFmain->clsAutoVoucherList 查询脚本[" & strsql & "]"
    Rst.Open strsql, DBconn, adOpenForwardOnly, adLockReadOnly
    Set GetSumRecord = Rst
End Function

Public Sub PrnBatch(VchLst As Object)
Dim i As Long
Dim sqlstr As String
Dim Cardnumber As String
Dim cls_print As Object
Dim rd As New ADODB.Recordset
Set cls_print = CreateObject("UF_Public_base.clsVoucher_Print")
    rd.CursorLocation = adUseClient
    sqlstr = "select isnull(parameters,'') from sa_menuconfig where menuId in(select menuId from sa_gotovoucherconfig where ckey='" & strKey & "')"
    rd.Open sqlstr, DBconn, 3, 4
    If Not rd.EOF Then Cardnumber = rd.Fields(0).value
    If Trim(Cardnumber) <> "" Then
        For i = 1 To VchLst.Rows - 1
            If VchLst.TextMatrix(i, VchLst.GridColIndex("selcol")) = "Y" Then
                cls_print.Voucher_Print m_login, Cardnumber, VchLst.TextMatrix(i, VchLst.GridColIndex("id")), , , , False
            End If
        Next i
    End If
 
End Sub

