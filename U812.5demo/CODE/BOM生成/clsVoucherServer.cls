VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVouchServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'单据业务类

Option Explicit

'当前页码
Public PageCurrent As Long

Public aa As Long

'设置单据摸板格式
Public Function SetVouchStyle(ByRef conn As Connection, ByRef Vouch As Object, CardNum As String, Optional ShowID As String) As Boolean
    On Error GoTo ErrHand
    Dim TempID As String
    Dim rs As New ADODB.Recordset
    Dim clr1 As Long
    Dim clr2 As Long
    Dim TempSrv As New clsVoucherTemplate

    If ShowID = "" Then
        TempID = GetVouchTempID(conn, CardNum)
    Else
        TempID = ShowID
    End If

    '取得Rule颜色
    Call TempSrv.GetRuleColor(conn, clr1, clr2)

    Call Vouch.SetRuleColor(clr1, clr2)

    '获得单据摸板数据
    Set rs = GetVouchTempRs(conn, CardNum, TempID)
    '设置单据摸板
    Vouch.Visible = False

    Call Vouch.SetTemplateData(rs)

    Vouch.Visible = True
    SetVouchStyle = True
    Exit Function
ErrHand:
    SetVouchStyle = False
End Function

'获的单据显示摸板号
Private Function GetVouchTempID(ByRef conn As Connection, CardNum As String) As String
    Dim strSql As String
    Dim rs As New ADODB.Recordset
    Dim TempID As String

    strSql = "SELECT VT_ID From VoucherTemplates WHERE (VT_CardNumber = '" & CardNum & "') AND (VT_TemplateMode = '0')"
    rs.Open strSql, conn
    If Not rs.EOF Then
        TempID = rs!VT_ID
    End If

    rs.Close
    Set rs = Nothing
    GetVouchTempID = TempID
End Function
'获的单据打印摸板号
Private Function GetVouchPrintTempID(ByRef conn As Connection, CardNum As String) As String
    Dim strSql As String
    Dim rs As New ADODB.Recordset
    Dim TempID As String

    strSql = "SELECT VT_ID From VoucherTemplates WHERE (VT_CardNumber = '" & CardNum & "') AND (VT_TemplateMode = '1')"
    rs.Open strSql, conn
    If Not rs.EOF Then
        TempID = rs!VT_ID
    End If

    rs.Close
    Set rs = Nothing
    GetVouchPrintTempID = TempID
End Function
'获的单据摸板数据集
Private Function GetVouchTempRs(ByRef conn As Connection, CardNum As String, TempID As String) As Recordset
    Dim TempSrv As New clsVoucherTemplate
    Set GetVouchTempRs = TempSrv.GetTemplateData2(conn, CardNum, TempID)
    'Set GetVouchTempRs = TempSrv.GetTemplateData2(conn, 17, 95)
End Function


'填充单据数据
Public Function SetVouchXML(ByRef conn As Connection, ByRef Vouch As Object, CardNum As String, PKField As String, PKValue As String) As Boolean
    On Error GoTo ErrHand
    Dim strWhere As String
    ' 表头XML对象
    Dim domHead As DOMDocument
    '表体XML对象
    Dim domBody As DOMDocument

    strWhere = PKField & "='" & PKValue & "'"
    Set domHead = GetDomHead(conn, CardNum, strWhere)
    Set domBody = GetDomBody(conn, CardNum, strWhere)

    Vouch.setVoucherDataXML domHead, domBody

    SetVouchXML = True
    Exit Function
ErrHand:
    SetVouchXML = False
End Function

''填充单据数据
'Public Function SetVouchData(ByRef Conn As Connection, ByRef Vouch As Object, CardNum As String, PKField As String, PKValue As String) As Boolean
'On Error GoTo ErrHand
'    Dim strWhere As String
'    ' 表头XML对象
'    Dim dataHead As CDataVO
'    '表体XML对象
'    Dim dataBody As CDataVO
'
'    strWhere = PKField & "='" & PKValue & "'"
'    Set dataHead = GetDataHead(Conn, CardNum, strWhere)
'    Set dataBody = GetDataBody(Conn, CardNum, strWhere)
'    FillVouchHead Vouch, dataHead
'    FillVouchBody Vouch, dataBody
'
'    SetVouchData = True
'    Exit Function
'ErrHand:
'    SetVouchData = False
'End Function
'
''填充单据表头
'Private Sub FillVouchHead(ByRef Vouch As Object, Data As CDataVO)
'    Dim I As Integer
'    If Data Is Nothing Then Exit Sub
'    If Data.Count = 0 Then Exit Sub
'    Dim sField As String
'    For I = 1 To Data.Item(1).Count
'        sField = Data.Item(1).Item(I).FieldCode
'        Vouch.headerText(sField) = Data.Item(1).Item(I).Value
'    Next I
'End Sub
'
''填充单据表体
'Private Sub FillVouchBody(ByRef Vouch As Object, Data As CDataVO)
'    Dim I As Integer, j As Integer
'    If Data Is Nothing Then Exit Sub
'    If Data.Count = 0 Then Exit Sub
'    Dim sField As String
'    For I = 1 To Data.Count
'        For j = 1 To Data.Item(I).Count
'            sField = Data.Item(I).Item(j).FieldCode
'            Vouch.bodyText(I, sField) = Data.Item(I).Item(j).Value
'        Next j
'    Next I
'End Sub
'获得表头XML数据
'原来是private
Public Function GetDomHead(ByRef conn As Connection, CardNum As String, strWhere As String) As DOMDocument
'表头视图
    Dim HeadView As String
    Dim strSql As String
    Dim rs As New ADODB.Recordset

    ' 表头XML对象
    Dim domHead As New DOMDocument

    strSql = "SELECT  BTQName"
    strSql = strSql & " From Vouchers"
    strSql = strSql & " WHERE (CardNumber = '" & CardNum & "') "
    rs.Open strSql, conn
    If Not rs.EOF Then
        HeadView = IIf(IsNull(rs!BTQName), "", rs!BTQName)
    Else
        Set GetDomHead = domHead
        Exit Function
    End If
    rs.Close
    Set rs = Nothing

    If Trim(HeadView) = "" Then
        Set GetDomHead = domHead
        Exit Function
    End If
    '获得表头XML数据
    strSql = "Select *,'' as editprop From  " & HeadView & " Where " & strWhere
    Set rs = conn.Execute(strSql)
    rs.Save domHead, adPersistXML
    rs.Close

    Set rs = Nothing
    Set GetDomHead = domHead
    Exit Function
ErrHand:
    Set GetDomHead = domHead
End Function

'获得表体XML数据
'原来为private
Public Function GetDomBody(ByRef conn As Connection, CardNum As String, strWhere As String) As DOMDocument
'表体视图
    Dim BodyView As String
    Dim strSql As String
    Dim rs As New ADODB.Recordset
    Dim HaveBodyGrid As String
    ' 表体XML对象
    Dim domBody As New DOMDocument

    strSql = "SELECT  BWQName, HaveBodyGrid"
    strSql = strSql & " From Vouchers"
    strSql = strSql & " WHERE (CardNumber = '" & CardNum & "') "
    rs.Open strSql, conn
    If Not rs.EOF Then
        BodyView = IIf(IsNull(rs!BWQName), "", rs!BWQName)
        HaveBodyGrid = IIf(IsNull(rs!HaveBodyGrid), "0", rs!HaveBodyGrid)
    Else
        Set GetDomBody = domBody
        Exit Function
    End If
    rs.Close
    Set rs = Nothing

    If Trim(BodyView) = "" Or HaveBodyGrid = "False" Then
        Set GetDomBody = domBody
        Exit Function
    End If

    '********************************870*******************************************
    'strSql = "Select * From  " & BodyView & " Where " & strWhere

    '================================861===========================================
    strSql = "Select *,'' as editprop From  " & BodyView & " Where " & strWhere

    Set rs = conn.Execute(strSql)
    rs.Save domBody, adPersistXML
    rs.Close
    Set rs = Nothing
    Set GetDomBody = domBody
    Exit Function
ErrHand:
    Set GetDomBody = domBody
End Function


'单据打印
Public Sub VouchPrint(conn As ADODB.Connection, Vouch As Object, CardNum As String, Optional PrintTempID As String)
    Dim o As Object
    Dim rsTemplate As Recordset
    Dim rsField As Recordset
    Set o = CreateObject("UFVoucherServer85.clsVoucherTemplate")
    If o Is Nothing Then Exit Sub   '如果创建对象Fail,则退出
    If PrintTempID = "" Then
        Dim PrintTmpID As String
        PrintTmpID = GetVouchPrintTempID(g_Conn, CardNum)
        Set rsTemplate = o.GetTemplateData2(g_Conn, CardNum, PrintTmpID)
        Set rsField = o.GetFixedData(conn.ConnectionString, PrintTmpID)
    Else
        Set rsTemplate = o.GetTemplateData2(conn, CardNum, PrintTempID)
        Set rsField = o.GetFixedData(conn.ConnectionString, PrintTempID)
    End If
    Call Vouch.PrintVoucher(rsTemplate, rsField, False)

End Sub
'单据预览
Public Sub VouchPreview(conn As ADODB.Connection, Vouch As Object, CardNum As String, Optional PrintTempID As String)
    Dim o As Object
    Dim rsTemplate As Recordset
    Dim rsField As Recordset
    Set o = CreateObject("UFVoucherServer85.clsVoucherTemplate")
    If o Is Nothing Then Exit Sub   '如果创建对象Fail,则退出
    If PrintTempID = "" Then
        Dim PrintTmpID As String
        PrintTmpID = GetVouchPrintTempID(g_Conn, CardNum)
        Set rsTemplate = o.GetTemplateData2(conn, CardNum, PrintTmpID)
        Set rsField = o.GetFixedData(conn.ConnectionString, PrintTmpID)
    Else
        Set rsTemplate = o.GetTemplateData2(conn, CardNum, PrintTempID)
        Set rsField = o.GetFixedData(conn.ConnectionString, PrintTempID)
    End If
    Call Vouch.PrintVoucher(rsTemplate, rsField, True)

End Sub
'单据输出
Public Sub VouchOutput(conn As ADODB.Connection, Vouch As Object, CardNum As String, Optional PrintTempID As String)
    Dim o As Object
    Dim rsTemplate As Recordset
    Dim rsField As Recordset
    Set o = CreateObject("UFVoucherServer85.clsVoucherTemplate")
    If o Is Nothing Then Exit Sub   '如果创建对象Fail,则退出
    If PrintTempID = "" Then
        Dim PrintTmpID As String
        PrintTmpID = GetVouchPrintTempID(g_Conn, CardNum)
        Set rsTemplate = o.GetTemplateData2(conn, CardNum, PrintTmpID)
        Set rsField = o.GetFixedData(conn.ConnectionString, PrintTmpID)
    Else
        Set rsTemplate = o.GetTemplateData2(conn, CardNum, PrintTempID)
        Set rsField = o.GetFixedData(conn.ConnectionString, PrintTempID)
    End If
    Call Vouch.ExportToFile(rsTemplate, rsField)

End Sub

'获得当前页的主键值
Public Function GetPagePKValue(conn As ADODB.Connection, PageNo As Long, CardNum As String, sPKFld As String, Optional ByVal sWhere As String, Optional ByVal sorder As String) As String
    Dim strSql As String, rs As New ADODB.Recordset
    Dim sTable As String
    GetPagePKValue = "0"
    'sTable的值是：视图HY_FZ_VHandWage
    sTable = GetViewHead(conn, CardNum)
    strSql = "SET NOCOUNT ON"
    strSql = strSql & vbCrLf & "DECLARE @TB TABLE(PAGEID INT IDENTITY(1,1),ID VARCHAR(30))"
    strSql = strSql & vbCrLf & "INSERT INTO @TB (ID) SELECT " & sPKFld & " FROM " & sTable
    If sWhere <> "" Then strSql = strSql & " WHERE " & sWhere
    If sorder <> "" Then strSql = strSql & " ORDER BY " & sorder

    strSql = strSql & vbCrLf & "select id from @tb where pageid='" & PageNo & "'"
    rs.Open strSql, conn
    If Not rs.EOF Then
        'GetPagePKValue 为3
        GetPagePKValue = IIf(IsNull(rs(0)), 0, rs(0))
    End If
    rs.Close
    Set rs = Nothing
End Function

'获得页数
Public Function GetPageCount(conn As ADODB.Connection, CardNum As String, sPKFld As String, Optional ByVal sWhere As String, Optional ByVal sorder As String) As String
    Dim strSql As String, rs As New ADODB.Recordset
    Dim sTable As String
    sTable = GetViewHead(conn, CardNum)
    strSql = "SET NOCOUNT ON"
    strSql = strSql & vbCrLf & "DECLARE @TB TABLE(PAGEID INT IDENTITY(1,1),ID VARCHAR(30))"
    
    '单据联查 时 按钮状态控制 by zhangwchb 20110809
'    If tmpLinkTbl = "" Then
        strSql = strSql & vbCrLf & "INSERT INTO @TB (ID) SELECT " & sPKFld & " FROM " & sTable
'    Else
'        strSql = strSql & vbCrLf & "INSERT INTO @TB (ID) SELECT " & sTable & "." & sPKFld & " FROM " & sTable
'        strSql = strSql & vbCrLf & " inner join " & tmpLinkTbl & " on " & sTable & "." & sPKFld & " = " & tmpLinkTbl & ".id"
'
'    End If
    
    If sWhere <> "" Then strSql = strSql & " WHERE " & sWhere
    If sorder <> "" Then strSql = strSql & " ORDER BY " & sorder

    strSql = strSql & vbCrLf & "select max(pageid) from @tb"
    strSql = strSql & vbCrLf & "set nocount off"
    rs.Open strSql, conn
    If Not rs.EOF Then
        GetPageCount = IIf(IsNull(rs(0)), 0, rs(0))
    End If
    rs.Close
    Set rs = Nothing
End Function



'获得表头视图
Public Function GetViewHead(ByRef conn As Connection, CardNum As String) As String
'表头视图
    Dim HeadView As String

    Dim rs As New ADODB.Recordset

    Dim strSql As String
    strSql = "SELECT  BTQName"
    strSql = strSql & " From Vouchers"
    strSql = strSql & " WHERE (CardNumber = '" & CardNum & "') "
    rs.Open strSql, conn
    If Not rs.EOF Then
        HeadView = IIf(IsNull(rs!BTQName), "", rs!BTQName)
    End If
    rs.Close
    Set rs = Nothing
    GetViewHead = HeadView
End Function


'获得表体视图
Public Function GetViewBody(ByRef conn As Connection, CardNum As String) As String
'表体视图
    Dim BodyView As String
    Dim strSql As String

    Dim rs As New ADODB.Recordset

    strSql = "SELECT  BWQName, HaveBodyGrid"
    strSql = strSql & " From Vouchers"
    strSql = strSql & " WHERE (CardNumber = '" & CardNum & "') "
    rs.Open strSql, conn
    If Not rs.EOF Then
        BodyView = IIf(IsNull(rs!BWQName), "", rs!BWQName)
    End If
    rs.Close
    Set rs = Nothing
    GetViewBody = BodyView
End Function




