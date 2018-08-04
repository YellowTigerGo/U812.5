VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVoucherUIPlugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private g_business As Object
Private m_uftoolbar As Object
'Private DBConn As ADODB.Connection
Public frmContractParent As Object

'初始化
Public Function init(ByVal objLogin As Object, ByVal objForm As Object, ByVal objVoucher As Object, msbar As Object)
    If DBconn Is Nothing Then
        Set DBconn = New ADODB.Connection
        DBconn.ConnectionString = objLogin.UfDbName
        DBconn.CursorLocation = adUseClient
        DBconn.Open
    End If
    If TypeName(objForm) = "frmContractParent" Then '初始化时如果是合同工作台主窗体
        Set frmContractParent = objForm
    End If
End Function

Public Function InitEx(uftoolbar As Object, other As Variant)
    '初始化功能，在窗体Toolbar初始化时执行一次
    'uftoolbar代表窗体上的UFToolbar对象
    Set m_uftoolbar = uftoolbar
'    m_uftoolbar.FireCommand "Add"
End Function

'执行按钮，objVoucher为单据或列表对象
'VarentValue为在表中预置的cVariant的值。
Public Function RunCommand(ByVal objLogin As Object, ByVal objForm As Object, ByVal objVoucher As Object, _
                            ByVal Skey As String, ByVal VarentValue As Variant, ByVal other As String)
    
    
    Dim i As Long
    Dim id As String
    Dim bClose As Boolean
    If DBconn Is Nothing Then
        Set DBconn = New ADODB.Connection
        DBconn.Open objLogin.UfDbName
    End If
    If g_oLogin Is Nothing Then
       Set g_oLogin = New U8Login.clsLogin
       Set g_oLogin = objLogin
    End If
    Dim cCmdLine As String
'    Set g_business = m_uftoolbar.Business
    If Skey = "_CUSTDEFbtnClose" Then
        bClose = True
    Else
        bClose = False
    End If
    For i = 1 To objVoucher.rows - 1
        If objVoucher.TextMatrix(i, objVoucher.GetColIndex("selcol")) = "Y" Then
            id = objVoucher.TextMatrix(i, objVoucher.GetColIndex("id"))
            If Not DoClose(bClose, id) Then
                Exit For
            End If
        End If
    Next
    m_uftoolbar.FireSysCommand enumButton, "tlbRefresh"
End Function

Private Function DoClose(bClose As Boolean, id As String) As Boolean
    Dim strsql As String
    On Error GoTo hErr
    If bClose Then
        strsql = "update RdRecord01 set cDefine13='是' where id='" & id & "'"
    Else
        strsql = "update RdRecord01 set cDefine13='否' where id='" & id & "'"
    End If
    DBconn.Execute strsql
    DoClose = True
    Exit Function
hErr:
    MsgBox Err.Description, vbCritical, "提示"
End Function

'在执行系统按钮只前调用,Cancel后会不再执行系统方法。
Public Function BeforeRunSysCommand(ByVal objLogin As Object, ByVal objForm As Object, ByVal objVoucher As Object, _
                            ByVal Skey As String, ByVal VarentValue As Variant, ByRef Cancel As Boolean, ByVal other As String)
                            
    Dim domHead As New DOMDocument
    Dim domBody As New DOMDocument
    Dim list As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim id As String
    Dim sError As String
    Dim i As Long
    
'    objForm.ctlVoucher.getVoucherDataXML domHead, domBody
'    Set list = domBody.selectNodes("//z:row")
    Select Case Skey
    
        Case "Sure"
            objVoucher.getVoucherDataXML domHead, domBody
            id = GetHeadItemValue(domHead, "ID")
            If Not CheckQC(id, sError) Then
                MsgBox sError, vbCritical, "提示"
                Cancel = True
            Else
                
            End If
'        Case "Save"
'            objVoucher.getVoucherDataXML domHead, domBody
'            ID = GetHeadItemValue(domHead, "ID")
'            Dim cprojectcode As String
'            Dim CPROORDERCODE As String
'            Dim IPROORDERAUTOID As String
'            cprojectcode = GetHeadItemValue(domHead, "CPROJECTCODE")
'            CPROORDERCODE = GetHeadItemValue(domHead, "CPROORDERCODE")
'            IPROORDERAUTOID = GetHeadItemValue(domHead, "IPROORDERAUTOID")
'            If Not Checkcprojectcode(cprojectcode, CPROORDERCODE, IPROORDERAUTOID, sError) Then
'                MsgBox sError, vbCritical, "提示"
'                Cancel = True
'            Else
'
'            End If
        Case "DoVerify"
'            Set domHead = objVoucher.GetListDom(True)
            For i = 1 To objVoucher.rows - 1
                If objVoucher.TextMatrix(i, objVoucher.GetColIndex("selcol")) = "Y" And objVoucher.TextMatrix(i, objVoucher.GetColIndex("CVERIFIER")) = "" Then
                    If InStr(1, id, "'" & objVoucher.TextMatrix(i, objVoucher.GetColIndex("id")) & "'", vbTextCompare) <= 0 Then
                        id = id & "'" & objVoucher.TextMatrix(i, objVoucher.GetColIndex("id")) & "',"
                    End If
                End If
            Next
            If Len(id) > 0 Then
                id = Left$(id, Len(id) - 1)
                If Not CheckQC(id, sError) Then
                    MsgBox sError, vbCritical, "提示"
                    Cancel = True
                End If
            End If
    End Select
    
    
    
ExitDo:
    Set list = Nothing
    Set node = Nothing
    Set domHead = Nothing
    Set domBody = Nothing

End Function

'检验质量检验方案
Private Function Checkcprojectcode(cprojectcode As String, CPROORDERCODE As String, IPROORDERAUTOID As String, ByRef sError As String) As Boolean
    On Error GoTo hErr
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    strsql = "select c.cprojectcode from mom_order t left outer join mom_orderdetail m on t.MoId =m.MoId " & _
            " left outer join EF_CustomerQC c on m.Define24=c.ccode" & _
            " where t.MoCode='" & CPROORDERCODE & "' and m.SortSeq='" & IPROORDERAUTOID & "'"
    rs.Open strsql, DBconn
    If Not rs.EOF And Not rs.BOF Then
        If cprojectcode <> rs!cprojectcode & "" Then
            sError = "当前质量检验方案与客户质量检验方案" & rs!cprojectcode & "不一致。"
            GoTo Finish
        End If
    End If
    Checkcprojectcode = True
    GoTo Finish
hErr:
    sError = Err.Description
Finish:
    If rs.State <> 0 Then
        rs.Close
    End If
    Set rs = Nothing
End Function

Private Function CheckQC(id As String, ByRef sError As String) As Boolean
    On Error GoTo hErr
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    strsql = "select q.CCHECKCODE,qs.CCHKGUIDECODE,qs.CCHECKVALUE,cs.FLOWERLIMIT,cs.FUPPERLIMIT from QMCheckVoucher q" & _
            " left outer join QMCheckVouchers qs on q.ID=qs.ID" & _
            " left outer join mom_orderdetail d on q.IPROORDERID=d.MoId and q.IPROORDERAUTOID=d.SortSeq " & _
            " left outer join EF_CustomerQC c on d.Define24=c.ccode" & _
            " left outer join EF_CustomerQCs cs on c.id=cs.id and qs.CCHKGUIDECODE=cs.CCHKGUIDECODE " & _
            " where ISNUMERIC(ISNULL(qs.CCHECKVALUE,0))=1 and q.ID in (" & id & ")  and( qs.CCHECKVALUE< cast(cs.FLOWERLIMIT as float) OR qs.CCHECKVALUE>cast(cs.FUPPERLIMIT as float)) "
    rs.Open strsql, DBconn
    While Not rs.EOF
        sError = sError & rs!CCHECKCODE & "检验指标" & rs!CCHKGUIDECODE & "检测值" & rs!CCHECKVALUE & "超出客户质量要求上下限(" & rs!FLOWERLIMIT & "-" & rs!FUPPERLIMIT & ")." & vbCrLf
        rs.MoveNext
    Wend
    If Len(sError) > 0 Then
        CheckQC = False
    Else
        CheckQC = True
    End If
    GoTo Finish
hErr:
    sError = Err.Description
Finish:
    If rs.State <> 0 Then
        rs.Close
    End If
    Set rs = Nothing
End Function


Private Function GetVouchid2(MoCode As String, SortSeq As String) As String
    On Error GoTo hErr
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    strsql = "select q.id from mom_order m left outer join mom_orderdetail b on m.MoId=b.MoId " & _
            " left outer join EF_CustomerQC q on b.Define24=q.ccode where m.MoCode='" & MoCode & "' and b.SortSeq='" & SortSeq & "'"
    rs.Open strsql, DBconn
    If Not rs.BOF And Not rs.EOF Then
        GetVouchid2 = rs!id & ""
    End If
    GoTo Finish
hErr:
Finish:
    If rs.State <> 0 Then
        rs.Close
    End If
    Set rs = Nothing
End Function
