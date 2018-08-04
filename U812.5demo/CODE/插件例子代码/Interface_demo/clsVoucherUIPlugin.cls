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
    Dim ele As IXMLDOMElement
    Dim editprop As String
    Dim id As String
    Dim sError As String
    Dim i As Long
    Dim Msg As String
    Dim autoid As String
    Dim citemcode As String
    Dim cInvCode As String
    Dim cmocode As String
    Dim imoseq As String
    Dim irowno As String
    Dim iquantity As Double
    Dim iOutQty As Double
    Dim iInQty As Double
    
    If DBconn Is Nothing Then
        Set DBconn = New ADODB.Connection
        DBconn.ConnectionString = objLogin.UfDbName
        DBconn.CursorLocation = adUseClient
        DBconn.Open
    End If
'    objForm.ctlVoucher.getVoucherDataXML domHead, domBody
'    Set list = domBody.selectNodes("//z:row")
    Select Case Skey
    
        Case "tlbSave"
            If VarentValue = "0413" Then
'                objVoucher.headerText("cdefine5") = "1"
                objVoucher.getVoucherDataXML domHead, domBody
'                domBody.save "C:\body.xml"
                Set list = domBody.selectNodes("//z:row")
                
                For Each ele In list
                        editprop = GetElementValue(ele, "editprop")
                        If editprop <> "D" Then
                            autoid = GetElementValue(ele, "autoid")
                            irowno = GetElementValue(ele, "irowno")
                            citemcode = GetElementValue(ele, "citemcode")
                            cInvCode = GetElementValue(ele, "cinvcode")
                            If IsXMJ(cInvCode) Then
                                cmocode = GetElementValue(ele, "cmocode")
                                imoseq = GetElementValue(ele, "imoseq")
                                iquantity = str2Dbl(GetElementValue(ele, "iquantity"))
                                If cmocode <> "" Then
                                    iInQty = GetInQty(citemcode, cInvCode, cmocode, imoseq)
                                    If editprop = "A" Then
                                        iOutQty = GetOutQty(citemcode, cInvCode, cmocode, imoseq, "")
                                    Else
                                        iOutQty = iquantity + GetOutQty(citemcode, cInvCode, cmocode, imoseq, autoid)
                                    End If
                                    If iOutQty > iInQty Then
                                        Msg = Msg & "行号:" & irowno & ",存货:" & cInvCode & ",项目:" & citemcode & ",生产订单号-行号:" & cmocode & "-" & imoseq & ",累计领料申请数量" & iOutQty & "大于采购入库量" & iInQty & vbCrLf
                                    End If
                                End If
                            End If
                        End If
                Next
                If Len(Msg) > 0 Then
                    FrmMsgBox.Text1.Text = Msg
                    FrmMsgBox.Show 1
'                    Cancel = True
                End If
            ElseIf VarentValue = "0412" Then
                objVoucher.getVoucherDataXML domHead, domBody
'                domBody.save "C:\body.xml"
                Set list = domBody.selectNodes("//z:row")
                
                For Each ele In list
                        editprop = GetElementValue(ele, "editprop")
                        If editprop <> "D" Then
                            autoid = GetElementValue(ele, "autoid")
                            irowno = GetElementValue(ele, "irowno")
                            citemcode = GetElementValue(ele, "citemcode")
                            cInvCode = GetElementValue(ele, "cinvcode")
                            If IsXMJ(cInvCode) Then
                                cmocode = GetElementValue(ele, "cmocode")
                                imoseq = GetElementValue(ele, "imoseq")
                                iquantity = str2Dbl(GetElementValue(ele, "iquantity"))
                                If cmocode <> "" Then
                                    iInQty = GetInQty(citemcode, cInvCode, cmocode, imoseq)
                                    If editprop = "A" Then
                                        iOutQty = GetOutQty2(citemcode, cInvCode, cmocode, imoseq, "")
                                    Else
                                        iOutQty = iquantity + GetOutQty2(citemcode, cInvCode, cmocode, imoseq, autoid)
                                    End If
                                    If iOutQty > iInQty Then
                                        Msg = Msg & "行号:" & irowno & ",存货:" & cInvCode & ",项目:" & citemcode & ",生产订单号-行号:" & cmocode & "-" & imoseq & ",累计出库数量" & iOutQty & "大于采购入库量" & iInQty & vbCrLf
                                    End If
                                End If
                            End If
                        End If
                Next
                If Len(Msg) > 0 Then
                    FrmMsgBox.Text1.Text = Msg
                    FrmMsgBox.Show 1
'                    Cancel = True
                End If
            ElseIf VarentValue = "MO21" Then
                objVoucher.getVoucherDataXML domHead, domBody
                Set list = domBody.selectNodes("//z:row")
                For Each ele In list
                        editprop = GetElementValue(ele, "editprop")
                        If editprop <> "D" Then
                            citemcode = GetElementValue(ele, "DInvFree_1")
'                            cmocode = GetElementValue(ele, "cmocode")
                            imoseq = GetElementValue(ele, "DSortSeq")
                            If citemcode <> "" Then
                                If ItemStop(citemcode) Then
                                    Msg = Msg & "行号:" & imoseq & "项目:" & citemcode & "已结算。" & vbCrLf
                                End If
                            End If
                        End If
                Next
                If Len(Msg) > 0 Then
                    FrmMsgBox.Text1.Text = Msg
                    FrmMsgBox.Show 1
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

'项目是否结算
Private Function ItemStop(citemcode As String) As Boolean
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo hErr
    strsql = "select * from fitemss97 where bclose=1 and citemcode='" & citemcode & "'"
    rs.Open strsql, DBconn
    If Not rs.EOF And Not rs.BOF Then
        ItemStop = True
    End If
    GoTo Finish
hErr:
    
Finish:
    If rs.State <> 0 Then
        rs.Close
    End If
    Set rs = Nothing
End Function

'是否项目件
Private Function IsXMJ(cInvCode As String) As Boolean
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo hErr
    strsql = "select * from Inventory where cinvdefine9='项目件' and cinvcode='" & cInvCode & "'"
    rs.Open strsql, DBconn
    If Not rs.EOF And Not rs.BOF Then
        IsXMJ = True
    End If
    GoTo Finish
hErr:
    
Finish:
    If rs.State <> 0 Then
        rs.Close
    End If
    Set rs = Nothing
End Function

Private Function GetInQty(citemcode As String, cInvCode As String, cmocode As String, imoseq As String) As Double
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo hErr
    strsql = "select sum(iQuantity) as iqty from rdrecords01" & _
            " where cItemCode='" & citemcode & "' and cInvCode='" & cInvCode & "' AND iordercode='" & cmocode & "' AND iorderseq='" & imoseq & "'"
    rs.Open strsql, DBconn
    If Not rs.EOF And Not rs.BOF Then
        GetInQty = str2Dbl(rs!iQty & "")
    End If
    GoTo Finish
hErr:
    
Finish:
    If rs.State <> 0 Then
        rs.Close
    End If
    Set rs = Nothing
End Function

'领料申请
Private Function GetOutQty(citemcode As String, cInvCode As String, cmocode As String, imoseq As String, autoid As String) As Double
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo hErr
    strsql = "select sum(iQuantity) as iqty from materialappvouchs" & _
            " where cItemCode='" & citemcode & "' and cInvCode='" & cInvCode & "' AND cmocode='" & cmocode & "' AND imoseq='" & imoseq & "'"
    If autoid <> "" Then
        strsql = strsql & " and autoid<>'" & autoid & "'"
    End If
    rs.Open strsql, DBconn
    If Not rs.EOF And Not rs.BOF Then
        GetOutQty = str2Dbl(rs!iQty & "")
    End If
    GoTo Finish
hErr:
    
Finish:
    If rs.State <> 0 Then
        rs.Close
    End If
    Set rs = Nothing
End Function


'材料出库
Private Function GetOutQty2(citemcode As String, cInvCode As String, cmocode As String, imoseq As String, autoid As String) As Double
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo hErr
    strsql = "select sum(iQuantity) as iqty from rdrecords11" & _
            " where cItemCode='" & citemcode & "' and cInvCode='" & cInvCode & "' AND cmocode='" & cmocode & "' AND imoseq='" & imoseq & "'"
    If autoid <> "" Then
        strsql = strsql & " and autoid<>'" & autoid & "'"
    End If
    rs.Open strsql, DBconn
    If Not rs.EOF And Not rs.BOF Then
        GetOutQty2 = str2Dbl(rs!iQty & "")
    End If
    GoTo Finish
hErr:
    
Finish:
    If rs.State <> 0 Then
        rs.Close
    End If
    Set rs = Nothing
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
