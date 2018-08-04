VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVoucher_Print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Private m_oDataSource As Object

Private WithEvents Voucher As UAPVoucherControl85.ctlVoucher
Attribute Voucher.VB_VarHelpID = -1


Public Sub Voucher_Print(M_login As Object, CardNumber As String, VoucherId As String, Optional DomHead As DOMDocument, Optional DomBody As DOMDocument, Optional vt_id As Long, Optional SetPrintDialog As Boolean = True, Optional VoucherIDname As String = "id")
 Dim HKEYstr As String
 Dim HKEYVale As String
 HKEYstr = "HKEY_LOCAL_MACHINE\SOFTWARE\Ufsoft\WF\V8.700\Portal\NoShowPrinterDlg"
    If SetPrintDialog Then
       Call regread(HKEYstr, HKEYVale)
       If Trim(HKEYVale) = "1" Then regwrite HKEYstr, "0"
    Else
       regwrite HKEYstr, "1"
    End If
    PrintVoucher M_login, CardNumber, VoucherId, DomHead, DomBody, False, vt_id, VoucherIDname
    
    '恢复打印原始设置
    regwrite HKEYstr, HKEYVale
    
End Sub

Public Sub Voucher_Preview(M_login As Object, CardNumber As String, VoucherId As String, Optional DomHead As DOMDocument, Optional DomBody As DOMDocument, Optional vt_id As Long, Optional VoucherIDname As String = "id")
    PrintVoucher M_login, CardNumber, VoucherId, DomHead, DomBody, True, vt_id, VoucherIDname
End Sub


Private Sub PrintVoucher(M_login As Object, CardNumber As String, VoucherId As String, DomHead As DOMDocument, DomBody As DOMDocument, Prints As Boolean, Optional Print_vt_id As Long = 0, Optional VoucherIDname As String = "id")
Dim frm As New Frm_voucher
Dim domh As New DOMDocument
Dim domB As New DOMDocument
Dim VoucherServer As New UFVoucherServer85.clsVoucherTemplate
Dim recPrint As ADODB.Recordset
Dim recTail As ADODB.Recordset
Dim rds  As New ADODB.Recordset
Dim sqlstr As String
On Error Resume Next '
    If Conn Is Nothing Then
        Set Conn = New ADODB.Connection
        Conn.Open M_login.UfDbName
    End If
    If m_logins Is Nothing Then
        Set m_logins = M_login
    End If
    '取得默认打印模板
    If Print_vt_id = 0 Then
        sqlstr = "select isnull(DEF_ID,0) as DEF_ID,isnull(DEF_ID_PRN,0) as DEF_ID_PRN   From  vouchers_base where cardnumber='" & CardNumber & "' "
        If rds.State <> 0 Then rds.Close
        rds.Open sqlstr, Conn, 3, 4
        Print_vt_id = rds.fields("DEF_ID_PRN").Value
    End If
    
    '根据主表关键字取得打印数据
    If Trim(VoucherId) <> "" Then
        sqlstr = "select isnull(BTQName,'') as BTQName,isnull(BWQName,'') as BWQName    From  vouchers_base where cardnumber='" & CardNumber & "' "
        If rds.State <> 0 Then rds.Close
        rds.Open sqlstr, Conn, 3, 4
        
        sqlstr = "select * from " & rds.fields("BTQName").Value & " where " & VoucherIDname & "='" & VoucherId & "'"
        If rds.State <> 0 Then rds.Close
        rds.Open sqlstr, Conn, 3, 4
        Set DomHead = New DOMDocument
        rds.save DomHead, adPersistXML
        
        sqlstr = "select * from " & rds.fields("BWQName").Value & " where " & VoucherIDname & "='" & VoucherId & "'"
        If rds.State <> 0 Then rds.Close
        rds.Open sqlstr, Conn, 3, 4
        Set DomBody = New DOMDocument
        rds.save DomBody, adPersistXML
    End If
     
    
    Set Voucher = Frm_voucher.ctlVoucher1
    Set m_oDataSource = CreateObject("IDataSource.DefaultDataSource")
'    If m_oDataSource Is Nothing Then
'        MsgBox "创建单据数据源对象Fail!", vbCritical, "错误"
'    End If
    Set m_oDataSource.SetLogin = M_login
    Voucher.LoginObj = M_login
    Voucher.InitDataSource
    Set recPrint = VoucherServer.GetTemplateData2(Conn, CardNumber, Print_vt_id)
    Set recTail = VoucherServer.GetFixedData(Conn, Print_vt_id)
    Voucher.setTemplateData recPrint

    Voucher.setVoucherDataXML DomBody, DomBody '
    Voucher.ExamineFlowAuditInfo = "GET"
    Voucher.PrintVoucher recPrint, recTail, Prints
End Sub

Private Sub Voucher_SaveSettingEvent(ByVal varDevice As Variant)
    Dim clsPrintVch As UFVoucherServer85.clsVoucherTemplate
    Set clsPrintVch = New UFVoucherServer85.clsVoucherTemplate
    If clsPrintVch.SaveDeviceCapabilities(Conn, vt_ids, varDevice) <> 0 Then
        MsgBox "打印设置保存失败！", vbCritical, Msg_Title
    End If
End Sub


Public Sub regwrite(ByVal HKEYstr As String, ByVal HKEYVale As String, Optional Types As String = "REG_SZ")
Dim wsh As Object
    Set wsh = CreateObject("Wscript.Shell") '利用wshshell写注册表
    wsh.regwrite HKEYstr, HKEYVale, Types
End Sub

Public Sub regread(ByVal HKEYstr As String, ByRef HKEYVale As String)
Dim wsh As Object
    Set wsh = CreateObject("Wscript.Shell") '利用wshshell读注册表
        HKEYVale = wsh.regread(HKEYstr)
'    End If
End Sub
