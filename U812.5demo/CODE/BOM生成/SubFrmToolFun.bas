Attribute VB_Name = "SubFrmToolFun"
'                       打印/打印预览表单
'
' Argument(s)
'       oConnection     (ADODB.Connection) 数据库连接对象
'       oVoucher        (ctlVoucher) 单据对象
'       sBillNumber     (String) 单据号
'       sTemplateID     (String) 模板号
'       bPreview        [Boolean, False] 标志是否显示预览界面

'单据打印预览功能
Public Sub ExecSubVoucherPrint( _
       ByRef oConnection As ADODB.Connection, _
       ByRef oVoucher As ctlVoucher, _
       ByVal sBillNumber As String, _
       ByVal sTemplateID As String, _
       Optional ByVal bPreview As Boolean = False)
    Dim oField As ADODB.Recordset                          ' 固定文本数据
    Dim oTemplate As ADODB.Recordset                       ' 单据模板数据
    Dim oVoucherTemplate As UFVoucherServer85.clsVoucherTemplate

    Dim sError As String
    Dim oDomHead As DOMDocument
    Dim oDomBody As DOMDocument

    'On Error GoTo Err_Handler


    ' *******************************************************
    ' * 进行本地打印
    '
    Set oVoucherTemplate _
    = CreateObject("UFVoucherServer85.clsVoucherTemplate")

    If oVoucherTemplate Is Nothing Then
        MsgBox GetString("U8.DZ.JA.Res1330"), vbCritical, GetString("U8.DZ.JA.Res030")
        GoTo Exit_Label
    End If

    Set oTemplate = oVoucherTemplate.GetTemplateData2( _
                    conn:=oConnection, _
                    sBillName:=sBillNumber, _
                    vTemplateID:=sTemplateID)

    Set oField = oVoucherTemplate.GetFixedData( _
                 conn:=oConnection.ConnectionString, _
                 vVtid:=sTemplateID)

    oVoucher.VoucherStatus = VSNormalMode
    oVoucher.AutoAggregate
    PrintVoucher = True

    Call oVoucher.PrintVoucher( _
         rsTemplate:=oTemplate, _
         rsField:=oField, _
         bShowPrintViewDlg:=bPreview)

Exit_Label:
    On Error GoTo 0
    Set oDomHead = Nothing
    Set oDomBody = Nothing
    Set oVoucherTemplate = Nothing

    If Not oField Is Nothing Then
        If oField.State = adStateOpen Then _
           Call oField.Close
    End If
    Set oField = Nothing

    If Not oTemplate Is Nothing Then
        If oTemplate.State = adStateOpen Then _
           Call oTemplate.Close
    End If
    Set oTemplate = Nothing

    Exit Sub
Err_Handler:
    Call ShowErrorInfo( _
         sHeaderMessage:=GetString("U8.DZ.JA.Res1340"), _
         lMessageType:=vbInformation, _
         lErrorLevel:=ufsELOnlyHeader _
                    )
    If g_blnDEBUG_MODE Then
        Call ShowDebugForm( _
             bErrorMode:=True, _
             sProcedure:="Sub VoucherPrint of Module modFuncL")
    End If
    GoTo Exit_Label
End Sub
' Precedure             ExportVoucherData2File
' Purpose
'                       导出单据单据数据到指定的文件
'
' Argument(s)
'       oConnection     (ADODB.Connection) 数据库连接对象
'       oVoucher        (ctlVoucher) 单据对象
'       sBillNumber     (String) 单据号
'       sTemplateID     (String) 模板号
'
' Author                Li Hongye
' Created               2005-06-03, 10:50
'
' Revision(s)
'
'   Author      Date        Action              E-mail
'   ----------------------------------------------------------
'   Li Hongye   2005-06-03  创建结构,编写代码   lhye@ufsoft.com.cn
'
Public Sub ExportVoucherDataToFile( _
       ByRef oConnection As ADODB.Connection, _
       ByRef oVoucher As ctlVoucher, _
       ByVal sBillNumber As String, _
       ByVal sTemplateID As String)

    Dim oField As ADODB.Recordset                          ' 固定文本数据
    Dim oTemplate As ADODB.Recordset                       ' 单据模板数据
    Dim oVoucherTemplate As Object

    'On Error GoTo Err_Handler

    Set oVoucherTemplate _
    = CreateObject("UFVoucherServer85.clsVoucherTemplate")

    If oVoucherTemplate Is Nothing Then
        MsgBox GetString("U8.DZ.JA.Res1330"), vbCritical, GetString("U8.DZ.JA.Res030")
        GoTo Exit_Label
    End If

    Set oTemplate = oVoucherTemplate.GetTemplateData2( _
                    conn:=oConnection, _
                    sBillName:=sBillNumber, _
                    vTemplateID:=sTemplateID)

    Set oField = oVoucherTemplate.GetFixedData( _
                 conn:=oConnection.ConnectionString, _
                 vVtid:=sTemplateID)

    Call oVoucher.ExportToFile( _
         rsTemplate:=oTemplate, _
         rsField:=oField)

Exit_Label:
    On Error GoTo 0
    Set oVoucherTemplate = Nothing

    If Not oField Is Nothing Then
        If oField.State = adStateOpen Then _
           Call oField.Close
    End If
    Set oField = Nothing

    If Not oTemplate Is Nothing Then
        If oTemplate.State = adStateOpen Then _
           Call oTemplate.Close
    End If
    Set oTemplate = Nothing

    Exit Sub
Err_Handler:
    Call ShowErrorInfo( _
         sHeaderMessage:=GetString("U8.DZ.JA.Res1350"), _
         lMessageType:=vbInformation, _
         lErrorLevel:=ufsELOnlyHeader _
                    )
    GoTo Exit_Label

End Sub




