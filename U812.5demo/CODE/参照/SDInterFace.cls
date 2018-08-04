VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SDInterFace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function RunCommand(ByVal objLogin As Object, ByVal objForm As Object, ByVal objVoucher As Object, _
                           ByVal sKey As String, ByVal VarentValue As Variant, ByVal other As String)

    Set g_Conn = Nothing
    g_Conn.ConnectionString = objLogin.UfDbName
    g_Conn.CursorLocation = adUseClient
    g_Conn.Open
    
    Call ButtonClick(sKey, "", objForm, objVoucher)

End Function

'初始化
Public Function Init(ByVal objLogin As Object, ByVal objForm As Object, ByVal objVoucher As Object, msbar As Object)
    Set g_Conn = Nothing
    g_Conn.ConnectionString = objLogin.UfDbName
    g_Conn.CursorLocation = adUseClient
    g_Conn.Open
End Function

Private Function Getxsddly() As String
    Dim rs As New ADODB.Recordset
    Dim strSql As String
    strSql = "select cValue from accinformation where csysid='mt' and cname='xsddly' "
    rs.Open strSql, g_Conn
    If Not rs.BOF And Not rs.EOF Then
        Getxsddly = rs!cValue
    End If
    Set rs = Nothing
End Function


'在执行系统按钮只前调用,Cancel后会不再执行系统方法。
Public Function BeforeRunSysCommand(ByVal objLogin As Object, ByVal objForm As Object, ByVal objVoucher As Object, _
                                    ByVal sKey As String, ByVal VarentValue As Variant, ByRef Cancel As Boolean, ByVal other As String)
    Dim oDomHead As New DOMDocument
    Dim oDomBody As New DOMDocument

    Set g_oLogin = objVoucher.LoginObj

    objVoucher.getVoucherDataXML oDomHead, oDomBody
'    oDomHead.Save "c:\head.xml"
'    oDomBody.Save "c:\body.xml"

    Set g_Conn = Nothing
    g_Conn.ConnectionString = g_oLogin.UfDbName
    g_Conn.CursorLocation = adUseClient
    g_Conn.Open
    
Finish:
    Set oDomHead = Nothing
    Set oDomBody = Nothing
    Set g_oLogin = Nothing
    Set g_Conn = Nothing
End Function

'在执行系统按钮只前调用,Cancel后会不再执行系统方法。
Public Function AfterRunSysCommand(ByVal objLogin As Object, ByVal objForm As Object, ByVal objVoucher As Object, _
                                    ByVal sKey As String, ByVal VarentValue As Variant, ByRef Cancel As Boolean, ByVal other As String)
    
    Set g_Conn = Nothing
    g_Conn.ConnectionString = objLogin.UfDbName
    g_Conn.CursorLocation = adUseClient
    g_Conn.Open
    
    Call ButtonClick(sKey, VarentValue, objForm, objVoucher)
End Function

Public Function ButtonClick(sButtonKey As String, VarentValue As Variant, oForm As Object, octlVoucher As Object)

    Dim bSuccess As Boolean
    Dim sWhere As String
    Dim oDomHead As New DOMDocument
    Dim oDomBody As New DOMDocument
    
    Set g_oLogin = octlVoucher.LoginObj
 
    octlVoucher.getVoucherDataXML oDomHead, oDomBody

    Select Case sButtonKey
        Case "_CUSTDEFbtnRefBom" '参照设计清单
            iSinvCZ = True
             If ReferVouch Then
                '调用添加按钮
                Call ProcessData(octlVoucher)
            End If
        Case "_CUSTDEFbtnXMXQJH" '参照项目需求计划
            If GetHeadItemValue(oDomHead, "cwhname") = "" Then
                MsgBox "请先录入仓库！", vbInformation, "提示"
            Else
                If ReferVouch2 Then
                    '调用添加按钮
                    Call ProcessData2(octlVoucher)
                End If
            End If
            
    End Select
    
    Set oDomHead = Nothing
    Set oDomBody = Nothing
End Function
