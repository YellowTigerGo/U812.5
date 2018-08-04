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

    Call ButtonClick(sKey, "", objForm, objVoucher)

End Function

'初始化
Public Function Init(ByVal objLogin As Object, ByVal objForm As Object, ByVal objVoucher As Object, msbar As Object)
'MsgBox "Init"
End Function

'在执行系统按钮只前调用,Cancel后会不再执行系统方法。
Public Function BeforeRunSysCommand(ByVal objLogin As Object, ByVal objForm As Object, ByVal objVoucher As Object, _
                                    ByVal sKey As String, ByVal VarentValue As Variant, ByRef Cancel As Boolean, ByVal other As String)
    MsgBox "Save is Canceled"
    Cancel = True
End Function

Public Function ButtonClick(sButtonKey As String, sMenuKey As String, oForm As Object, octlVoucher As Object)

    Dim bSuccess As Boolean
    Dim sWhere As String
    Dim oDomHead As New DOMDocument
    Dim oDomBody As New DOMDocument

    Set g_oLogin = octlVoucher.LoginObj

    octlVoucher.getVoucherDataXML oDomHead, oDomBody


'    Dim xoctlVoucher As Object
'
'    Set xoctlVoucher = CreateObject("UAPVoucherControl85.ctlVoucher")
'
'    xoctlVoucher.setVoucherDataXML oDomHead, oDomBody



    Set g_Conn = Nothing
    g_Conn.ConnectionString = g_oLogin.UfDbName
    g_Conn.CursorLocation = adUseClient
    g_Conn.Open

    ' GetMjInvDefine

    Select Case sButtonKey
    Case "btnpro", "_CUSTDEFbtnpro"
        '    获取权限字段
    Call getAuthString(g_Conn)
    iSinvCZ = True
        
    End Select
End Function
