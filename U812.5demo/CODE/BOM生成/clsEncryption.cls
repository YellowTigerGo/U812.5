VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEncryption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'/********加密模块***********/
'创建时间 2007-08-30
'创建人 王志军
'modify by chenliangc 2009-3-10


Option Explicit

Private Const Sub_ID = "ST"                '模块
Private Const EncryptCode = "JA"           '加密号

'判断莫开是否启用
Public Function bUsed() As Boolean
    Dim strSql As String
    Dim rs As New ADODB.Recordset
    bUsed = False
    strSql = "select cAcc_ID from ua_account_sub where cAcc_Id = '" & g_oLogin.cAcc_Id & "' and csub_id='" & Sub_ID & "' and dSubSysUsed is not NULL "
    If Len(strSql) > 0 Then
        Set rs = g_Conn.Execute(strSql)
        If Not rs.EOF Then
            If rs(0) > 0 Then
                bUsed = True
            End If
        End If
        rs.Close
    End If
End Function

'校验插件是否加密
Public Function IsOutOfLimit() As Boolean
 
    Dim strSql As String
    Dim rs As New ADODB.Recordset
    Dim lCount As Long
    Dim sMessage As String
     
    On Error GoTo Err_Handler
     
    'IsOutOfLimit = False
    IsOutOfLimit = True

    '检查控制。
    If g_oLogin.TaskExec(EncryptCode, 10) = False Then
    
    
        Dim ErrorInfor As Variant
        If Not g_oBusiness Is Nothing Then
            g_oBusiness.PortalMenuRename "ST02JC02", "#DEMOVERSION#", ErrorInfor
        ElseIf Not goBusiness Is Nothing Then
            goBusiness.PortalMenuRename "ST02JC02", "#DEMOVERSION#", ErrorInfor
        End If
        
       '此处添加对试用期的控制方案,如：未加密时，看板环路条数不能超过15条
        If rs Is Nothing Then Set rs = New ADODB.Recordset
        If rs.State = adStateOpen Then Call rs.Close

        Dim Count As Long 'strSql As String, rs As New ADODB.Recordset,
'        strSql = "select datediff(d,dateadd(m,3,min (dmDate)),max(dmDate)) from HY_DZ_BorrowOut"
'         select isnull(datediff( m,max(dmDate),min(dmDate)),-1) from HY_DZ_BorrowOut

        strSql = "select isnull(datediff(m,min(dmDate),max(dmDate)),-1) from HY_DZ_BorrowOut where  ccreatetype<>N'期初单据'"
        rs.Open strSql, g_Conn
        If Not rs.EOF Then
            Count = IIf(IsNull(rs(0)), 0, rs(0))
        End If
        rs.Close
        If Count > 1 Then
        Else
            IsOutOfLimit = False
            Exit Function
        End If
    Else
        IsOutOfLimit = False
    End If

    If rs.State = adStateOpen Then Call rs.Close
    Set rs = Nothing
    Exit Function
Err_Handler:
    sMessage = GetString("U8.DZ.JA.Res1360")
End Function
