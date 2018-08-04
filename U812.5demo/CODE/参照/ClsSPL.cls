VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSPL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExecSplCancelConfirm(conn As Connection, lngVoucherID As Long, strErr As String) As Boolean
    
    
    Dim sCreateType As String
    
    Dim oRd As New ADODB.Recordset
    oRd.CursorLocation = adUseClient
    
    oRd.Open "select isnull(cCreateType,'') as  cCreateType from HY_DZ_BorrowOut with (nolock) where id =" & lngVoucherID, conn, adOpenDynamic, adLockReadOnly
    If oRd.RecordCount > 0 Then
       sCreateType = vFieldVal(oRd.Fields("cCreateType"))
       oRd.Close
       Set oRd = Nothing
    Else
       ExecSplCancelConfirm = False
       strErr = "单据不存在!"
       oRd.Close
       Set oRd = Nothing
       Exit Function
    End If
    
    
    If g_Conn Is Nothing Then
        Set g_Conn = New ADODB.Connection
    End If
    
    If g_Conn.State = 0 Then
       Set g_Conn = conn
    End If
    
    If sCreateType = "期初单据" Then
    
       
       If VoucherIsCreate2(lngVoucherID) Then
            ExecSplCancelConfirm = False
            strErr = GetString("U8.DZ.JA.Res250")
            Exit Function
       End If
    Else
       
       If VoucherIsCreate(lngVoucherID) Then
            ExecSplCancelConfirm = False
            strErr = GetString("U8.DZ.JA.Res250")
            Exit Function
       End If
    
    End If
    
   
'    If VoucherIsCreate(lngVoucherID) Then
'        ExecSplCancelConfirm = False
'        strErr = GetString("U8.DZ.JA.Res250")
'        Exit Function
'    End If
    Dim strSql As String
    Dim lEffectedCount As Integer
    strSql = "update hy_dz_borrowout set iverifystate=1,iStatus=1 , cHandler='', dVeriDate=NULL where ID=" & lngVoucherID
    
    If sCreateType = "期初单据" Then
       ' strSql = strSql + " update hy_dz_borrowouts set iQtyOut =0,iQtyOut2 = 0 where id =" & lngVoucherID
    End If
    
    conn.Execute strSql, lEffectedCount
    If lEffectedCount > 0 Then
        ExecSplCancelConfirm = True
    Else
        ExecSplCancelConfirm = False
    End If
End Function
