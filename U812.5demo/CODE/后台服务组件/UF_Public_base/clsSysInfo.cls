VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSysInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim cn As Object 'ADODB.Connection

Public Function Init(oLogin As Object, ByRef strErr As String) As Boolean
    Init = False
    strErr = ""
    On Error GoTo Errhandler
    Set cn = CreateObject("ADODB.Connection")
    
    cn.Open oLogin.UfDbName
    
    Init = True
    Exit Function
Errhandler:
    strErr = Err.Description
    
End Function

'返回制定单据类型的数据表和视图名称
Public Function GetVoucherTable(strCardNum As String, ByRef strErr As String, Optional Conn As Object, Optional ByRef strHeadTable As String, Optional ByRef strBodyTable As String, Optional ByRef strHeadView As String, Optional ByRef strBodyView As String, Optional ByRef strListView As String, Optional ByRef blnHasBody As Boolean) As Boolean
    GetVoucherTable = False
    strErr = ""
    
    On Error GoTo Errhander
    
    Dim rs As New ADODB.Recordset
    
    If Not Conn Is Nothing Then
        Set cn = Conn
    End If
    
    rs.Open "select * from vouchers_base where cardnumber = '" & strCardNum & "'", cn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
        strHeadTable = "" & rs!BTTblName
        strHeadView = "" & rs!BTQName
        strBodyTable = "" & rs!BWTblName
        strBodyView = "" & rs!BWQName
        strListView = "" & rs!VchListQName
        blnHasBody = CBool(Val("" & rs!HaveBodyGrid))
    End If
    rs.Close
    
    Set rs = Nothing
    GetVoucherTable = True
    Exit Function
Errhander:
    strErr = Err.Description
End Function
