VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_ShowVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim Conn As New ADODB.Connection
Dim strConn As String
Dim rs As New ADODB.Recordset
Dim strSQL As String

Public Function RunCommand(ByVal objLogin As Object, ByVal objForm As Object, ByVal objVoucher As Object, _
                            ByVal sKey As String, ByVal VarentValue As Variant, ByVal other As String)
    
    Dim menuId As String
    Dim name As String
    Dim subSysId As String
    Dim authid As String
    Dim voucherID As String
    Dim row As Integer  '第几行
    Dim rowsel As Integer '选择多少行
    Dim bodyrows As String '实际数据有多少行
    Dim CBATCH As String '批号
    Dim CINVCODE As String '存货编码
    
    '菜单ID
    menuId = "QM020202"
    '单据名称
    name = "产品检验单"
    '子产品号
    subSysId = "QM"
    '权限ID
    authid = "QM02020201"
    '单据ID
    voucherID = 0
    
    If objVoucher.headerText("cdlcode") = "" Then
        MsgBox "当前无对应单据"
        Exit Function
    End If
    rowsel = objVoucher.rowsel
    row = objVoucher.row
    bodyrows = objVoucher.bodyrows
    CBATCH = objVoucher.bodyText(1, "CBATCH")
    CINVCODE = objVoucher.bodyText(1, "CINVCODE")
    
    If row <= bodyrows Then
        If rowsel > 1 Then Exit Function
        If CBATCH = "" Then
        MsgBox "该存货非批次管理，无法查找到对应的检验单"
        Exit Function
        Else
           voucherID = getQM(CBATCH, CINVCODE, objLogin)
           If voucherID = "" Then
            MsgBox "所选存货无对应的检验单，请检查......"
            Exit Function
           End If
           Call objForm.CTBCtrl1.Business.ResetCallFunctionFlag
           Dim strNewCommandline As String
           strNewCommandline = "ID:" & menuId & "&&&Name:" & name & "&&&SubSysID:" & subSysId & "&&&AuthID:" & authid & "&&&CmdLine:" & vbTab & voucherID
           Call objForm.CTBCtrl1.Business.RunCommand(strNewCommandline)
        End If
    End If
End Function

Function getQM(ByVal qm As String, ByVal code As String, ByVal objLogin As Object)
    strConn = objLogin.UFDbName
     '   使用客户端数据游标
    Conn.CursorLocation = adUseClient
    '   打开Access的连接
    Conn.Open strConn
    '   如果RecordSet的状态不是关闭状态,则关闭Recordset
    If rs.State <> adStateClosed Then rs.Close
    '   Recordser打开表People
    'Summary收发子表
    strSQL = "select id from QMCheckVoucher where ccheckcode in (select son.cCheckCode from rdrecords son LEFT OUTER JOIN rdrecord main ON son.id = main.id "
    strSQL = strSQL & " where main.cvouchtype = 10 and son.cinvcode ='" & code & "' and cbatch = '" & qm & "') and id in (select id from QMCheckVouchers where id in ("
    strSQL = strSQL & "select son.iCheckIdBaks from rdrecords son LEFT OUTER JOIN rdrecord main ON son.id = main.id where main.cvouchtype = 10 and son.cinvcode ='"
    strSQL = strSQL & code & "' and cbatch = '" & qm & "' ))"
'    Rs.Open "select * from QMCheckVoucher where CBATCH = '" & qm & "' and CINVCODE = " & "'" & code & "' and CSOURCE = '生产订单'", Conn, adOpenKeyset, adLockOptimistic
    rs.Open strSQL, Conn, adOpenKeyset, adLockOptimistic
   While Not rs.EOF
        getQM = rs("id")
        rs.MoveNext
   Wend
End Function

'初始化
Public Function Init(ByVal objLogin As Object, ByVal objForm As Object, ByVal objVoucher As Object, msbar As Object)
'MsgBox "Init"
End Function

'在执行系统按钮只前调用,Cancel后会不再执行系统方法。
Public Function BeforeRunSysCommand(ByVal objLogin As Object, ByVal objForm As Object, ByVal objVoucher As Object, _
                            ByVal sKey As String, ByVal VarentValue As Variant, ByRef Cancel As Boolean, ByVal other As String)
'    MsgBox "Save is Canceled"
    Dim ccode As String
    Dim rs As New ADODB.Recordset
    Dim sql As String
    
    On Error GoTo err
    
     '   使用客户端数据游标
    Conn.CursorLocation = adUseClient
    '   打开Access的连接
    Conn.Open objLogin.UFDbName
    
'    Conn = objLogin.UFDbName
    
    ccode = objVoucher.headerText("cdlcode")
    
    If sKey = "print" Then
        sql = "update DispatchList set cDefine5 = isnull(cDefine5,0) + 1 where cdlcode ='" & ccode & "'"
        Conn.Execute sql
        sql = "select cDefine5 from DispatchList where cdlcode ='" & ccode & "'"
        rs.Open sql, Conn, adOpenKeyset, adLockOptimistic
        If Not rs.BOF And Not rs.EOF Then
            objVoucher.headerText("cdefine5") = CStr(rs.Fields("cdefine5"))
            rs.Close
        End If
    End If
    Cancel = False
    Exit Function
err:
      MsgBox err.Description
      Cancel = True
End Function
