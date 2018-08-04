Attribute VB_Name = "modMain_PZ"
Option Explicit
Public frmZD1_Label1_Caption As String
Public Me_Caption As String

Function Msg(prompt As String, Optional buttons As VbMsgBoxStyle) As VbMsgBoxResult
    Dim sTitle As String
    Msg = MsgBox(prompt, buttons, sTitle)       'App.Title
End Function

Public Function FillZd(cType As String, StrZd As String, iAmount1 As String, iAmount2 As String) As Variant
    Dim Rst As ADODB.Recordset, iAmount As Currency
    Dim rst1 As New ADODB.Recordset
    Dim cSql As String
    Dim cFilter As String
    Dim Lp As Long
    Dim sExchRate As String
    Dim intCyc As Integer
    sExchRate = "#,##0."
    For intCyc = 1 To clsSAWeb.GetExchRateDec("美元")
        sExchRate = sExchRate & "0"
    Next
     Set Rst = OpenRecordset(StrZd)
    frmZD1.MousePointer = vbHourglass
    If Len(frmZD1_Label1_Caption) > 4 Then
        Lp = val(Mid(frmZD1_Label1_Caption, 2, Len(frmZD1_Label1_Caption) - 4))
    Else
        Lp = 0
    End If
    If iAmount1 <> "" Then
        cFilter = " And iAmount1 >= " & iAmount1
    End If
    If iAmount2 <> "" Then
        cFilter = cFilter & " And iAmount1 <= " & iAmount2
    End If
    If cFilter <> "" Then
        cFilter = Mid(cFilter, 6, Len(cFilter) - 5)
    End If
    frmZD1_Label1_Caption = "正在读取可制单信息..."
    DoEvents
    If Rst.State <> 0 Then Rst.Close
    Rst.Open StrZd, DBConn, adOpenStatic, adLockReadOnly
    If Rst.RecordCount > 0 Then
    frmZD1_Label1_Caption = "共 " & Rst.RecordCount & " 条"
    Lp = Rst.RecordCount
    End If
    If cFilter <> "" Then
        Rst.Filter = cFilter
    End If
    frmZD1_Label1_Caption = "共" & CStr(Lp) & "条"
    On Error Resume Next
    Dim j As Long
    With Rst
    frmZD1.Grid.rows = 1
    Do While Not .EOF
    
        '"^  选择标志  |^单据编号|^单据日期|^期间|^附单据数|^单据类型|^摘要|^图书编码|^图书名称|^版次|^印次|^总金额"

'        cSql = Chr(9) & frmZD1.cboSign.Text & Chr(9) & !coutsign & Chr(9) & !coutbillsign & Chr(9) & !coutno_id
'        cSql = cSql & Chr(9) & !doutdate & Chr(9) & !value & Chr(9) & !Value_f & Chr(9) & !Rate & Chr(9) & !cdept_id
'        cSql = cSql & Chr(9) & !cPerson_id & Chr(9) & !cItem_Class & Chr(9) & !cItemCode & Chr(9) & !cCode & Chr(9) & !ID
'
            
        cSql = Chr(9) & frmZD1.cboSign.Text & Chr(9) & !cCode & Chr(9) & !ddate & Chr(9) & !iPeriod
        cSql = cSql & Chr(9) & !ibillnum & Chr(9) & !cbillsign & Chr(9) & !cDigest & Chr(9) & !cInvCode & Chr(9) & !cinvname
        cSql = cSql & Chr(9) & !cfree1 & Chr(9) & !cfree2 & Chr(9) & !JE & Chr(9) & !ID
        
        frmZD1.Grid.AddItem cSql
        
        frmZD1.Grid.row = frmZD1.Grid.rows - 1
        For j = 0 To frmZD1.Grid.cols - 1
            frmZD1.Grid.col = j
            frmZD1.Grid.CellForeColor = 0
            
        Next j
    

NextRst:
        .MoveNext
    Loop
    End With
    Rst.Close
    Set Rst = Nothing
    frmZD1.MousePointer = vbDefault
End Function


Function InitAccount() As Boolean
On Error GoTo Err0
    Set Account.UfDbData = New UfDbKit.UfDatabase
    Account.UfDbData.OpenDatabase m_Login.UfDbName 'oAcc.ConnectionStr
    Set Account.dbMData = Account.UfDbData.DbConnect
    Account.dbMData.CommandTimeout = 0
    Account.dbMData.Execute "SET NOCOUNT OFF"
    Account.iExchRateDecDgt = 2
    InitAccount = True
Err0:
    InitAccount = False
End Function

'功能:  求凭证类别编号
'参数:  凭证类别名称
'返回:  凭证类别编号
Function PzlbNameToCode(ByVal cname As String) As String
    Dim Rst As ADODB.Recordset
    Set Rst = OpenRecordset("SELECT * FROM dsign WHERE ctext=N'" & cname & "'")
    If Not Rst.BOF > 0 Then
        PzlbNameToCode = Rst!cSign
    Else
        PzlbNameToCode = ""
    End If
    Rst.Close
End Function

Public Sub Init()
    On Error Resume Next
    If Pubzz Is Nothing Then
        Set Pubzz = New clsPub ' CreateObject("ZzPub.clsPub")
        Pubzz.InitPubs2 "FA", m_Login.UfSystemDb, Account.UfDbData, m_Login.cAcc_Id, _
            m_Login.cIYear, m_Login.cUserId, m_Login.CurDate, m_Login.SysPassword
    End If
    TmpZD1 = "tempdb..[TmpZD1_" & sysInfo.ComputerName & (Timer * 100) & "]"
    TmpZD2 = "tempdb..[TmpZD2_" & sysInfo.ComputerName & (Timer * 100) & "]"
    
    Dim Rst As New ADODB.Recordset
    If Rst.State <> 0 Then Rst.Close
    Rst.Open "select top 1 ccode  from code where cclass='资产' and ccode='1501'and bclose=0 and bend=1 and isnull(cother,'')=''", DBConn, adOpenStatic, adLockReadOnly
    'LDX 2009-05-17 add beg
    If Not Rst.EOF Then
        md_ccode = Rst.Fields(0)   '默认借方科目
    End If
    'LDX 2009-05-17 add end
    If Rst.State <> 0 Then Rst.Close
    Rst.Open "select top 1  ccode  from code where cclass='负债'and bclose=0 and bend=1 and isnull(cother,'')=''", DBConn, adOpenStatic, adLockReadOnly
    'LDX 2009-05-17 add beg
    If Not Rst.EOF Then
        mc_ccode = Rst.Fields(0)   '默认贷方科目
    End If
    'LDX 2009-05-17 add end
End Sub

