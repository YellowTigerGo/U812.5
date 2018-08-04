VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGetLastYearData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function DoGet()
    If LockItem("MT6091", True, True) Then
        Call DoGet2
        Call LockItem("MT6091", False, True)
    End If
End Function

Private Function DoGet2()
    On Error GoTo hErr
    
    Dim bTrans As Boolean
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sSQL = "select iCount=isnull(count(*),0) from mt_budget"
    rs.Open sSQL, DBConn, adOpenStatic, adLockReadOnly
    If rs("iCount") > 0 Then
        MsgBox "本年度已经进行了预算业务处理,不能结转上年设置！", vbExclamation
        Exit Function
    End If
    rs.Close
    
    Dim sLastYearDBName As String
    sLastYearDBName = "UFDATA_" & m_login.cAcc_Id & "_" & Format(CLng(m_login.cIYear - 1), "0000")
    
    sSQL = "SELECT isnull(count(*),0) as icount FROM master.dbo.sysdatabases WHERE name = N'" & sLastYearDBName & "'"
    rs.Open sSQL, DBConn, adOpenStatic, adLockReadOnly
    If rs("iCount") < 1 Then
        MsgBox "上年度账不存在,不能结转上年设置！", vbExclamation
        Exit Function
    End If
    rs.Close
    
    sSQL = "select isnull(count(*),0) as iCount from " & sLastYearDBName & ".dbo.sysobjects where name=N'MT_baseset' and xtype=N'U'"
    rs.Open sSQL, DBConn, adOpenStatic, adLockReadOnly
    If rs("iCount") < 1 Then
        MsgBox "上年度账中没有基础设置数据,不需要结转上年设置！", vbExclamation
        Exit Function
    End If
    rs.Close
    
    sSQL = "select isnull(count(*),0) as iCount from " & sLastYearDBName & ".dbo.MT_baseset"
    rs.Open sSQL, DBConn, adOpenStatic, adLockReadOnly
    If rs("iCount") < 1 Then
        MsgBox "上年度账中没有基础设置数据,不需要结转上年设置！", vbExclamation
        Exit Function
    End If
    rs.Close
    
    sSQL = "select iCount=isnull(count(*),0) from mt_baseset"
    rs.Open sSQL, DBConn, adOpenStatic, adLockReadOnly
    If rs("iCount") > 1 Then
        If MsgBox("已经做过结转或已经做过基础设置,继续结转将删除原有的设置,是否继续？", vbExclamation + vbDefaultButton2 + vbYesNo) = vbNo Then
            Exit Function
        End If
    End If
    rs.Close
    
    If MsgBox("确认要结转上年度基础设置？", vbInformation + vbOKCancel + vbDefaultButton2) = vbCancel Then
        Exit Function
    End If
    
    sSQL = "delete mt_baseset " & vbCrLf & _
           "Delete mt_basesets " & vbCrLf & _
           "insert into mt_baseset ([id],[cVouchType],[VT_ID],[dDate],[cCode],[chandler],[cmaker],[cDepCode],[cItemClass],[cItemCode],[cExpCode],[cExpCCode],[rate],[cDefine1],[cDefine2],[cDefine3],[cDefine4],[cDefine5],[cDefine6],[cDefine7],[cDefine8],[cDefine9],[cDefine10],[cDefine11],[cDefine12],[cDefine13],[cDefine14],[cDefine15],[cDefine16])  " & vbCrLf & _
           "select [id],[cVouchType],[VT_ID],[dDate],[cCode],[chandler],[cmaker],[cDepCode],[cItemClass],[cItemCode],[cExpCode],[cExpCCode],[rate],[cDefine1],[cDefine2],[cDefine3],[cDefine4],[cDefine5],[cDefine6],[cDefine7],[cDefine8],[cDefine9],[cDefine10],[cDefine11],[cDefine12],[cDefine13],[cDefine14],[cDefine15],[cDefine16] from " & sLastYearDBName & "..mt_baseset " & vbCrLf & _
           "insert into mt_basesets ([AutoID],[id],[ccode],[adds],[lenssen],[cExpCode],[rate],[cDefine22],[cDefine23],[cDefine24],[cDefine25],[cDefine26],[cDefine27],[cDefine28],[cDefine29],[cDefine30],[cDefine31],[cDefine32],[cDefine33],[cDefine34],[cDefine35],[cDefine36],[cDefine37]) " & vbCrLf & _
           "select [AutoID],[id],[ccode],[adds],[lenssen],[cExpCode],[rate],[cDefine22],[cDefine23],[cDefine24],[cDefine25],[cDefine26],[cDefine27],[cDefine28],[cDefine29],[cDefine30],[cDefine31],[cDefine32],[cDefine33],[cDefine34],[cDefine35],[cDefine36],[cDefine37] from " & sLastYearDBName & "..mt_basesets "
    
    With DBConn
        .BeginTrans
        bTrans = True
        .Execute sSQL
        .CommitTrans
        bTrans = False
    End With
    
    MsgBox "结转成功！", vbInformation
    Exit Function
hErr:
    On Error Resume Next
    If bTrans Then DBConn.RollbackTrans
    MsgBox "结转上年基础设置时出错:" & Err.Description
End Function
