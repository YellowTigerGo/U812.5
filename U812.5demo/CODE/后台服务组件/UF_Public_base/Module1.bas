Attribute VB_Name = "Module1"

Public m_logins As Object
Public Conn As ADODB.Connection

Public vt_ids As String




'取系统配置信息 chenliangc
Public Function getAccinformation(strSysID As String, strName As String, Conn As Object) As String
    Dim rst As New ADODB.Recordset

    rst.CursorLocation = adUseClient
    rst.Open "Select cValue from accinformation where cSysID=N'" & strSysID & "' and cName=N'" & strName & "'", Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rst.EOF Then
        getAccinformation = ""
    Else
        If IsNull(rst(0)) Then
            getAccinformation = ""
        Else
            getAccinformation = rst(0)
        End If
    End If
    rst.Close
    Set rst = Nothing
End Function

'更新插入系统信息
Public Sub UpdateAccinfo(strSysID As String, strName As String, strValue As String, Conn)
    Dim affeceted As Long
    Conn.Execute "Update accinformation set cValue=N'" & strValue & "' where cSysId=N'" & strSysID & "' and cname=N'" & strName & "'", affeceted
    If affeceted = 0 Then
        Conn.Execute "insert into accinformation(cValue,cSysId,cname) values(N'" & strValue & "' ,N'" & strSysID & "' ,N'" & strName & "')"
    End If
End Sub

Public Function str2Dbl(ByVal val As String) As Double
    On Error GoTo hErr
    If Len(val) > 0 Then
        str2Dbl = CDbl(val)
    End If
    Exit Function
hErr:
    str2Dbl = 0
End Function

