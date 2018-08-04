Attribute VB_Name = "ahzzd_Mdl"
'Option Explicit
'Public rsVouchlist As ADODB.Recordset
'Public ifalg As Boolean
'
' by ahzzd 2006/05/09 修改DomHead 里的值

'
'
''将str字符串中的substr 子串过滤掉
'Public Function leach_substr(str As String, substr As String) As String
'Dim tempstr As String
'Dim i As Long
'    Do
'        i = InStr(1, str, substr, vbTextCompare)
'        If i > 0 Then
'            tempstr = Left(str, i - 1) & Right(str, Len(str) - i)
'            str = tempstr
'        End If
'    Loop Until i < 1
'    leach_substr = str
'End Function
'
''取出str_str_key1 和 str_key2 关键子之间的字符
'Public Function get_midstr(ByVal str As String, str_key1 As String, str_key2 As String) As String
'Dim tempstr As String
'Dim tempstr1 As String
'Dim i As Long
'tempstr1 = str
'    If str_key1 <> "" Then
'        i = InStr(1, tempstr1, str_key1, vbTextCompare)
'        If i > 0 Then
'            i = i + Len(str_key1) - 1
'            tempstr = Right(tempstr1, Len(str) - i)
'            tempstr1 = tempstr
'        End If
'    End If
'
'    If str_key2 <> "" Then
'        i = InStr(1, tempstr1, str_key2, vbTextCompare)
'        If i > 0 Then
'            tempstr = Left(tempstr1, i - 1)
'            tempstr1 = tempstr
'        End If
'    End If
'    get_midstr = tempstr1
'End Function
'
'
''取出 str_str_key 关键子左边的字符
'Public Function get_leftstr(str As String, str_key As String) As String
'Dim tempstr As String
'Dim i As Long
'    If str_key <> "" Then
'        i = InStr(1, str, str_key, vbTextCompare)
'        If i > 0 Then
'            tempstr = Left(str, i - 1)
'        End If
'        get_leftstr = tempstr
'    Else
'        get_leftstr = ""
'    End If
'
'End Function
'
''取出 str_str_key 关键子右边的字符
'Public Function get_Rightstr(str As String, str_key As String) As String
'Dim tempstr As String
'Dim i As Long
'    If str_key <> "" Then
'        i = InStr(1, str, str_key, vbTextCompare)
'        If i > 0 Then
'            i = i + Len(str_key) - 1
'            tempstr = Right(str, Len(str) - i)
'        End If
'        get_Rightstr = tempstr
'    Else
'        get_Rightstr = ""
'    End If
'End Function
'
'
'
