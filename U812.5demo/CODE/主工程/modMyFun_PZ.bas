Attribute VB_Name = "modMyFun_PZ"
Option Explicit
  
Public Function LeftEx(ByVal Str As String, ByVal n As Long) As String
    If n < 0 Then Exit Function
    LeftEx = Left(Str, n) 'fMainForm.ComEx.LeftEx(Str, n)
End Function

Public Function RightEx(ByVal Str As String, ByVal n As Long) As String
    If n < 0 Then Exit Function
    RightEx = Right(Str, n) 'fMainForm.ComEx.RightEx(Str, n)
End Function

Public Function MidEx(ByVal Str As String, ByVal Start As Long, Optional n As Variant) As String
    If Not IsMissing(n) Then
        If n < 0 Then Exit Function
    End If
    MidEx = Mid(Str, Start, n) 'fMainForm.ComEx.MidEx(Str, Start, n)
End Function

Public Function InStrEx(ByVal Start As Long, ByVal str1 As String, ByVal Str2 As String) As Long
    InStrEx = InStr(Start, str1, Str2) 'fMainForm.ComEx.InStrEx(Start, Str1, Str2)
End Function

'���ڼ��EDIT�ؼ��������Ƿ�Ϸ�
'�ڶ������������Ƿ���Ҫ������ʾ
Public Function DateCheck(cDateExp As Variant, Optional IsShowErrorMsg As Boolean) As String
    Dim date1 As String, date2 As String, dat As String
    Dim l As Integer, M As Integer
    Dim cOperater As String
    dat = Trim(cDateExp)
    M = Len(dat)
    If dat = "" Then
        DateCheck = ""
'Result:Row=135 Col=35  Content="���ڲ���Ϊ��!" ID=65871d08-8019-4b4f-9408-448267f866bd
        If IsShowErrorMsg Then Msg "���ڲ���Ϊ��!", vbCritical
        Exit Function
    Else
        Do While l <> -1
            If InStr(dat, ".") Then
                cOperater = "."
                l = InStr(dat, cOperater)
                If l > 0 Then
                    date1 = Mid(dat, 1, l - 1)
                    date2 = Mid(dat, l + 1)
                    dat = date1 & "/" & date2
                End If
            Else
                l = -1
            End If
        Loop
    End If
    If IsDate(dat) Then
        If CDate(dat) < CDate("1753-1-1") Then
            DateCheck = ""
'Result:Row=156 Col=45  Content="���ڷǷ�!"     ID=3e576006-46bf-476a-bd91-2105c78fbc97
            If IsShowErrorMsg Then Msg "���ڷǷ�!", vbCritical
        Else
            DateCheck = Format(dat, "YYYY/MM/DD")
        End If
    Else
        DateCheck = ""
'Result:Row=163 Col=41  Content="���ڷǷ�!"     ID=bc0d1bd5-3546-43ef-b634-b2152773a250
        If IsShowErrorMsg Then Msg "���ڷǷ�!", vbCritical
    End If
End Function

'��ڣ��ַ�����ʱ�䣨�����Ѿ���DateCheck()������ʽ������
'����: ������ʱ��
Public Function StrToDate(Str As String) As Date
    StrToDate = DateSerial(val(Left(Str, 4)), val(Mid(Str, 6, 2)), val(Right(Str, 2)))
End Function




'' Ŀ��:��֤�����������ڻ�����Ƿ�Ϸ�
'Public Function bPeriod(CNN As ADODB.Connection, dDate As Date, objsys As clsSystem, arap As String) As Boolean
Public Function bPeriod(CNN As ADODB.Connection, dDate As Date, arap As String) As Boolean
    Dim curMonth As Integer
    Dim tmpMonth As Integer
    bPeriod = False
    curMonth = CurrentAccMonth(CNN, arap)
    
    '' 2001.07.30:  ������Ѿ�ȫ������
    If curMonth = 13 Then Exit Function
 
    tmpMonth = Month(dDate)  'sl�����Ƚ�ת���ô˴��������Ͼ�ע��
    If tmpMonth >= curMonth Or tmpMonth = 0 Then  'sl 2005/01/17 ��ʱ�޸� ԭ������Ϊ tmpmonth>=curmonth�������Ƚ�ת�˴��޸�Ϊtmpmonth>curmonth��
       bPeriod = True
    End If
End Function




'' ��ǰ�����
'' 2001.07.30: ����ϵͳ��ǰ�����ȡ��: GL_Mend ���� bFlag_SA=1�ļ�¼����·�(iPeriod)+1
'' 2001.09.18: Max(period) �ĳ� IsNULL(Max(iPeriod),0)
Public Function CurrentAccMonth(cn As ADODB.Connection, araps As String) As Integer
 Dim rs As New ADODB.Recordset
 Dim strsql As String
 
 
     strsql = " select IsNULL(Max(iPeriod),0) +1 As iMonth From GL_Mend where bFlag_" & araps & "=1"
     rs.ActiveConnection = cn
     rs.Open strsql, , 3, 1
     If Not (rs.EOF And rs.BOF) Then
        CurrentAccMonth = rs(0)
         
     End If
     
     If rs.State = 1 Then rs.Close
     strsql = " select isnull(cvalue,0) From accinformation where cname='dARNatStartDate' and csysid='" & araps & "'"
     rs.ActiveConnection = cn
     rs.Open strsql, , 3, 1
'     If Month(rs(0)) > CurrentAccMonth Then     'sl �޸� 2005/01/17  ���ж�ģ����������
'        CurrentAccMonth = Month(rs(0))
'     End If
     If rs.State = 1 Then rs.Close
     Set rs = Nothing
End Function


