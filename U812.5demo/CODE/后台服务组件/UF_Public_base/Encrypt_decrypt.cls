VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Encrypt_decrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



'简单的加密函数
Public Function Encryption(strEnc As String) As String

    Dim i As Integer 'for循环用

    Dim strTmp As String '临时字符存储

    Encryption = ""

    If (Len(strEnc) < 1) Then

        Encryption = ""

    Else

        For i = 1 To Len(strEnc)

            If Asc(Mid(strEnc, i, 1)) < -6000 And Mid(strEnc, i, 1) <> "?" Then

                strTmp = Chr(Asc(Mid(strEnc, i, 1)) - 5) '对gbk的汉字进行加密

            ElseIf Asc(Mid(strEnc, i, 1)) <= 126 And Asc(Mid(strEnc, i, 1)) >= 80 Then

                strTmp = Chr(Asc(Mid(strEnc, i, 1)) - 47) '对大于79的asc进行加密

            ElseIf Asc(Mid(strEnc, i, 1)) <= 79 And Asc(Mid(strEnc, i, 1)) >= 33 Then

                strTmp = Chr(Asc(Mid(strEnc, i, 1)) + 47) '对小于80的asc进行加密

            Else

                strTmp = Mid(strEnc, i, 1) '其它的字符不去转化，但是要显示的

            End If

            Encryption = Encryption & strTmp '加密后的字符

        Next

    End If

End Function

'简单的解密函数

Public Function Decipher(strDec As String) As String

    Dim i As Integer 'for循环用

    Dim strTmp As String '临时字符存储

    Decipher = ""

    If (Len(strDec) < 1) Then

        Decipher = ""

    Else

        For i = 1 To Len(strDec)

            If Asc(Mid(strDec, i, 1)) < -6000 And Mid(strDec, i, 1) <> "?" Then

                strTmp = Chr(Asc(Mid(strDec, i, 1)) + 5) '对汉定进行解密

            ElseIf Asc(Mid(strDec, i, 1)) <= 126 And Asc(Mid(strDec, i, 1)) >= 80 Then

                strTmp = Chr(Asc(Mid(strDec, i, 1)) - 47) '对大于79的asc进行解密

            ElseIf Asc(Mid(strDec, i, 1)) <= 79 And Asc(Mid(strDec, i, 1)) >= 33 Then

                strTmp = Chr(Asc(Mid(strDec, i, 1)) + 47) '对小于80的asc进行解密

            Else

                strTmp = Mid(strDec, i, 1) '其它的字符不去转化，但是要显示的

            End If

            

            Decipher = Decipher & strTmp '解密后的字符

        Next

    End If

End Function




'isDEMO=True 表示是 演示期内
'isDEMO=false 表示是 已超过演示期
Public Function isDEMO(dbc As Object, Errstr As String) As Boolean
Dim strsql As String
isDEMO = True
On Error GoTo Errstr:
Dim rds As New ADODB.Recordset
Dim Enddate As Date '失效日期

strsql = "select top 1 cValue  from AccInformation where [cSysID]='AA' and [cName]='UFCDC_Encrypt_decrypt' "

Enddate = "2000-01-01"
rds.CursorLocation = adUseClient
rds.Open strsql, dbc, 3, 4
Enddate = Decipher(CStr(rds.fields(0).Value))

If Now > CDate(Enddate) Then
    Errstr = "免费服务期限已到，请联系软件供应商！"
    isDEMO = False
End If
Exit Function
Errstr:
Errstr = Errstr & Err.Description
isDEMO = True

End Function
