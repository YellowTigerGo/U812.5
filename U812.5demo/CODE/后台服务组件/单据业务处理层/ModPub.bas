Attribute VB_Name = "ModPub"
Option Explicit
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Public LockSOCode As New Collection, LockDLID As New Collection
Public rds_M As New ADODB.Recordset
Dim m_Login As U8Login.clsLogin
Public VoucherRefAgain As Boolean '�������� �Ƿ�༭�������� ���ر�ͷ��Ϣ
Public VoucherRefDomH As DOMDocument '�������� �Ƿ�༭�������� ���ر�ͷ��Ϣ
Public VoucherRefDomB As DOMDocument '�������� �Ƿ�༭�������� ���ر�����Ϣ
Public refCbustype As String 'ҵ������

Private domEnumValue As New DOMDocument
'strReferString
'strFieldName ��ҪУ����ֶ�
'DomValue ��ͷ������Dom����
'domFieldConfig voucherfileditemconfig�е�ֵ
'strCardSection �Ƿ��Ǳ�
Public cls_Public As Object

Private domVouchers As New DOMDocument
Private domFieldConfig As New DOMDocument
Private domCellCheck As New DOMDocument
Private domFieldCaption As New DOMDocument
Private domBTTableField As New DOMDocument '��ͷ�ֶ�
Private domBWTableField As New DOMDocument '�����ֶ�
'��str�ַ����е�substr �Ӵ����˵�
Public Function leach_substr(str As String, substr As String) As String
Dim tempStr As String
Dim i As Long
    Do
        i = InStr(1, str, substr, vbTextCompare)
        If i > 0 Then
            tempStr = Left(str, i - 1) & Right(str, Len(str) - i)
            str = tempStr
        End If
    Loop Until i < 1
    leach_substr = str
End Function
'' �ַ������뺯��
'' ����: strInput: Դ�ַ��� ; strKey: �ؼ��ִ�; Separate:����ַ����з����־,���� ;,/�� ;
Public Function ExtractStr(strInput As String, strKey As String, Optional Separate As String = "") As String
 Dim strLeft As String, iPos As Long
     iPos = 1
     strLeft = strInput
     If strInput = "" Or InStr(1, strLeft, Separate) = 0 Then
        ExtractStr = ""
        Exit Function
     End If
     ''ȫ��ȡСд���ټ�飬�������Сд�����϶���ֹ
     iPos = InStr(1, LCase(strLeft), LCase(Separate))
     Do While iPos > 0  ''�ҷ��봮�Ŀ�ʼλ��
        If InStr(1, Mid(LCase(strLeft), 1, iPos), LCase(strKey)) > 0 Then
           strLeft = Left(strLeft, iPos - 1)
           If strKey <> "" Then   ''����йؼ��ִ�
              ''ȡ�������,����  a+b=3 ,a+bΪ�ؼ���,���3ȡ����
              ExtractStr = Mid(strLeft, Len(strKey) + 1, Len(strLeft) - Len(strKey))
              Exit Do
           Else
              ExtractStr = strLeft
              Exit Do
           End If
        End If
        strLeft = Mid(strLeft, iPos + 1, Len(strLeft) - iPos)
        iPos = InStr(1, LCase(strLeft), LCase(Separate))
        If iPos = 0 Then  ''����Ѿ������һ�������־
           If InStr(1, LCase(strLeft), LCase(strKey)) > 0 Then iPos = Len(strLeft) + 1
        End If
     Loop
End Function
Sub Main()

End Sub
''�������ڷ��ػ����
Public Function GetAccMonth(ddate As Date, objsys As clsSystem) As Integer
    Dim i As Integer
     ''С��1��
     If ddate < objsys.getBeginDate(1) Then
        GetAccMonth = -1
        Exit Function
     End If
     ''����12��
     If ddate > objsys.getEndDate(12) Then
        GetAccMonth = 0
        Exit Function
     End If
     For i = 1 To 12
        If ddate >= objsys.getBeginDate(i) And ddate <= objsys.getEndDate(i) Then
           GetAccMonth = i
           Exit For
        End If
     Next
End Function

'' ��ǰ�����
Public Function CurrentAccMonth(CN As ADODB.Connection) As Integer
 Dim Rs As New ADODB.Recordset
 Dim strSQL As String
 'by ahzzd 2006/05/29
      If Rs.State = 1 Then Rs.Close
     strSQL = " select IsNULL(Max(iPeriod),0)+1  As iMonth From GL_Mend where bflag_FA=1"
     Rs.ActiveConnection = CN
     Rs.Open strSQL, , 3, 1
     If Not (Rs.EOF And Rs.BOF) Then
        CurrentAccMonth = Rs(0)
     End If
     If Rs.State = 1 Then Rs.Close
     Set Rs = Nothing
End Function
 

Public Function NullToStr(vValue As ADODB.Field) As String
  NullToStr = IIf(IsNull(vValue.value), "NULL", "'" & vValue.value & "'")
End Function

Public Function NullToNull(vValue As ADODB.Field) As Variant
  NullToNull = IIf(IsNull(vValue.value), "NULL", vValue.value)
End Function

Public Function ToDBL(ByVal sValue As Variant) As Double
    If IsNumeric(sValue) Then
       ToDBL = CDbl(sValue)
    Else
       ToDBL = 0
    End If
End Function

Public Function Zero(x As Double) As Double
 Dim y As Double
    y = ToDBL(x)
    Zero = IIf(Abs(y) < 0.000001, 0, y)
End Function

' �������ַ���Ϊ��Ϊ�յ��ж�
Public Function NoBlank(strString As Variant) As Boolean
    On Error Resume Next
    If IsNull(strString) Then
        NoBlank = False
    Else
        If Len(strString) = 0 Then
            NoBlank = False
        Else
            NoBlank = True
        End If
    End If
End Function

Public Function IsBlank(strString As Variant) As Boolean
 On Error Resume Next
    If IsNull(strString) Then
       IsBlank = True
    Else
       If Len(strString) = 0 Then
          IsBlank = True
       Else
          IsBlank = False
       End If
    End If
End Function

''�ж�ָ������Ƿ�Ϊ˫����
''�̶���������ͬ���������
Public Function IsTwoUnit(CN, strInvCode As String, Optional ByVal bNewCollection As Boolean = False) As Boolean
    Dim rst As New ADODB.Recordset
    Dim strSeek As String
    IsTwoUnit = False
    strSeek = "Select cInvCode,iGroupType from Inventory where cInvCode='" & strInvCode & "'"
    If CN.State = 1 Then
        rst.ActiveConnection = CN
        rst.Open strSeek, , adOpenForwardOnly, adLockReadOnly
        If Not (rst.EOF And rst.BOF) Then
            IsTwoUnit = IIf(rst.Fields("iGroupType") = 2, True, False)
            If rst.State = 1 Then rst.Close
            Set rst = Nothing
        End If
        If bNewCollection Then
            CN.Close
            Set CN = Nothing
        End If
    End If
End Function

''������
Public Function ClearCol(Col As Collection)
  Dim iCount As Long, i As Long
      iCount = Col.Count
      For i = 1 To iCount
          Col.Remove 1
      Next
End Function

'' �����ϵͳ��ĳ��������Ƿ����
'' ����;strSubSys:��ϵͳ����; iMonth:�·�
Public Function CheckSubSysAcc(CN As ADODB.Connection, strSubSys As String, ByVal iMonth As Integer, Optional ByVal bNewCollection As Boolean = False) As Boolean
  Dim rec As New ADODB.Recordset
  Dim rec2 As New ADODB.Recordset
  Dim sKey As String, sKey2 As String, strSeek As String, strSeek2 As String
  Dim CN2 As New ADODB.Connection
  Dim strDBName As String
  CN2.Open CN.ConnectionString
      If CN2.State = 1 Then
         sKey = Trim("d" & strSubSys & "Startdate")       '' �������ڼ����ؼ���
         sKey2 = Trim("bFlag_" & strSubSys)             '' �Ƿ���˼����ؼ���
         strSeek = "Select cValue From Accinformation where cSysid='" & strSubSys & "' and cName='" & sKey & "'"
         strSeek2 = "Select " & sKey2 & " From Gl_mend where iPeriod=" & iMonth
         If rec.State = 1 Then rec.Close
         If rec2.State = 1 Then rec2.Close
         rec.CursorLocation = adUseClient
         rec2.CursorLocation = adUseClient
         rec.ActiveConnection = CN2
         rec2.ActiveConnection = CN2
         rec.Open strSeek, , adOpenForwardOnly, adLockReadOnly
         rec2.Open strSeek2, , adOpenForwardOnly, adLockReadOnly
         If Not (rec.BOF And rec2.EOF) Then
            If Not IsNull(rec.Fields(0)) Then
               If Not (rec2.EOF And rec2.BOF) Then
                  CheckSubSysAcc = CBool(rec2.Fields(0))
               Else
                  CheckSubSysAcc = False
               End If
            Else
               CheckSubSysAcc = False
            End If
         Else
            CheckSubSysAcc = False
         End If
         If rec.State = 1 Then rec.Close
         If rec2.State = 1 Then rec2.Close
         Set rec = Nothing
         Set rec2 = Nothing
         CN2.Close
         Set CN2 = Nothing
         Exit Function
      Else
         CheckSubSysAcc = False
         Exit Function
      End If
      Exit Function
End Function


Public Function CreateTempTable(objsys As clsSystem, Optional sPreFix As String) As String
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'�����ʱ�ļ�����
'sPreFix����ʱ��ǰ׺
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Dim i As Long
    Dim sTempName As String
    Dim sRnd As String
ReCreate:
    Randomize
    CreateTempTable = ""
    sTempName = NewTrim(objsys.sComputerName) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now)
    sRnd = Int((10000 * Rnd) + 1000)
    sTempName = NewTrim(sTempName & sRnd)
    sPreFix = sPreFix & sTempName
    CreateTempTable = sPreFix
End Function

Public Function NewTrim(s As String) As String
'---------------------------------------------------------------------------------------
'�������ƣ�NewTrim
'�������ܣ�����ַ��������еĿո�
'---------------------------------------------------------------------------------------
'����˵����
'  s��Ҫ������ַ�����
'---------------------------------------------------------------------------------------
    Dim i As Long
    NewTrim = ""
    If InStr(1, s, " ") = 0 Then
        NewTrim = s
        Exit Function
    Else
        For i = 1 To Len(s)
            If Mid(s, i, 1) <> " " Then
                NewTrim = NewTrim & Mid(s, i, 1)
            End If
        Next i
    End If
End Function

Public Function getDefineName(CN As ADODB.Connection, strFldName As String, Optional cItemName As String) As String
    Dim strID As String
    Dim rst As New ADODB.Recordset
    getDefineName = ""
    rst.CursorLocation = adUseClient
    strID = GetDefineID(strFldName)
    If rst.State = adStateOpen Then
        rst.Close
    End If
    On Error Resume Next
    rst.Open "Select isnull(cItemName,cItem),ISNULL(cItemName,'') AS cItemName from userdef where cID='" & strID & "'", CN, adOpenKeyset, adLockReadOnly
    If Not rst.EOF Then
        getDefineName = rst(0)
        cItemName = rst(1)
    End If
    rst.Close
    Set rst = Nothing
    
End Function
Public Function GetDefineID(strFldName As String) As String
    Select Case LCase(strFldName)
        Case "cdefine1"
            GetDefineID = "01"  '����ͷ�Զ�����1
        Case "cdefine2"
            GetDefineID = "02"  '����ͷ�Զ�����2
        Case "cdefine3"
            GetDefineID = "03"  '����ͷ�Զ�����3
        Case "cdefine4"
            GetDefineID = "04"  '����ͷ�Զ�����4
        Case "cdefine5"
            GetDefineID = "05"  '����ͷ�Զ�����5
        Case "cdefine6"
            GetDefineID = "06"  '����ͷ�Զ�����6
        Case "cdefine7"
            GetDefineID = "07"  '����ͷ�Զ�����7
        Case "cdefine8"
            GetDefineID = "08"  '����ͷ�Զ�����8
        Case "cdefine9"
            GetDefineID = "09"  '����ͷ�Զ�����9
        Case "cdefine10"
            GetDefineID = "10"  '����ͷ�Զ�����10
        Case "cdefine11"
            GetDefineID = "36"  '����ͷ�Զ�����11
        Case "cdefine12"
            GetDefineID = "37"  '����ͷ�Զ�����12
        Case "cdefine13"
            GetDefineID = "38"  '����ͷ�Զ�����13
        Case "cdefine14"
            GetDefineID = "39"  '����ͷ�Զ�����14
        Case "cdefine15"
            GetDefineID = "40"  '����ͷ�Զ�����15
        Case "cdefine16"
            GetDefineID = "41"  '����ͷ�Զ�����16
            
        Case "cdefine22"
            GetDefineID = "22"  '�������Զ�����1
        Case "cdefine23"
            GetDefineID = "23"  '�������Զ�����2
        Case "cdefine24"
            GetDefineID = "24"  '�������Զ�����3
        Case "cdefine25"
            GetDefineID = "25"  '�������Զ�����4
        Case "cdefine26"
            GetDefineID = "26"  '�������Զ�����5
        Case "cdefine27"
            GetDefineID = "27"  '�������Զ�����6
        Case "cdefine28"
            GetDefineID = "42"  '�������Զ�����7
        Case "cdefine29"
            GetDefineID = "43"  '�������Զ�����8
        Case "cdefine30"
            GetDefineID = "44"  '�������Զ�����9
        Case "cdefine31"
            GetDefineID = "45"  '�������Զ�����10
        Case "cdefine32"
            GetDefineID = "46"  '�������Զ�����11
        Case "cdefine33"
            GetDefineID = "47"  '�������Զ�����12
        Case "cdefine34"
            GetDefineID = "48"  '�������Զ�����13
        Case "cdefine35"
            GetDefineID = "49"  '�������Զ�����14
        Case "cdefine36"
            GetDefineID = "50"  '�������Զ�����15
        Case "cdefine37"
            GetDefineID = "51"  '�������Զ�����16
        Case "cinvdefine1"
            GetDefineID = "17"
        Case "cinvdefine2"
            GetDefineID = "18"
        Case "cinvdefine3"
            GetDefineID = "19"
        Case "cfree1"
            GetDefineID = "20"
        Case "cfree2"
            GetDefineID = "21"
        Case "cfree3"
            GetDefineID = "28"
        Case "cfree4"
            GetDefineID = "29"
        Case "cfree5"
            GetDefineID = "30"
        Case "cfree6"
            GetDefineID = "31"
        Case "cfree7"
            GetDefineID = "32"
        Case "cfree8"
            GetDefineID = "33"
        Case "cfree9"
            GetDefineID = "34"
        Case "cfree10"
            GetDefineID = "35"
        Case "cinvdefine4"       ''�Զ�����4
            GetDefineID = "52"
        Case "cinvdefine5"       ''�Զ�����5
            GetDefineID = "53"
        Case "cinvdefine6"       ''�Զ�����6
            GetDefineID = "54"
        Case "cinvdefine7"       ''�Զ�����7
            GetDefineID = "55"
        Case "cinvdefine8"       ''�Զ�����8
            GetDefineID = "56"
        Case "cinvdefine9"       ''�Զ�����9
            GetDefineID = "57"
        Case "cinvdefine10"       ''�Զ�����10
            GetDefineID = "58"
        Case "cinvdefine11"       ''�Զ�����11
            GetDefineID = "59"
        Case "cinvdefine12"       ''�Զ�����12
            GetDefineID = "60"
        Case "cinvdefine13"       ''�Զ�����13
            GetDefineID = "61"
        Case "cinvdefine14"       ''�Զ�����14
            GetDefineID = "62"
        Case "cinvdefine15"       ''�Զ�����15
            GetDefineID = "63"
        Case "cinvdefine16"       ''�Զ�����16
            GetDefineID = "64"
        Case "ccusdefine1"       ''�Զ�����7
            GetDefineID = "11"
        Case "ccusdefine2"       ''�Զ�����8
            GetDefineID = "12"
        Case "ccusdefine3"       ''�Զ�����9
            GetDefineID = "13"
        Case "ccusdefine4"       ''�Զ�����10
            GetDefineID = "65"
        Case "ccusdefine5"       ''�Զ�����11
            GetDefineID = "66"
        Case "ccusdefine6"       ''�Զ�����12
            GetDefineID = "67"
        Case "ccusdefine7"       ''�Զ�����13
            GetDefineID = "68"
        Case "ccusdefine8"       ''�Զ�����14
            GetDefineID = "69"
        Case "ccusdefine9"       ''�Զ�����15
            GetDefineID = "70"
        Case "ccusdefine10"       ''�Զ�����16
            GetDefineID = "71"
        Case "ccusdefine11"       ''�Զ�����11
            GetDefineID = "72"
        Case "ccusdefine12"       ''�Զ�����12
            GetDefineID = "73"
        Case "ccusdefine13"       ''�Զ�����13
            GetDefineID = "74"
        Case "ccusdefine14"       ''�Զ�����14
            GetDefineID = "75"
        Case "ccusdefine15"       ''�Զ�����15
            GetDefineID = "76"
        Case "ccusdefine16"       ''�Զ�����16
            GetDefineID = "77"
            
    End Select
End Function

'create by  zhaojp
'����˵�� ��֤�����в��������ֶκϷ�������֤ͨ�÷���
'�Ϸ������ݣ� �ֶ��� checksql �е����� ,����֤����ͨ�������ֶ��е�checksql��ֵ
'-----------------------------------------------------------------------------------------
'����ֶεĲ���������ͨ�������ֶε�ֵ���ı�ʱ
'����:�������Ͳ�ͬʱ����������Ĳ������Ϳ�����1 (����) Ҳ������5(�޲���)
'�������ֶα�����뵽 ColNoCheckFieldName �У���ȡ��ALLreferFiledCheck�����Ը��ֶ�ֵ�ĺϷ�����֤
'��ô����ֶ�ֵ�ĺϷ�����֤��Ҫ����ʵ��
'-----------------------------------------------------------------------------------------
'����˵��
'ColNoCheckFieldName ����Ҫͨ����֤���ֶ�������
Public Function ALLreferFiledCheck(cardnum As String, conn As ADODB.Connection, domHead As DOMDocument, domBody As DOMDocument, strError As String, Optional ColNoCheckFieldName As Collection) As Boolean
    '���е���checksql �ֶζ�Ҫ���
    Dim nodS As IXMLDOMNodeList
    Dim nod As IXMLDOMNode
    Dim eleMent As IXMLDOMNode
    Dim Row As Long
    Dim strCardSection As String
    Dim strFieldName  As String
    Dim fillsql As String
    Dim Attr As IXMLDOMAttribute
    SubInit cardnum, conn
    ALLreferFiledCheck = True
    Set nodS = domFieldConfig.selectNodes("//z:row")
    strError = ""
    If nodS Is Nothing Then Exit Function
    For Each nod In nodS
        strFieldName = nod.Attributes.getNamedItem("fieldname").nodeValue
        strCardSection = nod.Attributes.getNamedItem("cardsection").nodeValue
        'fillsql = nod.Attributes.getNamedItem("fillselectsql").nodeValue
        If BlnFieldCheck(strFieldName, strCardSection, ColNoCheckFieldName) Then
            If strCardSection = "T" Then
                strError = CellCheck(conn, cardnum, strFieldName, domHead, domBody, strCardSection)
            Else
                Row = 0
                For Each eleMent In domBody.selectNodes("//z:row")
                    'For Each Attr In eleMent.Attributes
                        strError = CellCheck(conn, cardnum, strFieldName, domHead, domBody, strCardSection, Row)
                        If strError <> "" Then
                            ALLreferFiledCheck = False
                            Exit Function
                        End If
                    'Next
                    Row = Row + 1
                    If strError <> "" Then
                        ALLreferFiledCheck = False
                        Exit Function
                    End If
                Next
            End If
            If strError <> "" Then
                ALLreferFiledCheck = False
                Exit Function
            End If
        End If
    Next
End Function
Private Sub SubInit(strCardNum As String, DBConn As ADODB.Connection)
    Dim rst As New ADODB.Recordset
    Dim strSQL As String
    rst.CursorLocation = adUseClient
    rst.Open "select * from vouchers_base where cardnumber='" + strCardNum + "'", DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rst.Save domVouchers, adPersistXML
    rst.Close
    rst.Open "select * from sa_voucherfieldconfig where cardnumber=N'" + strCardNum + "' and errresid<>'' order by cardsection", DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rst.Save domFieldConfig, adPersistXML
    rst.Close
    rst.Open "select distinct fieldname,cardsection,cardnum,carditemname from voucheritems_lang where cardnum='" & strCardNum & "'", DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rst.Save domFieldCaption, adPersistXML
    rst.Close
    
    Set rst = Nothing
    Call GetHeadBodyField(strCardNum, DBConn)
End Sub
'�z���ֶ��Ƿ���Ҫ���
Private Function BlnFieldCheck(strFieldName As String, strCardSection As String, Optional ColNoCheckFieldName As Collection) As Boolean
On Error Resume Next
Dim blncheck As Boolean
    blncheck = Not IsExitFieldName(strFieldName, ColNoCheckFieldName)
    If blncheck Then
        If strCardSection = "T" Then
            If Not domBTTableField.selectSingleNode("//z:row [@name ='" & strFieldName & "']") Is Nothing Then
                BlnFieldCheck = True
            Else
                BlnFieldCheck = False
            End If
        Else
            If Not domBWTableField.selectSingleNode("//z:row [@name ='" & strFieldName & "']") Is Nothing Then
                BlnFieldCheck = True
            Else
                BlnFieldCheck = False
            End If
    End If
    Else
        BlnFieldCheck = blncheck
    End If


End Function
Private Function CellCheck(conn As ADODB.Connection, cardnum As String, strFieldName As String, domHead As DOMDocument, _
         domBody As DOMDocument, strCardSection As String, Optional lngRow As Long, Optional strValues As String) As String
    Dim varRefer  As Variant
    Dim lngRowOld As Long
    Dim lngRowMax As Long
    Dim i As Long
    Dim nod As IXMLDOMNode
    Dim strSQL As String
    Dim strValue As String
    Dim blnOk As Boolean
    Dim strError As String
    Dim strReferName As String
    Dim lst As IXMLDOMNodeList
    Dim clsComp As New UsSaCompStr.clsCompStr

    Dim strReferType As String
    On Error Resume Next
    Set nod = domFieldConfig.selectSingleNode("//z:row[@cardsection='" + UCase(strCardSection) + "' and @fieldname='" + LCase(strFieldName) + "']")
    If nod Is Nothing Then Exit Function
    If nod.Attributes.getNamedItem("refertype") Is Nothing Then Exit Function
    strReferType = nod.Attributes.getNamedItem("refertype").Text
    Select Case strReferType
        Case "4"
            strValue = GetItemValue(domHead, strFieldName, strCardSection)
            Set nod = domEnumValue.documentElement.selectSingleNode("row[@name='" + LCase(strFieldName) + "']")
            If nod Is Nothing Then
                strError = "false"
            Else
                Dim ele As IXMLDOMElement
                Set ele = nod.selectSingleNode("value[@code='" + ReplaceSpecialCode(strValue) + "']")
                If ele Is Nothing Then
                    strError = "false"
                End If
            End If
        Case "1", "itemclass"
            Set nod = domFieldConfig.selectSingleNode("//z:row[@cardsection='" + UCase(strCardSection) + "' and @fieldname='" + LCase(strFieldName) + "']")
           
            If Not nod Is Nothing Then
                strReferName = nod.Attributes.getNamedItem("refername").Text
                strValue = GetItemValue(domHead, strFieldName, strCardSection)
                If strValue = "" Then
                    
                Else
               
                    strSQL = nod.Attributes.getNamedItem("cellchecksql").Text
                    blnOk = FillVoucherItems(conn, domHead, domBody, strFieldName, strReferName, strSQL, strCardSection, lngRow)
                    If Not blnOk Then
                        If Not nod.Attributes.getNamedItem("errresid") Is Nothing Then
                            strError = nod.Attributes.getNamedItem("errresid").Text
                            strError = GetErrorString(strError, strFieldName, strCardSection)
                            CellCheck = strError
                        End If
                    End If
                End If
            End If
        Case "free"
            strValue = GetItemValue(domHead, strFieldName, strCardSection)
            If strValue <> "" Then
           
            'Call FieldCheckFree(conn, GetItemValue(domHead, "cinvcode", strCardSection), strFieldName, GetItemValue(domHead, strFieldName, strCardSection), strError)
            'Call FieldCheckFree(conn, "", "cfree1", "dfdf", strError)
            If strError <> "" Then strError = "��" & lngRow + 1 & "���е�" & GetFieldName(strFieldName) & strError
            End If
        Case Else
    End Select
    CellCheck = strError
End Function



'�õ���ͷ�ͱ������Ҫ��֤�ֶ�
Private Sub GetHeadBodyField(strCardNum As String, DBConn As ADODB.Connection)
    Dim dom As DOMDocument
    Dim rst As New ADODB.Recordset
    Dim strSQL As String
    Dim BtTblname As String
    Dim BwTblname As String
    strSQL = "select BTtblname , BWTblName from vouchers where cardnumber='" & strCardNum & "'"
    rst.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rst Is Nothing Then Exit Sub
    
    BtTblname = rst.Fields("Bttblname").value
    BwTblname = IIf(IsNull(rst.Fields("BWTblName")), "", rst.Fields("BWTblName"))
    rst.Close
    strSQL = "select name  from syscolumns where id=object_id('" & BtTblname & "')"
    rst.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rst.Save domBTTableField, adPersistXML
    rst.Close
    If BwTblname <> "" Then
        strSQL = "select name  from syscolumns where id=object_id('" & BwTblname & "')"
        rst.Open strSQL, DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
        rst.Save domBWTableField, adPersistXML
        rst.Close
    End If
    Set rst = Nothing
End Sub

Public Function IsExitFieldName(StrField As String, ColNoCheckFieldName As Collection) As Boolean
    On Error GoTo err
    Dim i As Integer
    If ColNoCheckFieldName Is Nothing Then IsExitFieldName = False: Exit Function
    For i = 1 To ColNoCheckFieldName.Count
        If ColNoCheckFieldName.Item(i) = StrField Then
            IsExitFieldName = True
            Exit Function
        End If
    Next i
    IsExitFieldName = False
    Exit Function
    'If ColNoCheckFieldName.Item(StrField) <> "" Then IsExitFieldName = True: Exit Function
err:
    IsExitFieldName = False
End Function

Private Function GetItemValue(domValue As DOMDocument, strFieldName As String, strCardSection As String, Optional lngRow As Long) As String
    If strCardSection = "T" Then
        GetItemValue = GetHeadItemValue(domValue, strFieldName)
    Else
        GetItemValue = GetBodyItemValue(domValue, strFieldName, lngRow)
    End If
End Function
Private Function ReplaceSpecialCode(strSourceCode As String) As String
    Dim strTmp As String
    strTmp = Replace(strSourceCode, "'", "&apos;")
    strTmp = Replace(strTmp, """", "&quot;")
    strTmp = Replace(strTmp, ">", "&gt;")
    strTmp = Replace(strTmp, "<", "&lt;")
    strTmp = Replace(strTmp, "&", "&amp;")
    strTmp = Replace(strTmp, "\", "\\")
    ReplaceSpecialCode = strTmp
End Function
Private Function FillVoucherItems(DBConn As ADODB.Connection, domHead As DOMDocument, domBody As DOMDocument, strReferFieldName As String, strReferName As String, strSQL As String, strCardSection As String, Optional lngRow As Long) As Boolean
    
    Dim rst As New ADODB.Recordset
    rst.CursorLocation = adUseClient
    Dim lst As IXMLDOMNodeList
    Dim nod As IXMLDOMNode
    Dim strSelect As String
    Dim strSourceFieldName As String
    Dim strDesFldName As String
    Dim strAuth As String
    Dim varFilter As Variant
    Dim strFilter As String
    Dim strValue1 As String
    Dim strValue2 As String
    Dim i As Long
    Dim lstChange As IXMLDOMNodeList
    Dim eleChange As IXMLDOMElement
    Dim strMsg As String
    Dim blncheck As Boolean
    strSelect = strSQL
    On Error Resume Next
    strSelect = ReplaceVoucherItems(strSelect, domHead, domBody, lngRow)
    rst.Open ConvertSQLString(strSelect), DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    FillVoucherItems = True
    If Not rst.EOF Then
    Else
        FillVoucherItems = False
    End If
    rst.Close
    Set rst = Nothing
End Function
Private Function GetErrorString(strSource As String, strFieldName As String, strCardSection As String) As String
    Dim varError As Variant
    If InStr(strSource, ",") = 0 Then
        GetErrorString = GetString(strSource)
    Else
        varError = Split(strSource, ",")
        GetErrorString = GetStringPara(varError(0), GetFieldName(strFieldName))
    End If
End Function

'�ж����������
'����˵����
'cinvcodeValue ��Ӧ cinvcode ��ֵ
' StrFieldName_Lang �ֶζ�Ӧ�����Ա�
Private Function FieldCheckFree(conn As ADODB.Connection, cinvcodeValue As String, strFieldName As String, StrFieldvalue As String, strMsg) As Boolean
    Dim LngReturn As Long
    Dim clsDef As U8DefPro.clsDefPro
    Dim tmp_Login As U8Login.clsLogin
    Set tmp_Login = New U8Login.clsLogin
    Set clsDef = New U8DefPro.clsDefPro
    FieldCheckFree = True
    
    If Not clsDef.Init(False, conn.ConnectionString, "demo") Then
        FieldCheckFree = False
        If strMsg Then
            strMsg = GetString("U8.SA.xsglsql.01.frmbillvouch.00147") 'zh-CN����ʼ���Զ��������ʧ�ܣ�
        End If
        Set clsDef = Nothing
        Exit Function
    End If
    
    LngReturn = clsDef.ValidateFreeArEx(cinvcodeValue, "<Data />", strFieldName, StrFieldvalue, False)
    Select Case LngReturn
        Case 0
             strMsg = ""
             FieldCheckFree = True
        Case Else
             strMsg = GetString("U8.SA.xsglsql.01.frmbillvouch.00153") 'zh-CN��¼�벻�Ϸ�������
             FieldCheckFree = False
    End Select
    
    
End Function

Private Function GetFieldName(strFieldName As String) As String
    Dim sKey As String
     sKey = LCase(strFieldName)
    If Not domFieldCaption.selectSingleNode("//z:row[@fieldname='" & strFieldName & "']").Attributes.getNamedItem("carditemname") Is Nothing Then
        GetFieldName = domFieldCaption.selectSingleNode("//z:row[@fieldname='" & strFieldName & "']").Attributes.getNamedItem("carditemname").Text
    Else
        GetFieldName = strFieldName
    End If
End Function

Public Function ReplaceSysPara(strSource As String) As String
    Dim lngPos1 As Integer
    Dim lngPos2 As Integer
    Dim strFieldName As String
    Dim varField As Variant
    
    lngPos1 = InStr(1, strSource, "@")
    Do While lngPos1 > 0
        lngPos2 = InStr(lngPos1, strSource, "=")
        If lngPos2 = 0 Then
            strFieldName = Mid(strSource, lngPos1)
            If Right(strFieldName, 1) = ")" Then
                strFieldName = Left(strFieldName, Len(strFieldName) - 1)
            End If
            If Right(strFieldName, 1) = """" Then
                strFieldName = Left(strFieldName, Len(strFieldName) - 1)
            End If
        Else
            strFieldName = Mid(strSource, lngPos1, lngPos2 - lngPos1)
        End If
        If Right(strFieldName, 1) = """" Then strFieldName = Left(strFieldName, Len(strFieldName) - 1)
        strSource = Replace(strSource, strFieldName, GetGlobalVariant(CStr(strFieldName)))
        lngPos1 = InStr(1, strSource, "@")
    Loop
    ReplaceSysPara = strSource
End Function

Private Function ReplaceVoucherItems(strSQL As String, domHead As DOMDocument, domBody As DOMDocument, Optional lngRow As Long) As String
    Dim lngPos1 As Integer
    Dim lngPos2 As Integer
    Dim strFieldName As String
    Dim varField As Variant
    Dim strValue As String
    lngPos1 = InStr(1, strSQL, "[")
    Do While lngPos1 > 0
        lngPos2 = InStr(lngPos1, strSQL, "]")
        If lngPos2 <= 0 Then Exit Do
        strFieldName = Mid(strSQL, lngPos1 + 1, lngPos2 - lngPos1 - 1)
        varField = Split(strFieldName, ",")
        If UBound(varField) = 1 Then
            strValue = GetVoucherItemValue(domHead, domBody, CStr(varField(0)), CStr(varField(1)), lngRow)
            strSQL = Replace(strSQL, "[" + varField(0) + "," + varField(1) + "]", strValue)
        Else
            strSQL = Replace(strSQL, "[" + varField(0) & "]", varField(0) + "")
        End If
        lngPos1 = InStr(lngPos1 + Len(strValue), strSQL, "[")
    Loop
    ReplaceVoucherItems = strSQL
End Function

Private Function GetVoucherItemValue(domHead As DOMDocument, domBody As DOMDocument, strSection As String, strFieldName As String, Optional lngRow As Long) As String
    If strSection = "B" Then
        GetVoucherItemValue = GetBodyItemValue(domBody, strFieldName, lngRow)
    End If
    If strSection = "T" Then
        GetVoucherItemValue = GetHeadItemValue(domHead, strFieldName)
    End If
End Function
Public Function GetGlobalVariant(strName As String) As String
    Select Case LCase(strName)
        Case "@username"
            GetGlobalVariant = m_Login.cUserName
        Case "@curdate"
            GetGlobalVariant = m_Login.CurDate
'        Case "@ifloatraterule"
'            GetGlobalVariant = clsSAWeb.iFloatRateRule
'        Case "@bsaleprice"
'            GetGlobalVariant = IIf(clsSAWeb.bSalePrice, 1, 0)
'        Case "@bmostart"
'            GetGlobalVariant = IIf(clsSAWeb.bMOStart, 1, 0)
'        Case "@bbostart"
'            GetGlobalVariant = IIf(IsDate(clsSAWeb.GetSysDicOption("BO", "dBOFirstDate")), 1, 0)
'        Case "@bmpstart"
'            GetGlobalVariant = IIf(IsDate(clsSAWeb.GetSysDicOption("MP", "dMPFirstDate")), 1, 0)
'        Case "@bmqstart"
'            GetGlobalVariant = IIf(IsDate(clsSAWeb.GetSysDicOption("MQ", "dMQFirstDate")), 1, 0)
'        Case "@sastartdate"
'            GetGlobalVariant = clsSAWeb.SAStDate
'        Case "@bcusinvlimited"
'            GetGlobalVariant = IIf(clsSAWeb.bCusInvLimited, 1, 0)
        Case Else
            GetGlobalVariant = strName
    End Select
End Function
'�жϵ�λ����Ա�Ƿ�ƥ��
Public Function IsPersonMatch(DBConn As ADODB.Connection, Optional cDeptCode As String, Optional CcPersonCode As String) As Boolean
    On Error GoTo err
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset
    If cDeptCode = "" Or CcPersonCode = "" Then
        IsPersonMatch = True
    Else
        Rs.Open "select * from person where cdepcode ='" & cDeptCode & "' and cpersoncode ='" & CcPersonCode & "' ", DBConn, adOpenForwardOnly, adLockReadOnly
        If Not Rs.EOF Then
            IsPersonMatch = True
        Else
            IsPersonMatch = False
        End If
        Rs.Close
    End If
    Exit Function
err:
    IsPersonMatch = False
End Function
