Attribute VB_Name = "BaseFun"
'****************************************
'���̳��ù���˵����
'          ��������
'����ʱ�䣺2009-2-24
'�����ˣ�chenliangc
'****************************************
Option Explicit

Public Const HKEY_LOCAL_MACHINE As Long = &H80000002
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal lngRootKey As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

'����GUID��API��غ���
Private Type guid
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4(8) As Byte
End Type
Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As guid) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long

Public Const ERROR_SUCCESS As Long = 0
Public Const REG_SZ As Long = 1
Public Const KEY_QUERY_VALUE As Long = &H1

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'871��870ע����а�װ·����ֵ��ͬ
Public Const INSTALLKEY870 = "SOFTWARE\UfSoft\WF\V8.700\Install\CurrentInstPath"
Public Const INSTALLKEY871 = "SOFTWARE\UfSoft\WF\V8.700\Install\CurrentInstPath"
Public Const INSTALLITEM = ""


Public gU8Version As String           'U8�汾��
Public strAppPath As String           'U8��װĿ¼

'���������ļ�
Public Sub LoadHelpId(ByRef oForm As Object, ByVal sHelpID As String)
'Ĭ��ȡ871��װ·�������Ϊ�գ�ȡ870��װ·��
    strAppPath = RegRead(INSTALLKEY871, INSTALLITEM)

    If strAppPath = "" Then
        strAppPath = RegRead(INSTALLKEY870, INSTALLITEM)
    End If

    App.HelpFile = strAppPath & "\help\ST_" & g_oLogin.LanguageRegion & ".chm" ' HelpFile
    oForm.HelpContextID = sHelpID
End Sub

'��ȡע����еİ�װĿ¼
Public Function RegRead(ByVal cSubKey As String, ByVal cItem As String) As String
    RegRead = ""
    Dim hKey As Long
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, cSubKey, 0, KEY_QUERY_VALUE, hKey) = ERROR_SUCCESS Then    ' ��ע����
        Dim cTemp As String * 128
        Dim nTemp As Long
        Dim nType As Long
        nType = REG_SZ
        nTemp = 128
        If RegQueryValueEx(hKey, cItem, 0, nType, ByVal cTemp, nTemp) = ERROR_SUCCESS Then       ' ���/������ֵ
            RegRead = Left(cTemp, InStr(1, cTemp, Chr(0)) - 1)
        End If
        RegCloseKey (hKey)                                 ' �ر�ע����
    End If
End Function

'����ǰ�ļ��������
Function Pub_ReadSysCMP() As String
    Dim Strbuffer As String * 128
    Dim lnglstrlen As Long

    Strbuffer = String(128, " ")
    lnglstrlen = GetComputerName(Strbuffer, 128)
    Pub_ReadSysCMP = GetStringFromLPSTR(Strbuffer)
    Pub_ReadSysCMP = Pub_ReadSysCMP & Trim(getCurrentSession)
End Function

Public Function getCurrentSession() As String
    Dim objTerm As Object
    Set objTerm = CreateObject("TermMisc.Terminal")
    getCurrentSession = str(objTerm.GetSessionID)
    Set objTerm = Nothing
End Function

'ȡϵͳ������Ϣ chenliangc
Public Function getAccinformation(strSysID As String, strName As String, conn As Object) As String
    Dim rst As New ADODB.Recordset

    rst.CursorLocation = adUseClient
    rst.Open "Select cValue from accinformation where cSysID=N'" & strSysID & "' and cName=N'" & strName & "'", conn, adOpenForwardOnly, adLockReadOnly, adCmdText
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

'���²���ϵͳ��Ϣ
Public Sub UpdateAccinfo(strSysID As String, strName As String, strValue As String, conn)
    Dim affeceted As Long
    conn.Execute "Update accinformation set cValue=N'" & strValue & "' where cSysId=N'" & strSysID & "' and cname=N'" & strName & "'", affeceted
    If affeceted = 0 Then
        conn.Execute "insert into accinformation(cValue,cSysId,cname) values(N'" & strValue & "' ,N'" & strSysID & "' ,N'" & strName & "')"
    End If
End Sub

'����Ȩ��
Public Function ZwTaskExec(ologin As U8Login.clsLogin, cAuthId As String, bIsLock As Integer, Optional bCopy As Boolean = False) As Boolean


    Dim sMsgTitle As String
    sMsgTitle = GetString("U8.DZ.JA.Res030")
    If ologin.TaskExec(cAuthId, bIsLock) Then
        ZwTaskExec = True
    Else
        Select Case ologin.LogState
        Case 22
            ZwTaskExec = True
        Case Else
            ZwTaskExec = False
            '����ʱ�����û��Ȩ�ޣ��򲻵�����ʾ��
            If Not bCopy Then
                If Trim(ologin.ShareString) = "" Then
                    Dim conn As New ADODB.Connection
                    conn.ConnectionString = ologin.UfDbName
                    conn.Open
                        ReDim varArgs(0)
                        varArgs(0) = GetAuthName(cAuthId, conn)
                        MsgBox GetStringPara("U8.DZ.JA.Res490", varArgs(0)), vbCritical, sMsgTitle
    '                    MsgBox "[" & GetAuthName(cAuthId, conn) & "]������ʱ����ִ�У�", vbCritical, sMsgTitle
                    conn.Close
                    Set conn = Nothing
                Else
                    MsgBox ologin.ShareString, vbCritical, sMsgTitle
                End If
            End If
        End Select
        ologin.ClearError
    End If
End Function

Public Function GetAuthName(AuthID As String, conn As Object)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    sql = "select cauth_name from ufsystem..ua_auth where cauth_id='" & AuthID & "'"
    rs.Open sql, conn
    If Not rs.EOF Then
        GetAuthName = rs!cauth_name
    End If
    rs.Close
End Function

'��ñ�ͷ��ͼ
Public Function GetViewHead(ByRef conn As Connection, CardNum As String) As String
'��ͷ��ͼ
    Dim HeadView As String

    Dim rs As New ADODB.Recordset

    Dim strSql As String
    strSql = "SELECT  BTQName"
    strSql = strSql & " From Vouchers"
    strSql = strSql & " WHERE (CardNumber = '" & CardNum & "') "
    rs.Open strSql, conn
    If Not rs.EOF Then
        HeadView = IIf(IsNull(rs!BTQName), "", rs!BTQName)
    End If
    rs.Close
    Set rs = Nothing
    GetViewHead = HeadView
End Function

'��ñ�����ͼ
Public Function GetViewBody(ByRef conn As Connection, CardNum As String) As String
'������ͼ
    Dim BodyView As String
    Dim strSql As String

    Dim rs As New ADODB.Recordset

    strSql = "SELECT  BWQName, HaveBodyGrid"
    strSql = strSql & " From Vouchers"
    strSql = strSql & " WHERE (CardNumber = '" & CardNum & "') "
    rs.Open strSql, conn
    If Not rs.EOF Then
        BodyView = IIf(IsNull(rs!BWQName), "", rs!BWQName)
    End If
    rs.Close
    Set rs = Nothing
    GetViewBody = BodyView
End Function

'���U8�汾
Public Function GetU8Version(conn As Object) As String
    Dim sqlstr As String
    Dim rs As New ADODB.Recordset
    Dim sVa As String

    sqlstr = "select * from UFSystem..ua_version"
    rs.Open sqlstr, conn, 1, 1
    If Not rs.EOF Then
        sVa = Trim(rs!iVer)
    Else
        sVa = ""
    End If

    If sVa <> "" Then
        If InStr(sVa, "870") > 0 Or InStr(sVa, "871") > 0 Then
            GetU8Version = "871"
        Else
            GetU8Version = "872"
        End If
    Else
        GetU8Version = "871"
    End If
    rs.Close
    Set rs = Nothing
End Function

'�����Ƿ����
'##ModelId=476A5D9500BB
Public Function checkTableExist(conn As Object, sTableName As String) As Boolean
    Dim sqlstr As String
    Dim rs As New ADODB.Recordset
    Dim bVa As Boolean


    sqlstr = "select 1 From sysobjects  where  id = object_id('" & sTableName & "') and   type = 'U'"
    rs.Open sqlstr, conn, 1, 1
    If Not rs.EOF Then
        bVa = True
    Else
        ReDim varArgs(0)
        varArgs(0) = sTableName
        MsgBox GetStringPara("U8.DZ.JA.Res500", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")
        bVa = False
    End If

    rs.Close
    Set rs = Nothing
    checkTableExist = bVa
End Function

'����Ƿ����ý������

Public Function checkJCJYStart(conn As Object) As Boolean
    Dim sqlstr As String
    Dim rs As New ADODB.Recordset
    Dim bVa As Boolean


    sqlstr = "select *  From accinformation where csysid ='ST' and cname =N'bBorrowBusiness' and isnull(cvalue,'') =N'true'"
    rs.Open sqlstr, conn, 1, 1
    If rs.RecordCount > 0 Then
        checkJCJYStart = True
    Else
        checkJCJYStart = False
    End If

    rs.Close
    Set rs = Nothing
    
End Function

Function GetStringFromLPSTR(StrBuf As String) As String
'��LPSTR�ַ���ת��ΪLPCSTR�ַ���
    Dim i As Long
    i = InStr(1, StrBuf, Chr(0), vbTextCompare)
    If i <> 0 Then
        GetStringFromLPSTR = Left(StrBuf, i - 1)
    End If

End Function
'�ַ���ת��Ϊ����
Public Function ConvertStrToDbl(sVa As Variant) As Double
    If IsNull(sVa) Then
        sVa = ""
    End If

    If sVa <> "" And IsNumeric(sVa) Then
        ConvertStrToDbl = CDbl(sVa)
    Else
        ConvertStrToDbl = 0
    End If
End Function

'��֯���ȵ��ַ���
Public Function GetPrecision(ByVal iValue As Long) As String
    Dim i As Integer
    Dim str As String
    For i = 1 To iValue
        str = str & "0"
    Next
    GetPrecision = str
End Function

'�������� ȡdom������ֵ
Public Function GetNodeAtrVal(IXNOde As IXMLDOMNode, sKey As String) As String
'sKey = LCase(sKey)
    If IXNOde.Attributes.getNamedItem(sKey) Is Nothing Then
        GetNodeAtrVal = ""
    Else
        GetNodeAtrVal = IXNOde.Attributes.getNamedItem(sKey).nodeValue
    End If
End Function

Public Function GetTheLastID( _
       login As clsLogin, _
       ByVal oConnection As ADODB.Connection, _
       ByVal sTable As String, _
       ByVal sField As String, _
       Optional ByVal sDataNumFormat As String = "00000000", _
       Optional ByVal sWhereStatement As String = "") As Variant

    Dim sql As String
    Dim oRecordset As New Recordset
    Dim sWhere As String

    On Error GoTo Err_Handler

    GetTheLastID = "0"


    If Trim(sWhereStatement) <> "" Then
        sWhere = sWhere & "  (" & sWhereStatement & ") "
    End If


    '��ȡ����Ȩ��
    'R ��Ȩ��



    sql = "SELECT " & HeadPKFld & " FROM " & sTable & " " _
        & "WHERE " & IIf(sWhere = "", "1=1", sWhere) & " and " & sAuth_ALL _
        & "ORDER BY " & sField & ""


    oRecordset.Open sql, oConnection, adOpenStatic, adLockReadOnly, adCmdText



    If Not oRecordset.EOF Then
        oRecordset.MoveLast
        GetTheLastID = Format(CDbl(oRecordset.Fields(0).Value), sDataNumFormat)
    Else
        GetTheLastID = Format(0, sDataNumFormat)
    End If

Exit_Label:
    On Error GoTo 0
    If Not oRecordset Is Nothing Then
        If oRecordset.State = adStateOpen Then _
           Call oRecordset.Close
    End If
    Set oRecordset = Nothing

    Exit Function
Err_Handler:
    If g_blnDEBUG_MODE Then
        Call ShowDebugForm( _
             bErrorMode:=True, _
             sProcedure:="Function GetTheLastID of Module modFuncL")
    End If

    If Not oRecordset Is Nothing Then
        If oRecordset.State = adStateOpen Then _
           Call oRecordset.Close
    End If
    Set oRecordset = Nothing

    Err.Raise _
            Number:=Err.Number, _
            Source:="Function GetTheLastID of Module modFuncL", _
            Description:=Err.Description

End Function

'                       ��ʾ������Ϣ����
Public Sub ShowDebugForm( _
       Optional ByVal sDebugText As String, _
       Optional ByVal bErrorMode As Boolean = True, _
       Optional ByVal sProcedure As String)

    Dim sText As String

    If bErrorMode Then
        sText = sText _
              & "Debug Text:" & vbTab & sDebugText & vbCrLf & vbCrLf _
              & "Procedure:" & vbTab & sProcedure & vbCrLf _
              & "Error Number:" & vbTab & Err.Number & vbCrLf _
              & "Error Source:" & vbTab & Err.Source & vbCrLf _
              & "Error Description:" & vbTab & Err.Description & vbCrLf
    Else
        sText = sDebugText
    End If

    FrmMsgBox.Text1.Text = sText
    FrmMsgBox.Show vbModal
End Sub
'                       ���ݲ�ͬ����Ϣ���ͣ���ʾ�ض�����Ϣ����
Public Sub ShowErrorInfo( _
       Optional ByVal sHeaderMessage As String, _
       Optional ByVal lMessageType As VbMsgBoxStyle = vbInformation, _
       Optional ByVal lErrorLevel As ErrorLevelConstants = ufsELAllInfo)

    Dim sMessage As String

    sMessage = IIf( _
               Expression:=Trim(sHeaderMessage) = "", _
               TruePart:="", _
               FalsePart:=sHeaderMessage & vbCrLf & vbCrLf)

    Select Case lErrorLevel
    Case ufsELAllInfo
        sMessage = sMessage _
                & GetString("U8.DZ.JA.Res510") & vbTab & Err.Number & vbCrLf _
                & GetString("U8.DZ.JA.Res520") & vbTab & Err.Source & vbCrLf _
                & GetString("U8.DZ.JA.Res530") & vbTab & Err.Description
        
    Case ufsELHeaderAndDescription
        sMessage = sMessage _
               & Err.Description

    End Select

    Call MsgBox(Prompt:=sMessage, _
            Buttons:=lMessageType, Title:=GetString("U8.DZ.JA.Res030") _
                                            )
End Sub

' Precedure             Null2Something
' Purpose
'                       �� NULL ֵ�滻Ϊ "" �� 0 ��ָ����ֵ
'
' Argument(s)
'       vTarget         (Variant) Ҫ�����ж��Ƿ�Ϊ NULL ��ֵ
'       vReplace        [Variant, ""] �����滻��ֵ
'
' Return(s)
'       (Variant)
'       ���ش������ַ���
'
' Note(s)
'       1)              һ���ڶ�ȡ���ݿ�����ֶε�����ʱ����
'                       ���˺�����
Public Function Null2Something( _
       ByVal vTarget As Variant, _
       Optional ByVal vReplace As Variant = "") As Variant

'On Error GoTo Err_Handler

    Null2Something = vTarget

    If IsNull(vTarget) Then
        Null2Something = vReplace
    End If

Exit_Label:
    On Error GoTo 0

    Exit Function
Err_Handler:
    Err.Raise _
            Number:=Err.Number, _
            Source:="Function Null2Something of Module modFuncL", _
            Description:=Err.Description

End Function

'ʱ���m modify by chenliangc
Public Function GetTimeStamp(conn As Object, HeadT As String, PKValue As Long) As String
    On Error GoTo ErrHand
    Dim sql As String, rs As New ADODB.Recordset

    sql = "select convert(nchar,convert(money,ufts),2)  from " _
        & MainTable & " WHERE (" & HeadPKFld & " = " & PKValue & ") "

    Set rs = conn.Execute(sql)
    If Not rs.EOF Then
        GetTimeStamp = rs(0)
    Else
        GetTimeStamp = RecordDeleted    '��¼������,�������ѱ�ɾ��
    End If

    rs.Close
    Set rs = Nothing
    Exit Function

ErrHand:
    GetTimeStamp = RecordError
End Function

'ȡ��ͷDOM������
Public Function GetHeadItemValue(ByVal domHead As DOMDocument, ByVal sKey As String) As String
    If domHead.selectSingleNode("//z:row").Attributes.getNamedItem(sKey) Is Nothing Then
        GetHeadItemValue = ""
    Else
        GetHeadItemValue = domHead.selectSingleNode("//z:row").Attributes.getNamedItem(sKey).nodeValue
    End If
End Function

'�ӵ��ݿؼ���ñ������ݶ���
Public Function GetBodyVouchData(conn As Object, Vouch As Object, BTable As String) As CDataVO
    Dim i As Integer, j As Integer
    Dim DMO As CDMO
    Set DMO = New CDMO
    Dim DVO As CDataVO
    Set DVO = DMO.NewData(conn, BTable)
    Dim DataItem As CDataItem
    For j = 1 To Vouch.BodyRows
        If j = 1 Then
            For i = 1 To DVO.Item(1).Count
                DVO.Item(1).Item(i).Value = Vouch.bodyText(j, DVO.Item(1).Item(i).FieldCode)
            Next i
        Else
            Set DataItem = DMO.GetCloneDataItem(DVO.Item(1))
            For i = 1 To DataItem.Count
                DataItem.Item(i).Value = Vouch.bodyText(j, DataItem.Item(i).FieldCode)
            Next i
            DVO.Add DataItem
        End If
    Next j
    Set GetBodyVouchData = DVO

End Function
'�ӵ��ݿؼ���ñ�ͷ���ݶ���
Public Function GetHeadVouchData(conn As Object, Vouch As Object, HTable As String) As CDataVO
    Dim i As Integer
    Dim DMO As CDMO
    Set DMO = New CDMO
    Dim DVO As CDataVO
    Set DVO = DMO.NewData(conn, HTable)
    For i = 1 To DVO.Item(1).Count
        DVO.Item(1).Item(i).Value = Vouch.headerText(DVO.Item(1).Item(i).FieldCode)
    Next i
    Set GetHeadVouchData = DVO
End Function

'��ȡ������λ������:
'�̶�,����,��
Public Function GetGroupType(cGroupCode As String, conn As Object) As Integer
    On Error GoTo Err_Handler

    Dim sql As String
    Dim rs As New ADODB.Recordset

    sql = "SELECT iGroupType FROM ComputationGroup WHERE cGroupCode='" & cGroupCode & "'"
    rs.Open sql, conn, 1, 1

    If Not rs.EOF Then
        GetGroupType = Val(rs("iGroupType"))
    Else
        GetGroupType = 0
    End If

    rs.Close
    Set rs = Nothing
    Exit Function


Err_Handler:
    rs.Close
    Set rs = Nothing
    GetGroupType = 0
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Function

'���������ʼ������
'�˴�ȡ����������,Ӧ��ʱ�ɸ���ʵ��ҵ���ȡ����ģ������
'ture ����Ϊ��,false ����Ϊ��
Public Function GetFloatRateRule(conn As Object) As Boolean
    On Error GoTo Err_Handler

    Dim bFloatRateRule As Boolean
    Dim sql As String
    Dim rs As New ADODB.Recordset

    '���۹���
    'SQL = "select ChgBenchmark from mom_parameter"
    '������
    sql = "select cvalue from accinformation where csysid='AA' AND CName='iFloatRateRule'"
    rs.Open sql, conn, 1, 1
    If Not rs.EOF Then
        bFloatRateRule = IIf(rs(0) = 0, False, True)
    End If

    GetFloatRateRule = bFloatRateRule

    rs.Close
    Set rs = Nothing

    Exit Function


Err_Handler:
    rs.Close
    Set rs = Nothing
    GetFloatRateRule = True
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Function
''ȡ�õ��ݺ�
''bGetFormatOnly true:ֻȡ�õ��ݺ���ǰ׺��ʽ��
Public Function GetVoucherNO(conn As Object, domHead As DOMDocument, strVouchType As String, errMsg As String, strVouchNo As String, Optional DomFormat As DOMDocument, Optional bGetFormatOnly As Boolean, Optional bSelfFormat As Boolean, Optional bGetRealNo As Boolean) As Boolean
    GetVoucherNO = BOGetVoucherNO(conn, domHead, strVouchType, errMsg, strVouchNo, DomFormat, bGetFormatOnly, bSelfFormat, , bGetRealNo)
End Function

'����ʾ����ģ��ID
Public Function GetVoucherID(conn As Object, strCardNumber As String) As String
    Dim rs As New ADODB.Recordset
    Dim sqlstr As String
    Dim sVa As String

    sqlstr = "Select VT_ID From vouchertemplates where vt_cardnumber='" & strCardNumber & "' and vt_templatemode='0'"
    rs.Open sqlstr, conn, 1, 1
    If Not rs.EOF Then
        sVa = rs!VT_ID
    Else
        sVa = ""
    End If
    rs.Close
    Set rs = Nothing

    GetVoucherID = sVa
End Function

''ȡ�õ��ݺ�
''bGetFormatOnly true:ֻȡ�õ��ݺ���ǰ׺��ʽ��
Public Function BOGetVoucherNO(puConn As ADODB.Connection, domHead As DOMDocument, strVouchType As String, errMsg As String, strVouchNo As String, Optional DomFormat As DOMDocument, Optional bGetFormatOnly As Boolean, Optional bSelfFormat As Boolean, Optional bRepeatRedo As Boolean, Optional bGetRealNo As Boolean, Optional bSynchronization As Boolean = False) As Boolean
    Dim strCardNum As String
    Dim objBillNo As New UFBillComponent.clsBillComponent
    Dim ele As IXMLDOMElement
    Dim strTableName As String
    Dim strFieldName As String
    Dim strSeed As String

    'm_SysInfor = clsInfor.Information

    If IsMissing(DomFormat) = True Or DomFormat Is Nothing Then
        Set DomFormat = New DOMDocument
    End If
    On Error GoTo DoErr
    errMsg = ""


    If objBillNo.InitBill(puConn.ConnectionString, strVouchType) = False Then
        errMsg = GetString("U8.DZ.JA.Res540")
        BOGetVoucherNO = False
        Exit Function
    End If

    If IsMissing(bSelfFormat) = True Or bSelfFormat = False Then
        If DomFormat.loadXML(objBillNo.GetBillFormat) = False Then
            errMsg = GetString("U8.DZ.JA.Res550")
            BOGetVoucherNO = False
            Exit Function
        End If
    End If

    Set ele = DomFormat.selectSingleNode("//���ݱ��")
    bRepeatRedo = ele.getAttribute("�غ��Զ���ȡ") Or Not ele.getAttribute("�����ֹ��޸�")

    If bGetFormatOnly = True Then
        GoTo DoExit
    End If

    If IsMissing(bSelfFormat) = True Or bSelfFormat = False Then
        For Each ele In DomFormat.selectNodes("//���ݱ��/ǰ׺[@��������!=5]")
            If ele.Attributes.getNamedItem("��������").nodeValue <> "Զ�̺�" Then
                If LCase(Right(ele.Attributes.getNamedItem("Դ���ֶ���").nodeValue, 4)) = "date" Then
                    Dim strDate As String
                    strDate = GetHeadItemValue(domHead, ele.Attributes.getNamedItem("Դ���ֶ���").nodeValue)    '�������ֶ���
                    If InStr(strDate, "T") > 0 Then
                        strDate = Left(strDate, InStr(strDate, "T") - 1)
                    End If
                    ele.setAttribute "����", strDate
                Else
                    ele.setAttribute "����", GetHeadItemValue(domHead, ele.Attributes.getNamedItem("Դ���ֶ���").nodeValue)    '�������ֶ���
                End If
                '            Else
                '                ele.setAttribute "����", m_SysInfor.RemoteID
            End If
        Next
    End If

    Dim strTmpCode As String
    Dim sWhere As String
    strTmpCode = objBillNo.GetNumber(DomFormat.xml, False)
    If Not bSynchronization Then
        If strTmpCode = strVouchNo Or strVouchNo = "" Then
            'ע��,�˴�����ʹ��objBillNo.GetNumber,���򵥾ݺ���ˮ�Ų����Զ�����
            strVouchNo = objBillNo.GetNumberWithOutMTS(DomFormat.xml, bGetRealNo)
        Else
            strVouchNo = strTmpCode
        End If
    Else
        Select Case strVouchType
        Case "BO90"
            strTableName = "bom_chgmain"
            strFieldName = "cChgCode"
        Case "BO91"
            strTableName = "bom_batchChg"
            strFieldName = "cChgCode"
        End Select
        Set ele = DomFormat.selectSingleNode("//ǰ׺")
        If Not ele Is Nothing Then
            strSeed = ele.getAttribute("����")
        End If

        objBillNo.DataSynchronization strTableName, strFieldName, strVouchNo, strSeed, sWhere
        'bug fix  �Ƴ��� 2003-09-09
        'strVouchNo = objBillNo.GetNumber(DomFormat.xml, False)
        strTmpCode = objBillNo.GetNumber(DomFormat.xml, False)
        If strTmpCode = strVouchNo Then
            strVouchNo = objBillNo.GetNumber(DomFormat.xml, bGetRealNo)
        Else
            strVouchNo = strTmpCode
        End If
    End If
DoExit:
    BOGetVoucherNO = True
    Exit Function
DoErr:
    errMsg = "Function GetVoucherNO:" & Err.Description & objBillNo.GetLastErrorA
    BOGetVoucherNO = False
End Function

'********************************************
'2008-11-17
'Ϊƥ��872��LP���������۸��ٷ�ʽ�Ĵ���
'sSosID        ���۶�����ID
'sDemandType   ���۶�������
'sDemandCode   ���۶��������
Public Sub GetSoDemandType(sSosId As String, ByRef sDemandType As String, ByRef sDemandCode As String, conn As Object)
    Dim rstmp As New ADODB.Recordset
    Dim sqltmp As String

    If gU8Version = "872" Then
        sqltmp = "Select * From SO_SODetails Where iSOsID ='" & sSosId & "'"
        rstmp.Open sqltmp, conn, 1, 1
        If Not rstmp.EOF Then
            sDemandType = Null2Something(rstmp!idemandtype)
            Select Case sDemandType
            Case "0"
                sDemandCode = ""
            Case "1"   '�����и���
                sDemandCode = sSosId
            Case "4"   '���۶������ٷ�ʽΪ�������
                If Trim(rstmp!cDemandCode) = "" Then
                    sDemandCode = "Systemdefault"
                Else
                    sDemandCode = Null2Something(rstmp!cDemandCode)
                End If
            Case "5"   '���۶�������
                sDemandCode = Null2Something(rstmp!csocode)
            Case Else
                sDemandType = "0"
                sDemandCode = ""
            End Select
            '            sDemandType = "0"
            '            sDemandCode = ""
        Else
            sDemandType = "0"
            sDemandCode = ""
        End If
        rstmp.Close
        Set rstmp = Nothing
    Else
        sDemandType = "0"
        sDemandCode = "0"
    End If
End Sub

'���õ��ݿؼ���Ŀд״̬
Public Function SetVouchItemState(Voucher As Object, strFieldName As String, CardSection As SectionsConstants, bCanModify As Boolean) As Boolean

    On Error GoTo Err

    With Voucher
        If Not .ItemState(strFieldName, CardSection) Is Nothing Then
            If .ItemState(strFieldName, CardSection).bCanModify <> bCanModify Then
                If CardSection = siHeader Then
                    .EnableHead strFieldName, bCanModify
                Else
                    If Not .ItemState(strFieldName, CardSection) Is Nothing Then
                        .ItemState(strFieldName, CardSection).bCanModify = bCanModify
                    End If
                End If
            End If
        End If
    End With
    Exit Function
Err:
    MsgBox Err.Description, vbExclamation, GetString("U8.DZ.JA.Res030")
End Function

'��߽��ۿ���
Public Function bGetMPService(ByVal m_CardNum As String, ByVal domHead As DOMDocument, ByVal domBody As DOMDocument, conn As Object, login As clsLogin) As Boolean
    Dim objSCMMPCost As Object
    Dim iErr As Integer
    '2008-01-31 ��Ҫ�޸�����ģ��
    bGetMPService = False
    Set objSCMMPCost = CreateObject("U8SCMMpCostService.clsMPCostPassword")
    objSCMMPCost.Initial "PU", login, conn

    Call objSCMMPCost.bAnalyseMpCostDom(m_CardNum, domHead, domBody)
    Call objSCMMPCost.DoShowMpCostInputBox("", iErr)
    Set objSCMMPCost = Nothing
    If iErr = 0 Or iErr = -1 Then
        bGetMPService = False
    Else
        bGetMPService = True
    End If
End Function

'**********************************************
'*  ����˵���� �������
'**********************************************
Public Function getKL(conn As Object, ByVal vCusCode As String, ByVal vInvCode As String, ByVal vFree1 As String, ByVal vFree2 As String, ByVal vddate As String, ByVal strQuantity As Double, ByVal clsSys As Object, ByRef kl As Double) As Boolean
    Dim strPara As String, strErr As String, DOMPara As New DOMDocument, iQuotedPrice As Double, fCusMinPrice As Double, fSalecost As Double, SettleCost As Double

    strPara = "<Data> "
    strPara = strPara & "<row name='ccuscode' datatype='202' length='20' type='1' value ='" + vCusCode + "' />"
    strPara = strPara & "<row name='cinvcode' datatype='202' length='20' type='1' value ='" + vInvCode + "' />"
    If Mid(clsSys.FreePriceType, Val(Right("cfree1", 1)), 1) = "1" Then
        strPara = strPara & "<row name='sfree1' datatype='202' length='20' type='1' value ='" + vFree1 + "' />"
    Else
        strPara = strPara & "<row name='sfree1' datatype='202' length='20' type='1' value ='' />"
    End If
    If Mid(clsSys.FreePriceType, Val(Right("cfree2", 1)), 1) = "1" Then
        strPara = strPara & "<row name='sfree2' datatype='202' length='20' type='1' value ='" + vFree2 + "' />"
    Else
        strPara = strPara & "<row name='sfree2' datatype='202' length='20' type='1' value ='' />"
    End If
    strPara = strPara & "<row name='sdate' datatype='202' length='10' type='1' value ='" + vddate + "' />"
    strPara = strPara & "<row name='quantity' datatype='5' length='10' type='1' value ='" + CStr(Abs(strQuantity)) + "' />"
    strPara = strPara & "<row name='nInvPrice' datatype='5' length='10' type='2' value ='' />"
    strPara = strPara & "<row name='nSalePrice' datatype='5' length='10' type='2' value ='' />"
    strPara = strPara & "<row name='nInvNowPrice' datatype='5' length='10' type='2' value ='' />"
    strPara = strPara & "<row name='nDisRate' datatype='5' length='10' type='2' value ='' />"
    strPara = strPara & "<row name='minPrice' datatype='5' length='10' type='2' value ='' />"
    strPara = strPara & "<row name='sErr' datatype='202' length='200' type='2' value ='' />"
    strPara = strPara & " </Data>"

    DOMPara.loadXML strPara
    strErr = ""
    Dim fCusMinPriceTmp As Double
    If clsSys.bUseDatePrice = True Then
        'ȡ������
        ExecGetPriceProc conn, "SA_GetSalesPrice", DOMPara
        iQuotedPrice = Val(DOMPara.documentElement.selectSingleNode("row[@name='nInvPrice']").Attributes.getNamedItem("value").Text)
        fCusMinPrice = Val(DOMPara.documentElement.selectSingleNode("row[@name='minPrice']").Attributes.getNamedItem("value").Text)
        If Val(fCusMinPrice) <> 0 Then
            fCusMinPriceTmp = fCusMinPrice
        End If
    End If
    If iQuotedPrice = 0 Then
        ExecGetPriceProc conn, "SA_FetchPrice", DOMPara
    End If

    strErr = DOMPara.documentElement.selectSingleNode("row[@name='sErr']").Attributes.getNamedItem("value").Text
    'CountPrice = strErr
    iQuotedPrice = Val(DOMPara.documentElement.selectSingleNode("row[@name='nInvPrice']").Attributes.getNamedItem("value").Text)
    fSalecost = Val(DOMPara.documentElement.selectSingleNode("row[@name='nSalePrice']").Attributes.getNamedItem("value").Text)
    SettleCost = Val(DOMPara.documentElement.selectSingleNode("row[@name='nInvNowPrice']").Attributes.getNamedItem("value").Text)
    kl = Val(DOMPara.documentElement.selectSingleNode("row[@name='nDisRate']").Attributes.getNamedItem("value").Text)

    getKL = True
End Function
'**********************************************
'*  ����˵���� �������  (871��)
'**********************************************
Public Function getKL871(conn As Object, ByVal vCusCode As String, ByVal vInvCode As String, ByVal vFree1 As String, ByVal vFree2 As String, ByVal vddate As String, ByVal strQuantity As Double, ByVal strExchName As String, ByVal clsSys As Object, ByRef kl As Double) As Boolean
    Dim strPara As String, strErr As String, DOMPara As New DOMDocument, iQuotedPrice As Double, fCusMinPrice As Double, fSalecost As Double, SettleCost As Double
    'Dim DOMPara As New DOMDocument
    'Dim strPara As String
    '    Dim strQuantity As String
    Dim ele As IXMLDOMElement
    'Dim strErr As String

    '    If blnQtyPrice Then
    '        If Not IsMissing(domBody) Then
    '            strQuantity = GetBodyItemValue(domBody, "iquantity", 0)
    '        End If
    '    Else
    '        strQuantity = "0"
    '    End If

    strPara = "<Data> "
    strPara = strPara & "<row name='ccuscode' datatype='202' length='20' type='1' value ='" + vCusCode + "' />"
    strPara = strPara & "<row name='cinvcode' datatype='202' length='20' type='1' value ='" + vInvCode + "' />"
    If Mid(clsSys.FreePriceType, Val(Right("cfree1", 1)), 1) = "1" Then
        strPara = strPara & "<row name='sfree1' datatype='202' length='20' type='1' value ='" + vFree1 + "' />"
    Else
        strPara = strPara & "<row name='sfree1' datatype='202' length='20' type='1' value ='' />"
    End If
    If Mid(clsSys.FreePriceType, Val(Right("cfree2", 1)), 1) = "1" Then
        strPara = strPara & "<row name='sfree2' datatype='202' length='20' type='1' value ='" + vFree2 + "' />"
    Else
        strPara = strPara & "<row name='sfree2' datatype='202' length='20' type='1' value ='' />"
    End If
    strPara = strPara & "<row name='sdate' datatype='202' length='10' type='1' value ='" + vddate + "' />"
    strPara = strPara & "<row name='exchname' datatype='202' length='8' type='1' value ='" + strExchName + "' />"
    strPara = strPara & "<row name='quantity' datatype='5' length='10' type='1' value ='" + CStr(Abs(Val(strQuantity))) + "' />"
    strPara = strPara & "<row name='bsales' datatype='11' length='1' type='1' value ='1' />"
    strPara = strPara & "<row name='nInvPrice' datatype='5' length='10' type='2' value ='' />"
    strPara = strPara & "<row name='nSalePrice' datatype='5' length='10' type='2' value ='' />"
    strPara = strPara & "<row name='nInvNowPrice' datatype='5' length='10' type='2' value ='' />"
    strPara = strPara & "<row name='nDisRate' datatype='5' length='10' type='2' value ='' />"
    strPara = strPara & "<row name='minPrice' datatype='5' length='10' type='2' value ='' />"
    strPara = strPara & "<row name='sErr' datatype='202' length='200' type='2' value ='' />"
    strPara = strPara & " </Data>"
    DOMPara.loadXML strPara
    strErr = ""
    Dim fCusMinPriceTmp As Double
    If clsSys.bUseDatePrice = True Then
        'ȡ������
        ExecGetPriceProc871 conn, DOMPara, True, strExchName
        iQuotedPrice = Val(DOMPara.documentElement.selectSingleNode("row[@name='nInvPrice']").Attributes.getNamedItem("value").Text)
        fCusMinPrice = Val(DOMPara.documentElement.selectSingleNode("row[@name='minPrice']").Attributes.getNamedItem("value").Text)
    End If
    If iQuotedPrice = 0 Then
        ExecGetPriceProc871 conn, DOMPara, False, strExchName
        iQuotedPrice = Val(DOMPara.documentElement.selectSingleNode("row[@name='nInvPrice']").Attributes.getNamedItem("value").Text)
        If fCusMinPrice = 0 Then fCusMinPrice = Val(DOMPara.documentElement.selectSingleNode("row[@name='minPrice']").Attributes.getNamedItem("value").Text)
    End If
    If iQuotedPrice = 0 And clsSys.bUseDatePrice = True Then
        ExecGetPriceProc871 conn, DOMPara, True, ""
        iQuotedPrice = Val(DOMPara.documentElement.selectSingleNode("row[@name='nInvPrice']").Attributes.getNamedItem("value").Text)
        If fCusMinPrice = 0 Then fCusMinPrice = Val(DOMPara.documentElement.selectSingleNode("row[@name='minPrice']").Attributes.getNamedItem("value").Text)
    End If
    If iQuotedPrice = 0 Then
        ExecGetPriceProc871 conn, DOMPara, False, ""
        iQuotedPrice = Val(DOMPara.documentElement.selectSingleNode("row[@name='nInvPrice']").Attributes.getNamedItem("value").Text)
        If fCusMinPrice = 0 Then fCusMinPrice = Val(DOMPara.documentElement.selectSingleNode("row[@name='minPrice']").Attributes.getNamedItem("value").Text)
    End If
    strErr = DOMPara.documentElement.selectSingleNode("row[@name='sErr']").Attributes.getNamedItem("value").Text
    getKL871 = True
    iQuotedPrice = Val(DOMPara.documentElement.selectSingleNode("row[@name='nInvPrice']").Attributes.getNamedItem("value").Text)
    fSalecost = Val(DOMPara.documentElement.selectSingleNode("row[@name='nSalePrice']").Attributes.getNamedItem("value").Text)
    SettleCost = Val(DOMPara.documentElement.selectSingleNode("row[@name='nInvNowPrice']").Attributes.getNamedItem("value").Text)
    kl = Val(DOMPara.documentElement.selectSingleNode("row[@name='nDisRate']").Attributes.getNamedItem("value").Text)
    If SettleCost = 0 Then
        SettleCost = iQuotedPrice * kl / 100
    End If
    '            If fCusMinPrice = 0 Then
    '                fCusMinPrice = GetInvFieldValue("iinvlscost", cinvcode)
    '            End If
    Set DOMPara = Nothing
End Function

'**********************************************
'*  ����˵���� �������  (872��)       xin 2008-10-22
'**********************************************
Public Function getKL872(conn As Object, ByVal vCusCode As String, ByVal vInvCode As String, ByVal vFree1 As String, ByVal vFree2 As String, ByVal vFree3 As String, ByVal vFree4 As String, ByVal vFree5 As String, ByVal vFree6 As String, ByVal vFree7 As String, ByVal vFree8 As String, ByVal vFree9 As String, ByVal vFree10 As String, ByVal vddate As String, ByVal strQuantity As Double, ByVal strExchName As String, ByVal clsSys As Object, ByRef kl As Double) As Boolean
    Dim strPara As String, strErr As String, DOMPara As New DOMDocument, iQuotedPrice As Double, fCusMinPrice As Double, fSalecost As Double, SettleCost As Double
    'Dim DOMPara As New DOMDocument
    'Dim strPara As String
    '    Dim strQuantity As String
    Dim ele As IXMLDOMElement
    'Dim strErr As String

    '    If blnQtyPrice Then
    '        If Not IsMissing(domBody) Then
    '            strQuantity = GetBodyItemValue(domBody, "iquantity", 0)
    '        End If
    '    Else
    '        strQuantity = "0"
    '    End If

    strPara = "<Data> "
    strPara = strPara & "<row name='ccuscode' datatype='202' length='20' type='1' value ='" + vCusCode + "' />"
    strPara = strPara & "<row name='cinvcode' datatype='202' length='20' type='1' value ='" + vInvCode + "' />"
    If Mid(clsSys.FreePriceType, Val(Right("cfree1", 1)), 1) = "1" Then
        strPara = strPara & "<row name='sfree1' datatype='202' length='20' type='1' value ='" + vFree1 + "' />"
    Else
        strPara = strPara & "<row name='sfree1' datatype='202' length='20' type='1' value ='' />"
    End If
    If Mid(clsSys.FreePriceType, Val(Right("cfree2", 1)), 1) = "1" Then
        strPara = strPara & "<row name='sfree2' datatype='202' length='20' type='1' value ='" + vFree2 + "' />"
    Else
        strPara = strPara & "<row name='sfree2' datatype='202' length='20' type='1' value ='' />"
    End If
    If Mid(clsSys.FreePriceType, Val(Right("cfree3", 1)), 1) = "1" Then
        strPara = strPara & "<row name='sfree3' datatype='202' length='20' type='1' value ='" + vFree3 + "' />"
    Else
        strPara = strPara & "<row name='sfree3' datatype='202' length='20' type='1' value ='' />"
    End If
    If Mid(clsSys.FreePriceType, Val(Right("cfree4", 1)), 1) = "1" Then
        strPara = strPara & "<row name='sfree4' datatype='202' length='20' type='1' value ='" + vFree4 + "' />"
    Else
        strPara = strPara & "<row name='sfree4' datatype='202' length='20' type='1' value ='' />"
    End If
    If Mid(clsSys.FreePriceType, Val(Right("cfree5", 1)), 1) = "1" Then
        strPara = strPara & "<row name='sfree5' datatype='202' length='20' type='1' value ='" + vFree5 + "' />"
    Else
        strPara = strPara & "<row name='sfree5' datatype='202' length='20' type='1' value ='' />"
    End If
    If Mid(clsSys.FreePriceType, Val(Right("cfree6", 1)), 1) = "1" Then
        strPara = strPara & "<row name='sfree6' datatype='202' length='20' type='1' value ='" + vFree6 + "' />"
    Else
        strPara = strPara & "<row name='sfree6' datatype='202' length='20' type='1' value ='' />"
    End If
    If Mid(clsSys.FreePriceType, Val(Right("cfree7", 1)), 1) = "1" Then
        strPara = strPara & "<row name='sfree7' datatype='202' length='20' type='1' value ='" + vFree7 + "' />"
    Else
        strPara = strPara & "<row name='sfree7' datatype='202' length='20' type='1' value ='' />"
    End If
    If Mid(clsSys.FreePriceType, Val(Right("cfree8", 1)), 1) = "1" Then
        strPara = strPara & "<row name='sfree8' datatype='202' length='20' type='1' value ='" + vFree8 + "' />"
    Else
        strPara = strPara & "<row name='sfree8' datatype='202' length='20' type='1' value ='' />"
    End If
    If Mid(clsSys.FreePriceType, Val(Right("cfree9", 1)), 1) = "1" Then
        strPara = strPara & "<row name='sfree9' datatype='202' length='20' type='1' value ='" + vFree9 + "' />"
    Else
        strPara = strPara & "<row name='sfree9' datatype='202' length='20' type='1' value ='' />"
    End If
    If Mid(clsSys.FreePriceType, Val(Right("cfree10", 2)), 1) = "1" Then
        strPara = strPara & "<row name='sfree10' datatype='202' length='20' type='1' value ='" + vFree10 + "' />"
    Else
        strPara = strPara & "<row name='sfree10' datatype='202' length='20' type='1' value ='' />"
    End If

    strPara = strPara & "<row name='sdate' datatype='202' length='10' type='1' value ='" + vddate + "' />"
    strPara = strPara & "<row name='exchname' datatype='202' length='8' type='1' value ='" + strExchName + "' />"
    strPara = strPara & "<row name='quantity' datatype='5'  length='10' type='1' value ='" + CStr(Abs(Val(strQuantity))) + "' />"
    strPara = strPara & "<row name='bsales' datatype='11' length='1' type='1' value ='1' />"
    strPara = strPara & "<row name='nInvPrice' datatype='5' length='10' type='2' value ='' />"
    strPara = strPara & "<row name='nSalePrice' datatype='5' length='10' type='2' value ='' />"
    strPara = strPara & "<row name='nInvNowPrice' datatype='5' length='10' type='2' value ='' />"
    strPara = strPara & "<row name='nDisRate' datatype='5' length='10' type='2' value ='' />"
    strPara = strPara & "<row name='minPrice' datatype='5' length='10' type='2' value ='' />"
    strPara = strPara & "<row name='sErr' datatype='202' length='200' type='2' value ='' />"
    strPara = strPara & " </Data>"
    DOMPara.loadXML strPara
    strErr = ""
    Dim fCusMinPriceTmp As Double
    If clsSys.bUseDatePrice = True Then
        'ȡ������
        ExecGetPriceProc872 conn, DOMPara, True, strExchName
        iQuotedPrice = Val(DOMPara.documentElement.selectSingleNode("row[@name='nInvPrice']").Attributes.getNamedItem("value").Text)
        fCusMinPrice = Val(DOMPara.documentElement.selectSingleNode("row[@name='minPrice']").Attributes.getNamedItem("value").Text)
    End If
    If iQuotedPrice = 0 Then
        ExecGetPriceProc872 conn, DOMPara, False, strExchName
        iQuotedPrice = Val(DOMPara.documentElement.selectSingleNode("row[@name='nInvPrice']").Attributes.getNamedItem("value").Text)
        If fCusMinPrice = 0 Then fCusMinPrice = Val(DOMPara.documentElement.selectSingleNode("row[@name='minPrice']").Attributes.getNamedItem("value").Text)
    End If
    If iQuotedPrice = 0 And clsSys.bUseDatePrice = True Then
        ExecGetPriceProc872 conn, DOMPara, True, ""
        iQuotedPrice = Val(DOMPara.documentElement.selectSingleNode("row[@name='nInvPrice']").Attributes.getNamedItem("value").Text)
        If fCusMinPrice = 0 Then fCusMinPrice = Val(DOMPara.documentElement.selectSingleNode("row[@name='minPrice']").Attributes.getNamedItem("value").Text)
    End If
    If iQuotedPrice = 0 Then
        ExecGetPriceProc872 conn, DOMPara, False, ""
        iQuotedPrice = Val(DOMPara.documentElement.selectSingleNode("row[@name='nInvPrice']").Attributes.getNamedItem("value").Text)
        If fCusMinPrice = 0 Then fCusMinPrice = Val(DOMPara.documentElement.selectSingleNode("row[@name='minPrice']").Attributes.getNamedItem("value").Text)
    End If
    strErr = DOMPara.documentElement.selectSingleNode("row[@name='sErr']").Attributes.getNamedItem("value").Text
    getKL872 = True
    iQuotedPrice = Val(DOMPara.documentElement.selectSingleNode("row[@name='nInvPrice']").Attributes.getNamedItem("value").Text)
    fSalecost = Val(DOMPara.documentElement.selectSingleNode("row[@name='nSalePrice']").Attributes.getNamedItem("value").Text)
    SettleCost = Val(DOMPara.documentElement.selectSingleNode("row[@name='nInvNowPrice']").Attributes.getNamedItem("value").Text)
    kl = Val(DOMPara.documentElement.selectSingleNode("row[@name='nDisRate']").Attributes.getNamedItem("value").Text)
    If SettleCost = 0 Then
        SettleCost = iQuotedPrice * kl / 100
    End If
    '            If fCusMinPrice = 0 Then
    '                fCusMinPrice = GetInvFieldValue("iinvlscost", cinvcode)
    '            End If
    Set DOMPara = Nothing
End Function


'**********************************************
'*  ����˵���� �������2
'**********************************************
Public Function getKL2(conn As Object, strCusCode As String, strInvCode As String, iquantity As Double, ByRef kl2 As Double, ByRef errMsg As String, Optional sKey As String) As Boolean
    Dim cmdGetKL As New ADODB.Command
    Dim parGetKl As New ADODB.Parameter

    On Error GoTo DoErr

    cmdGetKL.CommandText = "SA_FetchQuantityDisRate"
    cmdGetKL.CommandType = adCmdStoredProc
    Set parGetKl = cmdGetKL.CreateParameter("sCusCode", adVarChar, adParamInput, 20, strCusCode)
    cmdGetKL.Parameters.Append parGetKl
    Set parGetKl = cmdGetKL.CreateParameter("sInvCode", adVarChar, adParamInput, 20, strInvCode)
    cmdGetKL.Parameters.Append parGetKl
    Set parGetKl = cmdGetKL.CreateParameter("nQuantity", adDouble, adParamInput, , str(iquantity))
    cmdGetKL.Parameters.Append parGetKl
    Set parGetKl = cmdGetKL.CreateParameter("nDisRate", adDouble, adParamOutput)
    cmdGetKL.Parameters.Append parGetKl
    cmdGetKL.ActiveConnection = conn
    cmdGetKL.Execute
    kl2 = IIf(IsNull(cmdGetKL.Parameters("nDisRate").Value), 0, cmdGetKL.Parameters("nDisRate").Value)
    'kl2 = CDbl(FormatNum(kl2, clsSys.cKLBit))
    Set cmdGetKL = Nothing
    getKL2 = True
    On Error GoTo 0
    Exit Function
DoErr:
    errMsg = Err.Description
    getKL2 = False
    Set cmdGetKL = Nothing
    On Error GoTo 0
End Function

'****************************************************
'** ����˵�������������۵���
'****************************************************
Public Function GetPrice2(ByRef bdom As DOMDocument) As Boolean
    Dim dataXmlNode As IXMLDOMNode, rowNode As IXMLDOMElement
    Dim viquantity As Double, vfSalePrice As Double, vfSaleCost As Double
    Dim i As Long
    Dim retvalue As Boolean

    retvalue = True
    '�õ�������
    Set dataXmlNode = bdom.selectSingleNode("//rs:data")
    If Not (dataXmlNode Is Nothing) Then
        'ѭ����������¼
        For i = 0 To dataXmlNode.childNodes.Length - 1
            '�õ�����
            Set rowNode = dataXmlNode.childNodes(i).selectSingleNode("//z:row")
            If Not (rowNode Is Nothing) Then
                If Not ((rowNode.Attributes.getNamedItem("fsaleprice") Is Nothing) Or (rowNode.Attributes.getNamedItem("iquantity") Is Nothing)) Then
                    If IsNumeric(rowNode.Attributes.getNamedItem("fsaleprice").nodeValue) And IsNumeric(rowNode.Attributes.getNamedItem("iquantity").nodeValue) Then
                        vfSalePrice = rowNode.Attributes.getNamedItem("fsaleprice").nodeValue     '���۽��
                        viquantity = rowNode.Attributes.getNamedItem("iquantity").nodeValue      '����
                        If viquantity > 0 Then
                            rowNode.setAttribute "fsalecost", Format(vfSalePrice / viquantity, m_sPriceFmtSA)       '���۵���
                        Else
                            retvalue = False
                        End If
                    Else
                        retvalue = False
                    End If
                Else
                    retvalue = False
                End If
            Else
                retvalue = False
            End If
        Next
    Else
        retvalue = False
    End If

    GetPrice2 = retvalue
End Function

'
Public Function ExecGetPriceProc871(conn As ADODB.Connection, DOMPara As DOMDocument, blnSales As Boolean, strExchName As String) As String
    Dim ele As IXMLDOMElement
    Dim strName As String
    Dim StrType As String
    Dim strDataType As String
    Dim strLen As String
    Dim strValue As String

    Set ele = DOMPara.selectSingleNode("//row[@name='exchname']")
    ele.Attributes.getNamedItem("value").nodeValue = strExchName
    Set ele = DOMPara.selectSingleNode("//row[@name='bsales']")
    ele.Attributes.getNamedItem("value").nodeValue = IIf(blnSales, "1", "0")
    ExecGetPriceProc conn, "SA_FetchPrice", DOMPara
End Function

'872 xin 2008-10-22
Public Function ExecGetPriceProc872(conn As ADODB.Connection, DOMPara As DOMDocument, blnSales As Boolean, strExchName As String) As String
    Dim ele As IXMLDOMElement
    Dim strName As String
    Dim StrType As String
    Dim strDataType As String
    Dim strLen As String
    Dim strValue As String

    Set ele = DOMPara.selectSingleNode("//row[@name='exchname']")
    ele.Attributes.getNamedItem("value").nodeValue = strExchName
    Set ele = DOMPara.selectSingleNode("//row[@name='bsales']")
    ele.Attributes.getNamedItem("value").nodeValue = IIf(blnSales, "1", "0")
    ExecGetPriceProc conn, "SA_FetchPrice", DOMPara
End Function


Public Function ExecGetPriceProc(conn As ADODB.Connection, strProcName As String, DOMPara As DOMDocument) As String
    Dim cmdPrice As New ADODB.Command, parPrice As New ADODB.Parameter
    '    Dim domPara As New DOMDocument
    Dim lst As IXMLDOMNodeList
    Dim ele As IXMLDOMElement
    Dim strName As String
    Dim StrType As String
    Dim strDataType As String
    Dim strLen As String
    Dim strValue As String

    '    domPara.loadXML strPara
    cmdPrice.CommandText = strProcName
    cmdPrice.CommandType = adCmdStoredProc
    Set lst = DOMPara.documentElement.childNodes
    For Each ele In lst
        strName = ele.Attributes.getNamedItem("name").Text
        strDataType = ele.Attributes.getNamedItem("datatype").Text
        StrType = ele.Attributes.getNamedItem("type").Text
        strLen = ele.Attributes.getNamedItem("length").Text
        strValue = ele.Attributes.getNamedItem("value").Text
        If Val(StrType) = 1 Then
            Set parPrice = cmdPrice.CreateParameter(strName, Val(strDataType), Val(StrType), strLen, strValue)
        Else
            Set parPrice = cmdPrice.CreateParameter(strName, Val(strDataType), Val(StrType), strLen)
        End If
        cmdPrice.Parameters.Append parPrice
    Next
    cmdPrice.ActiveConnection = conn
    cmdPrice.Execute
    Set lst = DOMPara.selectNodes("//row[@type='2']")
    For Each ele In lst
        strName = ele.Attributes.getNamedItem("name").Text
        strValue = IIf(IsNull(cmdPrice.Parameters(strName).Value), "", cmdPrice.Parameters(strName).Value)
        ele.setAttribute "value", strValue
    Next
    Set cmdPrice = Nothing
    '    Set domPara = Nothing
End Function


'���ܣ�ʹ�ô�����롢��Ӧ�̱�������˰���ۡ���˰�����Լ�˰�ʣ�ȡ�۳ɹ�����true
'���������
'������룺InvCode
'��Ӧ�̱��룺VendorCode
'���ؽ����boolean
'
'���ղɹ�ȡ�ۺ�����
'�����õķ�ʽ����ȡ�۽ӿڣ��õ���˰����TaxCost����˰����UnitCost��˰��TaxRate

'##ModelId=42645FEB019E
Public Function GetInvBuyPrice(login As clsLogin, conn As Object, SysInfo As Object, VendorCode As String, InvCode As String, UnitCost As Variant, Quantity As Variant, TaxCost As Variant, TaxRate As Variant, bTaxCost As Boolean, sXmlx As String) As Boolean
    Dim oAlgorithmManager As Object
    Set oAlgorithmManager = CreateObject("Algorithm.clsAlgorithmManager")
    Dim oCostQueryAlgorithm As Object
    Set oCostQueryAlgorithm = CreateObject("Algorithm.ICostQueryAlgorithm")

    'ʹ�õ������ͣ��ɹ�ѡ����ݿ����ӡ�login�����ʼ������õ�ǰȡ���㷨
    'cardnum ="88" ��ʾ�ɹ�����
    Set oCostQueryAlgorithm = oAlgorithmManager.GetCostQueryAlgorithm("88", SysInfo, conn, login)

    '�ж�ȡ�۷����Ƿ���ã�ĳЩȡ�۷����㷨���봫����Ҫ�Ĳ������ܷ��ض��󣬷��򷵻ؿ�
    If Not oCostQueryAlgorithm Is Nothing Then
        'ʹ�ô�����롢��Ӧ�̱��롢��Ӧ���͡����֡����ڡ��������Լ���ǰ�ı���ֶα�ʶ�����˰���ۡ���˰�����Լ�˰�ʣ�ȡ�۳ɹ�����true
        If oCostQueryAlgorithm.GetCost(InvCode, VendorCode, 1, "�����", login.CurDate, Quantity, UnitCost, TaxCost, TaxRate, bTaxCost, "cinvcode", sXmlx) Then
            'ʹ��ȡ�õ���˰����unitcost����˰����taxcost��˰�� taxrate
            GetInvBuyPrice = True
        End If
    End If
End Function

'�ж���Դ�����Ƿ��޸�
Public Function SourceIsChanged(ufts As String, strTable, strcCode As String) As Boolean
    Dim tmpRst As ADODB.Recordset
    Dim tmpstr As String
    SourceIsChanged = False
    On Error GoTo ErrHandle
    Set tmpRst = New ADODB.Recordset
    tmpstr = "  select convert(char,convert(money,ufts),2) as ufts from " & strTable & _
           "  where cCode='" & strcCode & "'"
    tmpRst.Open tmpstr, g_Conn, adOpenStatic, adLockReadOnly
    If tmpRst.RecordCount > 0 Then
        If ufts <> tmpRst.Fields("ufts") Then
            SourceIsChanged = True
        End If
    Else
        SourceIsChanged = True
    End If
    tmpRst.Close
exitfun:
    If Not tmpRst Is Nothing Then
        Set tmpRst = Nothing
    End If
    Exit Function
ErrHandle:
    GoTo exitfun
End Function

Public Function CheckVoucherStatus(ByVal strID As Long, StrType As String) As String
    On Error GoTo lerr
    Dim i As Integer
    Dim strSql As String
    Dim rstJC As New ADODB.Recordset
    '
    '    strSql = " select distinct  ccode,istatus,iQtyOutSum = case when isnull(iQtyOutSum,0)>0 then 1 else 0 end " & vbCrLf & _
         '             " From " & VoucherList & " where ID = " & strID

    strSql = " select ccode,istatus From " & MainView & " where ID = " & strID
    If rstJC Is Nothing Then Set rstJC = CreateObject("ADODB.Recordset")
    If rstJC.State = adStateOpen Then Call rstJC.Close
    Call rstJC.Open(strSql, g_Conn, adOpenStatic, adLockReadOnly, adCmdText)

    If rstJC.RecordCount > 0 Then
        CheckVoucherStatus = rstJC.Fields("istatus").Value
        If CheckVoucherStatus <> "�ر�" Then
            'enum by modify
            If StrType = "�ڳ�����" Then
                If VoucherIsCreate2(strID) Then CheckVoucherStatus = "����"
            Else
                If VoucherIsCreate(strID) Then CheckVoucherStatus = "����"
            End If
        End If
    Else
        CheckVoucherStatus = ""
    End If

    Set rstJC = Nothing
    'CheckVoucherStatus = True
    Exit Function
lerr:
End Function


'�ж���Դ�����Ƿ������� ByVal strTable As String, ByVal dblQty As Double
Public Function VoucherIsCreate(ByVal lngID As Long) As Boolean
    Dim tmpRst As ADODB.Recordset
    Dim tmpstr As String
    VoucherIsCreate = False
    
    
    On Error GoTo ErrHandle
    Set tmpRst = New ADODB.Recordset
    tmpRst.CursorLocation = adUseClient
    
    
    
    '  tmpstr = "  select iQtyOut from " & strTable & "  where ID=" & lngID & " and " & dblQty & " > 0"
    tmpstr = "select palcode from HY_FYSL_Collections with(nolock) where  palcode in (select ccode from HY_FYSL_Payment with(nolock) where id = " & lngID & ")"
    tmpRst.Open tmpstr, g_Conn, adOpenStatic, adLockReadOnly
    
   
    
    If tmpRst.RecordCount > 0 Then
        VoucherIsCreate = True
 
    End If
    
    
exitfun:
    If Not tmpRst Is Nothing Then
        Set tmpRst = Nothing
    End If
    Exit Function
ErrHandle:
    GoTo exitfun
End Function

'�ж���Դ�����Ƿ������� ByVal strTable As String, ByVal dblQty As Double
Public Function VoucherIsCreate2(ByVal lngID As Long) As Boolean
    Dim tmpRst As ADODB.Recordset
    Dim tmpstr As String
    VoucherIsCreate2 = False
    On Error GoTo ErrHandle
    Set tmpRst = New ADODB.Recordset
    tmpRst.CursorLocation = adUseClient
    
    
    '  tmpstr = "  select iQtyOut from " & strTable & "  where ID=" & lngID & " and " & dblQty & " > 0"
    tmpstr = " select iQtyOut from HY_DZ_BorrowOuts where ID=" & lngID & _
           " and  (iQtyBack > 0 or iQtyBack2 >0 or iQtyCOut>0 or iQtyCOut2>0 " & _
           " or iQtyCSale>0 or iQtyCSale2>0 or iQtyCFree>0 or iQtyCFree2>0 " & _
           " or iQtyCOver>0 or iQtyCOver2>0 )"
    tmpRst.Open tmpstr, g_Conn, adOpenStatic, adLockReadOnly
    If tmpRst.RecordCount > 0 Then
        VoucherIsCreate2 = True
    Else
        tmpstr = "select * from Ap_Vouch where cPluginsourcetype='������õ�'and cPluginsourceautoid='" & lngID & "'"
        tmpRst.Close
        tmpRst.Open tmpstr, g_Conn, adOpenStatic, adLockReadOnly
        If Not tmpRst.BOF Or Not tmpRst.EOF Then
               VoucherIsCreate2 = True
        Else
            tmpstr = " select b.id  from HY_DZ_BorrowOuts a " & _
                    "inner join  HY_DZ_BorrowOutbacks b on a.autoid=b.upautoid " & _
                    " where a.ID=" & lngID
                     tmpRst.Close
        tmpRst.Open tmpstr, g_Conn, adOpenStatic, adLockReadOnly
        If Not tmpRst.BOF Or Not tmpRst.EOF Then
               VoucherIsCreate2 = True
        Else
              VoucherIsCreate2 = False
              End If
        End If
      '  VoucherIsCreate2 = False
    End If
    tmpRst.Close
    If VoucherIsCreate2 = False Then
        tmpstr = "select top 1 iSourceId as id  from salepayvoucht with (nolock) where (actvt_id is null or actvt_Id='') and csourcetype in('1660','1690') and iSourceId =" & lngID
        tmpstr = tmpstr & " union all "
        tmpstr = tmpstr & " select top 1 SourceVoucherID as id  from NE_CostApply with (nolock) where  SourceType  in(1660,1690) and  SourceVoucherID =" & lngID
        tmpstr = tmpstr & " union all "
        tmpstr = tmpstr & " select top 1 sourcecode  as id  from tc_borrowfeedback  with (nolock)  where  sourcecode= " & lngID
        
         tmpRst.Open tmpstr, g_Conn, adOpenStatic, adLockReadOnly
         If tmpRst.RecordCount > 0 Then
            VoucherIsCreate2 = True
         End If
         tmpRst.Close
    End If
exitfun:
    If Not tmpRst Is Nothing Then
        Set tmpRst = Nothing
    End If
    Exit Function
ErrHandle:
    GoTo exitfun
End Function

'�ж���Դ�����Ƿ��ѹ黹���
Public Function VoucherIsAllBack(ByVal lngID As Long) As Boolean
    Dim tmpRst As ADODB.Recordset
    Dim tmpstr As String
    On Error GoTo ErrHandle
    
    Set tmpRst = New ADODB.Recordset
    tmpstr = "select * from V_HY_DZ_BorrowOutsSD where V_HY_DZ_BorrowOutsSD.ID= " & lngID & " and  (iquantityUpSD >0  or case when igrouptype=2 then inumUpSD else 0 end >0) and (1=1) "
    tmpRst.Open tmpstr, g_Conn, adOpenStatic, adLockReadOnly
    If tmpRst.RecordCount > 0 Then
        VoucherIsAllBack = False
    Else
        VoucherIsAllBack = True
    End If
    tmpRst.Close

exitfun:
    If Not tmpRst Is Nothing Then
        Set tmpRst = Nothing
    End If
    Exit Function
ErrHandle:
    VoucherIsAllBack = False
    GoTo exitfun
End Function

'�ж���Դ�����Ƿ�δ����
Public Function VoucherIsOut(ByVal lngID As Long) As Boolean
    Dim tmpRst As ADODB.Recordset
    Dim tmpstr As String
    On Error GoTo ErrHandle
    
    Set tmpRst = New ADODB.Recordset
    tmpstr = "select * from HY_DZ_BorrowOuts where HY_DZ_BorrowOuts.ID= " & lngID & " and  isnull(iQtyOut,0) > 0 "
    tmpRst.Open tmpstr, g_Conn, adOpenStatic, adLockReadOnly
    If tmpRst.RecordCount > 0 Then
        VoucherIsOut = True
    Else
        VoucherIsOut = False
    End If
    tmpRst.Close

exitfun:
    If Not tmpRst Is Nothing Then
        Set tmpRst = Nothing
    End If
    Exit Function
ErrHandle:
    VoucherIsOut = False
    GoTo exitfun
End Function

Public Function CreateGUID(Optional strRemoveChars As String = "", Optional bRemove As Boolean = True) As String
    Dim udtGUID As guid
    Dim strGUID As String
    Dim bytGUID() As Byte
    Dim lngLen As Long
    Dim lngRetVal As Long
    Dim lngPos As Long

    'Initialize
    lngLen = 40
    bytGUID = String(lngLen, 0)

    'Create the GUID
    CoCreateGuid udtGUID

    'Convert the structure into a displayable string
    lngRetVal = StringFromGUID2(udtGUID, VarPtr(bytGUID(0)), lngLen)
    strGUID = bytGUID
    If (Asc(Mid$(strGUID, lngRetVal, 1)) = 0) Then
        lngRetVal = lngRetVal - 1
    End If

    'Trim the trailing characters
    strGUID = Left$(strGUID, lngRetVal)

    'Remove the unwanted characters
    For lngPos = 1 To Len(strRemoveChars)
        strGUID = Replace(strGUID, Mid(strRemoveChars, lngPos, 1), "")
    Next
    If bRemove Then
        strGUID = Replace(strGUID, "-", "")
        strGUID = Replace(strGUID, "{", "")
        strGUID = Replace(strGUID, "}", "")
    End If
    CreateGUID = strGUID
End Function

Public Function DropTable(strTableName As String)
    On Error Resume Next
    g_Conn.Execute "drop table " & strTableName
End Function

Public Function CreateTableName(TblName As String) As String
    Dim a As Object
    Set a = CreateObject("Wscript.Network")
    CreateTableName = TblName & Replace(Replace(a.ComputerName, ".", ""), "-", "")
End Function

Public Function MsgBox(ByVal Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String, Optional HelpFile As String) As VbMsgBoxResult
    MsgBox = VBA.MsgBox(Prompt, Buttons, "U8")
End Function
