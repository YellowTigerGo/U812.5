Attribute VB_Name = "ModPublic"
Option Explicit
Public Declare Function htmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public Const HH_DISPLAY_topic = &H0
Public Const HH_HELP_CONTEXT = &HF
Public strMenuID As String

Public clsSAWeb As EFVoucherMo.clsSystem
 
Public domFilter As DOMDocument
Public VoucherRefAgain As Boolean '�������� �Ƿ�༭�������� ���ر�ͷ��Ϣ
Public VoucherRefDomH As DOMDocument '�������� �Ƿ�༭�������� ���ر�ͷ��Ϣ
Public VoucherRefDomB As DOMDocument '�������� �Ƿ�༭�������� ���ر�����Ϣ
Public refCbustype As String 'ҵ������
'Public Const Msg_Title = "������ҵ���"            'Add by TTH 2001.5.21
Public Const Msg_Title = ""            'Add by TTH 2001.5.21
Public blnOnEdit As Boolean


Public m_bInvAuth As Boolean
Public m_bDepAuth As Boolean
Public m_bVenAuth As Boolean
Public m_bCusAuth As Boolean
Public m_bPerAuth As Boolean
Public m_bUseAuth As Boolean
Public strAuthStrForLogin As String

Public objFilterNew As New UFGeneralFilter.FilterSrv
Public sTmpTableName As String '��λ/�б�λ ��ʱ����

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////
'��������ťͼ��
Public Const IDB_Print = 314                '��ӡ
Public Const IDB_Preview = 312              'Ԥ��
Public Const IDB_Output = 308               '���
Public Const IDB_Copy = 318                 '����
Public Const IDB_Add = 323                  '����
Public Const IDB_Modify = 324               '�޸�
Public Const IDB_Delete = 326               'ɾ��
Public Const IDB_Save = 988                '����
Public Const IDB_Discard = 316              '����
Public Const IDB_Addrecord = 343            '����
Public Const IDB_Deleterecord = 347         'ɾ��
Public Const IDB_AlterPO = 321              '���
Public Const IDB_Close = 353                '�ر�
Public Const IDB_Open = 351                 '��
Public Const IDB_Confirm = 1100             '���
Public Const IDB_Cancelconfirm = 341        '����
Public Const IDB_Payment = 377              '�ָ�
Public Const IDB_Cancelpayment = 377        '����
Public Const IDB_Locate = 309               '��λ
Public Const IDB_First = 1174               '����
Public Const IDB_Previous = 1139            '����
Public Const IDB_Next = 1133                '����
Public Const IDB_Last = 1117                'ĩ��
Public Const IDB_Refresh = 154              'ˢ��
Public Const IDB_Help = 396                 '����
Public Const IDB_Exit = 1118                '�˳�
Public Const IDB_BatchOpen = 394                '����
Public Const IDB_BatchClose = 392              '����
Public Const IDB_BatchVeri = 400                '����
Public Const IDB_BatchUnVeri = 393             '����
Public Const IDB_Filter = 1120             '����
Public Const IDB_Set = 8             '����
Public Const IDB_FilterSet = 991            '����
Public Const IDB_BatchPrint = 395            '����
Public Const IDB_BatchJust = 389            '����
Public Const IDB_SelectAll = 336            'ȫѡ
Public Const IDB_UnSelectAll = 334            'ȫ��
Public Const IDB_Calendar = 1108            '����
Public Const IDB_Calc = 335            '������
Public Const IDB_InValid = 327            '����
Public Const IDB_Switch = 368            '�л�

'��������ť�ؼ���
Public Const sKey_Batchprint = "PrintBatch"         '����
Public Const sKey_Print = "Print"                   '��ӡ
Public Const sKey_Preview = "Preview"               'Ԥ��
Public Const sKey_Output = "Output"                 '���
Public Const sKey_Add = "Add"                       '����
Public Const sKey_Modify = "Modify"                 '�޸�
Public Const sKey_Delete = "Delete"                 'ɾ��
Public Const sKey_Save = "Save"                     '����
Public Const sKey_Discard = "Discard"               '����
Public Const sKey_Addrecord = "AddRecord"           '����
Public Const sKey_Deleterecord = "DeleteRecord"     'ɾ��
Public Const sKey_AlterPO = "AlterPO"               '���
Public Const sKey_Close = "Close"                   '�ر�
Public Const sKey_Open = "Open"                     '��
Public Const sKey_Confirm = "Confirm"               '���
Public Const sKey_Cancelconfirm = "Cancelconfirm"   '����
Public Const sKey_Payment = "Payment"               '�ָ�
Public Const sKey_Cancelpayment = "CancelPayment"   '����
Public Const sKey_Locate = "Locate"                 '��λ
Public Const sKey_LocateSet = "LocateSet"           '��λ����
Public Const sKey_First = "First"                   '����
Public Const sKey_Previous = "Previous"             '����
Public Const sKey_Next = "Next"                     '����
Public Const sKey_Last = "Last"                     'ĩ��
Public Const sKey_Refresh = "Refresh"               'ˢ��
Public Const sKey_Help = "Help"                     '����
Public Const sKey_Exit = "Exit"                     '�˳�
'�����б�����������Դ�Ĺؼ����ַ���
Public Const strKprintbill = "printbill"  '��ӡ����
Public Const strKfilter = "filter"   '����
Public Const strKfind = "find"    '����
Public Const strKsetfield = "setfield"   '������ʾ�ֶ�
Public Const strKsort = "sort"  '����
Public Const strKhelp = "help"    '����
Public Const strKclose = "close"   '�˳�
Public Const strKCard = "card" '����
Public Const strKSelectAll = "SelectAll" 'ȫѡ
Public Const strKUnSelectAll = "UnSelectAll" 'ȫ��

'��������ť��ʾ����
Public Const strBatchprint = "����"
Public Const strBatchOpen = "����"
Public Const strBatchClose = "����"
Public Const strBatchVeri = "����"
Public Const strBatchUnVeri = "����"
Public Const strPrint = "��ӡ"
Public Const strPreview = "Ԥ��"
Public Const strOutput = "���"
Public Const strCopy = "����"
Public Const strAdd = "����"
'Public Const strAdd = "����"
Public Const strModify = "�޸�"
'  860sp������861�޸Ĵ� ע��
'��U861�е��ݸ����ṩ�˱��渽���Ĺ���
Public Const strchenged = "����"

'Public Const strchenged = "���"

Public Const strDelete = "ɾ��"
Public Const strSave = "����"
Public Const strDiscard = "����"
Public Const strAddrecord = "����"
Public Const strDeleterecord = "ɾ��"
Public Const strAlterPO = "���"
Public Const strClose = "�ر�"
Public Const strOpen = "��"
Public Const strConfirm = "���"
Public Const strCancelconfirm = "����"
Public Const strPayment = "�ָ�"
Public Const strCancelpayment = "����"
Public Const strLocate = "��λ"
Public Const strLocateSet = "��λ����"
Public Const strFirst = "����"
Public Const strPrevious = "����"
Public Const strNext = "����"
Public Const strLast = "ĩ��"
Public Const strRefresh = "ˢ��"
Public Const strHelp = "����"
Public Const addvouth = "����"
Public Const tm_Print = "����"

Public Const strFilter = "����"
Public Const strfiltersetting = "����"
Public Const strExit = "�˳�"
Public Const strColumn = "��Ŀ"
Public Const strBatchJust = "����"
Public Const strSelectAll = "ȫѡ"
Public Const strUnSelectAll = "ȫ��"
Public Const strconSelectAll = "��ѡ"
Public Const strmake_sure = "ȷ��"

Public TBLStyle             As TBLType

Public sysInfo As New SystemInfo.cSysInfo
Declare Function apiFindExecutable Lib "shell32.dll" Alias _
    "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory _
    As String, ByVal lpResult As String) As Long
Public gbInvSort As Boolean     '�Ƿ񵥾ݸ��ݴ����������
'////////////////////////////////////////////////////////////////////////////////////////////////////////////
Global gcSales As New clsSales
Public BillPrnVTID As String

Public bLoadmain As Boolean  '�������Ƿ����
Public Const sVersionCode = "V8.600"
Public sUFSetUpPath As String
Public cModeCode As String '����ģ��Ĵ���
Public gs_Version As String
Public blnNotDemo As Boolean   ''�Ƿ���ʾ�汾
Public gbUseOrder As Boolean
Public Const SM_CXVSCROLL = 2
Public Const SM_CXHSCROLL = 21
Public lngClr1 As Long, lngClr2 As Long ''���ݿؼ��Ĳ��ɱ༭��ɫ
Public iSAReferType As Integer  ''ģ����������
Public strMainCaption As String
Public gsConnectString   As String  'Ϊ AX_Case ���õ����ݿ����Ӵ�
Global blnSureOp As Boolean
Global intSureSel As Integer
Global bCancelFlag As Boolean
Global lnID As Long '���ݲ鿴����(frmlook)�õ���ID
Global strVNum As String '���ݲ鿴����(frmlook)�õ��ݺ�
Global SBVID As Long '���ݲ鿴����(frmlook)�õ��ݺ�
Global dlDisVal As Double  '�ۿ۶�
Global dlDisRate As Double  '�ۿ���
Global dlTotal As Double '�����ܱ���
Global iCostRefMode As Integer    ''�۸����ģʽ(0-���ͻ�; 1-�����)
Public frmTmp As Form             ''��ʱ����(����ָ���ͻ��۸�)
Public sFreePriceType As String     '��ǰ���ļ���������Ƽ�
Public intCostRefType As Integer
Public blnCostRefCustomer As Boolean
Global gcAccount As New clsAccount
Global gcReport As New clsReport
Global MyRes As New USSARes.cResUtil
Global ogszzPub As clsPub
Global gsAppPath As String          'Ӧ�ó���·��
Global gsWindowPath As String       'Windows ·��
Global gsComputerName As String     '�û����������
Public blnfrmFarIsShow As Boolean

Public Const WM_SETTINGCHANGE = &H1A
Public Const LOCALE_SSHORTDATE = &H1F
Public Const HWND_BROADCAST = &HFFFF&
Public Const DATE_SHORTDATE = &H1
Public Const DATE_LONGDATE = &H2
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Public Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As String) As Long
Public Type ColProertyType
    ColNo       As Long
    ReferType   As ButtonType
    ColProerty  As EditStyle
    CanEdit     As Boolean
End Type

Public Type RowProertyType
    RowNo As Long
    NodeLevel As Integer
    RowKey As String
    RowFatherKey As String
    RowFatherRow As Long
    RowBrother As Integer
End Type
Public Enum SelectResultType:
    Normal = 0          '����
    Level1 = 1          '��һ��
    Level2 = 2          '�ڶ���
    Level3 = 3          '������
    Level4 = 4          '���ļ�
    Level5 = 5          '���弶
    Endlevel = 6        '����ϸ��
End Enum

Enum TBLType
    TBLText
    TBLPicture
    TBLNormal
End Enum

 
Public intTab() As Integer
Public sMsgText() As String
Public ColProerty() As ColProertyType
Public RowProerty() As RowProertyType
Public m_login As U8Login.clsLogin
'Public m_login As Object
Public DBconn As ADODB.Connection

Public cls_Public As Object
  

Public Enum ToWhere
    ToBack = 0
    ToNext = 1
End Enum
Public Enum VSs
    ����
    �޸�
    ɾ��
End Enum

Global Const gsPassKey      As String = "uf97******"    '��������ַ���
    '����״̬
    Public Enum OptStatus
        iClose = 0          '�ر�
        iOpen = 1           '��
        iView = 2           '�鿴
        iQuery = 3          '��ѯ
        iNew = 4            '����
        iEdit = 5           '�༭
        iDelete = 6         'ɾ��
        iPrint = 7          '��ӡ
        iEvaluate = 8       '����
        iAdd = 9            '¼��ԭʼ
        iWriteZW = 10       '����
        iDepr = 11          '�۾ɼ���
        iMonthEnd = 12      '��ĩ����
        iYearEnd = 13       '��ĩ��ת
        iEditNew = 14       '�༭�¿�Ƭ
        iEditAdd = 15       '�༭ԭʼ��Ƭ
        iEarase = 16        'ע��
        iDeleteEarase = 17  'ɾ����ע����Ƭ
        iRestoreEarase = 18 '�ָ���ע����Ƭ
        iEditZW = 19        '�༭ƾ֤
        iDeleteZW = 20      'ɾ��ƾ֤
        iUnDeleteZW = 21    '�ָ�ɾ��ƾ֤
        iRevZW = 22         '���ֳ���ƾ֤
        iViewZW = 23        '�鿴ƾ֤
        iRelView = 24       '����
    End Enum
      
    Enum WaitType:
        iWait = 0               '�ȴ�
    End Enum

    '�߿�����
    Enum LineStyle
        LineNone = 0
        LineThin = 1
        LineMidBold = 2
        LineDash = 3
        LineDot = 4
        LineBold = 5
        LineDouble = 6
        LineDashDot = 7
    End Enum
Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Public Const SRCCOPY = &HCC0020
Public Const MF_BYPOSITION = &H400&
Public Const MF_BITMAP = &H4&
Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const ERROR_SUCCESS = 0
Public Const REG_SZ = 1                         ' Unicode �� Null ��β���ַ���
Public Const REG_DWORD = 4                      ' 32-λ����
Public Const gREGKEYSYSINFOLOC = "Display\Settings"
Public Const gREGVALSYSINFOLOC = "MSINFO"
Public Const gREGKEYSYSINFO = "Display\Settings"
Public Const gREGVALSYSINFO = "DPILogicalX"
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Const SM_CXSCREEN = 0

Public Const SM_CYSCREEN = 1
Public Const HELP_CONTEXT = &H1
Public BillPrnSet As Integer
Public VouchPrnFD As Integer
Public BillInvName As String
Public strErrMsg As String
Public g_bLoginSuccess As Boolean
Public g_FormbillShow As Boolean '�Ƿ�ɹ���ʾ�˵��ݴ���
Public bVerifyMsg As Boolean     '������������Ϣ�ж�

Public Function ObjectExists(pColl As Object, sMemName As String) As Boolean
    Dim pObj As Object
  
    On Error Resume Next
    Err = 0
    Set pObj = pColl(sMemName)
    ObjectExists = (Err = 0)
End Function


Public Sub ShowError(strErrDesc As String)
    If strErrDesc <> "" Then MsgBox strErrDesc, vbCritical, App.Title
End Sub
   
'==================================================================
'�������ƣ�CodeToName
'���ã��ɱ���������
'������
'    Code:
'        ����
'    Style:
         
'����ֵ�����ơ������벻�����򷵻ؿ�ֵ��
'==================================================================

Public Function CodeToName(Code As String, Style As String) As String
    On Error GoTo ErrorHandle
    
    Select Case Style
    Case "INVENTORY"
        CodeToName = FromTo("cInvCode", Code, "cInvName", "Inventory")
    Case "INVENTORYC"
        CodeToName = FromTo("cInvCode", Code, "cInvCCode", "Inventory")
    Case "INVENTORYstd"
        CodeToName = FromTo("cInvCode", Code, "cInvStd", "Inventory")
    Case "INVENTORYAdd"
        CodeToName = FromTo("cInvCode", Code, "cInvAddCode", "Inventory")
    Case "VENDOR"
        CodeToName = FromTo("cVenCode", Code, "cVenAbbName", "Vendor")
    Case "CUSTOMER"
        CodeToName = FromTo("cCusCode", Code, "cCusAbbName", "Customer")
    Case "WAREHOUSE"
        CodeToName = FromTo("cWhCode", Code, "cWhName", "Warehouse")
    Case "RDSTYLE"
        CodeToName = FromTo("cRdCode", Code, "cRdName", "Rd_Style")
    Case "PERSON"
        CodeToName = FromTo("cPersonCode", Code, "cPersonName", "Person")
    Case "SETTLESTYLE"
        CodeToName = FromTo("cSSCode", Code, "cSSName", "SettleStyle")
    Case "INVENTORYCLASS"
        CodeToName = FromTo("cInvCCode", Code, "cInvCName", "InventoryClass")
    Case "VENDORCLASS"
        CodeToName = FromTo("cSCCode", Code, "cSCName", "VendorClass")
    Case "CUSTOMERCLASS"
        CodeToName = FromTo("cCCCode", Code, "cCCName", "CustomerClass")
    Case "BANK"
        CodeToName = FromTo("cBCode", Code, "cBName", "Bank")
    Case "FOREIGNCURRENCY"
        CodeToName = FromTo("cExchCode", Code, "cExchName", "ForeignCurrency")
    Case "DEPARTMENT"
        CodeToName = FromTo("cDepCode", Code, "cDepName", "Department")
    Case "EXPENSEITEM"
        CodeToName = FromTo("cExpCode", Code, "cExpName", "ExpenseItem")
    Case "ECOCLASS"
        CodeToName = FromTo("cEcoCode", Code, "cEcoName", "EcoClass")
    Case "PURCHASETYPE"
        CodeToName = FromTo("cPTCode", Code, "cPTName", "PurchaseType")
    Case "CUSTOMERALL"
        CodeToName = FromTo("cCusCode", Code, "cCusAbbName", "Customer")
    Case "SALETYPE"
        CodeToName = FromTo("cSTCode", Code, "cSTName", "SaleType")
    Case "SHIPPINGCHOICE"
        CodeToName = FromTo("cSCCode", Code, "cSCName", "ShippingChoice")
    Case "PAYCONDITION"
        CodeToName = FromTo("cPayCode", Code, "cPayName", "PayCondition")
    Case Else
        CodeToName = ""
    End Select
    Exit Function
ErrorHandle:
    CodeToName = ""
End Function

'======================================================================
'��    �̣�Private Function FromTo
'Ŀ    �ģ�����һ�������Լ��ת��
'��    ����strFrom    Դ�ֶ�
'          strFromValue  ����Դ�ֶε�ֵ
'          strTo      Ŀ���ֶ�
'          strTable   ����
'��    �����ɹ�ת�����Ŀ���ֶε�ֵ���������ΪNull
'======================================================================
'
Public Function FromTo(strFrom As String, strFromValue As String, strTo As String, strTable As String) As Variant
    Dim strsql As String
    Dim rec As New ADODB.Recordset
    strsql = "select " & strTo & " as name from " & strTable & " where " & _
              strFrom & "='" & strFromValue & "'"
    
    rec.Open strsql, DBconn, adOpenForwardOnly, adLockReadOnly
    If rec.EOF = True And rec.BOF = True Then
        FromTo = Null
    Else
        FromTo = Trim(rec!Name)
    End If
    rec.Close
End Function

Public Function trimLastLetter(oldstr As String) As String
     Dim StrLen As Long
     StrLen = Len(oldstr)
     If StrLen = 0 Then Exit Function
     trimLastLetter = Mid(oldstr, 1, StrLen - 1)
End Function


 

Public Sub Main()

    If Not HaveSufficeResources() Then Exit Sub
    
    Dim i As Long, lDiskSpace As Long
    Dim CmdLine As String
    Dim cmdlnlen As Long
    Dim bSuccess As Boolean
    Dim bQD As Boolean
    Dim strData As String
    Dim LngLCID As Long
    Dim strStartDate As String
    On Error GoTo ErrMain
'    SAVer = "SA"
    bQD = False
    LngLCID = GetSystemDefaultLCID
    strData = String(255, vbNullChar)
    GetLocaleInfo LngLCID, LOCALE_SSHORTDATE, strData, 255
    strData = Left(strData, InStr(1, strData, Chr(0)) - 1)    '������������ø�ʽ
    
    If LCase(Trim(strData)) <> "yyyy-mm-dd" Then
        SetLocaleInfo LngLCID, LOCALE_SSHORTDATE, "yyyy-MM-dd"    '��д���������ø�ʽ
        SendMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0&, 0&      '������Ϣ������
    End If
    
    gsConnectString = m_login.UfDbName
    If DBconn Is Nothing Then
        Set DBconn = New ADODB.Connection
        DBconn.ConnectionTimeout = 600
        DBconn.CommandTimeout = 1200
    End If
    If DBconn.State = 1 Then DBconn.Close
    DBconn.Open m_login.UfDbName
LB:
    bSuccess = True
    Set clsSAWeb = New EFVoucherMo.clsSystem
    Dim tmpStr As String
    If clsSAWeb.Init(m_login, tmpStr) = False Then
        MsgBox "��ʼ����̨����������" & tmpStr
    End If
    clsSAWeb.cModeCode = cModeCode
    
    Screen.MousePointer = vbHourglass
    gsComputerName = sysInfo.ComputerName
    bQD = True
    gsWindowPath = GetWindowsDirectoryFromSystem
    Load frmMain
    g_bLoginSuccess = True
    bLoadmain = True
    Exit Sub
ErrMain:
    If bQD Then

        Screen.MousePointer = vbDefault
    End If
    MsgBox Err.Description
    If m_login.ShareString <> "" Then
        MsgBox m_login.ShareString
    End If

    
End Sub
Public Function IsSaleStart() As Boolean
    
End Function

Public Sub CenterByForm(ctlOcx As Control, frmForm As Form)
'---------------------------------------------------------------------------------------
'�������ƣ�CenterByForm
'�������ܣ�ʹ�ؼ��ڴ������
'---------------------------------------------------------------------------------------
'����˵����
'---------------------------------------------------------------------------------------
'������
'   �������ߣ�
'   ����ʱ�䣺
'   �������ͣ�
'   �� �� ֵ��
'   �������ͣ�
'   �� �� �ˣ�
'   �޸�ʱ�䣺
'   �޸�ԭ��
'---------------------------------------------------------------------------------------
    ctlOcx.Left = (frmForm.Width - ctlOcx.Width) \ 2
End Sub

Public Sub CenterByCOM(ctlOcx1 As Control, ctlOcx2 As Control)
'---------------------------------------------------------------------------------------
'�������ƣ�CenterByCOM
'�������ܣ�ʹ�ؼ��������ؼ�����
'---------------------------------------------------------------------------------------
'����˵����
'---------------------------------------------------------------------------------------
'������
'   �������ߣ�
'   ����ʱ�䣺
'   �������ͣ�
'   �� �� ֵ��
'   �������ͣ�
'   �� �� �ˣ�
'   �޸�ʱ�䣺
'   �޸�ԭ��
'---------------------------------------------------------------------------------------
    ctlOcx1.Left = (ctlOcx2.Width - ctlOcx1.Width) \ 2
    ctlOcx1.Top = (ctlOcx2.Height - ctlOcx1.Height) \ 2
End Sub


Public Function GetSplitCharNum(s As String, sChar As String) As Integer
'---------------------------------------------------------------------------------------
'�������ƣ�GetSplitCharNum
'�������ܣ��õ��ָ������
'---------------------------------------------------------------------------------------
    Dim i As Integer
    
    GetSplitCharNum = 0
    If InStr(1, s, sChar) = 0 Then Exit Function
    
    For i = 1 To Len(s)
        If Mid(s, i, 1) = sChar Then
            GetSplitCharNum = GetSplitCharNum + 1
        End If
    Next i
End Function


'*************************************************************
'�������ƣ�AddGraph
'���̹��ܣ���ImageList�ؼ������λͼ��ͼ�ꡢ���
'���������img:�ؼ�����
'          iNum:ͼ��˳���
'          cKey:ͼ��ؼ�������
'          iResNum:��Դ��
'          Format:���ݸ�ʽ. 0:λͼ��Դ,1:ͼ����Դ,2:�����Դ
'*************************************************************
Private Sub AddGraph(Img As ImageList, iNum As Integer, cKey As String, iResNum As Integer, iFormat As PicType)
'    img.ListImages.Add iNum, cKey,  LoadResPicture(iResNum, iFormat)
    Img.ListImages.Add iNum, cKey, gcSales.GetResPicture(iResNum, iFormat)
End Sub

Private Sub AddGraphNew(Img As MSComctlLib.ImageList, iNum As Integer, cKey As String, iResNum As Integer, iFormat As Integer)
    Img.ListImages.Add iNum, cKey, MyRes.LoadResPic(iResNum, iFormat)
End Sub


Public Function CreateTempTable(Optional sPreFix As String) As String
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
    sTempName = NewTrim(gsComputerName) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now)
    sRnd = Int((10000 * Rnd) + 1000)
    sTempName = NewTrim(sTempName & sRnd)
    sPreFix = "tempdb.." & sPreFix & sTempName
    CreateTempTable = RemoveSpecialChar(sPreFix)
    
End Function

Private Function RemoveSpecialChar(ByVal strName As String) As String
    'ȥ�������ַ�
    Dim i As Long
    
    RemoveSpecialChar = vbNullString
    For i = 1 To Len(strName)
        
        Select Case Mid(strName, i, 1)
        Case "-", "%", "#", "'", ";", "`", "|", "\", "/", "+", "*", ",", """", " "
        
        Case Else
            RemoveSpecialChar = RemoveSpecialChar & Mid(strName, i, 1)
        End Select
        
    Next i
    
End Function

Public Function Replace(sExistString As String, sFindString As String, sReplaceString As String, Optional bReplaceAll As Boolean = True) As String

'---------------------------------------------------------------------------------------
'�������ƣ�Replace
'�������ܣ��ַ����滻����
'---------------------------------------------------------------------------------------
'����˵����
'   sExiststring��     ԭ�ַ�����
'   sFindString��      Ҫ�滻���ַ�����
'   sReplaceString��   �滻����ַ�����
'   bReplaceAll��      �Ƿ�ȫ���滻��
'---------------------------------------------------------------------------------------

    Const DEFREPLACESTR = "$$$"
    Dim i As Integer
    Dim iLocate As Integer
    Dim iLength As String
    Dim s1 As String, s2 As String
    
    Replace = sExistString
    i = InStr(1, Replace, sFindString)
    If i > 0 Then
        Do
            Replace = Left(Replace, i - 1) & DEFREPLACESTR & Mid(Replace, i + Len(sFindString), Len(Replace))
            If bReplaceAll Then
                i = InStr(1, Replace, sFindString)
            Else
                Exit Do
            End If
        Loop Until i = 0
    End If
    iLocate = InStr(1, Replace, DEFREPLACESTR)
    If iLocate = 0 Then Exit Function
    If bReplaceAll = True Then
        Do
            s1 = "": s2 = ""
            s1 = Left(Replace, iLocate - 1): s2 = Mid(Replace, iLocate + Len(DEFREPLACESTR), Len(Replace))
            Replace = s1 & sReplaceString & s2
            iLocate = InStr(1, Replace, DEFREPLACESTR)
        Loop Until iLocate = 0
    Else
        s1 = "": s2 = ""
        s1 = Left(Replace, iLocate - 1): s2 = Mid(Replace, iLocate + Len(DEFREPLACESTR), Len(Replace))
        Replace = s1 & sReplaceString & s2
    End If

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


Public Function LockItem(PowerID As String, bLock As Integer, Optional bMsg As Boolean = True) As Boolean
    Dim s As String
    Dim lDiskSpace  As Long
    
    On Error Resume Next
    If PowerID = "" Then
        LockItem = True
        Exit Function
    End If
    s = "T"
    If s = "T" Then
        m_login.ClearError
        Select Case CBool(bLock)
            Case True
                LockItem = m_login.TaskExec(PowerID, True)
                If Not LockItem And bMsg Then
                    MsgBox IIf(Trim(m_login.ShareString) = "", "��������ʧ��,������", m_login.ShareString), vbCritical
                End If
            Case False
                LockItem = m_login.TaskExec(PowerID, False)
                If Not LockItem Then
                  m_login.ClearError
                  LockItem = m_login.TaskExec(PowerID, False)
                End If
            Case Else
                LockItem = m_login.TaskExec(PowerID, 1)
        End Select
    Else
        LockItem = True
    End If
    

End Function


Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' ѭ��ָ��
    Dim rc As Long                                          ' ���ش���
    Dim hKey As Long                                        ' �򿪵�ע����ľ��
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' ע�������������
    Dim tmpVal As String                                    ' ע�������ʱ�洢��
    Dim KeyValSize As Long                                  ' ע��������Ĵ�С
    '------------------------------------------------------------
    ' �ڸ��� {HKEY_LOCAL_MACHINE...} �´�ע���
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' ��ע���
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' �������...
    
    tmpVal = String$(1024, 0)                             ' ��������ռ�
    KeyValSize = 1024                                       ' ��Ǳ�����С
    
    '------------------------------------------------------------
    ' ����ע���ֵ...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' ���/������ֵ
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' �������
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 ����� Null ��β���ַ���...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null �ҵ������ַ�����ȡ
    Else                                                    ' WinNT ����Ҫ�� Null �����ַ���...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null δ�ҵ��� ����ȡ�ַ���
    End If
    '------------------------------------------------------------
    ' Ϊ��ת����������ֵ����..
    '------------------------------------------------------------
    Select Case KeyValType                                  ' ������������...
    Case REG_SZ                                             ' �ַ�����ע�����������
        KeyVal = tmpVal                                     ' �����ַ���ֵ
    Case REG_DWORD                                          ' ˫����ע�����������
        For i = Len(tmpVal) To 1 Step -1                    ' ת��ÿһλ
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' һ���ַ�һ���ַ��ؽ���ֵ
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' ת��˫����Ϊ�ַ�����
    End Select
    
    GetKeyValue = True                                      ' ���سɹ�
    rc = RegCloseKey(hKey)                                  ' �ر�ע���
    Exit Function                                           ' �˳�
    
GetKeyError:      ' ������������...
    KeyVal = ""                                             ' ���÷���ֵΪ���ַ���
    GetKeyValue = False                                     ' ����ʧ��
    rc = RegCloseKey(hKey)                                  ' �ر�ע���
End Function

Public Function GetCurrActiveForms() As Boolean
    GetCurrActiveForms = True
End Function

Public Sub ResizeBar(Frm As Form, StBar As Object, WhichIsMax As Integer)
    Dim sngDefSize As Single, sngSize As Single
    
    Dim i As Integer
    On Error Resume Next
    sngDefSize = Frm.Width / (StBar.Panels.Count * 2)
    sngSize = 0
    If WhichIsMax = 1 Then
        For i = WhichIsMax + 1 To StBar.Panels.Count
            StBar.Panels(i).Width = IIf(Trim(StBar.Panels(i).Text) = "", sngDefSize, Frm.TextWidth("WW") * Len(StBar.Panels(i).Text) * 1.5)
            If StBar.Panels(i).Width < sngSize Then
                StBar.Panels(i).Width = sngSize
            End If
            sngSize = sngSize + StBar.Panels(i).Width
        Next i
        StBar.Panels(WhichIsMax).Width = Frm.Width - sngSize
    Else
'        Frm.s
        For i = 1 To WhichIsMax - 1
            StBar.Panels(i).Width = IIf(Trim(StBar.Panels(i).Text) = "", sngDefSize, Frm.TextWidth("WW") * Len(StBar.Panels(i).Text) * 1.5)
            If StBar.Panels(i).Width < sngSize Then
                StBar.Panels(i).Width = sngSize
            End If
            sngSize = sngSize + StBar.Panels(i).Width
        Next i
        For i = WhichIsMax + 1 To StBar.Panels.Count
            StBar.Panels(i).Width = IIf(Trim(StBar.Panels(i).Text) = "", sngDefSize, Frm.TextWidth("WW") * Len(StBar.Panels(i).Text) * 1.5)
            If StBar.Panels(i).Width < sngSize Then
                StBar.Panels(i).Width = sngSize
            End If
            sngSize = sngSize + StBar.Panels(i).Width
        Next i
        StBar.Panels(WhichIsMax).Width = Frm.Width - sngSize
    End If
End Sub

Public Function GetLoginErrStr(ByVal nErrNo As Integer) As String
    GetLoginErrStr = m_login.ShareString
End Function



Public Function SumtoChinessEX(cSum As String, iSection As Long) As String

  Dim l As Long
  Dim C As String
  
  On Error Resume Next
  
  SumtoChinessEX = ""
  l = Len(Trim(cSum)) + 1
  If iSection >= l Then
    SumtoChinessEX = "��"
    Exit Function
  End If
  C = Mid(cSum, l - iSection, 1)
  If C = "0" Then
    SumtoChinessEX = "��"
  Else
    SumtoChinessEX = gcSales.Number2Chinese(C)
  End If
  
End Function

Public Sub FormatCellPrice(Grid As Object, ColNum As Long, RowNum As Long)
     
    Grid.col = ColNum
    Grid.row = RowNum
    Grid.Text = Format(Grid.Text, "###,###,###,###,##0.00")
End Sub

Public Function RetLastDate(dNowDate As Date, iDays As Integer) As Date
'-----------------------------------------------------------------------------------------------------------------------
'���ĳ���ڼ� n ��ǰ������
'-----------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    
    i = Month(dNowDate)
    Select Case i
        Case 1, 3, 5, 7, 8, 10, 12
            RetLastDate = Year(dNowDate) & "-" & i & "-31"
        Case 4, 6, 9, 11
            RetLastDate = Year(dNowDate) & "-" & i & "-30"
        Case 2
            RetLastDate = Year(dNowDate) & "-" & i & "-28"
    End Select
End Function



Public Sub ShowHelpConText(Frm As Form, ByVal ConTextID As Long)
   Screen.MousePointer = 11
   htmlHelp Frm.hwnd, App.HelpFile, HH_DISPLAY_topic, 0
'   WinHelp frm.hwnd, IIf(Right(App.Path, 1) <> "\", App.Path & "\", App.Path) & "\Sales.hlp", HELP_CONTEXT, ConTextID
   Screen.MousePointer = 1
End Sub

'ȡ�ò���ֵ
Public Function getAccinformation(strSysID As String, strName As String, Optional cID As String = "") As String
    Dim Rst As New ADODB.Recordset
    Dim strsql As String
    If cID = "" Then
        strsql = "Select cValue from accinformation where cSysID='" & strSysID & "' and cName='" & strName & "'"
    Else
        strsql = "select cvalue from accinformation where cSysid='" & strSysID & "' and cID='" & cID & "' and cname='" & strName & "'"
    End If
    Rst.Open strsql, DBconn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Rst.EOF Then
        getAccinformation = ""
    Else
        If IsNull(Rst(0)) Then
            getAccinformation = ""
        Else
            getAccinformation = Rst(0)
        End If
    End If
    Rst.Close
    Set Rst = Nothing
End Function

' ���õ�Ԫ��Ķ��뷽ʽ
'  add:  2001.11.5
Public Sub SetAlignment(Grid As Object, row As Long, col As Long, AlignMode As Integer)
    
    With Grid
         .row = row
         .col = col
         .CellAlignment = AlignMode
    End With
    
  End Sub


' ��ʽ������(����ĵ�Ԫ��)
''�޸�:2000-05-10  ֱ��ʹ�� TextMatirx ����
Public Function FormatCellUPrice(Grid As Object, RowNum As Long, ColNum As Long, iPointNum As Integer) As String
  
  With Grid
       If iPointNum <> 0 Then
          FormatCellUPrice = Format(.TextMatrix(RowNum, ColNum), "###,###,###,###,##0." + String(iPointNum, "0"))
       Else
          FormatCellUPrice = Format(.TextMatrix(RowNum, ColNum), "###,###,###,###,##0")
       End If
  End With

End Function

'' ���ָ���Ĵ����Ƿ��Ѿ�������
'' ���ݴ������ƻ���Tag ���ָ���Ĵ����Ƿ����

Public Function CheckFrmExist(frmName As String, frmHwnd As Long, iFormIndex As Integer, Optional frmTag As String = "") As Boolean
 Dim frmChk As Form
 Dim frmIndex As Integer
 
  CheckFrmExist = False
  If frmTag = "" Then   ''�������
     For Each frmChk In Forms
         If LCase(Trim(frmChk.Name)) = LCase(Trim(frmName)) Then
            CheckFrmExist = True
            frmHwnd = frmChk.hwnd
            If frmChk.WindowState = vbMinimized Then
                frmChk.WindowState = 2
            End If
            Exit For
         End If
     Next
  Else  ''���Tag
'    For Each frmChk In Forms
'        If LCase(Trim(frmChk.Tag)) = LCase(Trim(frmTag)) Then
'           CheckFrmExist = True
'           frmHwnd = frmChk.hwnd
'           Exit For
'        End If
'    Next
     
     For frmIndex = Forms.Count To 1 Step -1
         If UCase(Forms(frmIndex - 1).Tag) = UCase(frmTag) Then
            CheckFrmExist = True
            iFormIndex = frmIndex - 1
            If Forms(frmIndex - 1).WindowState = vbMinimized Then
                Forms(frmIndex - 1).WindowState = 2
            End If
            Exit For
          End If
     Next
  End If
  
End Function


'' ͨ�õ�ADO��������
'' 2001.03
Public Function DealError(cn As ADODB.Connection) As String
 Dim strErrMsg As String
 Dim objErr As ADODB.Error
 
     For Each objErr In cn.Errors
         strErrMsg = strErrMsg & " Error #" & objErr.Number & vbCr & objErr.Description & vbCr _
                   & " (Source: " & objErr.Source & ")" & vbCr _
                   & " (SQL State: " & objErr.SQLState & ")" & vbCr _
                   & " (Native Error:" & objErr.NativeError & ")" & vbCr
          
         If objErr.HelpFile = "" Then
            strErrMsg = strErrMsg & "   No Help file available" & vbCr & vbCr
         Else
            strErrMsg = strErrMsg & "  (HelpFile: " & objErr.HelpFile & ")" & vbCr _
                                & "  (HelpContext: " & objErr.HelpContext & ")" & vbCr & vbCr
         End If
         
         ''If objErr.NativeError = 2601 Then Exit For  ''Υ��һ����
            
     Next
     
     If strErrMsg <> "" Then DealError = strErrMsg
     
End Function
''��ȡDOMHead
''
Public Function GetHeadItemValue(ByVal domHead As DOMDocument, ByVal Skey As String) As String
    Skey = LCase(Skey)
    If Not domHead.selectSingleNode("//z:row") Is Nothing Then
        If Not domHead.selectSingleNode("//z:row").Attributes.getNamedItem(Skey) Is Nothing Then
            GetHeadItemValue = domHead.selectSingleNode("//z:row").Attributes.getNamedItem(Skey).nodeValue
        End If
    Else
        GetHeadItemValue = ""
    End If
End Function
Public Function GetBodyItemValue(ByVal domBody As DOMDocument, ByVal Skey As String, ByVal R As Long) As String
    Skey = LCase(Skey)
    If Not domBody.selectNodes("//z:row").Item(R).Attributes.getNamedItem(Skey) Is Nothing Then
        GetBodyItemValue = domBody.selectNodes("//z:row").Item(R).Attributes.getNamedItem(Skey).nodeValue
    Else
        GetBodyItemValue = ""
    End If
End Function

'Public Function ShowVouch(VouchList As VouchList, ByVal skey As String, ByVal cols As U8colset.clsCols, Optional sText As String)
'    Dim Frm As New frmVouchNew
'    Dim sType As String
'    Dim sNegative As String
'    'Dim sText As String
'        Select Case skey
'
'            Case "MT003"
'                sText = VouchList.TextMatrix(VouchList.row, cols("id").iColPos)
'                Call Frm.ShowVoucher(MT03, sText)
'
'            Case "MT004"
'                sText = VouchList.TextMatrix(VouchList.row, cols("id").iColPos)
'                Call Frm.ShowVoucher(MT04, sText)
'
'            Case "MT005"
'                sText = VouchList.TextMatrix(VouchList.row, cols("id").iColPos)
'                Call Frm.ShowVoucher(MT05, sText)
'
'
'            Case "MT006"
'                sText = VouchList.TextMatrix(VouchList.row, cols("id").iColPos)
'                Call Frm.ShowVoucher(MT06, sText)
'
'            Case "MT007"
'                sText = VouchList.TextMatrix(VouchList.row, cols("id").iColPos)
'                Call Frm.ShowVoucher(MT07, sText)
'
'            Case "MT008"
'                sText = VouchList.TextMatrix(VouchList.row, cols("id").iColPos)
'                Call Frm.ShowVoucher(MT08, sText)
'
'            Case "MT009"
'                sText = VouchList.TextMatrix(VouchList.row, cols("id").iColPos)
'                Call Frm.ShowVoucher(MT09, sText)
'
'            Case "MT001"
'                sText = VouchList.TextMatrix(VouchList.row, cols("id").iColPos)
'                Call Frm.ShowVoucher(MT01, sText)
'
'
'
'        End Select
'        If Frm.Visible = False Then
'            Unload Frm
'        Else
'            Frm.ZOrder
'        End If
'End Function

''���ݲ���ģ��ƥ�����÷��ز�ѯ��
Public Function GetLikeString(strFieldName As String, strFieldValue As String) As String
    
    If strFieldValue <> "" Then
    Select Case iSAReferType    ''ģ����������
        Case 0      '0  ���ڻ�����ȷƥ��
            GetLikeString = strFieldName & " ='" & strFieldValue & "'"
        Case 1      '1  ���ڻ������ģ��ƥ��
            GetLikeString = strFieldName & " like '" & strFieldValue & "%'"
        Case 2      '2  ���ڻ�����ǰģ��ƥ��
            GetLikeString = strFieldName & " like '%" & strFieldValue & "'"
        Case 3      '3  ���ڻ���ǰ��ģ��ƥ��
            GetLikeString = strFieldName & " like '%" & strFieldValue & "%'"
        Case 4      '4  ��ѯȫ������ģ��ƥ��
            GetLikeString = " 1=1 "
    End Select
    Else
        GetLikeString = " 1=1 "
    End If
End Function
''���ز��չ�������
Public Function getReferString(strFieldName As String, strFieldValue As String) As String
    
    If strFieldValue = "" Then
        getReferString = ""
        Exit Function
    End If
    
    Select Case LCase(strFieldName)
        Case "coppcode"
            getReferString = GetLikeString(strFieldName, strFieldValue)
        Case "csscode"   ''���㷽ʽ
            getReferString = GetLikeString(strFieldName, strFieldValue) + " or  " + GetLikeString("cSSName", strFieldValue)
        Case "imassdate"       '������
        Case "iquotedprice"       '����
        Case "cmemo"       '��    ע
        Case "cbname"      ''��������
            getReferString = GetLikeString(strFieldName, strFieldValue) + " or  " + GetLikeString("cbaccount", strFieldValue) + " or  " + GetLikeString("cbcode", strFieldValue)
        Case "cexch_name"       '��    ��
            getReferString = GetLikeString(strFieldName, strFieldValue) + " or  " + GetLikeString("cexch_code", strFieldValue)
        Case "cdefine11"       '��ͷ�Զ���11
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cdefine12"       '��ͷ�Զ���12
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cdefine13"       '��ͷ�Զ���13
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cdefine14"       '��ͷ�Զ���14
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cdefine15"       '��ͷ�Զ���15
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cdefine16"       '��ͷ�Զ���16
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cwhname"       '�ֿ�
            getReferString = GetLikeString("cwhname", strFieldValue) + " or  " + GetLikeString("cwhcode", strFieldValue)
        Case "cscname"       '���˷�ʽ
            getReferString = GetLikeString("cscname", strFieldValue) + " or  " + GetLikeString("csccode", strFieldValue)
        Case "cpayname"       '��������
            getReferString = GetLikeString("cpayname", strFieldValue) + " or  " + GetLikeString("cpaycode", strFieldValue)
        'Case "cinvcode"       '������
            
        Case "cinvname"       '��������
            getReferString = GetLikeString("cinvname", strFieldValue) + " or  " + GetLikeString("cinvcode", strFieldValue) + " or " + GetLikeString("cInvAddCode", strFieldValue)
        Case "ccusabbname"       '�ͻ�����
            getReferString = GetLikeString("ccusabbname", strFieldValue) + " or  " + GetLikeString("ccuscode", strFieldValue) + " or  " + GetLikeString("ccusname", strFieldValue)
        Case "ccusdefine1"       '�ͻ��Զ���1
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "ccusdefine10"       '�ͻ��Զ���10
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "ccusdefine11"       '�ͻ��Զ���11
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "ccusdefine12"       '�ͻ��Զ���12
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "ccusdefine13"       '�ͻ��Զ���13
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "ccusdefine14"       '�ͻ��Զ���14
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "ccusdefine15"       '�ͻ��Զ���15
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "ccusdefine16"       '�ͻ��Զ���16
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "ccusdefine2"       '�ͻ��Զ���2
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "ccusdefine3"       '�ͻ��Զ���3
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "ccusdefine4"       '�ͻ��Զ���4
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "ccusdefine5"       '�ͻ��Զ���5
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "ccusdefine6"       '�ͻ��Զ���6
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "ccusdefine7"       '�ͻ��Զ���7
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "ccusdefine8"       '�ͻ��Զ���8
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "ccusdefine9"       '�ͻ��Զ���9
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cbatch"       '����
        Case "inufts"       '��ⵥʱ���
        Case "bgsp"       '�Ƿ����
        Case "citemcode"       '��Ŀ����
            getReferString = GetLikeString("citemcode", strFieldValue) + " or  " + GetLikeString("citemname", strFieldValue)
        Case "citem_class"       '��Ŀ�������
            getReferString = GetLikeString("citem_class", strFieldValue) + " or  " + GetLikeString("citem_name", strFieldValue)
        Case "cdepname"       '�ʲ�����
            getReferString = GetLikeString("cdepname", strFieldValue) + " or  " + GetLikeString("cdepcode", strFieldValue)
        Case "cstname"       '�ʲ�����
            getReferString = GetLikeString("cstname", strFieldValue) + " or  " + GetLikeString("cstcode", strFieldValue)
        Case "cpersonname"       'ҵ �� Ա
            getReferString = GetLikeString("cpersonname", strFieldValue) + " or  " + GetLikeString("cpersoncode", strFieldValue)
        Case "cinvm_unit"       '��������λ
        Case "cdefine22"       '�Զ�����1
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cdefine1"       '�Զ�����1
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cdefine10"       '�Զ�����10
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cdefine23"       '�Զ�����2
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cdefine2"       '�Զ�����2
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cdefine24"       '�Զ�����3
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cdefine3"       '�Զ�����3
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cdefine25"       '�Զ�����4
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cdefine8"       '�Զ�����8
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cdefine9"       '�Զ�����9
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cfree1"       '������1
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cfree10"       '������10
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cfree2"       '������2
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cfree3"       '������3
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cfree4"       '������4
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cfree5"       '������5
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cfree6"       '������6
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cfree7"       '������7
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cfree8"       '������8
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "cfree9"       '������9
            getReferString = GetLikeString("cvalue", strFieldValue) + " or  " + GetLikeString("cAlias", strFieldValue)
        Case "ccusinvcode", "ccusinvname"
            getReferString = GetLikeString("ccusinvcode", strFieldValue) + " or  " + GetLikeString("ccusinvname", strFieldValue)
        Case Else
            getReferString = GetLikeString(strFieldName, strFieldValue)
    End Select
     If getReferString <> "" Then
         getReferString = "( " + getReferString + ")"
     
     End If
End Function



'�����˺��� zcy
Public Sub BatchOperation(ByVal VouchList As Variant, ByVal soperation As String, ByVal Skey As String, ByVal cols As U8colset.clsCols)

End Sub
Public Function CardNumberToVouchType(strCardNumber As String) As String
    Select Case strCardNumber
        Case "24"
            CardNumberToVouchType = "01"        '�ɹ���ⵥ
 
        Case Else
            CardNumberToVouchType = ""
    End Select
End Function
Public Function GetTableFromCardNum(strCardNumber As String, bHeader As Boolean) As String
    Dim rstTmp As ADODB.Recordset
    
    If strCardNumber = "" Then
        GetTableFromCardNum = ""
    Else
        Set rstTmp = New ADODB.Recordset
        rstTmp.Open "select isnull(bttblname,'') as bttblname,isnull(bwtblname,'') as bwtblname from vouchers where cardnumber='" & strCardNumber & "'", DBconn, adOpenForwardOnly, adLockReadOnly
        If Not rstTmp.EOF Then
            
            GetTableFromCardNum = IIf(bHeader, rstTmp(0), IIf(rstTmp(1) <> "", rstTmp(1), rstTmp(0)))
        End If
        rstTmp.Close
        Set rstTmp = Nothing
    End If
End Function
Public Function GetNodeAtrVal(IXNOde As IXMLDOMNode, Skey As String) As String
    Skey = LCase(Skey)
    If IXNOde.Attributes.getNamedItem(Skey) Is Nothing Then
        GetNodeAtrVal = ""
    Else
        GetNodeAtrVal = IXNOde.Attributes.getNamedItem(Skey).nodeValue
    End If
End Function


''����ַ�������
Public Function GetStrTrueLenth(sUnicode As String) As Integer
    Dim strAnsi As String
    strAnsi = StrConv(sUnicode, vbFromUnicode)
    GetStrTrueLenth = LenB(strAnsi)
End Function

Public Function MsgBox(ByVal sPrompt As String, Optional ByVal enumButtons As VbMsgBoxStyle = vbOKOnly, Optional ByVal cTitle As String = "", Optional ByVal cHelpFile As String = "", Optional ByVal Context = "") As VbMsgBoxResult
    On Error Resume Next
    Screen.MousePointer = vbDefault
    MsgBox = VBA.MsgBox(sPrompt, enumButtons, "�������") ', cHelpFile, Context)
End Function

Public Function U8cBool(val As Variant) As Boolean
    On Error Resume Next
    If val = "��" Then
        U8cBool = True
        Exit Function
    ElseIf val = "��" Then
        U8cBool = False
        Exit Function
    End If
    U8cBool = CBool(val)
End Function
Public Function GetEleAtrVal(ele As IXMLDOMElement, Skey As String) As String
    Dim skey2 As String
    skey2 = LCase(Skey)
    If IsNull(ele.getAttribute(skey2)) Then
        GetEleAtrVal = ""
    Else
        GetEleAtrVal = ele.getAttribute(skey2)
    End If
End Function
'Զ���ն˴���
Public Function getCurrentSessionID() As String
    
    Dim objTerm As Object
    If IsWindow9X Then
        getCurrentSessionID = ""
    Else
        Set objTerm = CreateObject("TermMisc.Terminal")
        getCurrentSessionID = Trim(Str(objTerm.GetSessionID))
    End If
    Set objTerm = Nothing
End Function
Public Sub InitGrdCol(Grid As Object)
    With Grid
        .BackColor = &HFFFFFF
        .BackColorBkg = &HFFFFFF
        .BackColorFixed = &HFFE3C6
        .BackColorSel = &H9F6646
        .ForeColorSel = &HFFFFFF
        .GridColor = &HAEAEAE
        .GridColorFixed = &H888888
    End With

End Sub
Public Sub initFormColor(Frm As Form)
    Frm.BackColor = &HFFFFFF
End Sub

Public Function gszzPub() As ZzPub.clsPub
    If ogszzPub Is Nothing Then
        Set ogszzPub = New ZzPub.clsPub
        ogszzPub.InitPubs1 m_login.UfSystemDb, m_login.UfDbName, m_login.cacc_id, m_login.cIYear, m_login.cUserId, m_login.CurDate, m_login.SysPassword
        Set gszzPub = ogszzPub
    Else
        Set gszzPub = ogszzPub
    End If
End Function



Public Function GetWindowsDirectoryFromSystem() As String
    Dim strPath As String
    strPath = Space$(1024)
    Call GetSystemDirectory(strPath, 1024)
    strPath = Left(strPath, InStr(1, strPath, vbNullChar) - 1)
    Dim i As Long
    Dim length As Long
    length = Len(strPath)
    For i = length To 1 Step -1
        If Mid(strPath, i, 1) = "\" Then
            length = i
            Exit For
        End If
    Next
    GetWindowsDirectoryFromSystem = Left(strPath, length - 1)
End Function



'����Ƿ��½�
Public Function check_is_yj() As Boolean
Dim rdtemp As New ADODB.Recordset
On Error GoTo Err
    rdtemp.Open "select * from GL_mend where iperiod=" & m_login.iMonth, DBconn, adOpenStatic, adLockReadOnly
    If rdtemp!bflag_FA Then
        check_is_yj = True
    Else
        check_is_yj = False
    End If
    Set rdtemp = Nothing
    Exit Function
Err:
    check_is_yj = True
    Set rdtemp = Nothing
End Function



Public Sub ChangeOneFormTbr(Frm As Form, objTbl As Control, objU8Tbl As Control, Optional strCardNum As String)
    
    '//U872�������չ��ť�Ĺ���
    If strCardNum <> "" Then
        Call objU8Tbl.InitExternalButton(strCardNum, m_login)
    End If
    
    objU8Tbl.SetToolbar objTbl
    objU8Tbl.SetDisplayStyle TextOnly
    objTbl.Visible = False
    objU8Tbl.Visible = True
    objU8Tbl.Left = objTbl.Left
    objU8Tbl.Top = objTbl.Top
    objU8Tbl.Width = Frm.Width - 6 * Screen.TwipsPerPixelX
    objU8Tbl.Height = objTbl.Height
End Sub


Public Sub SendPortalMessage(strFormGuid As String, strCardNumber As String, strID As String, _
                               Optional strMessageType As String = "CurrentDocChanged", _
                               Optional strMaker As String, Optional ufts As String = "", Optional cVoucherCode As String = "", Optional mVoucherType As String = "", Optional bReturnFlag As Boolean = False)
    Dim tsb As Object
    Dim strXml As String
    Dim mystrCardNumber As String
    Select Case strCardNumber
    Case "01", "03", "05", "06", "28"
        mystrCardNumber = "01"
    Case "07", "13", "14", "15", "07Red", "13Red", "14Red", "15Red"
        mystrCardNumber = "07"
    Case "02", "04"
        mystrCardNumber = "02"
    Case Else
        mystrCardNumber = strCardNumber
    End Select
    
    Dim Authid As String
    Dim AbbAuthid As String
    Call GetAuthIdForWf(mVoucherType, bReturnFlag, Authid, AbbAuthid)
    If Not (g_business Is Nothing) Then
        Set tsb = g_business.GetToolbarSubjectEx(strFormGuid)
    End If
    strXml = "<?xml version='1.0' encoding='UTF-8'?>"
    strXml = strXml & "<Message type='" & strMessageType & "'>"
    strXml = strXml & "<Selection context='EF:" + strCardNumber + "'>"
    strXml = strXml & "<Element typeName='Voucher' cVoucherId='" & strID & "' cMaker='" & strMaker & "' cCardNum='" & mystrCardNumber & "' Ufts='" & ufts & "' cVoucherCode='" & cVoucherCode & "'  AuditAuthId ='" & Authid & "'  AbandonAuthId='" & AbbAuthid & "' />"
    strXml = strXml & "</Selection>"
    strXml = strXml & "</Message>"
    If Not (tsb Is Nothing) Then
        If bVerifyMsg Then
            Call g_business.AsyncTransMessage(strFormGuid, strXml)
        Else
            Call tsb.TransMessage(strFormGuid, strXml)
        End If
    End If
    Set tsb = Nothing
End Sub

Public Sub GetAuthIdForWf(mVoucherType As String, bReturnFlag As Boolean, AuditAuthId As String, ByRef AbandonAuthId As String)
    Select Case LCase(mVoucherType)
        Case "sa18"
            AuditAuthId = "SA03120204"
          AbandonAuthId = "SA03120205"
        Case "sa19"
            AuditAuthId = "SA03120304"
          AbandonAuthId = "SA03120305"
        Case LCase("EFBWGL020301")
            AuditAuthId = "EFBWGL02040103"
            AbandonAuthId = "EFBWGL02040104"
'            AuditAuthId = "EFBWGL020301"
'          AbandonAuthId = "EFBWGL02030101"
    End Select
End Sub

Public Sub ChangeToolbar()
    Dim Frm As Form
    Dim Obj As Control
    Dim objTbl As Control
    Dim objU8Tbl As Control

    For Each Frm In Forms
        For Each Obj In Frm
            If TypeName(Obj) = "Toolbar" Then Set objTbl = Obj
            If TypeName(Obj) = "CTBCtrl" Then Set objU8Tbl = Obj
            'by lg070314 ����U870�˵��ں�,UFToolBar�Ĵ���
            If LCase(TypeName(Obj)) = "uftoolbar" Then Set objU8Tbl = Obj
            If Not (objTbl Is Nothing) And Not (objU8Tbl Is Nothing) Then
                ChangeOneFormTbr Frm, objTbl, objU8Tbl
                Set objTbl = Nothing
                Set objU8Tbl = Nothing
                Exit For
            End If
        Next
    Next
End Sub
 

Public Function Unload_frms(FrmNames As String, Optional frmquantity As Long = 2) As Boolean
Dim frmX  As Form
Dim i As Long
i = 0
For Each frmX In Forms
    If LCase(frmX.Name) = LCase(FrmNames) Then
        i = i + 1
        If i > frmquantity Then
             Unload frmX
        End If
    End If
Next
End Function

Public Function bVerifyCanModify(CardNum As String, Mid As String, MCode As String, tmpAuthId As String) As Boolean
    On Error GoTo Errhandle
    Dim AuditServiceProxy As Object
    Dim tmpbCanModify As Boolean
    Dim tmpIdName As String
    Dim tmpCodeName As String
    Dim sErr As String
    Dim callerCtx As Object
    Set callerCtx = CreateObject("UFSoft.U8.Framework.LoginContext.CalledContext")
    callerCtx.subid = m_login.cSub_Id
    callerCtx.TaskID = m_login.TaskID
    callerCtx.token = m_login.userToken
    Set AuditServiceProxy = CreateObject("UFIDA.U8.Audit.ServiceProxy.AuditServiceProxy")
    tmpbCanModify = AuditServiceProxy.IsChangeableVoucher(Mid, CardNum, MCode, callerCtx, sErr)
    If tmpbCanModify = True Then
        If LockItem(tmpAuthId, 1) = False Then
            bVerifyCanModify = False
            Call LockItem(tmpAuthId, 0)
            Exit Function
        End If
        Call LockItem(tmpAuthId, 0)
    Else
        MsgBox GetString("U8.pu.prjpu860.04753") '�����в������޸�
        bVerifyCanModify = False
        Exit Function
    End If
    bVerifyCanModify = True
    Exit Function
Errhandle:
    bVerifyCanModify = False
End Function

'�ύ�����Ĵ���
'CbillType ί�д����㵥���Ƿ� רƱ ��Ʊ
Public Function DoUndoSubmit(m_Handle As Boolean, m_CardNumber As String, m_Mid As String, m_TablName As String, m_ufts As String, isWfcontrolled As Boolean, ByRef strErr As String, Optional cVoucherCode As String, Optional CbillType As String = "") As Boolean
    On Error GoTo ErrHandler
    Dim objCalledContext As Object
    Set objCalledContext = CreateObject("UFSoft.U8.Framework.LoginContext.CalledContext")
    objCalledContext.subid = m_login.cSub_Id
    objCalledContext.TaskID = m_login.TaskID
    objCalledContext.token = m_login.userToken
    Dim clsSub As Object
    Set clsSub = CreateObject("EFWorkFlowSrv.clsSAWorkFlowSrv")
    
    Dim Obj As Object
    Set Obj = CreateObject("UFLTMService.clsService")
    Obj.Start DBconn.ConnectionString
    Obj.BeginTransaction
    
    If m_Handle Then
        DoUndoSubmit = clsSub.DoSubmit(m_CardNumber, m_CardNumber & ".Submit", m_Mid, "", objCalledContext, m_ufts, isWfcontrolled, strErr, m_login, CbillType)
    Else
        DoUndoSubmit = clsSub.UndoSubmit(m_CardNumber, m_CardNumber & ".Submit", m_Mid, m_CardNumber, objCalledContext, m_ufts, isWfcontrolled, strErr, cVoucherCode, m_login)
    End If
    If DoUndoSubmit Then
        Obj.Commit
    Else
        Obj.Rollback
    End If
    Obj.Finish
    Set Obj = Nothing
    Exit Function
ErrHandler:
    strErr = VBA.Err.Description
    DoUndoSubmit = False
End Function


Public Function IsInForms(strFormName As String, Optional iIndex As Long) As Boolean
    Dim frmIndex As Long
    For frmIndex = Forms.Count To 1 Step -1
        If Forms(frmIndex - 1).Caption = strFormName Then
           IsInForms = True
           If Forms(frmIndex - 1).WindowState = 1 Then Forms(frmIndex - 1).WindowState = 2
           iIndex = frmIndex - 1
           Exit Function
        End If
    Next
    IsInForms = 0
End Function
Public Function UA_Task(ByVal TaskID As String) As Boolean
    On Error GoTo Errhandle
    Dim sStr As String
    
    
    If Not m_login Is Nothing Then
        m_login.ClearError
        If m_login.TaskExec(Trim(TaskID), -1, m_login.cIYear) Then
            UA_Task = True
            Exit Function
        Else
            If TaskID = "GS050103" Or TaskID = "GS050104" Then
              UA_Task = False
              Exit Function
            End If
            If m_login.ShareString <> "" Then
                MsgBox m_login.ShareString, 64, Msg_Title
            Else
                MsgBox "����(����)��ͻ��û�д��������Ȩ�ޣ����Ժ����ԡ�", 64, Msg_Title
            End If
            m_login.ClearError
            UA_Task = False
            Exit Function
        End If
    Else
        MsgBox "ϵͳ�����ע�����������쳣,���ܽ��й�������,�������绷����", vbCritical, Msg_Title
        UA_Task = False
        Exit Function
    End If
'    UA_Task = True
    Exit Function
 
Errhandle:
    MsgBox Err.Description, vbExclamation, Msg_Title
  
End Function

Public Function UA_FreeTask(ByVal TaskID As String) As Boolean
 On Error GoTo Errhandle
 
 If Not m_login Is Nothing Then
    m_login.ClearError
     If m_login.TaskExec(TaskID, 0, m_login.cIYear) Then
        UA_FreeTask = True
     Else
        m_login.ClearError
        UA_FreeTask = False
     End If
 Else
     MsgBox "ϵͳ�����ע�����������쳣,���ܽ��й����ͷ�,�������绷����", vbCritical, Msg_Title
     UA_FreeTask = False
     Exit Function
 End If
'    UA_FreeTask = True
     Exit Function
Errhandle:
  MsgBox Err.Description, vbExclamation, Msg_Title
End Function
Public Sub LoadHelpId(HelpID As String, appForm As Form)
    appForm.HelpContextID = HelpID
End Sub
  
Public Function bVerifyCanModifyByTaskInfo(CardNum As String, Mid As String, MCode As String, tmpAuthId As String) As Boolean
On Error GoTo Errhandle
    Dim rs As ADODB.Recordset
    Dim tmpbCanModify  As Boolean
    Dim tmpvalue As String
    Dim sql As String
    Set rs = New ADODB.Recordset
    bVerifyCanModifyByTaskInfo = True
    sql = "select top 1 * from V_WF_WFTaskInfo where vouchertype='" & CardNum & "' and vouchercode='" & MCode & "' and voucherid = '" & Mid & "'  order by ccreatetime desc"
    rs.Open sql, DBconn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            If rs.Fields("CtaskType") & "" = 5 Then
                tmpbCanModify = True
            Else
                tmpbCanModify = False
            End If
            
        End If
    End If
    If tmpbCanModify = True Then
        bVerifyCanModifyByTaskInfo = True
        Exit Function
    Else
        bVerifyCanModifyByTaskInfo = False
        Exit Function
    End If
    Exit Function
Errhandle:
    bVerifyCanModifyByTaskInfo = True
End Function
 
  
'�õ������ļ�
Public Sub GetHelpFile()
Dim dom As New DOMDocument
Dim GetHelpFileName As String
    If dom.Load(App.Path & "\" & App.EXEName & ".XML") Then
        If Not dom.selectSingleNode("//ProductFacade_Information/Main/HelpFile") Is Nothing Then
            GetHelpFileName = dom.selectSingleNode("//ProductFacade_Information/Main/HelpFile").Text
            App.HelpFile = m_login.GetIstallPath & "\Help\" & GetHelpFileName
            Wrtlog "�����ļ�װ�سɹ�!"
        End If
    End If
End Sub

''��־�ļ�
Public Sub Wrtlog(Msg As String, Optional sLogName As String = "")
    On Error Resume Next
    Dim fs As Object
    Dim oLogFile As Object
    If Trim(sLogName) = "" Then sLogName = App.EXEName & ".log"
    
    sLogName = App.EXEName & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Dir(App.Path & "\" + sLogName) = "" Then
        Set oLogFile = fs.CreateTextFile(App.Path & "\" + sLogName, True, True)
    Else
        Set oLogFile = fs.OpenTextFile(App.Path & "\" + sLogName, ForAppending, False, TristateTrue)
        If FileLen(App.Path & "\" + sLogName) > 100000000 Then
            Set oLogFile = fs.CreateTextFile(App.Path & "\" + sLogName, True, True)
        End If
    End If
    Msg = Now & "  " & Msg
    Call oLogFile.WriteLine(Msg)
    oLogFile.Close
    Set oLogFile = Nothing
End Sub



Public Function SetHeadItemValue(ByVal domHead As DOMDocument, ByVal Skey As String, ByVal value As Variant) As Boolean
    Skey = LCase(Skey)
    If Not domHead.selectSingleNode("//z:row").Attributes.getNamedItem(Skey) Is Nothing Then
        domHead.selectSingleNode("//z:row").Attributes.getNamedItem(Skey).nodeValue = value
        SetHeadItemValue = True
    Else
        SetHeadItemValue = False
    End If
End Function

