Attribute VB_Name = "Mod"
'****************************************
'���̻�������˵����
'               1��ʵ�ֵ��ݵĻ�������
'               2��ʵ���б�Ļ�������

'����ʱ�䣺2008-11-21
'�����ˣ�xuyan
'****************************************
Option Explicit
' * ���ݺ�
Public gstrCardNumber As String '=ua_menu���е�cMenu_Id�ֶε�ֵ������CARDNUM
Public gstrCardNumberlist As String '�б�id

Public g_oBusiness As Object '����business����,���Ż�����,�ö�������ã�һ��������ȡLogin���󣬶��������Դ������ʾ���رա�����Ƚ��в�����
Public g_oLogin As U8Login.clsLogin '����Login����,��ԭ861��Login����ͬ�����������Ż�����,����ͨ��business�����ӵõ�������"u8ע�Ჿ��[v8.71]"
Public g_bLogined As Boolean '����

Public goBusiness As Object '�б�business����,���Ż�����,�ö�������ã�һ��������ȡLogin���󣬶��������Դ������ʾ���رա�����Ƚ��в�����
Public goLogin As U8Login.clsLogin '�б�Login����,��ԭ861��Login����ͬ�����������Ż�����,����ͨ��business�����ӵõ�������"u8ע�Ჿ��[v8.71]"
Public gbLogined As Boolean '�б�

Public g_Conn As New ADODB.Connection '��������
Public gConn As New ADODB.Connection '�б�����

Public clsbill As Object ' USERPCO.VoucherCO '?����ʱ��ȡsql��
Public mologin As USCOMMON.login '?

Public gsGUIDForVouch As String                '���ݵ�GUID
Public gsGUIDForList As String                 '�б��GUID

Public idtmp As String
Public gMoCode As String
    
'����������õı�ͷ��������DOM
Public gDomReferHead As DOMDocument
Public gDomReferBody As DOMDocument
Public moneytmp As Double
Public numbertmp As Double
Public numappprice As Double
Public isfyflg As Boolean

Public iSinvCZ As Boolean

Public m_OK As Boolean
 
Public tmpTableName As String
Public bSwFlag As Boolean
  
 

'�Ƿ���ƴ��Ȩ��,���Ȩ�޷��� 11-7-12 chenliangc
Public bInv_ControlAuth As Boolean
Public sAuth_invR As String, sAuth_invW As String
Attribute sAuth_invW.VB_VarUserMemId = 1073741841

'�Ƿ����ҵ��ԱȨ��,ҵ��ԱȨ���ַ��� 11-7-12 chenliangc
Public bPerson_ControlAuth As Boolean
Attribute bPerson_ControlAuth.VB_VarUserMemId = 1073741843
Public sAuth_personR As String, sAuth_personW As String
Attribute sAuth_personR.VB_VarUserMemId = 1073741844
Attribute sAuth_personW.VB_VarUserMemId = 1073741844

'�Ƿ���ƹ�Ӧ��Ȩ��,��Ӧ��Ȩ���ַ��� 11-7-13 chenliangc
Public bVendor_ControlAuth As Boolean
Attribute bVendor_ControlAuth.VB_VarUserMemId = 1073741846
Public sAuth_vendorR As String, sAuth_vendorW As String
Attribute sAuth_vendorR.VB_VarUserMemId = 1073741847
Attribute sAuth_vendorW.VB_VarUserMemId = 1073741847

'�Ƿ���Ʋ���Ȩ��,����Ȩ���ַ��� 11-7-13 chenliangc
Public bDep_ControlAuth As Boolean
Attribute bDep_ControlAuth.VB_VarUserMemId = 1073741849
Public sAuth_depR As String, sAuth_depW As String
Attribute sAuth_depR.VB_VarUserMemId = 1073741850
Attribute sAuth_depW.VB_VarUserMemId = 1073741850

'user auth
Public bCheckUser As Boolean
Public isfromcon As Boolean

'�Ƿ���ƿͻ�Ȩ��,�ͻ�Ȩ���ַ��� 11-7-13 chenliangc
Public bCus_ControlAuth As Boolean
Attribute bCus_ControlAuth.VB_VarUserMemId = 1073741852
Public sAuth_CusR As String, sAuth_CusW As String
Attribute sAuth_CusR.VB_VarUserMemId = 1073741853
Attribute sAuth_CusW.VB_VarUserMemId = 1073741853

'�Ƿ���Ʋֿ�Ȩ��,�ֿ�Ȩ���ַ��� 11-7-13 chenliangc
Public bWareHouse_ControlAuth As Boolean
Attribute bWareHouse_ControlAuth.VB_VarUserMemId = 1073741855
Public sAuth_WareHouseR As String, sAuth_WareHouseW As String
Attribute sAuth_WareHouseR.VB_VarUserMemId = 1073741856
Attribute sAuth_WareHouseW.VB_VarUserMemId = 1073741856

'�Ƿ���ƻ�λȨ��,��λȨ���ַ��� 11-7-13 chenliangc
Public bPosition_ControlAuth As Boolean
Attribute bPosition_ControlAuth.VB_VarUserMemId = 1073741858
Public sAuth_PositionR As String, sAuth_PositionW As String
Attribute sAuth_PositionR.VB_VarUserMemId = 1073741859
Attribute sAuth_PositionW.VB_VarUserMemId = 1073741859

'�Ƿ���Ʋ���ԱȨ��,����ԱȨ���ַ��� 11-7-12 chenliangc
Public sAuth_cmaker As String
Attribute sAuth_cmaker.VB_VarUserMemId = 1073741861

'�ɼ����� 11-7-12 chenliangc
Public sAuth_ALL As String
Attribute sAuth_ALL.VB_VarUserMemId = 1073741862
Public sAuth_AllList As String

Public sMakeAuth_ALL As String

Public sAuth_UnitR As String

Public pageCount As Long   '��ҳ��
Attribute pageCount.VB_VarUserMemId = 1073741863
Public PageCurrent As Long  '��ǰҳ
Attribute PageCurrent.VB_VarUserMemId = 1073741864
Public lngVoucherID As Long ' ��������IDֵ
Attribute lngVoucherID.VB_VarUserMemId = 1073741865
Public MainTable As String '��������
Attribute MainTable.VB_VarUserMemId = 1073741866
Public DetailsTable As String '�����ֱ�
Attribute DetailsTable.VB_VarUserMemId = 1073741867
Public HeadPKFld As String '���������ֶ�
Attribute HeadPKFld.VB_VarUserMemId = 1073741868
Public MainView As String '��ͷ��ͼ
Attribute MainView.VB_VarUserMemId = 1073741869
Public DetailsView As String '������ͼ
Attribute DetailsView.VB_VarUserMemId = 1073741870

Public TblName As String
Public ViewDetailName As String
Public ViewMainName As String
Public VoucherList As String '�б���ͼ
Attribute VoucherList.VB_VarUserMemId = 1073741871
Public VoucherList2 As String '�б���ͼ

Public conid As String
'��¼��ǰ����״̬�������ӣ��޸ģ�����
Public mOpStatus As OpStatus

'���ݱ��,�Ƶ���,���������ֶ���
Public strcCode, StrcMaker, StrdDate As String
Attribute strcCode.VB_VarUserMemId = 1073741872
Attribute StrcMaker.VB_VarUserMemId = 1073741872
Attribute StrdDate.VB_VarUserMemId = 1073741872
'�����,����,״̬,�ر���,�ر�����,������,��������
Public StrcHandler, StrdVeriDate, StriStatus, StrCloseUser, StrdCloseDate, StrIntoUser, StrdIntoDate As String
Attribute StrcHandler.VB_VarUserMemId = 1073741875
Attribute StrdVeriDate.VB_VarUserMemId = 1073741875
Attribute StriStatus.VB_VarUserMemId = 1073741875
Attribute StrCloseUser.VB_VarUserMemId = 1073741875
Attribute StrdCloseDate.VB_VarUserMemId = 1073741875
Attribute StrIntoUser.VB_VarUserMemId = 1073741875
Attribute StrdIntoDate.VB_VarUserMemId = 1073741875

Public sID, sAutoId As Long '��������id,���ӱ�autoid
Attribute sID.VB_VarUserMemId = 1073741882
Attribute sAutoId.VB_VarUserMemId = 1073741882

Public sTmpTableName As String '��λ/�б�λ ��ʱ����
Public sGUID As String '��ʱ������Ҫ�õ���guid

Public strCellCode As String '���ر��������
Attribute strCellCode.VB_VarUserMemId = 1073741884
Public strCellName As String '���ر��������
Attribute strCellName.VB_VarUserMemId = 1073741885
Public symbol As String '�������㷽ʽ
Attribute symbol.VB_VarUserMemId = 1073741886

Public strWhere As String '��������
Attribute strWhere.VB_VarUserMemId = 1073741887
Public strWhere2 As String '��������

Public TimeStamp As String   'ʱ���
Attribute TimeStamp.VB_VarUserMemId = 1073741888
Public OldTimeStamp As String    '��ǰ���ݵ�ʱ��� chenliangc��Ҳ���ڵ���ģ��Ԥ��ufts��ʹ��voucher.headertext("ufts")
Attribute OldTimeStamp.VB_VarUserMemId = 1073741889

'Public Const HelpFile = "\Help\���۹���_zh-CN.chm"    '�����ļ�·��
'Public Const HelpFile = "\HY\client\HY_DZ_JA_JYGH\������ҵ�������.chm"    '�����ļ�·��
Public Const HelpFile = "\Help\ST_zh-CN.chm"  '�����ļ�·��


'����Ȩ��
Public Const AuthBrowse = "FYSL02050301" '���
Public Const AuthBrowselist = "PD01030101" '���


'ʱ����жϽ��
Public Const RecordDeleted = 1 '�ѱ������û�ɾ��
Public Const RecordModified = 2 '�ѱ������û��޸�
Public Const RecordNoChanged = 0 '����,���Բ���
Public Const RecordError = -1     '�쳣


'U8ϵͳ��Ϣ
Public m_SysInfor As clsSystemInfo
Attribute m_SysInfor.VB_VarUserMemId = 1073741890
Public clsInfor As Object 'Info_PU.ClsS_Infor
Attribute clsInfor.VB_VarUserMemId = 1073741891

'��ʽ�������
Public m_sQuantityFmt As String '������ʽ
Attribute m_sQuantityFmt.VB_VarUserMemId = 1073741892
Public m_sNumFmt As String      ' ��ֵ��ʽ
Attribute m_sNumFmt.VB_VarUserMemId = 1073741893
Public m_iExchRateFmt As String   ' ������
Attribute m_iExchRateFmt.VB_VarUserMemId = 1073741894
Public m_iRateFmt As String   ' ˰��
Attribute m_iRateFmt.VB_VarUserMemId = 1073741895
Public m_sPriceFmt As String  ' ����ʽ
Attribute m_sPriceFmt.VB_VarUserMemId = 1073741896
Public m_sPriceFmtSA As String  ' ����ʽ�������ã�
Attribute m_sPriceFmtSA.VB_VarUserMemId = 1073741897

Public gcCreateType As String     '���� ��������
Attribute gcCreateType.VB_VarUserMemId = 1073741898
'Public tmpLinkTbl As String '������������ʱccmdline����������ʱ���� '�������� ʱ ��ť״̬���� by zhangwchb 20110809


' * �Ƿ��ַ�
'##ModelId=42F6FF0701F4
Public Const gstrBAD_STRING As String = " ~`!@#$%^&*()-:;+={}[]'\|<>?,./"

'����״̬
Public Enum OpStatus
    ADD_MAIN = 1              '��������
    ADD_SUB = 2               '�����Ӽ�
    MODIFY_MAIN = 3           '�޸�����
    MODIFY_SUB = 4            '�޸��Ӽ�
    DELETE_MAIN = 5           'ɾ������
    DELETE_SUB = 6            'ɾ���Ӽ�
    SHOW_MAIN = 7             '��������ֻ����ʾ
    SHOW_SUB = 8              '�Ӽ�����ֻ����ʾ
    SHOW_ALL = 9              '����ֻ����ʾ
    ADD_MAIN_AFTER = 10       '������������
    ADD_SUB_AFTER = 11        '�����Ӽ�����
    MODIFY_MAIN_AFTER = 12    '�޸���������
    MODIFY_SUB_AFTER = 13     '�޸��Ӽ�����
    DELETE_MAIN_AFTER = 14    'ɾ����������
    DELETE_SUB_AFTER = 15     'ɾ���Ӽ�ʾ����
    SHOW_NOTHING = 16         'û�ж�Ӧ����
End Enum
' ***********************************************************
' * ����ģʽ
'
#If DEBUG_MODE = 1 Then
    Public Const g_blnDEBUG_MODE As Boolean = True
#Else
    Public Const g_blnDEBUG_MODE As Boolean = False
#End If

' ***********************************************************
' * ���󼶱�
' *
' * (���ڶ��ƴ�����Ϣ����Ϣ��)
Public Enum ErrorLevelConstants
    ufsELAllInfo = 0                ' ����������Ϣ(�Ѻ���Ϣ������š� ����Դ����������)
    ufsELOnlyHeader = 1             ' ֻ�����Ѻ���ʾ��Ϣ(�Զ���)
    ufsELHeaderAndDescription = 2   ' ֻ���������������Ѻ���Ϣ
End Enum
' *
' * ���嶯��
Public Enum FormActionConstants
    ufsFANew = 1    ' ����
    ufsFAEdit = 2   ' �༭
    ufsFAView = 0   ' ���
End Enum

Public Enum BillType       '�Ƶ�����
    ���� = 0
    �ɹ� = 1
    ��� = 2
    Ӧ�� = 3
End Enum

Public Enum SaleVoucherType   '�����Ƶ�ģʽ
    ר�÷�Ʊ = 0
    ����ר�÷�Ʊ = 1
    ��ͨ��Ʊ = 2
    ������ͨ��Ʊ = 3
    ���۷����� = 9
    �˻��� = 10
    ���۶��� = 12
    ί�д��������� = 15
    ί�д����˻��� = 16
    ���۱��۵� = 21
    
End Enum

Public Enum PUVoucherType   '�ɹ��Ƶ�ģʽ
    �ɹ��빺�� = 0
    �ɹ����� = 1
    �ɹ������� = 2
    �ɹ���Ʊ = 4
End Enum

Public Enum COVoucherType   '����Ƶ�ģʽ
    �ɹ���ⵥ = 1
    �ɹ���Ʊ = 4
    ������ⵥ = 8
    �������ⵥ = 9
    ����Ʒ��ⵥ = 10
    ���ϳ��ⵥ = 11
    ������ = 12
    ��װ�� = 13
    ��ж�� = 14
    ��̬ת���� = 15
    �̵㵥 = 18
    ���۳��ⵥ = 32
    ���ϸ�Ʒ���� = 46
    ������ϸ�Ʒ = 55
    �������뵥 = 62
End Enum


'��ȡȨ���������
Public Sub getAuthString(conn As ADODB.Connection)
    sAuth_ALL = "(1=1)"
    sMakeAuth_ALL = "(1=1)"
    sAuth_AllList = "(1=1)"
    Dim sauth_unit As String                               '��ͷ��λȨ��-�п����ǿͻ�����Ӧ�̣�����
    sauth_unit = "(1=2)"

    bPerson_ControlAuth = False
    bInv_ControlAuth = False
    bVendor_ControlAuth = False
    bDep_ControlAuth = False
    bCus_ControlAuth = False
    bWareHouse_ControlAuth = False
    bPosition_ControlAuth = False
    bCheckUser = False

'    '���Ȩ��
'    sAuth_invR = ""
'    sAuth_invW = ""
'    If LCase(getAccinformation("ST", "bInventoryCheck", conn)) = "true" Then
'        bInv_ControlAuth = True
'        sAuth_invR = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Inventory", , "R")
'        If sAuth_invR = "1=2" Then sAuth_invR = "-1"
'        sAuth_invW = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Inventory", , "W")
'        If sAuth_invW = "1=2" Then sAuth_invW = "-1"
'
'        sAuth_ALL = sAuth_ALL & IIf(sAuth_invR = "", "", " and (not exists(select top 1 iid from HY_DZ_BorrowOuts a INNER JOIN inventory b ON a.cInvCode=b.cinvcode where a.id=HY_DZ_BorrowOut.id AND isnull(b.iid,0) not in (" & sAuth_invR & ")))")
'        sMakeAuth_ALL = sMakeAuth_ALL & IIf(sAuth_invW = "", "", " and (exists(select top 1 iid from HY_DZ_BorrowOuts a INNER JOIN inventory b ON a.cInvCode=b.cinvcode where a.id=HY_DZ_BorrowOut.id AND isnull(b.iid,0) in (" & sAuth_invW & ")))")
'        sAuth_AllList = sAuth_AllList & IIf(sAuth_invR = "", "", " and iid in (" & sAuth_invR & ")")
'
'    End If
'
'    '�Ƶ���Ȩ��
'    sAuth_cmaker = ""
'    If LCase(getAccinformation("ST", "bOperatorCheck", conn)) = "true" Then
'        bCheckUser = True
'        sAuth_cmaker = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "user", , "R")
'        sAuth_ALL = sAuth_ALL & IIf(sAuth_cmaker = "", "", " and (cmaker in (" & sAuth_cmaker & ")) ")
'        sMakeAuth_ALL = sMakeAuth_ALL & IIf(sAuth_cmaker = "", "", " and (cmaker in (" & sAuth_cmaker & ")) ")
'        sAuth_AllList = sAuth_AllList & IIf(sAuth_cmaker = "", "", " and (cmaker in (" & sAuth_cmaker & ")) ")
'    End If

    '    'ҵ��ԱȨ��
    '    sAuth_personR = ""
    '    sAuth_personW = ""
    '    If LCase(getAccinformation("ST", "bCheckPersonAuth", conn)) = "true" Then
    '        bPerson_ControlAuth = True
    '        sAuth_personR = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Person", , "R")
    '        sAuth_personW = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Person", , "W")
    '
    '        sAuth_ALL = sAuth_ALL & IIf(sAuth_personR = "", "", " and ( isnull(cpersoncode,'')='' or cpersoncode in (" & sAuth_personR & ")) ")
    '
    '    End If


    '����Ȩ��
    sAuth_depR = ""
    sAuth_depW = ""
    If LCase(getAccinformation("ST", "bDepartmentCheck", conn)) = "true" Then
        bDep_ControlAuth = True
        sAuth_depR = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Department", , "R")
        sAuth_depW = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Department", , "W")
        If sAuth_depR = "1=2" Then sAuth_depR = "'|'"
        If sAuth_depW = "1=2" Then sAuth_depW = "'|'"
        sAuth_ALL = sAuth_ALL & IIf(sAuth_depR = "", "", " and ( isnull(chdepartcode,'')='' or chdepartcode in (" & sAuth_depR & ")) ")
        sMakeAuth_ALL = sMakeAuth_ALL & IIf(sAuth_depR = "", "", " and ( isnull(chdepartcode,'')='' or chdepartcode in (" & sAuth_depR & ")) ")
        sAuth_AllList = sAuth_AllList & IIf(sAuth_depR = "", "", " and ( isnull(chdepartcode,'')='' or chdepartcode in (" & sAuth_depR & ")) ")
        
        sauth_unit = sauth_unit

    Else
        sauth_unit = sauth_unit

    End If

'    '�ͻ�Ȩ��
'    sAuth_CusR = ""
'    sAuth_CusW = ""
'    If LCase(getAccinformation("ST", "bCustomerCheck", conn)) = "true" Then
'        bCus_ControlAuth = True
'        sAuth_CusR = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Customer", , "R")
'        sAuth_CusW = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Customer", , "W")
'
'        '11.0 �ͻ�iid�ĳ�nvarchar������
'        If sAuth_CusR = "1=2" Then sAuth_CusR = "'-1'"
'        If sAuth_CusW = "1=2" Then sAuth_CusW = "'-1'"
'        sauth_unit = sauth_unit & " or " & IIf(sAuth_CusR = "", "(ctype='�ͻ�')", " (ctype='�ͻ�' and ( isnull(bObjectCode,'')='' or bObjectCode in (select ccuscode from customer where iid in (" & sAuth_CusR & ")))) ")
'    Else
'        sauth_unit = sauth_unit & " or " & " (ctype='�ͻ�') "
'    End If
'
'    '��Ӧ��Ȩ��
'    sAuth_vendorR = ""
'    sAuth_vendorW = ""
'    If LCase(getAccinformation("ST", "bVendorCheck", conn)) = "true" Then
'        bVendor_ControlAuth = True
'        sAuth_vendorR = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Vendor", , "R")
'        sAuth_vendorW = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Vendor", , "W")
'        If sAuth_vendorR = "1=2" Then sAuth_vendorR = "-1"
'        If sAuth_vendorW = "1=2" Then sAuth_vendorW = "-1"
'        sauth_unit = sauth_unit & " or " & IIf(sAuth_vendorR = "", "(ctype='��Ӧ��')", " (ctype='��Ӧ��' and ( isnull(bObjectCode,'')='' or bObjectCode in (select cvencode from vendor where iid in (" & sAuth_vendorR & ")))) ")
'    Else
'        sauth_unit = sauth_unit & " or " & " (ctype='��Ӧ��') "
'    End If
'
'    '�ֿ�Ȩ��
'    sAuth_WareHouseR = ""
'    sAuth_WareHouseW = ""
'    If LCase(getAccinformation("ST", "bWarehouseCheck", conn)) = "true" Then
'        bWareHouse_ControlAuth = True
'        sAuth_WareHouseR = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Warehouse", , "R")
'        sAuth_WareHouseW = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Warehouse", , "W")
'        If sAuth_WareHouseR = "1=2" Then sAuth_WareHouseR = "'|'"
'        If sAuth_WareHouseW = "1=2" Then sAuth_WareHouseW = "'|'"
'        sAuth_ALL = sAuth_ALL & IIf(sAuth_WareHouseR = "", "", " and (not exists(select top 1 1 from HY_DZ_BorrowOuts a where a.id=HY_DZ_BorrowOut.id and isnull(a.cwhcode,'')<>'' and a.cwhcode not in (" & sAuth_WareHouseR & ")))")
'        sMakeAuth_ALL = sMakeAuth_ALL & IIf(sAuth_WareHouseW = "", "", " and (exists(select top 1 1 from HY_DZ_BorrowOuts a where a.id=HY_DZ_BorrowOut.id and (ISNULL(a.cwhcode,N'')=N'' OR a.cwhcode in (" & sAuth_WareHouseW & "))))")
'
'        sAuth_AllList = sAuth_AllList & IIf(sAuth_WareHouseR = "", "", " and ( isnull(cwhcode,'')='' or cwhcode in (" & sAuth_WareHouseR & ")) ") '" and (not exists(select top 1 1 from HY_DZ_BorrowOuts a where a.id=HY_DZ_BorrowOut.id and a.cwhcode not in (" & sAuth_WareHouseR & ")))")
'    End If
'
'    '��λȨ��
'    sAuth_PositionR = ""
'    sAuth_PositionW = ""
'    If LCase(getAccinformation("ST", "bPostionCheck", conn)) = "true" Then
'        bPosition_ControlAuth = True
'        sAuth_PositionR = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Position", , "R")
'        sAuth_PositionW = GetRowAuth(conn.ConnectionString, g_oLogin.cUserId, "Position", , "W")
'        If sAuth_PositionR = "1=2" Then sAuth_PositionR = "'|'"
'        If sAuth_PositionW = "1=2" Then sAuth_PositionW = "'|'"
'        sAuth_ALL = sAuth_ALL & IIf(sAuth_PositionR = "", "", " and (not exists(select top 1 1 from HY_DZ_BorrowOuts a where a.id=HY_DZ_BorrowOut.id and (isnull(a.cposition,'')<>'' and a.cposition not in (" & sAuth_PositionR & "))))")
'        sMakeAuth_ALL = sMakeAuth_ALL & IIf(sAuth_PositionW = "", "", " and ( exists(select top 1 1 from HY_DZ_BorrowOuts a where a.id=HY_DZ_BorrowOut.id and (isnull(a.cposition,'')='' or a.cposition in (" & sAuth_PositionW & "))))")
'
'        sAuth_AllList = sAuth_AllList & IIf(sAuth_PositionR = "", "", " and ( isnull(cposition,'')='' or cposition in (" & sAuth_PositionR & ")) ")
'    End If
'
'    '12.0�����������λ����֧�������û�
'    sauth_unit = sauth_unit & " or " & " (ctype='�����û�') "
    
    '12.0������Ʒ��������������ɴ����ƣ�
    'sauth_unit = "(" & sauth_unit & ") and isnull(cborrowouttype,0)!=1"
    
    If sauth_unit = "(1=2)" Then sauth_unit = "(1=1)"
    sAuth_ALL = sAuth_ALL & " and (" & sauth_unit & ")"
    sAuth_AllList = sAuth_AllList & " and (" & sauth_unit & ")"
    sMakeAuth_ALL = sMakeAuth_ALL & " and (" & sauth_unit & ")"
    sAuth_UnitR = sauth_unit
    
End Sub

'������ԱȨ�� CheckUserAuth
'connstr ���ݿ����Ӵ�
'selfuserid ����Ĳ���Աuserid ,ֱ��ȡ login.cuserid
'objuserid ���ݵ��Ƶ��ˣ�ȡ�����ϵ��Ƶ��ˣ��˴�ֱ����username
'cfunctionid ����Ȩ���Ѿ�ͨ��getauthstring������, ��Ҫ����ɾ��-��W�������-��V��,����-��U�����ر�-��C��������-��A��
Public Function CheckUserAuth(connstr As String, selfuserid As String, objuserid As String, cfunctionid As String) As Boolean
'    Dim authsrv As Object
'    Set authsrv = CreateObject("U8RowAuthsvr.clsRowAuth")
    If bCheckUser = False Then
        CheckUserAuth = True
        Exit Function
    End If

    Dim authsrv As New U8RowAuthsvr.clsRowAuth
    If Not authsrv.Init(connstr, selfuserid, False, "ST") Then Exit Function
    CheckUserAuth = authsrv.IsHoldAuth("user", objuserid, , cfunctionid)
End Function


Public Function IsWFControlled() As Boolean
 
    On Error GoTo ErrHandle
    Dim cBizObjectId As String
    Dim bWFControlled As Boolean
    Dim errMsg As String
    Dim o As Object
    cBizObjectId = "HYJCGH001"
    IsWFControlled = False
    Set o = CreateObject("SCMWorkFlowCommon.clsWFController")
    'If o.getIsWfControl(mologin.AccountConnection, cBizObjectId, cBizObjectId & ".Submit", mologin.OldLogin.cIYear, mologin.OldLogin.cAcc_Id, bWFControlled, errMsg) Then
    If o.getIsWFHasActivated(mologin.AccountConnection, cBizObjectId, cBizObjectId & ".Submit", bWFControlled, errMsg) Then
       IsWFControlled = bWFControlled
    End If
    Set o = Nothing
    Exit Function
    
ErrHandle:
    IsWFControlled = False
    Exit Function
End Function
