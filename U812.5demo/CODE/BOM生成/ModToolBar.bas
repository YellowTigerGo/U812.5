Attribute VB_Name = "ModToolBar"
Option Explicit

'u8���ÿ�ݼ�

'����F5
'�޸�F8
'ɾ��DEL
'���Ctrl+U
'����ALT+U
'����F6
'����Ctrl+Z
'����Ctrl+N
'ɾ��Ctrl+D
'��ӡCtrl+P
'Ԥ��Ctrl+V
'���Alt+E
'ˢ��Ctrl+R
'��λCtrl+F3
'��ҳAlt+PageUp
'��һҳPageUp
'��һҳPageDown
'ĩҳAlt+PageDown
'����Ctrl+G
'�ر�Alt+C
'��Alt+O

Public varArgs() As Variant
'��������ť�ؼ���
Public Const sKey_Batchprint = "PrintBatch"         '����
Public Const sKey_Print = "Print"                   '��ӡ
Public Const sKey_Preview = "Preview"               'Ԥ��
Public Const sKey_Output = "Output"                 '���
Public Const sKey_Export = "Export"                 '���
Public Const sKey_Copy = "Copy"                     '����
Public Const sKey_Add = "Add"                       '����
Public Const sKey_Modify = "Modify"                 '�޸�
Public Const sKey_Delete = "Delete"                 'ɾ��
Public Const sKey_Save = "Save"                     '����
Public Const sKey_Discard = "Discard"               '����
Public Const sKey_Addrecord = "AddRecord"           '����
Public Const sKey_InsertRecord = "InsertRecord"     '����
Public Const sKey_Deleterecord = "DeleteRecord"     'ɾ��
Public Const sKey_AlterPO = "AlterPO"               '���
Public Const skey_ReportCheck = "ReportCheck"       '����
Public Const sKey_Close = "Close"                   '�ر�
Public Const sKey_Open = "Open"                     '��
Public Const sKey_Confirm = "Confirm"               '���
Public Const sKey_Cancelconfirm = "Cancelconfirm"   '����
Public Const sKey_Return = "Return"                 '�黹
Public Const sKey_BatchReturn = "BatchReturn"       '�����黹
Public Const sKey_QueryConfirm = "QueryConfirm"     '����(��������ѯ)
Public Const sKey_Payment = "Payment"               '�ָ�
Public Const sKey_Cancelpayment = "CancelPayment"   '����
Public Const sKey_BatchBV = "BatchBV"               '�������ɷ�Ʊ
Public Const sKey_Settle = "Settle"                 '����
Public Const sKey_Locate = "Locate"                 '��λ
Public Const sKey_LocateSet = "LocateSet"           '��λ����
Public Const sKey_Load = "Load"                     '����
Public Const sKey_First = "First"                   '����
Public Const sKey_Previous = "Previous"             '����
Public Const sKey_Next = "Next"                     '����
Public Const sKey_Last = "Last"                     'ĩ��
Public Const sKey_RefVoucher = "RefVoucher"         '�������� by zhangwchb 20110718
Public Const sKey_Refresh = "Refresh"               'ˢ��
Public Const sKey_Help = "Help"                     '����
Public Const sKey_Exit = "Exit"                     '�˳�
Public Const sKey_Lock = "Lock"                     '����
Public Const sKey_RLock = "removelock"              '����
Public Const sKey_Acc = "Accessories"               '����
Public Const sKey_Link = "Link"                     '����
Public Const sKey_Column = "Column"                 '��Ŀ
Public Const sKey_Fetchprice = "Fetchprice"         'ȡ��
'��������ť chenliangc
Public Const sKey_Submit = "Submit"                 '�ύ
Public Const sKey_Unsubmit = "Unsubmit"             '����
Public Const sKey_Resubmit = "Resubmit"             '�����ύ
Public Const sKey_ViewVerify = "ViewVerify"         '����
'����
Public Const sKey_ReferVoucher = "RererVoucher"     '����(һ��һ)
Public Const sKey_ReferVouchers = "RererVouchers"   '����(���һ)
Public Const sKey_CreateVoucher = "CreateVoucher"       '�Ƶ�
Public Const sKey_CreateSAVoucher = "CreateSAVoucher"       '�����۵�
Public Const sKey_CreatePUVoucher = "CreatePUVoucher"       '�Ʋɹ���
Public Const sKey_CreateSCVoucher = "CreateSCVoucher"       '�ƿ�浥
Public Const sKey_CreateAPVoucher = "CreateAPVoucher"       '��Ӧ����

'��ѡ
Public Const sKey_ReverseSelection = "ReverseSelection"       '��ѡ
Public Const sKey_VoucherDesign = "VoucherDesign"    '���ݸ�ʽ����
Public Const sKey_SaveVoucherDesign = "SaveVoucherDesign"    '���ݸ�ʽ����



'##ModelId=431947B203CD
Public Const gstrHelpCode As String = "Help"
'##ModelId=431947B203D6
Public gstrHelpText As String  '= "����"
'##ModelId=431947B203E1
Public gstrHelpTip As String  '= "����"
'##ModelId=431947B30002
Public Const gintHelpImg As Integer = 145


'�����б�����������Դ�Ĺؼ����ַ���
Public Const strKprintbill = "printbill"  '��ӡ����
Public Const strKfilter = "filter"   '����
Public Const strKfind = "find"    '����
Public Const strKsetfield = "setfield"   '������ʾ�ֶ�
Public Const strKsort = "sort"  '����
Public Const strKhelp = "help"    '����
Public Const strKclose = "close"   '�˳�
Public Const strKCard = "card"    '����
Public Const strKSelectAll = "SelectAll"    'ȫѡ
Public Const strKUnSelectAll = "UnSelectAll"    'ȫ��

Public Const strKComparePrice = "ComparePrice"    '�ȼ�



Public Const strKLock = "lock"
Public Const strKRLock = "removelock"


Public Const sKey_Add1 = "Add1" ' "Add1"                       '���� �ڳ�
Public strAdd1 As String  ' "�ڳ�"

Public Const sKey_Add2 = "Add2" ' "Add2"                       '���� ���ڳ�
Public strAdd2 As String  ' "����"
'��������ť��ʾ����
Public strBatchprint As String  ' "����"
Public strBatchOpen As String  ' "����"
Public strBatchClose As String  ' "����"
Public strBatchVeri As String  ' "����"
Public strBatchUnVeri As String  ' "����"
Public strPrint As String  ' "��ӡ"
Public strPreview As String  ' "Ԥ��"
Public strOutput As String  ' "���"
Public strCopy As String  ' "����"
Public strAdd As String  ' "����"
Public strModify As String  ' "�޸�"
Public strdelete As String  ' "ɾ��"
Public strSave As String  ' "����"
Public strDiscard As String  ' "����"
Public strAddrecord As String  ' "����"
Public strDeleterecord As String  ' "ɾ��"
Public strAlterPO As String  ' "���"
Public strReportCheck As String  ' "����"
Public strClose As String  ' "�ر�"
Public strOpen As String  ' "��"
Public strConfirm As String  ' "���"
Public strCancelconfirm As String  ' "����"
Public strQueryConfirm As String  ' "����"
Public strPayment As String  ' "�ָ�"
Public strCancelpayment As String  ' "����"
Public strBatchBV As String  ' "����"
Public strSettle As String  ' "����"
Public strLocate As String  ' "��λ"
Public strLocateSet As String  ' "��λ����"
Public strFirst As String  ' "����"
Public strPrevious As String  ' "����"
Public strNext As String  ' "����"
Public strLast As String  ' "ĩ��"
Public strRefVoucher As String  ' "��������"         '�������� by zhangwchb 20110718
Public strRefresh As String  ' "ˢ��"
Public strHelp As String  ' "����"
Public strFilter As String  ' "����"
Public strExit As String  ' "�˳�"
Public strColumn As String  ' "��Ŀ"
Public strSelectAll As String  ' "ȫѡ"
Public strUnSelectAll As String  ' "ȫ��"
Public strLock As String  ' "����"
Public strRLock As String  ' "����"
Public strBatchLock As String  ' "����"
Public strBatchRLock As String  ' "����"
Public strAcc As String  ' "����"
Public strLink As String  ' "����"
Public strFetchprice As String  ' "ȡ��"
'������ chenliangc
Public strSubmit As String  ' "�ύ"                 '�ύ
Public strUnsubmit As String  ' "����"             '����
Public strResubmit As String  ' "�����ύ"             '�����ύ
Public strViewVerify As String  ' "����"         '����

Public strReferVoucher As String  ' "����(һ��һ)"     '����(һ��һ)
Public strReferVouchers As String  ' "����(���һ)"   '����(���һ)
Public strCreateVoucher As String  ' "�Ƶ�"       '�Ƶ�
Public strCreateSAVoucher As String  ' "���۵���"       '�Ƶ�
Public strCreatePUVoucher As String  ' "�ɹ�����"       '�Ƶ�
Public strCreateSCVoucher As String  ' "��浥��"       '�Ƶ�
Public strCreateAPVoucher As String  ' "Ӧ������"       '�Ƶ�

Public strReverseSelection As String  ' "��ѡ"
Public strVoucherDesign As String  ' "��ʽ����"         '���ݸ�ʽ����
Public strSaveVoucherDesign As String  ' "���沼��"     '���ݸ�ʽ����

Public Sub InitMulText()
    strBatchprint = GetString("U8.DZ.JA.btn010")    '") '����"
    strBatchOpen = GetString("U8.DZ.JA.btn020")    '����"
    strBatchClose = GetString("U8.DZ.JA.btn030")    '����"
    strBatchVeri = GetString("U8.DZ.JA.btn035")    '����"
    strBatchUnVeri = GetString("U8.DZ.JA.btn040")    '����"
    strPrint = GetString("U8.DZ.JA.btn045")    '��ӡ"
    strPreview = GetString("U8.DZ.JA.btn050")    'Ԥ��"
    strOutput = GetString("U8.DZ.JA.btn055")    '���"
    strCopy = GetString("U8.DZ.JA.btn060")    '����"
    strAdd = GetString("U8.DZ.JA.btn065")    '����"
    strAdd1 = GetString("U8.DZ.JA.btn760")
    strAdd2 = GetString("U8.DZ.JA.btn065")    '����"
    strModify = GetString("U8.DZ.JA.btn070")    '�޸�"
    strdelete = GetString("U8.DZ.JA.btn075")    'ɾ��"
    strSave = GetString("U8.DZ.JA.btn080")    '����"
    strDiscard = GetString("U8.DZ.JA.btn090")    '����"
    strAddrecord = GetString("U8.DZ.JA.btn100")    '����"
    strDeleterecord = GetString("U8.DZ.JA.btn110")    'ɾ��"
    strAlterPO = GetString("U8.DZ.JA.btn120")    '���"
    strReportCheck = GetString("U8.DZ.JA.btn130")    '����"
    strClose = GetString("U8.DZ.JA.btn140")    '�ر�"
    strOpen = GetString("U8.DZ.JA.btn150")    '��"
    strConfirm = GetString("U8.DZ.JA.btn155")    '���"
    strCancelconfirm = GetString("U8.DZ.JA.btn160")    '����"
    strQueryConfirm = GetString("U8.DZ.JA.btn170")    '����"
    strPayment = GetString("U8.DZ.JA.btn180")    '�ָ�"
    strCancelpayment = GetString("U8.DZ.JA.btn190")    '����"
    strBatchBV = GetString("U8.DZ.JA.btn200")    '����"
    strSettle = GetString("U8.DZ.JA.btn210")    '����"
    strLocate = GetString("U8.DZ.JA.btn220")    '��λ"
    strLocateSet = GetString("U8.DZ.JA.btn230")    '��λ����"
    strFirst = GetString("U8.DZ.JA.btn240")    '����"
    strPrevious = GetString("U8.DZ.JA.btn250")    '����"
    strNext = GetString("U8.DZ.JA.btn260")    '����"
    strLast = GetString("U8.DZ.JA.btn270")    'ĩ��"
    strRefresh = GetString("U8.DZ.JA.btn280")    'ˢ��"
    strHelp = GetString("U8.DZ.JA.btn290")    '����"
    strFilter = GetString("U8.DZ.JA.btn300")    '����"
    strExit = GetString("U8.DZ.JA.btn310")    '�˳�"
    strColumn = GetString("U8.DZ.JA.btn320")    '��Ŀ"
    strSelectAll = GetString("U8.DZ.JA.btn330")    'ȫѡ"
    strUnSelectAll = GetString("U8.DZ.JA.btn340")    'ȫ��"
    strLock = GetString("U8.DZ.JA.btn350")    '����"
    strRLock = GetString("U8.DZ.JA.btn360")    '����"
    strBatchLock = GetString("U8.DZ.JA.btn370")    '����"
    strBatchRLock = GetString("U8.DZ.JA.btn380")    '����"
    strAcc = GetString("U8.DZ.JA.btn390")   '����"
    strVoucherDesign = GetString("U8.DZ.JA.btn540")
    strSaveVoucherDesign = GetString("U8.DZ.JA.btn550")
    strRefVoucher = GetString("U8.DZ.JA.btn620")
    gstrHelpText = GetString("U8.DZ.JA.btn290")   '"����"
    gstrHelpTip = GetString("U8.DZ.JA.btn290")   '"����"

    strLink = GetString("U8.DZ.JA.btn400")
    strFetchprice = GetString("U8.DZ.JA.btn410")  '"ȡ��"
    '������ chenliangc
    strSubmit = GetString("U8.DZ.JA.btn420")  '"�ύ"                 '�ύ
    strUnsubmit = GetString("U8.DZ.JA.btn430")    '"����"             '����
    strResubmit = GetString("U8.DZ.JA.btn440")    ' "�����ύ"             '�����ύ
    strViewVerify = GetString("U8.DZ.JA.btn170")  '"����"         '����

    strReferVoucher = GetString("U8.DZ.JA.btn460")    ' "����(һ��һ)"     '����(һ��һ)
    strReferVouchers = GetString("U8.DZ.JA.btn470")  '"����(���һ)"   '����(���һ)
    strCreateVoucher = GetString("U8.DZ.JA.btn480")    '"�Ƶ�"       '�Ƶ�
    strCreateSAVoucher = GetString("U8.DZ.JA.btn490")    ' "���۵���"
    strCreatePUVoucher = GetString("U8.DZ.JA.btn500")    ' "�ɹ�����"
    strCreateSCVoucher = GetString("U8.DZ.JA.btn510")    '"��浥��"
    strCreateAPVoucher = GetString("U8.DZ.JA.btn520")    '"Ӧ������"
    '
    strReverseSelection = GetString("U8.DZ.JA.btn530")    ' "��ѡ"
End Sub


'�ϲ�������
Public Sub ChangeOneFormTbr(Frm As Form, objTbl As Toolbar, objU8Tbl As Control)
    Dim rs As New ADODB.Recordset
    Dim strSql As String
    strSql = "select * from UFMeta_" + g_oLogin.cAcc_Id + "..AA_CustomerButton  where cFormKey='" & gstrCardNumber & "' and cLocaleID='zh-cn' order by cButtonID"
    Set rs = g_Conn.Execute(strSql)
'    Set objU8Tbl.Business = g_oBusiness
'    With objTbl
'
'        .Buttons(sKey_Print).Tag = g_oBusiness.createportaltoolbartag("print", "ICOMMON", "PortalToolbar")
'        .Buttons(sKey_Preview).Tag = g_oBusiness.createportaltoolbartag("print preview", "ICOMMON", "PortalToolbar")
'        .Buttons(sKey_Output).Tag = g_oBusiness.createportaltoolbartag("Output", "ICOMMON", "PortalToolbar")
'
'        .Buttons(sKey_First).Tag = g_oBusiness.createportaltoolbartag("first page", "ICOMMON", "PortalToolbar")
'        .Buttons(sKey_Previous).Tag = g_oBusiness.createportaltoolbartag("previous page", "ICOMMON", "PortalToolbar")
'        .Buttons(sKey_Next).Tag = g_oBusiness.createportaltoolbartag("next page", "ICOMMON", "PortalToolbar")
'        .Buttons(sKey_Last).Tag = g_oBusiness.createportaltoolbartag("last page", "ICOMMON", "PortalToolbar")
'        .Buttons(sKey_RefVoucher).Tag = g_oBusiness.createportaltoolbartag("query", "ICOMMON", "PortalToolbar")  'zhangwchb
'
'        .Buttons(sKey_Add).Tag = g_oBusiness.createportaltoolbartag("add", "IEDIT", "PortalToolbar")
'        .Buttons(sKey_Modify).Tag = g_oBusiness.createportaltoolbartag("modify", "IEDIT", "PortalToolbar")
'        .Buttons(sKey_Delete).Tag = g_oBusiness.createportaltoolbartag("delete", "IEDIT", "PortalToolbar")
'
'
'        .Buttons(sKey_ReferVoucher).Tag = g_oBusiness.createportaltoolbartag("create", "IDEAL", "PortalToolbar")
'        .Buttons(sKey_Copy).Tag = g_oBusiness.createportaltoolbartag("Copy", "IEDIT", "PortalToolbar")
'        .Buttons(sKey_Save).Tag = g_oBusiness.createportaltoolbartag("Save", "IEDIT", "PortalToolbar")
'        .Buttons(sKey_Discard).Tag = g_oBusiness.createportaltoolbartag("back", "IEDIT", "PortalToolbar")
'
'
'        .Buttons(sKey_Confirm).Tag = g_oBusiness.createportaltoolbartag("column", "ICOMMON", "PortalToolbar")
'        .Buttons(sKey_Cancelconfirm).Tag = g_oBusiness.createportaltoolbartag("Unapprove", "ICOMMON", "PortalToolbar")
'        .Buttons(sKey_Open).Tag = g_oBusiness.createportaltoolbartag("Open", "ICOMMON", "PortalToolbar")
'        .Buttons(sKey_Close).Tag = g_oBusiness.createportaltoolbartag("Close", "ICOMMON", "PortalToolbar")
'
'
'        .Buttons(sKey_Locate).Tag = g_oBusiness.createportaltoolbartag("location", "ICOMMON", "PortalToolbar")
'        .Buttons(sKey_Refresh).Tag = g_oBusiness.createportaltoolbartag("refresh", "ICOMMON", "PortalToolbar")
'        .Buttons(gstrHelpCode).Tag = g_oBusiness.createportaltoolbartag("help", "ICOMMON", "PortalToolbar")
'
'
'
'
'        .Buttons(sKey_Addrecord).Tag = g_oBusiness.createportaltoolbartag("add a row", "IEDIT", "PortalToolbar")
'        .Buttons(sKey_Deleterecord).Tag = g_oBusiness.createportaltoolbartag("delete row", "IEDIT", "PortalToolbar")
'        .Buttons(sKey_Fetchprice).Tag = g_oBusiness.createportaltoolbartag("price", "IEDIT", "PortalToolbar")  'ȡ��
'        .Buttons(sKey_Acc).Tag = g_oBusiness.createportaltoolbartag("accessories", "IEDIT", "PortalToolbar")    '����
'
'        '������ chenliangc
'        .Buttons(sKey_Submit).Tag = g_oBusiness.createportaltoolbartag("Submit", "ICOMMON", "PortalToolbar")
'        .Buttons(sKey_Unsubmit).Tag = g_oBusiness.createportaltoolbartag("recover", "ICOMMON", "PortalToolbar")
'        .Buttons(sKey_Resubmit).Tag = g_oBusiness.createportaltoolbartag("Submit", "ICOMMON", "PortalToolbar")    '
'        .Buttons(sKey_ViewVerify).Tag = g_oBusiness.createportaltoolbartag("Relate query", "ICOMMON", "PortalToolbar")    '��
'        '����
'        .Buttons(sKey_ReferVoucher).Tag = g_oBusiness.createportaltoolbartag("create", "IEDIT", "PortalToolbar")
'        .Buttons(sKey_CreateVoucher).Tag = g_oBusiness.createportaltoolbartag("creating", "ICOMMON", "PortalToolbar")
'        If Not rs.EOF Then
'            Do While Not rs.EOF
'                .Buttons(CStr(rs!cButtonkey)).Tag = g_oBusiness.createportaltoolbartag(CStr(rs!cImage), CStr(rs!cGroup), "PortalToolbar")
'                rs.MoveNext
'            Loop
'        End If
'        'U810.0���䣬 ��Ӹ�ʽ���úͱ��水ť
'        .Buttons(sKey_VoucherDesign).Tag = g_oBusiness.createportaltoolbartag("format", "IEDIT", "PortalToolbar")
'        .Buttons(sKey_SaveVoucherDesign).Tag = g_oBusiness.createportaltoolbartag("format", "IEDIT", "PortalToolbar")
'
'    End With
    'InitToolBar�������Ѿ����ù�
    'objU8Tbl.SetToolbar objTbl
    objU8Tbl.SetDisplayStyle 0    'TextOnly
    objTbl.Visible = False
    objU8Tbl.Visible = True
    objU8Tbl.Left = objTbl.Left
    objU8Tbl.Top = objTbl.Top
    objU8Tbl.Width = Frm.Width - 6 * Screen.TwipsPerPixelX
    objU8Tbl.Height = objTbl.Height
End Sub

'------------------------------------------------------------
'��ʼ���������ؼ�
'------------------------------------------------------------
Public Sub Init_Toolbar(tlbObj As Toolbar)
    Dim rs As New ADODB.Recordset
    Dim strSql As String
    strSql = "select * from UFMeta_" + g_oLogin.cAcc_Id + "..AA_CustomerButton  where cFormKey='" & gstrCardNumber & "' and cLocaleID='zh-cn' order by cButtonID"
    Set rs = g_Conn.Execute(strSql)


    With tlbObj.Buttons
        .Clear

        .Add , sKey_Print, strPrint
        .Item(sKey_Print).ToolTipText = strPrint + "Ctrl+P"

        .Add , sKey_Preview, strPreview
        .Item(sKey_Preview).ToolTipText = strPreview + "Ctrl+W"

        .Add , sKey_Output, strOutput
        .Item(sKey_Output).ToolTipText = strOutput + "Alt+E"

        .Add , "S1", , tbrSeparator


        .Add , sKey_Confirm, strConfirm
        .Item(sKey_Confirm).ToolTipText = strConfirm + " Ctrl+U"

        .Add , sKey_Cancelconfirm, strCancelconfirm
        .Item(sKey_Cancelconfirm).ToolTipText = strCancelconfirm + "Alt+U"
        
         .Add , sKey_ViewVerify, strViewVerify
        .Item(sKey_ViewVerify).ToolTipText = strViewVerify
        
                '������ chenliangc
        .Add , sKey_Submit, strSubmit
        .Item(sKey_Submit).ToolTipText = strSubmit + " Ctrl+J"

        .Add , sKey_Resubmit, strResubmit
        .Item(sKey_Resubmit).ToolTipText = strResubmit

        .Add , sKey_Unsubmit, strUnsubmit
        .Item(sKey_Unsubmit).ToolTipText = strUnsubmit + "Alt+J"
     
           .Add , sKey_Open, strOpen
        .Item(sKey_Open).ToolTipText = strOpen + "Alt+O"

        .Add , sKey_Close, strClose
        .Item(sKey_Close).ToolTipText = strClose + "Alt+C"
        
          .Add , "S2", , tbrSeparator




        .Add , sKey_First, strFirst
        .Item(sKey_First).ToolTipText = strFirst + "Alt+PageUp"

        .Add , sKey_Previous, strPrevious
        .Item(sKey_Previous).ToolTipText = strPrevious + "PageUp"

        .Add , sKey_Next, strNext
        .Item(sKey_Next).ToolTipText = strNext + " PageDown"

        .Add , sKey_Last, strLast
        .Item(sKey_Last).ToolTipText = strLast + "Alt+PageDown"



        .Add , "S3", , tbrSeparator


        .Add , sKey_Locate, strLocate
        .Item(sKey_Locate).ToolTipText = strLocate + "Ctrl+F3"
        
        .Add , sKey_RefVoucher, strRefVoucher 'zhangwchb
        .Item(sKey_RefVoucher).ToolTipText = strRefVoucher

        .Add , sKey_Refresh, strRefresh
        .Item(sKey_Refresh).ToolTipText = strRefresh + "Ctrl+R"
        
  
        .Add , gstrHelpCode, gstrHelpText
        .Item(gstrHelpCode).ToolTipText = gstrHelpText + " F1"


        .Add , "S4", , tbrSeparator


        .Add , sKey_Add, strAdd
        .Item(sKey_Add).ToolTipText = strAdd + " F5"
        .Item(sKey_Add).Style = tbrDropdown
        .Item(sKey_Add).ButtonMenus.Add , sKey_Add2, strAdd2
        .Item(sKey_Add).ButtonMenus.Add , sKey_Add1, strAdd1
       

        .Add , sKey_Modify, strModify
        .Item(sKey_Modify).ToolTipText = strModify + " F8"

        .Add , sKey_Delete, strdelete
        .Item(sKey_Delete).ToolTipText = strdelete + " Delete"

        '���� chenliangc
        .Add , sKey_ReferVoucher, strReferVoucher
        .Item(sKey_ReferVoucher).ToolTipText = strReferVoucher
        .Item(sKey_ReferVoucher).Style = tbrDropdown
        Call .Item(sKey_ReferVoucher).ButtonMenus.Add(, sKey_ReferVouchers, strReferVouchers)

        '�Ƶ� chenliangc
        .Add , sKey_CreateVoucher, strCreateVoucher
        .Item(sKey_CreateVoucher).ToolTipText = strCreateVoucher
        .Item(sKey_CreateVoucher).Style = tbrDropdown

        .Item(sKey_CreateVoucher).ButtonMenus.Add , sKey_CreateSAVoucher, strCreateSAVoucher
        .Item(sKey_CreateVoucher).ButtonMenus.Add , sKey_CreatePUVoucher, strCreatePUVoucher
        .Item(sKey_CreateVoucher).ButtonMenus.Add , sKey_CreateSCVoucher, strCreateSCVoucher
        .Item(sKey_CreateVoucher).ButtonMenus.Add , sKey_CreateAPVoucher, strCreateAPVoucher

        .Add , sKey_Copy, strCopy
        .Item(sKey_Copy).ToolTipText = strCopy + "Ctrl+F5"

        .Add , sKey_Save, strSave
        .Item(sKey_Save).ToolTipText = strSave + " F6"

        .Add , sKey_Discard, strDiscard
        .Item(sKey_Discard).ToolTipText = strDiscard + " Ctrl+Z"


        .Add , "S5", , tbrSeparator


        .Add , sKey_Addrecord, strAddrecord
        .Item(sKey_Addrecord).ToolTipText = strAddrecord + "Ctrl+N"

        .Add , sKey_Deleterecord, strDeleterecord
        .Item(sKey_Deleterecord).ToolTipText = strDeleterecord + "Ctrl+D"
        
         '����
        .Add , sKey_Acc, strAcc
        .Item(sKey_Acc).ToolTipText = strAcc

        'ȡ��
        With .Add(, sKey_Fetchprice, strFetchprice, tbrDropdown)
            Call .ButtonMenus.Add(, "rowprice", GetString("U8.DZ.JA.btn770"))
            Call .ButtonMenus.Add(, "allprice", GetString("U8.DZ.JA.btn780"))
        End With
        'U810 ����  ��� ���ݸ�ʽ���úͱ��水ť  2011/03/04   LEW
        .Add , sKey_VoucherDesign, strVoucherDesign                  '���ݸ�ʽ���ð�ť
        .Item(sKey_VoucherDesign).ToolTipText = strVoucherDesign


        .Add , sKey_SaveVoucherDesign, strSaveVoucherDesign          '��ʽ���水ť
        .Item(sKey_SaveVoucherDesign).ToolTipText = strSaveVoucherDesign

       
        If Not rs.EOF Then
            Do While Not rs.EOF
                .Add , CStr(rs!cButtonkey), CStr(rs!cCaption)
                '               .Item(sKey_Deleterecord).ToolTipText = Rs!cCaption + Rs!cHotKey
                .Item(CStr(rs!cButtonkey)).ToolTipText = rs!cCaption + rs!cHotKey
                rs.MoveNext
            Loop
        End If


    End With
End Sub

'���� ��ذ�ť ��ʾ���
Public Sub SetWFControlBrnsList(login As clsLogin, myConn As ADODB.Connection, Toolbar As Object, UFToolbar As Object, cardnumber As String)
    Dim rstfilter As String

    '   .Buttons(sKey_Batchprint).Tag = goBusiness.createportaltoolbartag("Print", "ICOMMON", "PortalToolbar")
    '   .Buttons(strKSelectAll).Tag = goBusiness.createportaltoolbartag("Select All", "ICOMMON", "PortalToolbar")
    '   .Buttons(sKey_ReverseSelection).Tag = goBusiness.createportaltoolbartag("Revise", "ICOMMON", "PortalToolbar")
    '   .Buttons(strKUnSelectAll).Tag = goBusiness.createportaltoolbartag("Select none", "ICOMMON", "PortalToolbar")
    '   .Buttons(sKey_Link).Tag = goBusiness.createportaltoolbartag("relate query", "ICOMMON", "PortalToolbar")
    '   .Buttons(sKey_Confirm).Tag = goBusiness.createportaltoolbartag("Approve", "ICOMMON", "PortalToolbar")
    '   .Buttons(sKey_Cancelconfirm).Tag = goBusiness.createportaltoolbartag("Unapprove", "ICOMMON", "PortalToolbar")
    '   .Buttons(sKey_Open).Tag = goBusiness.createportaltoolbartag("Open", "ICOMMON", "PortalToolbar")
    '   .Buttons(sKey_Close).Tag = goBusiness.createportaltoolbartag("Close", "ICOMMON", "PortalToolbar")


    Toolbar.Buttons(sKey_CreateVoucher).Visible = False
    Toolbar.Buttons(sKey_Close).Visible = False
    Toolbar.Buttons(sKey_Open).Visible = False
    '    Toolbar.Buttons(sKey_Cancelconfirm).Visible = False
    '    Toolbar.Buttons(sKey_Confirm).Visible = False
    '    Toolbar.Buttons(strKUnSelectAll).Visible = False
    '    Toolbar.Buttons(strKSelectAll).Visible = False
    '    Toolbar.Buttons(sKey_ReverseSelection).Visible = False

    UFToolbar.RefreshVisible
End Sub

'�ϲ�������
Public Sub ChangeOneFormTbrlist(Frm As Form, objTbl As Toolbar, objU8Tbl As Control)
    Set objU8Tbl.Business = goBusiness
    With objTbl

        .Buttons(sKey_Print).Tag = goBusiness.createportaltoolbartag("print", "ICOMMON", "PortalToolbar")
        .Buttons(sKey_Preview).Tag = goBusiness.createportaltoolbartag("print preview", "ICOMMON", "PortalToolbar")
        .Buttons(sKey_Output).Tag = goBusiness.createportaltoolbartag("Output", "ICOMMON", "PortalToolbar")

        .Buttons(sKey_Locate).Tag = goBusiness.createportaltoolbartag("Location", "ICOMMON", "PortalToolbar")
        .Buttons(strKfilter).Tag = goBusiness.createportaltoolbartag("filter", "ICOMMON", "PortalToolbar")
        .Buttons(sKey_Column).Tag = goBusiness.createportaltoolbartag("column", "ICOMMON", "PortalToolbar")
        .Buttons(sKey_Batchprint).Tag = goBusiness.createportaltoolbartag("Print", "ICOMMON", "PortalToolbar")


        .Buttons(strKSelectAll).Tag = goBusiness.createportaltoolbartag("Select All", "ICOMMON", "PortalToolbar")
        .Buttons(sKey_ReverseSelection).Tag = goBusiness.createportaltoolbartag("Revise", "ICOMMON", "PortalToolbar")
        .Buttons(strKUnSelectAll).Tag = goBusiness.createportaltoolbartag("Select none", "ICOMMON", "PortalToolbar")
       ' .Buttons(sKey_Link).Tag = goBusiness.createportaltoolbartag("relate query", "ICOMMON", "PortalToolbar")


        .Buttons(sKey_Confirm).Tag = goBusiness.createportaltoolbartag("Approve", "ICOMMON", "PortalToolbar")
        .Buttons(sKey_Cancelconfirm).Tag = goBusiness.createportaltoolbartag("Unapprove", "ICOMMON", "PortalToolbar")
        .Buttons(sKey_Open).Tag = goBusiness.createportaltoolbartag("Open", "ICOMMON", "PortalToolbar")
        .Buttons(sKey_Close).Tag = goBusiness.createportaltoolbartag("Close", "ICOMMON", "PortalToolbar")
        '�Ƶ� chenliangc
        .Buttons(sKey_CreateVoucher).Tag = goBusiness.createportaltoolbartag("creating", "ICOMMON", "PortalToolbar")
        .Buttons(sKey_Refresh).Tag = goBusiness.createportaltoolbartag("refresh", "ICOMMON", "PortalToolbar")
        .Buttons(gstrHelpCode).Tag = goBusiness.createportaltoolbartag("help", "ICOMMON", "PortalToolbar")

    End With
    'InitToolBar�������Ѿ����ù�
    'objU8Tbl.SetToolbar objTbl
    objU8Tbl.SetDisplayStyle 0    'TextOnly
    objTbl.Visible = False
    objU8Tbl.Visible = True
    objU8Tbl.Left = objTbl.Left
    objU8Tbl.Top = objTbl.Top
    objU8Tbl.Width = Frm.Width - 6 * Screen.TwipsPerPixelX
    objU8Tbl.Height = objTbl.Height
End Sub

'------------------------------------------------------------
'��ʼ���������ؼ�
'------------------------------------------------------------
Public Sub Init_Toolbarlist(tlbObj As Toolbar)

    With tlbObj.Buttons
        .Clear

        .Add , sKey_Print, strPrint
        .Item(sKey_Print).ToolTipText = strPrint

        .Add , sKey_Preview, strPreview
        .Item(sKey_Preview).ToolTipText = strPreview

        .Add , sKey_Output, strOutput
        .Item(sKey_Output).ToolTipText = strOutput

        '        .Add , "S1", , tbrSeparator

        .Add , strKfilter, strFilter
        .Item(strKfilter).ToolTipText = strFilter
        .Item(strKfilter).Visible = False
        .Add , sKey_Locate, strLocate
        .Item(sKey_Locate).ToolTipText = strLocate

'        .Add , sKey_Link, strLink
'        .Item(sKey_Link).ToolTipText = strLink


        .Add , sKey_Column, strColumn
        .Item(sKey_Column).ToolTipText = strColumn

        .Add , sKey_Batchprint, strBatchprint
        .Item(sKey_Batchprint).ToolTipText = strBatchprint





        .Add , strKSelectAll, strSelectAll
        .Item(strKSelectAll).ToolTipText = strSelectAll
        .Item(strKSelectAll).Visible = False

        .Add , sKey_ReverseSelection, strReverseSelection
        .Item(sKey_ReverseSelection).ToolTipText = strReverseSelection
        .Item(sKey_ReverseSelection).Visible = False
        
        .Add , strKUnSelectAll, strUnSelectAll
        .Item(strKUnSelectAll).ToolTipText = strUnSelectAll
        .Item(strKUnSelectAll).Visible = False



        .Add , sKey_Confirm, strBatchVeri
        .Item(sKey_Confirm).ToolTipText = strBatchVeri

        .Add , sKey_Cancelconfirm, strBatchUnVeri
        .Item(sKey_Cancelconfirm).ToolTipText = strBatchUnVeri

        .Add , sKey_Open, strBatchOpen
        .Item(sKey_Open).ToolTipText = strBatchOpen

        .Add , sKey_Close, strBatchClose
        .Item(sKey_Close).ToolTipText = strBatchClose

        '         .Add , "S3", , tbrSeparator



        .Add , sKey_Refresh, strRefresh
        .Item(sKey_Refresh).ToolTipText = strRefresh
        .Item(sKey_Refresh).Visible = False
        
        .Add , gstrHelpCode, gstrHelpText
        .Item(gstrHelpCode).ToolTipText = gstrHelpText + " F1"


        '        .Add , "S4", , tbrSeparator
        '�Ƶ� chenliangc
        .Add , sKey_CreateVoucher, strCreateVoucher, tbrDropdown
        .Item(sKey_CreateVoucher).ToolTipText = strCreateVoucher
        .Item(sKey_CreateVoucher).ButtonMenus.Add , sKey_CreateSAVoucher, strCreateSAVoucher
        .Item(sKey_CreateVoucher).ButtonMenus.Add , sKey_CreatePUVoucher, strCreatePUVoucher
        .Item(sKey_CreateVoucher).ButtonMenus.Add , sKey_CreateSCVoucher, strCreateSCVoucher
        .Item(sKey_CreateVoucher).ButtonMenus.Add , sKey_CreateAPVoucher, strCreateAPVoucher

        .Item(sKey_Batchprint).Visible = False   '�ݲ�֧������
    End With
End Sub


'������״̬
'������״̬
Public Sub SetCtlStyle(Frm As Form, Voucher As Object, Toolbar As Toolbar, UFToolbar As UFToolbar, mOpStatus As OpStatus)
    On Error Resume Next
    Dim sql As String
    Dim strSql As String
    Dim rs As New ADODB.Recordset

    With Toolbar
        Select Case mOpStatus
            '����
        Case ADD_MAIN
            .Buttons(sKey_Print).Enabled = False  '��ӡ
            .Buttons(sKey_Preview).Enabled = False    'Ԥ��
            .Buttons(sKey_Output).Enabled = False    '���

            .Buttons(sKey_First).Enabled = False    '��ҳ
            .Buttons(sKey_Previous).Enabled = False    '��һҳ
            .Buttons(sKey_Next).Enabled = False    '��һҳ
            .Buttons(sKey_Last).Enabled = False    'ĩҳ

            .Buttons(sKey_Add).Enabled = False    '����
            .Buttons(sKey_Modify).Enabled = False    '�޸�
            .Buttons(sKey_Delete).Enabled = False    'ɾ��


            .Buttons(sKey_ReferVoucher).Enabled = True    '����
            .Buttons(sKey_Copy).Enabled = False    '����
            .Buttons(sKey_Save).Enabled = True    '����
            .Buttons(sKey_Discard).Enabled = True    '����
            .Buttons(sKey_Fetchprice).Enabled = True    'ȡ��

            .Buttons(sKey_Submit).Enabled = False  '�ύ
            .Buttons(sKey_Resubmit).Enabled = False    '�����ύ
            .Buttons(sKey_Unsubmit).Enabled = False   '����
            .Buttons(sKey_ViewVerify).Enabled = False    '����
            .Buttons(sKey_Confirm).Enabled = False    '���
            .Buttons(sKey_Cancelconfirm).Enabled = False    '����
            .Buttons(sKey_Return).Enabled = False   '�黹
            .Buttons(sKey_CreateVoucher).Enabled = False    '�Ƶ�
            .Buttons(sKey_Open).Enabled = False    '��
            .Buttons(sKey_Close).Enabled = False    '�ر�

            .Buttons(sKey_Locate).Enabled = False    '��λ
            .Buttons(sKey_Refresh).Enabled = False    'ˢ��
            .Buttons(gstrHelpCode).Enabled = True    '����
            .Buttons(sKey_VoucherDesign).Enabled = False

            .Buttons(sKey_Addrecord).Enabled = True    '����
            .Buttons(sKey_Deleterecord).Enabled = True    'ɾ��
            .Buttons(sKey_Acc).Enabled = True    '����
            .Buttons("Prorefer1").Enabled = False
            .Buttons("Prorefer2").Enabled = False
            .Buttons("Prorefer3").Enabled = False
            
            




            Voucher.VoucherStatus = VSeAddMode

            Frm.ComTemplatePRN.Visible = False
            Frm.ComTemplateShow.Visible = True
            Frm.LblTemplate.Caption = GetString("U8.DZ.JA.Res020")
            Frm.LblTemplate.Visible = True

            '�޸�
        Case MODIFY_MAIN

            .Buttons(sKey_Print).Enabled = False    '��ӡ
            .Buttons(sKey_Preview).Enabled = False    'Ԥ��
            .Buttons(sKey_Output).Enabled = False    '���
            .Buttons("Prorefer").Enabled = False

            .Buttons(sKey_First).Enabled = False    '��ҳ
            .Buttons(sKey_Previous).Enabled = False    '��һҳ
            .Buttons(sKey_Next).Enabled = False    '��һҳ
            .Buttons(sKey_Last).Enabled = False    'ĩҳ

            .Buttons(sKey_Add).Enabled = False    '����
            .Buttons(sKey_Modify).Enabled = False    '�޸�
            .Buttons(sKey_Delete).Enabled = False    'ɾ��
            .Buttons(sKey_Fetchprice).Enabled = True    'ȡ��

            .Buttons(sKey_ReferVoucher).Enabled = False    '����
            .Buttons(sKey_Copy).Enabled = False    '����
            .Buttons(sKey_Save).Enabled = True    '����
            .Buttons(sKey_Discard).Enabled = True    '����

            .Buttons(sKey_Submit).Enabled = False  '�ύ
            .Buttons(sKey_Resubmit).Enabled = False    '�����ύ
            .Buttons(sKey_Unsubmit).Enabled = False   '����
            .Buttons(sKey_ViewVerify).Enabled = False    '����
            .Buttons(sKey_Confirm).Enabled = False    '���
            .Buttons(sKey_Cancelconfirm).Enabled = False    '����
            .Buttons(sKey_Return).Enabled = False   '�黹
            .Buttons(sKey_CreateVoucher).Enabled = False    '�Ƶ�
            .Buttons(sKey_Confirm).Enabled = False    '���
            .Buttons(sKey_Cancelconfirm).Enabled = False    '����
            .Buttons(sKey_Open).Enabled = False    '��
            .Buttons(sKey_Close).Enabled = False    '�ر�
            
            .Buttons("Prorefer1").Enabled = False
            .Buttons("Prorefer2").Enabled = False
            .Buttons("Prorefer3").Enabled = False


            .Buttons(sKey_Locate).Enabled = False    '��λ
            .Buttons(sKey_Refresh).Enabled = False    'ˢ��
            .Buttons(gstrHelpCode).Enabled = True    '����


            .Buttons(sKey_Addrecord).Enabled = True    '����
            .Buttons(sKey_Deleterecord).Enabled = True    'ɾ��
            .Buttons(sKey_VoucherDesign).Enabled = False
            .Buttons(sKey_Acc).Enabled = True    '����

            Voucher.VoucherStatus = VSeEditMode

            Frm.ComTemplatePRN.Visible = False
            Frm.ComTemplateShow.Visible = True
            Frm.LblTemplate.Caption = GetString("U8.DZ.JA.Res020")
            Frm.LblTemplate.Visible = True

            '��ʾ
        Case SHOW_ALL

            .Buttons(sKey_Print).Enabled = True    '��ӡ
            .Buttons(sKey_Preview).Enabled = True    'Ԥ��
            .Buttons(sKey_Output).Enabled = True    '���

            .Buttons(sKey_First).Enabled = True    '��ҳ
            .Buttons(sKey_Previous).Enabled = True    '��һҳ
            .Buttons(sKey_Next).Enabled = True    '��һҳ
            .Buttons(sKey_Last).Enabled = True    'ĩҳ

            .Buttons(sKey_Add).Enabled = True    '����
            .Buttons(sKey_Modify).Enabled = True    '�޸�
            .Buttons(sKey_Delete).Enabled = True    'ɾ��
            .Buttons("Prorefer").Enabled = True

            .Buttons(sKey_ReferVoucher).Enabled = False    '����
            .Buttons(sKey_Copy).Enabled = True    '����
            .Buttons(sKey_Save).Enabled = False    '����
            .Buttons(sKey_Discard).Enabled = False    '����
            .Buttons(sKey_Fetchprice).Enabled = False    'ȡ��
            .Buttons(sKey_VoucherDesign).Enabled = True

            .Buttons(sKey_Confirm).Enabled = True    '���
            .Buttons(sKey_Cancelconfirm).Enabled = True    '����
            
            .Buttons(sKey_Locate).Enabled = True    '��λ
            .Buttons(sKey_Refresh).Enabled = True    'ˢ��
            .Buttons(gstrHelpCode).Enabled = True    '����


            .Buttons(sKey_Addrecord).Enabled = False    '����
            .Buttons(sKey_Deleterecord).Enabled = False    'ɾ��
            .Buttons(sKey_Acc).Enabled = False    '����
           ' .Buttons(sKey_Acc).Visible = False
           .Buttons("Prorefer1").Enabled = True
            .Buttons("Prorefer2").Enabled = True
            .Buttons("Prorefer3").Enabled = True



            '��ҳ��ť
            If pageCount <= 1 Then
                .Buttons(sKey_First).Enabled = False    '��ҳ
                .Buttons(sKey_Previous).Enabled = False    '��һҳ
                .Buttons(sKey_Next).Enabled = False    '��һҳ
                .Buttons(sKey_Last).Enabled = False    'ĩҳ
            ElseIf PageCurrent = pageCount Then
                .Buttons(sKey_First).Enabled = True    '��ҳ
                .Buttons(sKey_Previous).Enabled = True    '��һҳ
                .Buttons(sKey_Next).Enabled = False    '��һҳ
                .Buttons(sKey_Last).Enabled = False    'ĩҳ
            ElseIf PageCurrent = 1 Then
                .Buttons(sKey_First).Enabled = False    '��ҳ
                .Buttons(sKey_Previous).Enabled = False    '��һҳ
                .Buttons(sKey_Next).Enabled = True    '��һҳ
                .Buttons(sKey_Last).Enabled = True    'ĩҳ
            Else
                .Buttons(sKey_First).Enabled = True    '��ҳ
                .Buttons(sKey_Previous).Enabled = True    '��һҳ
                .Buttons(sKey_Next).Enabled = True    '��һҳ
                .Buttons(sKey_Last).Enabled = True    'ĩҳ

            End If


            'modify by chenliangc ��ӹ�������ť��ʾ

            '           SQL = "SELECT iStatus,iswfcontrolled,iverifystate,DownStreamcode,case when isnull(closeuser,N'')=N'' then 1 else 0 end as Closed FROM HY_DZ_BorrowOutChange WHERE ID=" & lngVoucherID
            sql = "SELECT iStatus,iswfcontrolled,iverifystate,case when isnull(closeuser,N'')=N'' then 1 else 0 end as Closed FROM " & MainTable & " WHERE ID=" & lngVoucherID
            Set rs = New ADODB.Recordset
            rs.Open sql, g_Conn, 1, 1
            If Not rs.EOF Then
            
                .Buttons(sKey_ViewVerify).Enabled = True    '����
                
                Select Case rs("iStatus")
                    '����
                Case 1
                    .Buttons(sKey_Open).Enabled = False    '��
                    .Buttons(sKey_Close).Enabled = False    '�ر�
                    .Buttons(sKey_CreateVoucher).Enabled = False    '�Ƶ�
                    .Buttons(sKey_ReferVoucher).Enabled = False    '����
                    .Buttons(sKey_Cancelconfirm).Enabled = False    '����
                    .Buttons(sKey_Return).Enabled = False   '�黹
                    If Not CBool(Null2Something(rs("iswfcontrolled"), "0")) Then    '�����빤����
                        .Buttons(sKey_Submit).Enabled = False  '�ύ
                        .Buttons(sKey_Resubmit).Enabled = False    '�����ύ
                        .Buttons(sKey_Unsubmit).Enabled = False   '����
                        
                    ElseIf CInt(Null2Something(rs("iverifystate"), 0)) <= 0 Then  '���������Ƶ�δ�ύ
                        .Buttons(sKey_Submit).Enabled = True  '�ύ
                        .Buttons(sKey_Resubmit).Enabled = False    '�����ύ
                        .Buttons(sKey_Unsubmit).Enabled = False   '����
                        .Buttons(sKey_Confirm).Enabled = False    '���
                        .Buttons(sKey_Cancelconfirm).Enabled = False    '����
                        .Buttons(sKey_Return).Enabled = False   '�黹
                    ElseIf CInt(rs("iverifystate")) = 1 Then   '�������������ύ
                        .Buttons(sKey_Submit).Enabled = False  '�ύ
                        .Buttons(sKey_Resubmit).Enabled = True    '�����ύ
                        .Buttons(sKey_Unsubmit).Enabled = True   '����
                        .Buttons(sKey_Confirm).Enabled = True    '���
                        .Buttons(sKey_Cancelconfirm).Enabled = True    '����
                        .Buttons(sKey_Return).Enabled = True   '�黹
                        .Buttons(sKey_Delete).Enabled = False    'ɾ��
                    End If

                    '��� ����˲�Ϊ��
                Case 2
           
                    .Buttons(sKey_Open).Enabled = False    '��
                    .Buttons(sKey_Close).Enabled = True    '�ر�
                    .Buttons(sKey_Submit).Enabled = False  '�ύ
                    .Buttons(sKey_Resubmit).Enabled = True    '�����ύ
                    .Buttons(sKey_Unsubmit).Enabled = True   '����
                    .Buttons(sKey_Modify).Enabled = False    '�޸�
                    .Buttons(sKey_Delete).Enabled = False    'ɾ��
                    .Buttons(sKey_Confirm).Enabled = False    '���
                    .Buttons(sKey_Submit).Enabled = False    '�ύ
                    .Buttons(sKey_Unsubmit).Enabled = True    '����
                    .Buttons(sKey_Resubmit).Enabled = True    '�����ύ
                    .Buttons(sKey_Cancelconfirm).Enabled = True    '����
                    .Buttons(sKey_CreateVoucher).Enabled = True    '�Ƶ�
                    .Buttons(sKey_Return).Enabled = True   '�黹
                    
                  If Not CBool(Null2Something(rs("iswfcontrolled"), "0")) Then    '�����빤����
                        .Buttons(sKey_Submit).Enabled = False  '�ύ
                        .Buttons(sKey_Resubmit).Enabled = False    '�����ύ
                        .Buttons(sKey_Unsubmit).Enabled = False   '����
'                        .Buttons(sKey_ViewVerify).Enabled = False    '����
                    End If

                    '����
                Case 3
                    .Buttons(sKey_Submit).Enabled = False  '�ύ
                    .Buttons(sKey_Resubmit).Enabled = False    '�����ύ
                    .Buttons(sKey_Unsubmit).Enabled = False   '����
                    .Buttons(sKey_Modify).Enabled = False    '�޸�
                    .Buttons(sKey_Delete).Enabled = False    'ɾ��
                    .Buttons(sKey_Confirm).Enabled = False    '���
                    .Buttons(sKey_Cancelconfirm).Enabled = False    '����
                    .Buttons(sKey_Open).Enabled = False    '��
                    .Buttons(sKey_Close).Enabled = True    '�ر�
                    .Buttons(sKey_Submit).Enabled = False    '�ύ
                    .Buttons(sKey_Unsubmit).Enabled = False    '����
                    .Buttons(sKey_Resubmit).Enabled = False    '�����ύ
                    .Buttons(sKey_CreateVoucher).Enabled = False    '�Ƶ�
                    .Buttons(sKey_Return).Enabled = True   '�黹


                    '�ر�
                Case 4
                    .Buttons(sKey_Submit).Enabled = False  '�ύ
                    .Buttons(sKey_Resubmit).Enabled = False    '�����ύ
                    .Buttons(sKey_Unsubmit).Enabled = False   '����
                    .Buttons(sKey_Modify).Enabled = False    '�޸�
                    .Buttons(sKey_Delete).Enabled = False    'ɾ��
                    .Buttons(sKey_ReferVoucher).Enabled = False    '����
                    .Buttons(sKey_Confirm).Enabled = False    '���
                    .Buttons(sKey_Cancelconfirm).Enabled = False    '����
                    .Buttons(sKey_Open).Enabled = True    '��
                    .Buttons(sKey_Close).Enabled = False    '�ر�
                    .Buttons(sKey_Submit).Enabled = False    '�ύ
                    .Buttons(sKey_Unsubmit).Enabled = False    '����
                    .Buttons(sKey_Resubmit).Enabled = False    '�����ύ
                    .Buttons(sKey_CreateVoucher).Enabled = False    '�Ƶ�
                    .Buttons(sKey_Return).Enabled = False   '�黹


                End Select
                 
                        
                'û�м�¼
            Else
                .Buttons(sKey_Print).Enabled = False    '��ӡ
                .Buttons(sKey_Preview).Enabled = False    'Ԥ��
                .Buttons(sKey_Output).Enabled = False    '���

                .Buttons(sKey_First).Enabled = False    '��ҳ
                .Buttons(sKey_Previous).Enabled = False    '��һҳ
                .Buttons(sKey_Next).Enabled = False    '��һҳ
                If pageCount >= 1 Then
                    .Buttons(sKey_Last).Enabled = True    'ĩҳ
                Else
                    .Buttons(sKey_Last).Enabled = False    'ĩҳ
                End If
                .Buttons(sKey_Add).Enabled = True    '����
                .Buttons(sKey_Modify).Enabled = False    '�޸�
                .Buttons(sKey_Delete).Enabled = False    'ɾ��


                .Buttons(sKey_ReferVoucher).Enabled = False    '����
                .Buttons(sKey_CreateVoucher).Enabled = False    '�Ƶ�

                .Buttons(sKey_Copy).Enabled = False    '����
                .Buttons(sKey_Save).Enabled = False    '����
                .Buttons(sKey_Discard).Enabled = False    '����


                .Buttons(sKey_Submit).Enabled = False    '�ύ
                .Buttons(sKey_Unsubmit).Enabled = False    '����
                .Buttons(sKey_Resubmit).Enabled = False    '�����ύ
                .Buttons(sKey_ViewVerify).Enabled = False    '����
                .Buttons(sKey_Confirm).Enabled = False    '���
                .Buttons(sKey_Cancelconfirm).Enabled = False    '����
                .Buttons(sKey_Open).Enabled = False    '��
                .Buttons(sKey_Close).Enabled = False    '�ر�
                .Buttons(sKey_Return).Enabled = False   '�黹
                .Buttons("Prorefer1").Enabled = True
                .Buttons("Prorefer2").Enabled = True
                .Buttons("Prorefer3").Enabled = True
    

                .Buttons(sKey_Locate).Enabled = True    '��λ
                .Buttons(sKey_Refresh).Enabled = False    'ˢ��
                .Buttons(gstrHelpCode).Enabled = True    '����


                .Buttons(sKey_Addrecord).Enabled = False    '����
                .Buttons(sKey_Deleterecord).Enabled = False    'ɾ��
                .Buttons(sKey_CreateVoucher).Enabled = False    '�Ƶ�
            End If

            Voucher.VoucherStatus = VSNormalMode

            Frm.ComTemplatePRN.Visible = True
            Frm.ComTemplateShow.Visible = False
            Frm.LblTemplate.Caption = GetString("U8.DZ.JA.Res010")
            Frm.LblTemplate.Visible = True


            rs.Close
            Set rs = Nothing


        End Select
    End With

    strSql = "select * from UFMeta_" + g_oLogin.cAcc_Id + "..AA_CustomerButton  where cFormKey='" & gstrCardNumber & "' and cLocaleID='zh-cn'   order by cButtonID"
    Set rs = g_Conn.Execute(strSql)
    If Not rs.EOF Then
        Do While Not rs.EOF
            Toolbar.Buttons(CStr(rs!cButtonkey)).Visible = Toolbar.Buttons(CStr(rs!cVisibleAsKey)).Visible
            Toolbar.Buttons(CStr(rs!cButtonkey)).Visible = Toolbar.Buttons(CStr(rs!cEnableAsKey)).Visible
            rs.MoveNext
        Loop
    End If

    If Voucher.headerText("cborrowouttype") = "2" Then
        '����������ܸ���
        Toolbar.Buttons(sKey_Copy).Enabled = False
    End If

    Toolbar.Refresh
    UFToolbar.RefreshEnable    'ע��,�˴�������RefreshEnable����,Refresh��������


End Sub
