VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H8000000A&
   ClientHeight    =   4320
   ClientLeft      =   345
   ClientTop       =   1905
   ClientWidth     =   7500
   Icon            =   "������.frx":0000
   LinkTopic       =   "MDIForm1"
   ScaleHeight     =   4320
   ScaleWidth      =   7500
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "ϵͳ(&S)"
      Begin VB.Menu miPrintSet 
         Caption         =   "��ӡ����"
         Visible         =   0   'False
      End
      Begin VB.Menu miReLogin 
         Caption         =   "����ע��(&R)"
      End
      Begin VB.Menu miExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "����(&T)"
      Begin VB.Menu miNote 
         Caption         =   "���²�"
         Shortcut        =   {F11}
      End
      Begin VB.Menu miCalc 
         Caption         =   "������"
         Shortcut        =   {F9}
      End
      Begin VB.Menu miCalendar 
         Caption         =   "�������"
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "����(&W)"
      WindowList      =   -1  'True
      Begin VB.Menu miPP 
         Caption         =   "ˮƽƽ��(&H)"
      End
      Begin VB.Menu miWPT 
         Caption         =   "��ֱƽ��(&V)"
      End
      Begin VB.Menu miCD 
         Caption         =   "���(&C)"
      End
      Begin VB.Menu miIcon 
         Caption         =   "����ͼ��(&A)"
      End
      Begin VB.Menu miWBar1 
         Caption         =   "-"
      End
      Begin VB.Menu miStatusbar 
         Caption         =   "��ʾ״̬��"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuShowCmdWin 
         Caption         =   "��ʾ�����"
         Checked         =   -1  'True
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuWindowBigbutton 
         Caption         =   "��ʾ��ť"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu miHelpS 
         Caption         =   "�������۹���ϵͳ����  F1"
      End
      Begin VB.Menu miHBar2 
         Caption         =   "-"
      End
      Begin VB.Menu miInter 
         Caption         =   "��������"
      End
      Begin VB.Menu miHBar3 
         Caption         =   "-"
      End
      Begin VB.Menu miAbout 
         Caption         =   "���� �������۹���ϵͳ"
      End
   End
   Begin VB.Menu mnuKK 
      Caption         =   "KKKKK"
      Visible         =   0   'False
      Begin VB.Menu miAddRow 
         Caption         =   "����һ��"
      End
      Begin VB.Menu miDelRow 
         Caption         =   "ɾ��һ��"
      End
      Begin VB.Menu miCopyRow 
         Caption         =   "���Ƶ�ǰ��"
      End
      Begin VB.Menu miAutoAddcBatch 
         Caption         =   "�Զ�ָ������"
      End
      Begin VB.Menu mibatchwh 
         Caption         =   "������д�ֿ�"
      End
      Begin VB.Menu miCopyBlue 
         Caption         =   "�������ֵ���"
      End
      Begin VB.Menu miSDiscount 
         Caption         =   "������ҵ�ۿ�"
         Visible         =   0   'False
      End
      Begin VB.Menu miTDiscount 
         Caption         =   "�ܶ��̯��ҵ�ۿ�"
      End
      Begin VB.Menu miCancelTDiscount 
         Caption         =   "ȡ���ܶ��̯��ҵ�ۿ�"
      End
      Begin VB.Menu fg1 
         Caption         =   "-"
      End
      Begin VB.Menu miCurXCL 
         Caption         =   "�鿴�ִ���"
      End
      Begin VB.Menu mixyye 
         Caption         =   "��������"
      End
      Begin VB.Menu atofgedit 
         Caption         =   "-"
      End
      Begin VB.Menu miAtoConfig 
         Caption         =   "ATOѡ��"
      End
      Begin VB.Menu miAtoDelconfig 
         Caption         =   "ATOɾ��ѡ��"
      End
      Begin VB.Menu miAtoQueryConfig 
         Caption         =   "ATO��ѯѡ��"
      End
      Begin VB.Menu miPtoConfig 
         Caption         =   "PTOѡ��"
      End
      Begin VB.Menu miPtoDelconfig 
         Caption         =   "PTOɾ��ѡ��"
      End
      Begin VB.Menu miPtoQueryConfig 
         Caption         =   "PTO��ѯ/�޸�ѡ��"
      End
   End
   Begin VB.Menu mnuKK2 
      Caption         =   "KKKKK2"
      Visible         =   0   'False
      Begin VB.Menu miAddRow2 
         Caption         =   "����һ��"
      End
      Begin VB.Menu miDelRow2 
         Caption         =   "ɾ��һ��"
      End
   End
   Begin VB.Menu mnu_JJ 
      Caption         =   "JJJJJ"
      Visible         =   0   'False
      Begin VB.Menu miO 
         Caption         =   "�������"
         Visible         =   0   'False
      End
      Begin VB.Menu miCloseCurFHD 
         Caption         =   "�رյ�ǰ��������¼"
      End
      Begin VB.Menu miCurFXCL 
         Caption         =   "�鿴�ִ���"
      End
      Begin VB.Menu miCurFHDML 
         Caption         =   "��ǰ������Ԥ��ë��"
      End
      Begin VB.Menu miCurWTFHDML 
         Caption         =   "��ǰί�д���������Ԥ��ë��"
      End
      Begin VB.Menu miCurWTJSDML 
         Caption         =   "��ǰί�д������㵥Ԥ��ë��"
      End
      Begin VB.Menu miCurFHDSettle 
         Caption         =   "�鿴��ǰ��������Ʊ���"
      End
      Begin VB.Menu miCurWTFHDSettle 
         Caption         =   "�鿴��ǰί�д����������������"
      End
      Begin VB.Menu fgfhd1 
         Caption         =   "-"
      End
      Begin VB.Menu miCurFHCM 
         Caption         =   "��ѯ��ǰ��������Ӧ�ĺ�ͬ"
      End
      Begin VB.Menu miCurWTFHCM 
         Caption         =   "��ѯ��ǰί�з�������Ӧ�ĺ�ͬ"
      End
      Begin VB.Menu miCurWTJSCM 
         Caption         =   "��ѯ��ǰί�д������㵥��Ӧ�ĺ�ͬ"
      End
      Begin VB.Menu miCurFH2SO 
         Caption         =   "�鿴��ǰ��������Ӧ����"
      End
      Begin VB.Menu miCurFHDTH 
         Caption         =   "�鿴��ǰ��������Ӧ�˻���"
      End
      Begin VB.Menu miCurTHDFH 
         Caption         =   "�鿴��ǰ�˻�����Ӧ������"
      End
      Begin VB.Menu miCurWTJSFH 
         Caption         =   "�鿴��ǰί�н��㵥��Ӧ������"
      End
      Begin VB.Menu miCurWTFHJS 
         Caption         =   "�鿴��ǰί�з�������Ӧ���㵥"
      End
      Begin VB.Menu miCurWTJSDFP 
         Caption         =   "�鿴��ǰί�н��㵥��Ӧ��Ʊ"
      End
      Begin VB.Menu miCurFHDFP 
         Caption         =   "�鿴��ǰ��������Ӧ��Ʊ"
      End
      Begin VB.Menu miCurFHCK 
         Caption         =   "�鿴��ǰ��������Ӧ���ⵥ"
      End
      Begin VB.Menu miCurJSCK 
         Caption         =   "�鿴��ǰί�з�������Ӧ���ⵥ"
      End
      Begin VB.Menu atofgfh 
         Caption         =   "-"
      End
      Begin VB.Menu miPtoQueryConfigJJ 
         Caption         =   "PTO��ѯ/�޸�ѡ��"
      End
   End
   Begin VB.Menu mnuOrder 
      Caption         =   "Order"
      Visible         =   0   'False
      Begin VB.Menu miCloseCurOrder 
         Caption         =   "�رյ�ǰ������¼"
      End
      Begin VB.Menu miCurOXCL 
         Caption         =   "�鿴�ִ���"
      End
      Begin VB.Menu miCurOrdML 
         Caption         =   "�鿴��ǰ����Ԥ��ë��"
      End
      Begin VB.Menu fgso1 
         Caption         =   "-"
      End
      Begin VB.Menu miCurOrdPU 
         Caption         =   "�鿴��ǰ�����ۼƲɹ����"
      End
      Begin VB.Menu miCurOrdFH 
         Caption         =   "�鿴��ǰ�����ۼƷ������"
      End
      Begin VB.Menu miCurOrdKP 
         Caption         =   "�鿴��ǰ�����ۼƿ�Ʊ���"
      End
      Begin VB.Menu fgso2 
         Caption         =   "-"
      End
      Begin VB.Menu miCurOrdCM 
         Caption         =   "��ѯ��ǰ������Ӧ�ĺ�ͬ"
      End
      Begin VB.Menu miCurSOToQuo 
         Caption         =   "�鿴��ǰ������Ӧ���۵�"
      End
      Begin VB.Menu miCurOrdFHD 
         Caption         =   "�鿴��ǰ������Ӧ������"
      End
      Begin VB.Menu miCurOrdFP 
         Caption         =   "�鿴��ǰ������Ӧ��Ʊ"
      End
      Begin VB.Menu atofgso 
         Caption         =   "-"
      End
      Begin VB.Menu miAtoQueryConfigOrder 
         Caption         =   "ATO��ѯѡ��"
      End
      Begin VB.Menu miPtoQueryConfigOrder 
         Caption         =   "PTO��ѯ/�޸�ѡ��"
      End
   End
   Begin VB.Menu mnuSettle 
      Caption         =   "Settle"
      Visible         =   0   'False
      Begin VB.Menu miCloseFHD 
         Caption         =   "�رյ�ǰ��Ʊ��Ӧ������"
      End
      Begin VB.Menu miCurSXCL 
         Caption         =   "�鿴�ִ���"
      End
      Begin VB.Menu miCurSettleML 
         Caption         =   "��ǰ��ƱԤ��ë��"
      End
      Begin VB.Menu miCurSettleSK 
         Caption         =   "�鿴��ǰ��Ʊ�տ�������"
      End
      Begin VB.Menu fpfg1 
         Caption         =   "-"
      End
      Begin VB.Menu miCurSettleCM 
         Caption         =   "��ѯ��ǰ��Ʊ��Ӧ�ĺ�ͬ"
      End
      Begin VB.Menu miCurFP2SO 
         Caption         =   "�鿴��ǰ��Ʊ��Ӧ����"
      End
      Begin VB.Menu miCurSettleFH 
         Caption         =   "�鿴��ǰ��Ʊ��Ӧ������"
      End
      Begin VB.Menu miCurPurBill 
         Caption         =   "�鿴��ǰ��Ʊ��Ӧ�ɹ���Ʊ"
      End
      Begin VB.Menu miCurSettleWTJS 
         Caption         =   "�鿴��ǰ��Ʊ��Ӧί�д������㵥"
      End
      Begin VB.Menu miCurSettleCK 
         Caption         =   "�鿴��ǰ��Ʊ��Ӧ���ⵥ"
      End
      Begin VB.Menu miCurSettleSP 
         Caption         =   "�鿴��ǰ��Ʊ��Ӧ���۷���֧����"
      End
      Begin VB.Menu miCurSettleEXP 
         Caption         =   "�鿴��ǰ��Ʊ��Ӧ������õ�"
      End
      Begin VB.Menu atofgfp 
         Caption         =   "-"
      End
      Begin VB.Menu miPtoQueryConfigSett 
         Caption         =   "PTO��ѯ/�޸�ѡ��"
      End
   End
   Begin VB.Menu mnuQuo 
      Caption         =   "quo"
      Visible         =   0   'False
      Begin VB.Menu miCurQuoToSO 
         Caption         =   "�鿴��ǰ���۵���Ӧ����"
      End
      Begin VB.Menu atofgquo 
         Caption         =   "-"
      End
      Begin VB.Menu miAtoQueryConfigQuo 
         Caption         =   "ATO��ѯѡ��"
      End
      Begin VB.Menu miPtoQueryConfigQuo 
         Caption         =   "PTO��ѯ/�޸�ѡ��"
      End
   End
   Begin VB.Menu mnuSortPrice 
      Caption         =   "�۸�����"
      Visible         =   0   'False
      Begin VB.Menu miAscPrice 
         Caption         =   "����"
      End
      Begin VB.Menu miDescPrice 
         Caption         =   "����"
      End
   End
   Begin VB.Menu mnuSortPriceLog 
      Caption         =   "���ۼ�¼����"
      Visible         =   0   'False
      Begin VB.Menu miAscPriceLog 
         Caption         =   "����"
      End
      Begin VB.Menu midescPriceLog 
         Caption         =   "����"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Sub Form_Load()
    FormInit
    GetMenuConfig
    GetHelpFile
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim frmUnload As Form
    Dim b As Integer
    
    On Error Resume Next
    For Each frmUnload In Forms
        If LCase(frmUnload.Name) <> "frmmain" Then
            b = Forms.Count
            Unload frmUnload         '.Name
            If b = Forms.Count Then
                Cancel = 3
                Exit Sub
            End If
        End If
    Next
    Set gcAccount = Nothing
    Unload Me
    g_bCanExit = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Set oColSet = Nothing
    Set ogszzPub = Nothing
    Set gcAccount = Nothing
    Exit Sub
End Sub

Public Sub OnCommand(ByVal cmenuid As String, ByVal cmenuname As String, ByVal cAuthId As String, ByVal cCmdLine As String)
    Dim vouchID() As String
    Dim strVouchID As String
    Dim strCurrentRow As String
 
    If Not HaveSufficeResources() Then Exit Sub
    strVouchID = ""
    If cCmdLine <> "" Then
        vouchID = Split(cCmdLine, vbTab, -1, vbTextCompare)
        If UBound(vouchID) >= 1 Then
            strVouchID = vouchID(1)
        End If
        If UBound(vouchID) >= 3 Then
            strCurrentRow = vouchID(2)
        End If
    End If
    Select Case UCase(cmenuid)
       Case "EFPBBASE0114" '��������
              TBLStyle = TBLText
              Call mnuFileMange20_Click
        Case "EFFYGL040101"  '����ѡ��
            frmFyglOption.Show 1
        Case "EFFYGL040203" '����Ԥ������������
            frmZdCX.bShowForm = True
            frmZdCX.Show 1
        Case Else
            MenuClick cmenuid, cAuthId, strVouchID, strCurrentRow, True, cmenuname
    End Select
End Sub

Public Sub MenuClick(strMenuID As String, strAuthID As String, Optional strVoucherID As String = "", Optional strCurrentRow, Optional blnMenuClick As Boolean = False, Optional cmenuname As String)
    Dim lst As IXMLDOMNodeList
    Dim nod As IXMLDOMElement
    Dim strType As String
    Dim strError As String
    Dim strPara As String
'    Dim pRep As ReportService.clsReportManager
'    Set pRep = New ReportService.clsReportManager
    Dim pRep As Object
    Dim other_frm As Object
    On Error Resume Next
    
    Set pRep = CreateObject("ReportService.clsReportManager")
    
    Set lst = domMenu.selectNodes("//z:row[@menuid='" + strMenuID + "']")
    If lst.length > 0 Then
        For Each nod In lst
            strType = nod.Attributes.getNamedItem("functionid").Text
            Select Case strType
                Case "voucherlist"
                    ShowVoucherList nod, strError, strAuthID, strVoucherID, cmenuname
                Case "uapreport"
                    strPara = nod.Attributes.getNamedItem("parameters").Text
                    cls_Public.WrtDBlog DBconn, m_login.cUserId, "EFmain", "EFmain->uapreport[" & strPara & "]����ʼ�򿪣�"
                    Call pRep.OpenReport(strPara, m_login, "")
'                Case "report"
'                    strPara = nod.Attributes.getNamedItem("parameters").Text
'                    OpenNewReport strPara, strAuthID
                Case "oldvoucher"
                    nod.setAttribute "setcurrentrow", strCurrentRow
                    nod.setAttribute "voucherid", strVoucherID
                    If strVoucherID = "" And blnMenuClick Then
                        nod.setAttribute "voucherautoid", ""
                    End If
                    ShowOldVoucher nod, strError, strAuthID
                    
                Case "other"
                    Set other_frm = CreateObject(nod.Attributes.getNamedItem("parameters").Text)
                    Call other_frm.Init(g_business, m_login)
                    Call other_frm.Show(m_login)
                    
                
            End Select
        Next
    Else
        If Not blnMenuClick Then OnCommand strMenuID, "", strAuthID, "" & vbTab & strVoucherID & vbTab & ""
    End If
End Sub

Private Function ShowVoucherList(nod As IXMLDOMNode, strErrorResId As String, Optional strAuthID As String, Optional strHeadKeys As String, Optional cmenuname As String) As Boolean
    
    Dim strColumnSetKey As String
    Dim frmVoucherList As New frmAutoVoucherList
    Dim clsVoucherList As New clsAutoVoucherList
    Dim strToolBar As String
    Dim strHelpId As String
    Dim strDefaultStr As String
    Dim strGUID As String
    
    If strAuthID <> "" Then
        If Not LockItem(strAuthID, True, True) Then
            strErrorResId = "�������벻�ɹ�"
            Set frmVoucherList = Nothing
            Set clsVoucherList = Nothing
            Exit Function
        End If
    End If
    strColumnSetKey = nod.Attributes.getNamedItem("parameters").Text
    If Not nod.Attributes.getNamedItem("toolbarname") Is Nothing Then
        strToolBar = nod.Attributes.getNamedItem("toolbarname").Text
    End If
    If Not nod.Attributes.getNamedItem("helpid") Is Nothing Then
        strHelpId = nod.Attributes.getNamedItem("helpid").Text
    End If
    If Not nod.Attributes.getNamedItem("defaultstr") Is Nothing Then
        strDefaultStr = nod.Attributes.getNamedItem("defaultstr").Text
    End If
    clsVoucherList.Init strColumnSetKey, strErrorResId
    If strDefaultStr <> "" Then
        clsVoucherList.strDefaultFilter = IIf(clsVoucherList.strDefaultFilter = "", strDefaultStr, "(" & clsVoucherList.strDefaultFilter & ") and (" & strDefaultStr & ")")
    End If
    Dim blnFilter As Boolean
    If strHeadKeys <> "" Then
        Dim strMainTbl As String
        Dim strMainKey As String
            
        strMainTbl = clsVoucherList.GetVoucherListSet("maintbl")
        strMainKey = clsVoucherList.GetVoucherListSet("mainkey")
        If strMainTbl <> "" And strMainKey <> "" Then
            strHeadKeys = strMainTbl & "." & strMainKey & " in (" & strHeadKeys & ")"
        End If
        clsVoucherList.strFilter = strHeadKeys
        blnFilter = True
    Else
'        blnFilter = clsVoucherList.ShowFilter()
        blnFilter = True
    End If
    If blnFilter Then
        Set frmVoucherList.clsVoucherLst = clsVoucherList
        frmVoucherList.strColumnSetKey = strColumnSetKey
        frmVoucherList.strAuthID = strAuthID
        frmVoucherList.strToolBarName = strToolBar
        frmVoucherList.strHelpId = strHelpId
        frmVoucherList.formCaption = cmenuname
        If Not (g_business Is Nothing) Then
            strGUID = ShowPortalForm(frmVoucherList, False)
            frmVoucherList.strFormGuid = strGUID
        Else
            frmVoucherList.Show
        End If
        Set clsVoucherList = Nothing
    Else
        Set frmVoucherList = Nothing
        Set clsVoucherList = Nothing
        If strAuthID <> "" Then LockItem strAuthID, False, False
    End If
End Function


Private Function ShowOldVoucher(nod As IXMLDOMNode, strErrorResId As String, Optional strAuthID As String) As Boolean
    If Not nod.Attributes.getNamedItem("authid") Is Nothing Then
        strAuthID = nod.Attributes.getNamedItem("authid").nodeValue
    End If
    ShowOldVoucher = True
    miBJDNew_Click nod, strAuthID
End Function
Public Sub miBJDNew_Click(nod As IXMLDOMNode, strTaskId As String, Optional imode As Integer, Optional SBVID As String, Optional cSBVCode As String, Optional mDom As DOMDocument)
    Dim strToolBar As String
    Dim domDefault As New DOMDocument
    Dim strVouchtype As String
    Dim blnFirst As Boolean
    Dim blnRetrunFlag As Boolean
    Dim strCardNumber As String
    Dim strHelpId As String
    
    blnFirst = False
    blnRetrunFlag = False
    
    '//�������ID
    If Not nod.Attributes.getNamedItem("helpid") Is Nothing Then
        strHelpId = nod.Attributes.getNamedItem("helpid").Text
    End If
    
    '//ģ����
    If Not nod.Attributes.getNamedItem("parameters") Is Nothing Then
        strCardNumber = nod.Attributes.getNamedItem("parameters").Text
    End If
    
    '//��������ʶ
    If Not nod.Attributes.getNamedItem("toolbarname") Is Nothing Then
        strToolBar = nod.Attributes.getNamedItem("toolbarname").Text
    End If
    
    Dim strCurrentRow As String
    If Not nod.Attributes.getNamedItem("setcurrentrow") Is Nothing Then
        strCurrentRow = nod.Attributes.getNamedItem("setcurrentrow").Text
    End If
    
    '//����ID
    Dim strVouchMainId As String
    If Not nod.Attributes.getNamedItem("voucherid") Is Nothing Then
        strVouchMainId = nod.Attributes.getNamedItem("voucherid").Text
    End If
    
    Dim strType As String
    If Not nod.Attributes.getNamedItem("condition") Is Nothing Then
        strType = nod.Attributes.getNamedItem("condition").Text
    End If
    
    '//Ȩ��ID
    If Not nod.Attributes.getNamedItem("authid") Is Nothing Then
        strTaskId = nod.Attributes.getNamedItem("authid").Text
    End If
    
    '//Ĭ�ϵ������͡��ڳ���ʶ�������ֱ�ʶ
    If Not nod.Attributes.getNamedItem("defaultstr") Is Nothing Then
        domDefault.loadXML nod.Attributes.getNamedItem("defaultstr").Text
        If Not domDefault.documentElement.Attributes.getNamedItem("cvouchtype") Is Nothing Then
            strVouchtype = domDefault.documentElement.Attributes.getNamedItem("cvouchtype").nodeValue
        End If
        If Not domDefault.documentElement.Attributes.getNamedItem("first") Is Nothing Then
            blnFirst = CBool(domDefault.documentElement.Attributes.getNamedItem("first").nodeValue)
        End If
        If Not domDefault.documentElement.Attributes.getNamedItem("retrunflag") Is Nothing Then
            blnRetrunFlag = CBool(domDefault.documentElement.Attributes.getNamedItem("retrunflag").nodeValue)
        End If
    End If
    Set domDefault = Nothing
    
    Dim frmDD As Form
    For Each frmDD In Forms
        If LCase(frmDD.Name) = LCase("frmBillVouchNew") Then
            frmDD.Voucher.ProtectUnload2
        End If
    Next
    
    If LockItem(strTaskId, True) Then
        
        Select Case UCase(strCardNumber)
            Case "EFYZGL030301"   '���⴦��

            Case Else
                Set frmDD = New frmVouchNew
        End Select
        If imode = 2 Then
            frmDD.strSBVID = strVouchMainId
            frmDD.strSBVCode = cSBVCode
            frmDD.hDOM = mDom
        End If
        frmDD.strToolBarName = strToolBar
        frmDD.strCardNum = strCardNumber
        frmDD.strVouchtype = strVouchtype
        frmDD.bReturnFlag = blnRetrunFlag
        frmDD.bFirst = blnFirst
        frmDD.FormVisible = False
        frmDD.strHelpId = strHelpId
        
        If strVouchMainId <> "" And imode <> 2 Then
            frmDD.ShowVoucher strCardNumber, strVouchMainId, 1, strCurrentRow
        Else
            frmDD.ShowVoucher strCardNumber, , imode
        End If
        
        If frmDD.FormVisible = True Then
            frmDD.UFTaskID = strTaskId
            If imode = 2 Then
                frmDD.ButtonClick "Add", "����"
            End If
            frmDD.ZOrder 0
        Else
            frmDD.UFTaskID = strTaskId
            Unload frmDD
            Exit Sub
        End If
    End If
End Sub
'
Private Sub FormInit()
    On Error Resume Next
    gcAccount.Version = 1
LBStart:

    m_login.AuthString = strAuthStrForLogin
    Screen.MousePointer = vbDefault
    InitAccount
    Init
End Sub
  
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

Private Function IsInFormsByTag(strTag As String) As Boolean
    Dim frmTmp As Form
    IsInFormsByTag = False
    For Each frmTmp In Forms
        If frmTmp.Tag = strTag Then
            frmTmp.ZOrder 0
            IsInFormsByTag = True
            Exit For
        End If
    Next
End Function
Private Sub mnuFileMange20_Click()
    'Ԥ���������
     Dim DbGSP As UfDatabase
    Set DbGSP = New UfDatabase
    DbGSP.OpenDatabase m_login.UfDbName, False, False, ";PWD=" & m_login.SysPassword
    If UA_Task("KI200701") Then
        Dim cls As New EFClass.IINterface
        Set cls.o_business = g_business
        cls.putWnd frmMain.hwnd
        cls.putPath App.HelpFile
        cls.Show DbGSP, m_login, 11, "Ԥ���������", TBLStyle, "KI200702", "KI200702", "KI2007", "", ""
        Set cls = Nothing
    End If
    UA_FreeTask "KI2007"
End Sub


