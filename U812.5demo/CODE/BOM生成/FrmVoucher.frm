VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{201FB79B-5556-47A4-AD9C-A46BA0C45A44}#6.42#0"; "UFToolBarCtrl.ocx"
Object = "{D2B3369D-2E6C-45DE-A705-14481242A2BE}#1.12#0"; "UFMenu6U.ocx"
Object = "{005151D4-33F6-471B-B651-222D4E411832}#4.5#0"; "UFFormPartner.ocx"
Object = "{AF8BBBB7-94C6-4772-B826-624478C37D6A}#1.5#0"; "UFKEYHOOK.ocx"
Object = "{9FD12F62-6922-47E1-B1AC-3615BBD3D7A5}#1.0#0"; "UFLabel.ocx"
Object = "{4C2F9AC0-6D40-468A-8389-518BB4F8C67D}#1.0#0"; "UFComboBox.ocx"
Object = "{456334B9-D052-4643-8F5F-2326B24BE316}#6.96#0"; "UAPvouchercontrol85.ocx"
Begin VB.Form FrmVoucher 
   Caption         =   "Form1"
   ClientHeight    =   5925
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9210
   Icon            =   "FrmVoucher.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin UFFormPartner.UFFrmCaption UFFrmCaptionMgr 
      Left            =   1200
      Top             =   5040
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "Form1"
      DebugFlag       =   0   'False
      SkinStyle       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin cPopMenu6.PopMenu PopMenu1 
      Left            =   360
      Top             =   3120
      _ExtentX        =   1058
      _ExtentY        =   1058
      HighlightCheckedItems=   0   'False
      TickIconIndex   =   0
      HighlightStyle  =   2
      ActiveMenuForeColor=   -2147483641
      MenuBackgroundColor=   16119285
      HookInSubClassMenu=   0   'False
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox PicTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   360
      ScaleHeight     =   15
      ScaleWidth      =   9615
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   9615
      Begin UFCOMBOBOXLib.UFComboBox ComTemplatePRN 
         Height          =   300
         Left            =   720
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   120
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   556
         _StockProps     =   196
         Text            =   ""
         Style           =   2
         ForeColor       =   2
      End
      Begin UFCOMBOBOXLib.UFComboBox ComTemplateShow 
         Height          =   300
         Left            =   6480
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   120
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   556
         _StockProps     =   196
         Text            =   ""
         Style           =   2
         ForeColor       =   2
      End
      Begin UFLABELLib.UFLabel LabTitle 
         Height          =   300
         Left            =   4095
         TabIndex        =   5
         Top             =   135
         Width           =   1200
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   661
         _StockProps     =   111
         Caption         =   "��������"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSize        =   -1  'True
         BackStyle       =   0
      End
      Begin UFLABELLib.UFLabel LblTemplate 
         Height          =   225
         Left            =   5490
         TabIndex        =   4
         Top             =   135
         Width           =   945
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   111
         Caption         =   "��ӡģ�棺"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSize        =   -1  'True
         BackStyle       =   0
      End
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin UAPVoucherControl85.ctlVoucher Voucher 
      Height          =   3015
      Left            =   1440
      TabIndex        =   1
      Top             =   1800
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5318
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10446406
      TitleForecolor  =   16050403
      DisabledColor   =   16777215
      ColAlignment0   =   9
      Rows            =   20
      Cols            =   20
      TitleForecolor  =   16050403
      ControlScrollBars=   0
      ControlAutoScales=   0
      BaseOfVScrollPoint=   0
      ShowSorter      =   0   'False
      ShowFixColer    =   0   'False
   End
   Begin UFToolBarCtrl.UFToolbar UFToolbar 
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UFKeyHook.UFKeyHookCtrl UFKeyHookCtrl1 
      Left            =   4440
      Top             =   4995
      _ExtentX        =   1905
      _ExtentY        =   529
   End
End
Attribute VB_Name = "FrmVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************
'ģ�鹦��˵����
'1��ʵ�ֵ��ݵĻ�����ť����:��ӡ�������Ԥ�����������޸ġ�ɾ�������С�ɾ�С������С����ơ����桢�������ύ�������ύ����������ˡ����󡢴򿪡��رա�ˢ�¡���ҳ����λ���������ϲ���ʾ���Ҽ��˵��������ģ��Ҽ��˵�������¼��λ���Ҽ��˵����������������Ƶ���ȡ�ۣ����У���������
'2��Ԥ�Ƶ���ģ�壺�������ݱ�ͷ����Ļ�����Ŀ����1-16��ͷ�Զ����1-10��������1-16�����Զ������Ŀ���롢���š���ֵ�ڵ�
'3  ���ݺ�����:     ��ˮ�š����š��ֿ⡢�������ڡ��Ƶ���
'4����ͷ������Ŀ�Ĳ��ա���Ч��У��
'5�����������Ŀ�Ĳ��ա���Ч��У��
'6������ʱ����Ч��У�� , ��: ������ṹ���Ƿ�Ϸ� , �����ڴ���ı�����У��, ���δ���͸����ʹ������������ⵥ��У�顢
'7����������
'8����������������
'9��֧�ֱ������򡢶����ʾ����ӡģ��
'10��֧�ִ�ӡ���ñ���
'11��֧�ֹ���Ȩ������
'12��֧������Ȩ��(��¼��Ȩ�ޣ��ݲ�֧�ִ��Ȩ��)
'13��֧��ȡ�۹��ܣ��۸���գ�������
'14��֧��������
'15��֧��վ�����
'16��֧�ֵ���ģ�����б����ݾ�������
'����ʱ�䣺2008-11-21
'�����ˣ�xuyan
'****************************************

Option Explicit

'���ݷ������
Private VchSrv As New clsVouchServer

Private m_Cancel, m_UnloadMode As Integer

' * �������Ա�������
Private m_strVT_ID As String
Private m_strVT_PRN_ID As String

'by liwqa Template
Private dicTemplate As New Dictionary                      '��¼������ʾģ����combox�Ķ�Ӧ��ϵ
Private dicTemplatePrint As New Dictionary
Private bInitForm As Boolean

Private objVoucherTemplate As New UFVoucherServer85.clsVoucherTemplate    'UFVoucherServer85.clsVoucherTemplate
Private objVoucher85 As UFVoucherD85.clsVoucher85
Private objBill As UFBillComponent.clsBillComponent
'����Ȩ��
Private Const AuthBrowse = "FYSL02050301"                  '���
Private Const AuthAdd = "FYSL02050302"                     '����
Private Const AuthModify = "FYSL02050303"                  '�޸�
Private Const AuthDelete = "FYSL02050304"                  'ɾ��
Private Const AuthVerify = "FYSL02050305"                  '���
Private Const AuthUnVerify = "FYSL02050306"                '����
Private Const AuthOpen = "FYSL02050310"                    '��
Private Const AuthClose = "FYSL02050311"                   '�ر�
Private Const AuthPrint = "FYSL02050307"                   '��ӡ
Private Const AuthOut = "FYSL02050308"                     '���
Private Const AuthProrefer = "FYSL02050312"                '����
Private Const AuthProrefer1 = "FYSL02050313"                '����

Dim wfcBack As Integer                                     '�Ƿ�������ˢ�½����ʾ
Dim inited As Boolean                                      '�����ʼ�����

'by zhangwchb 20110718 ��չ�ֶ�
Dim sExtendField As String
Dim sExtendJoinSQL As String
Dim sExtendBodyField As String
Dim sExtendBodyJoinSQL As String
Dim oExtend As Object

Dim objRelation As Object

Private mobjSubServ As New ScmPublicSrv.clsAutoFill        '20110822 by zhangwchb

'��ƽ̨������Ϣ�ӿ� 20110812
Private WithEvents m_mht As UFPortalProxyMessage.IMessageHandler
Attribute m_mht.VB_VarHelpID = -1

Private isSavedOK As Boolean                               '�����Ƿ񱣴�ɹ� 'by zhangwchb 20110829 �ύ����

'������״̬�£�����û�������ʾģ�塣��ʱ�жϱ����Ƿ�������
Dim sPreVTID As Integer
Public bexitload As Boolean '����ģ������Ȩ�޿���
Private mdOldSelRow As Long
Private mdOldSelCol As Long
Public bAlter As Boolean

'ɾ��
Private Sub ExecSubDelete()
    On Error GoTo Err_Handler
    Dim sMessage, sSource As String
    
    '12.0֧����չ�Զ�����
    Dim skeyfld As String
    Dim skeySubfld As String
    Dim objExtend As Object
    Set objExtend = CreateObject("VoucherExtendService.ClsExtendServer")
    Dim oDomHead As New DOMDocument
    Dim oHeadElement As IXMLDOMElement
    Dim strSql As String
    Dim rs As New ADODB.Recordset
       Dim strXML As String
    Dim objDoc As New MSXML2.DOMDocument
    Dim ErrDesc As String
'    'enum by modify
'    If Voucher.headerText("cCreateType") = "ת������" Then
'        MsgBox GetString("U8.DZ.JA.Res040"), vbInformation, GetString("U8.DZ.JA.Res030")
'        Exit Sub
'    End If

    If MsgBox(GetString("U8.DZ.JA.Res050"), vbYesNo, GetString("U8.DZ.JA.Res030")) = vbNo Then
        Exit Sub
    End If

    'by liwqa ����
    Dim m_bTrans As Boolean
    Dim lngAct As Long
    g_Conn.BeginTrans
    m_bTrans = True
 
    
    g_Conn.Execute "delete from " & MainTable & " where convert(nchar,convert(money,ufts),2)  = " & OldTimeStamp & " and  " & HeadPKFld & "=" & lngVoucherID, lngAct

    If lngAct = 0 Then

        If m_bTrans Then
            g_Conn.RollbackTrans
            m_bTrans = False
        End If
        MsgBox GetString("U8.DZ.JA.Res290"), vbInformation, GetString("U8.DZ.JA.Res030")
        GoTo Exit_Label
    End If
    
    
'      '************************************
'     '������Ŀ�������ۼ������ͽ��
'
'     strSql = " update HY_FYSL_Contract set totalappmoney= " & Null2Something(Voucher.headerText("totalappmoney"), 0) - Null2Something(Voucher.headerText("appprice"), 0) & " where  ccode= '" & Voucher.headerText("concode") & "'"
'       g_Conn.Execute strSql
'
'     '***********************************
'
    
    
      '����ɾ��
    '*******************************
'         strXML = Voucher.GetAccessoriesInfo(ErrDesc)
'        If ErrDesc <> "" Then
'                    If m_bTrans Then
'            g_Conn.RollbackTrans
'            m_bTrans = False
'        End If
'            MsgBox ErrDesc, vbInformation, GetString("U8.DZ.JA.Res030")
'            Exit Sub
'        End If
'        Dim m_oServer As New UFVoucherServer85.clsVoucherTemplate
'
'        Call objDoc.loadXML(strXML)
'        Call objDoc.documentElement.setAttribute("AccessoriesChanged", "1")
'        Call objDoc.documentElement.setAttribute("Deleted", "1")
'        strXML = objDoc.xml
'        Set objDoc = Nothing
'        Call m_oServer.SaveAccessories(strXML, g_Conn, ErrDesc)
'        If ErrDesc <> "" Then
'                   If m_bTrans Then
'            g_Conn.RollbackTrans
'            m_bTrans = False
'        End If
'            MsgBox ErrDesc, vbInformation, GetString("U8.DZ.JA.Res030")
'            Exit Sub
'        End If

  '*******************************
     
        g_Conn.CommitTrans
        m_bTrans = False
    '��һ�ŵ��ݵ� ID
    lngVoucherID = GetTheLastID(login:=g_oLogin, _
            oConnection:=g_Conn, _
            sTable:=MainTable, _
            sField:=HeadPKFld & " asc", _
            sDataNumFormat:="0", _
            sWhereStatement:="" & IIf(PageCurrent > 1, HeadPKFld & " < " & lngVoucherID & "", ""))



    '**********************************************************
    '����(����/����),ɾ������ȫ�ֱ���
    '���ӱ���,��ȡlngvoucherid,�޸ı��治����ȡ
    '**********************************************************

    pageCount = pageCount - 1
    If PageCurrent > 1 Then PageCurrent = PageCurrent - 1


    '**********************************************************
    '����,�޸�,ɾ�� ���µ���״̬
    '**********************************************************
    mOpStatus = SHOW_ALL
    Voucher.VoucherStatus = VSNormalMode
    Call ExecSubRefresh
    Call SetCtlStyle(Me, Voucher, Me.Toolbar, Me.UFToolbar, mOpStatus)


Exit_Label:
    On Error GoTo 0
    Exit Sub
Err_Handler:

    If m_bTrans Then
        g_Conn.RollbackTrans
        m_bTrans = False
    End If

    sMessage = GetString("U8.DZ.JA.Res060")

    ' * ��ʾ�ѺõĴ�����Ϣ
    Call ShowErrorInfo( _
            sHeaderMessage:=sMessage, _
            lMessageType:=vbExclamation, _
            lErrorLevel:=ufsELHeaderAndDescription _
            )

    ' * �����������Դ
    If Left(Err.Source, 3) = "***" Then
        sSource = Err.Source
    Else
        sSource = "***Sub DeleteVoucher of Form frmVoucher"
    End If

    ' * ����ģʽʱ����ʾ���Դ��ڣ����ڸ��ٴ���
    If g_blnDEBUG_MODE Then
        Call ShowDebugForm( _
                bErrorMode:=True, _
                sProcedure:=sSource)
    End If
    GoTo Exit_Label
End Sub

'�޸�
Private Sub ExecSubModify()
    On Error GoTo Err_Handler
    Dim sMessage, sSource As String


    '    If Voucher.headerText("iswfcontrolled") = 1 And Voucher.headerText("iverifystate") = "1" Then
    '        MsgBox GetString("U8.DZ.JA.Res070"), vbInformation, GetString("U8.DZ.JA.Res030")
    '        GoTo Exit_Label
    '    End If
    
    bAlter = False
    
    If Trim(Voucher.headerText(HeadPKFld)) = "" Then
        MsgBox GetString("U8.DZ.JA.Res080"), vbExclamation, GetString("U8.DZ.JA.Res030")
        GoTo Exit_Label
    End If


    mOpStatus = MODIFY_MAIN
    Voucher.VoucherStatus = VSeEditMode
     numappprice = 0
     moneytmp = Val(Null2Something(Voucher.headerText("progressmoney"), 0))
     numbertmp = Val(Null2Something(Voucher.headerText("progressqty"), 0))
     numappprice = Val(Null2Something(Voucher.headerText("appprice"), 0))

    'by liwqa Template
'    Call setTemplate(Voucher.headerText("ivtid"))
'    If Voucher.headerText("ivtid") = "" Then
'        Voucher.headerText("ivtid") = m_strVT_ID
'    End If

    '���õ��ݱ���Ƿ�ɱ༭
    Dim manual As Boolean                                  ' �Ƿ���ȫ�ֹ����
    Call SetVouchCodeEnable(manual)


    Call SetCtlStyle(Me, Voucher, Me.Toolbar, Me.UFToolbar, mOpStatus)

    'by zhangwchb 20110829 �ύ����
'    If bInWorkFlow = True And Val(Voucher.headerText("iverifystate")) = 0 Then
'        Me.Toolbar.Buttons("Submit").Enabled = True
'        UFToolbar.RefreshEnable
'    Else
'        Me.Toolbar.Buttons("Submit").Enabled = False
'        UFToolbar.RefreshEnable
'    End If

Exit_Label:
    On Error GoTo 0

    Exit Sub
Err_Handler:
    sMessage = GetString("U8.DZ.JA.Res090")

    ' * ��ʾ�ѺõĴ�����Ϣ
    Call ShowErrorInfo( _
            sHeaderMessage:=sMessage, _
            lMessageType:=vbExclamation, _
            lErrorLevel:=ufsELHeaderAndDescription _
            )

    ' * �����������Դ
    If Left(Err.Source, 3) = "***" Then
        sSource = Err.Source
    Else
        sSource = "***Sub DoEdit of Form frmVoucher"
    End If

    ' * ����ģʽʱ����ʾ���Դ��ڣ����ڸ��ٴ���
    If g_blnDEBUG_MODE Then
        Call ShowDebugForm( _
                bErrorMode:=True, _
                sProcedure:=sSource)
    End If
    GoTo Exit_Label

End Sub

'����
Private Sub ExecSubSave()

    On Error GoTo Err_Handler

    Dim HeadData     As New CDataVO

    Dim BodyData     As New CDataVO

    Dim strSql       As String

    Dim DMO          As CDMO

    Dim Rult         As CResult

    Dim rs           As New ADODB.Recordset

    Dim i            As Integer

    Dim sID2         As Long

    Dim strVoucherNo As String

    Dim bRepeatRedo  As Boolean
    
    Dim bReGetNO     As Boolean
    
    'by liwqa ����
    Dim m_bTrans     As Boolean

    Dim lngAct       As Long
    
    '12.0֧����չ�Զ�����
    Dim skeyfld      As String

    Dim skeySubfld   As String

    Dim objExtend    As Object

    Dim fldextends   As ADODB.Fields

    Set objExtend = CreateObject("VoucherExtendService.ClsExtendServer")

    Dim oDomBody       As New DOMDocument

    Dim oHeadElement   As IXMLDOMElement

    Dim oBodyElement() As IXMLDOMElement

    '    Me.Voucher.RemoveEmptyRow   '�������

    isSavedOK = False                                      'by zhangwchb 20110829 �ύ����

CHECK:
    '    For i = 1 To Voucher.BodyRows
    '        '        Voucher.BodyRowIsEmpty i
    '        If Trim(Voucher.bodyText(i, "cinvcode")) = "" Then
    '            Voucher.row = i
    '            Voucher.DelLine
    '            GoTo CHECK
    '        End If
    '    Next

    '��Ч��У��
    If ExecFunSaveCheck(Voucher) = False Then Exit Sub

    '�޸ı��棬��Ҫ�Ƚ�ʱ�����������ֲ���
    If Voucher.VoucherStatus = VSeEditMode Then
        If ExecFunCompareUfts = False Then Exit Sub
    End If

    '��ȡ�����ֶκ�����
    ' ��ȡ�ֶ�
    Set HeadData = GetHeadVouchData(g_Conn, Voucher, MainTable)
    '    Set BodyData = GetBodyVouchData(g_Conn, Voucher, DetailsTable)

    g_Conn.BeginTrans
    m_bTrans = True

    '��������,������id,autoid
    '�޸ı���ʱ,����Ҫ����id
    'dxb  2009 6 15

    '    g_Conn.Execute "update ufsystem..ua_identity set ifatherid=(select isnull(max(ID),1) from " & DetailsTable & _
    '                   ") , ichildid=(select isnull(max(autoid),1) from " & DetailsTable & _
    '                   ") where cvouchtype='" & gstrCardNumber & "' and cAcc_Id ='" & g_oLogin.cAcc_Id & "'"

    '    If Voucher.VoucherStatus = VSeAddMode Then
    '      Call GetMaxID
    '    End If
    '    strSql = "select * from ufsystem..ua_identity where cvouchtype='" & gstrCardNumber & "' and cAcc_Id ='" & g_oLogin.cAcc_Id & "'"
    '    Set rs = g_Conn.Execute(strSql)
    '    sID = rs.Fields("iFatherId").Value
    '    sAutoId = rs.Fields("iChildID").Value

    If Voucher.VoucherStatus = VSeAddMode Then
        Call GetMaxID
        '        Call GetMaxID
        HeadData.Item(1).Item("ID").Value = sID
    End If

    '���µ��ݺ���ˮ��
    Dim oDomHead   As New DOMDocument

    Dim oDomFormat As DOMDocument

    Dim sError     As String

    Set oDomHead = Voucher.GetHeadDom

    If Not BOGetVoucherNO(g_Conn, oDomHead, gstrCardNumber, sError, Voucher.headerText(strcCode), oDomFormat, False, , , True) Then
        If m_bTrans Then
            g_Conn.RollbackTrans
            m_bTrans = False
        End If

        Err.Raise 0, GetString("U8.DZ.JA.Res100")
        GoTo Err_Handler
    End If
    
    '12.0��ͷ��չ�Զ�����
    '    Set oHeadElement = oDomHead.selectSingleNode("//z:row")
    '    skeyfld = objExtend.getKeyid(g_Conn, MainTable & "_ExtraDefine")
    '    skeySubfld = objExtend.getKeyid(g_Conn, DetailsTable & "_ExtraDefine")

    '�����ͷ
    HeadData.Item(1).Item("dmDate").Value = g_oLogin.CurDate    'Now() '�Ƶ�ʱ�� Format(Now(), "YYYY-MM-DD HH:MM:SS")
    HeadData.Item(1).Item("iStatus").Value = 1             '״̬

    If bAlter Then
        HeadData.Item(1).Item("iStatus").Value = 2             '״̬
    End If
    
    Set DMO = New CDMO

    '����
    If Voucher.VoucherStatus = VSeAddMode Then
        
        Set Rult = DMO.Insert(g_Conn, HeadData)
        
              
       
'       strSql = " update HY_FYSL_Contract set totalappmoney=isnull(totalappmoney,0)+" & Voucher.headerText("appprice") & " where  ccode= '" & Voucher.headerText("concode") & "'"
'       g_Conn.Execute strSql
        
        
        '        '12.0��ͷ��չ�Զ�����-����
        '        oHeadElement.setAttribute skeyfld, sID
        '        Set fldextends = objExtend.getVoucherExtendSaveInfo(g_Conn, MainTable & "_ExtraDefine")
        '        objExtend.SavebyInsert oHeadElement, MainTable & "_ExtraDefine", g_Conn, fldextends, , skeyfld
        '
        '�޸�
    Else
        '�޸�ǰ������
        g_Conn.Execute "update " & MainTable & " set " & HeadPKFld & " = " & HeadPKFld & vbCrLf & " where convert(nchar,convert(money,ufts),2)  = " & OldTimeStamp & " and " & HeadPKFld & "=" & lngVoucherID, lngAct

        If lngAct = 0 Then

            If m_bTrans Then
                g_Conn.RollbackTrans
                m_bTrans = False
            End If

            MsgBox GetString("U8.DZ.JA.Res290"), vbInformation, GetString("U8.DZ.JA.Res030")

            Exit Sub

        End If
        
        '        If rs.State = adStateOpen Then rs.Close
        '        rs.Open "select " & strcCode & " from " & MainTable & " where " & HeadPKFld & " = " & lngVoucherID, g_Conn, adOpenForwardOnly, adLockReadOnly
        '        If Not rs.EOF Then
        '            If Voucher.headerText(strcCode) <> rs.Fields(strcCode).Value Then
        '                cSysBarCode = objBillBarCode.GeneralBarCode(g_Conn, g_oLogin.cAcc_Id, gstrCardNumber, Voucher.headerText(strcCode), "jc", "st1", cSplit)
        '                HeadData.Item(1).Item("csysbarcode").Value = cSysBarCode
        '            Else
        '                cSysBarCode = objBillBarCode.GeneralBarCode(g_Conn, g_oLogin.cAcc_Id, gstrCardNumber, Voucher.headerText(strcCode), "jc", "st1", cSplit)
        '                cSysBarCode = Voucher.headerText("csysbarcode")
        '            End If
        '
        '        End If
        
'        strSql = "select isnull(appprice,0) as appprice  from  HY_FYSL_Payment where id='" & Voucher.headerText("id") & "'"
'        Set rs = New ADODB.Recordset
'        rs.Open strSql, g_Conn
'        If Not rs.EOF Then
'
'
'       strSql = " update HY_FYSL_Contract set totalappmoney=isnull(totalappmoney,0)+" & Voucher.headerText("appprice") - rs.Fields("appprice") & " where  ccode= '" & Voucher.headerText("concode") & "'"
'       g_Conn.Execute strSql
'
'        End If
        
        Set Rult = DMO.Update(g_Conn, HeadData)
        
        '12.0��ͷ��չ�Զ�����-update
        '        Set fldextends = objExtend.getVoucherExtendSaveInfo(g_Conn, MainTable & "_ExtraDefine")
        '        objExtend.SavebyUpdate oHeadElement, MainTable & "_ExtraDefine", skeyfld, g_Conn, fldextends
    End If

    If Rult.Succeed = False Then

        If m_bTrans Then
            g_Conn.RollbackTrans
            m_bTrans = False
        End If

        MsgBox Rult.MsgCode & "," & GetString("U8.DZ.JA.Res110"), vbInformation, GetString("U8.DZ.JA.Res030")

        Exit Sub

    End If    '*************************
    '��д��ʷ����
      

     
     
    '*************************



   
    

    '��ͷ��������
    Dim ErrDesc    As String

    Dim blnsaveAcc As Boolean

    Dim strXML     As String

    Dim objDoc     As New MSXML2.DOMDocument

    strXML = Voucher.GetAccessoriesInfo(ErrDesc)

    If ErrDesc <> "" Then

        If m_bTrans Then
            g_Conn.RollbackTrans
            m_bTrans = False
        End If

        MsgBox ErrDesc, vbInformation, GetString("U8.DZ.JA.Res030")

        Exit Sub

    End If

    Dim m_oServer As New UFVoucherServer85.clsVoucherTemplate

    If mOpStatus = ADD_MAIN Then
   
        Call objDoc.loadXML(strXML)
        Call objDoc.documentElement.setAttribute("VoucherTypeID", gstrCardNumber)
        Call objDoc.documentElement.setAttribute("VoucherID", sID)
        strXML = objDoc.xml
        Set objDoc = Nothing
    End If

    Call m_oServer.SaveAccessories(strXML, g_Conn, ErrDesc)

    If ErrDesc <> "" Then

        If m_bTrans Then
            g_Conn.RollbackTrans
            m_bTrans = False
        End If

        MsgBox ErrDesc, vbInformation, GetString("U8.DZ.JA.Res030")

        Exit Sub

    End If

    'End If

    'update HY_DZ_Borrowins set iqtyout=iquantity,iqtyout2=inum where id=11
    'enum by modify
    '    If gcCreateType = "�ڳ�����" Then
    '        strSql = "update " & DetailsTable & " set iqtyout=iquantity,iqtyout2=inum where id = " & sID    '" and cCreateType = '�ڳ�����'"
    '        g_Conn.Execute strSql
    '    End If

    '���鵥�ݱ��
    strVoucherNo = Voucher.headerText("cCODE")
    Set oDomHead = Voucher.GetHeadDom

    bReGetNO = False
    
retry:

    If rs.State = adStateOpen Then rs.Close
    If mOpStatus = ADD_MAIN Then
        rs.Open "select cCode from " & MainTable & " where cCode='" & strVoucherNo & "' and id<>" & sID, g_Conn, 1, 1
    ElseIf mOpStatus = MODIFY_MAIN Then
        rs.Open "select cCode from " & MainTable & " where cCode='" & strVoucherNo & "' and ID<>" & Voucher.headerText("ID"), g_Conn, 1, 1
    End If

    If Not rs.EOF Then
        If Not BOGetVoucherNO(g_Conn, oDomHead, gstrCardNumber, sError, strVoucherNo, oDomFormat, False, False, bRepeatRedo, False) Then
            If m_bTrans Then
                g_Conn.RollbackTrans
                m_bTrans = False
            End If

            MsgBox sError, vbInformation, GetString("U8.DZ.JA.Res030")

            Exit Sub

        End If

        If bRepeatRedo Then
            If Voucher.headerText("cCODE") <> strVoucherNo Then
                Voucher.headerText("cCODE") = strVoucherNo
                g_Conn.Execute "Update " & MainTable & " set cCode='" & strVoucherNo & "' where ID='" & sID & "'"
            End If

            GoTo retry:
        Else

            If m_bTrans Then
                g_Conn.RollbackTrans
                m_bTrans = False
            End If

            MsgBox GetString("U8.DZ.JA.Res1080"), vbInformation, GetString("U8.DZ.JA.Res030")

            Exit Sub

        End If

        '        Voucher.headerText("cCODE") = strVoucherNo
        '        g_Conn.Execute "Update " & MainTable & " set cCode='" & strVoucherNo & "' where ID='" & sID & "'"
        '        Call BOGetVoucherNO(g_Conn, oDomHead, gstrCardNumber, sError, strVoucherNo, oDomFormat, False, , , True)

        bReGetNO = True
    Else

        If Not BOGetVoucherNO(g_Conn, oDomHead, gstrCardNumber, sError, strVoucherNo, oDomFormat, False, False, True, True) Then
            If m_bTrans Then
                g_Conn.RollbackTrans
                m_bTrans = False
            End If

            MsgBox sError, vbInformation, GetString("U8.DZ.JA.Res030")

            Exit Sub

        Else
            '            If Voucher.headerText("cCODE") <> strVoucherNo Then
            '                Voucher.headerText("cCODE") = strVoucherNo
            '                g_Conn.Execute "Update " & MainTable & " set cCode='" & strVoucherNo & "' where ID='" & sID & "'"
            '            End If
        End If
    End If

    rs.Close

    Dim tmpDomHead As New DOMDocument

    Dim tmpDomBody As New DOMDocument

    Dim MaxRowNO   As Long

    Dim t          As Integer
    

  
    If m_bTrans Then
        g_Conn.CommitTrans
        m_bTrans = False
    End If

    bAlter = False

    MsgBox GetString("U8.DZ.JA.Res120"), vbInformation, GetString("U8.DZ.JA.Res030")

    '**********************************************************
    '����(����/����),ɾ������ȫ�ֱ���
    '����,��ȡlngvoucherid
    '**********************************************************

    If mOpStatus = ADD_MAIN Then
        lngVoucherID = sID
        pageCount = pageCount + 1
        PageCurrent = PageCurrent + 1
    End If

    '**********************************************************
    '����,�޸�,ɾ�� ���µ���״̬
    '**********************************************************
    mOpStatus = SHOW_ALL
    Voucher.VoucherStatus = VSNormalMode
    Call ExecSubRefresh
    Call SetCtlStyle(Me, Voucher, Me.Toolbar, Me.UFToolbar, mOpStatus)

    isSavedOK = True                                       'by zhangwchb 20110829 �ύ����

    Exit Sub

Err_Handler:

    If m_bTrans Then
        g_Conn.RollbackTrans
        m_bTrans = False
    End If

    MsgBox Err.Description & "," & GetString("U8.DZ.JA.Res110"), vbInformation, GetString("U8.DZ.JA.Res030")

End Sub

Private Sub ComTemplatePRN_Click()

    m_strVT_PRN_ID = CStr(ComTemplatePRN.ItemData(ComTemplatePRN.ListIndex)) 'Left(Me.ComTemplatePRN.Text, InStr(1, Me.ComTemplatePRN.Text, " ") - 1)

End Sub


'by liwqa
Private Sub ComTemplateShow_Click()

    '����ģ��δ�仯ֱ���˳�
    'If m_strVT_ID = Mid(ComTemplateShow.Text, 1, InStr(1, ComTemplateShow.Text, " ") - 1) Or bInitForm = True Then Exit Sub
    If CStr(m_strVT_ID) = CStr(ComTemplateShow.ItemData(ComTemplateShow.ListIndex)) Or bInitForm = True Then Exit Sub
    
    If Voucher.BodyRows > 0 And Voucher.VoucherStatus <> VSNormalMode Then
        MsgBox GetResString("U8.ST.USKCGLSQL.frmbill.03350"), vbOKOnly + vbExclamation, GetResString("U8.ST.USKCGLSQL.modmain.03048") '�����Ѿ���������,���������ģ���л�!
        ComTemplateShow.ListIndex = sPreVTID
        UFToolbar.RefreshCombobox
        Exit Sub
    End If
    
    Dim domHead As New DOMDocument, domBody As New DOMDocument
    Screen.MousePointer = vbHourglass
    m_strVT_ID = ComTemplateShow.ItemData(ComTemplateShow.ListIndex) 'Mid(ComTemplateShow.Text, 1, InStr(1, ComTemplateShow.Text, " ") - 1)
    Voucher.headerText("ivtid") = m_strVT_ID
    Me.Voucher.getVoucherDataXML domHead, domBody
    ' * �������ݺ�̨�������
    '
    If objVoucherTemplate Is Nothing Then _
            Set objVoucherTemplate = _
            New UFVoucherServer85.clsVoucherTemplate
    SetTemplateData
    Me.Voucher.setVoucherDataXML domHead, domBody
    SetStamp Voucher
    Screen.MousePointer = vbDefault
    Call Form_Resize
End Sub


Private Sub Form_Activate()

    Call SetLayOut

End Sub

'����������ʾ
Private Sub SetLayOut()

    Me.UFToolbar.Move 0, 0, Me.ScaleWidth

    '��Ҫ�������Զ�������С
    Voucher.ControlAutoScales = AutoBoth

    PicTitle.Move 0, 0, Me.ScaleWidth
    Me.ComTemplatePRN.Move Me.PicTitle.Width - Me.ComTemplatePRN.Width, (Me.PicTitle.Height - Me.ComTemplatePRN.Height) / 2
    Me.ComTemplateShow.Move Me.PicTitle.Width - Me.ComTemplateShow.Width, (Me.PicTitle.Height - Me.ComTemplateShow.Height) / 2
    Me.LblTemplate.Move Me.PicTitle.Width - Me.ComTemplatePRN.Width - Me.LblTemplate.Width, Me.ComTemplatePRN.Top + 20
    LabTitle.Move (Me.PicTitle.Width - Me.LabTitle.Width) / 2, (Me.PicTitle.Height - Me.LabTitle.Height) / 2
    PicTitle.Visible = False


    Me.Voucher.Move 0, Me.PicTitle.Height, Me.ScaleWidth, Me.ScaleHeight - Me.PicTitle.Height


    If wfcBack > 0 Then
        wfcBack = wfcBack - 1
        If wfcBack = 0 Then
            Call ExecSubRefresh
        End If
    End If
End Sub

'��ݼ����� -chenliangc
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim strSql As String
'    Dim rs As New ADODB.Recordset
'
'
'    If Shift = vbAltMask Then
'        '��ݼ�ALT + E���
'        If KeyCode = vbKeyE Then
'            If Me.Toolbar.Buttons(sKey_Output).Enabled = True Then
'                ButtonClick sKey_Output
'            End If
'            '����
'        ElseIf KeyCode = vbKeyU Then
'            If Me.Toolbar.Buttons(sKey_Cancelconfirm).Enabled = True Then
'                ButtonClick sKey_Cancelconfirm
'            End If
'            '����
'        ElseIf KeyCode = vbKeyJ Then
'            If Me.Toolbar.Buttons(sKey_Unsubmit).Enabled = True Then
'                ButtonClick sKey_Unsubmit
'            End If
'        ElseIf KeyCode = vbKeyPageUp Then                  '��ҳ
'            If Me.Toolbar.Buttons(sKey_First).Enabled = True Then
'                ButtonClick sKey_First
'            End If
'        ElseIf KeyCode = vbKeyPageDown Then                'ĩҳ
'            If Me.Toolbar.Buttons(sKey_Last).Enabled = True Then
'                ButtonClick sKey_Last
'            End If
'        ElseIf KeyCode = vbKeyC Then                       '�ر�
'            If Me.Toolbar.Buttons(sKey_Close).Enabled = True Then
'                ButtonClick sKey_Close
'            End If
'        ElseIf KeyCode = vbKeyO Then                       '��
'            If Me.Toolbar.Buttons(sKey_Open).Enabled = True Then
'                ButtonClick sKey_Open
'            End If
'        End If
'    ElseIf Shift = vbCtrlMask Then
'        '��ݼ�Ctrl + P��ӡ
'        If KeyCode = vbKeyP Then
'            If Me.Toolbar.Buttons(sKey_Print).Enabled = True Then
'                ButtonClick sKey_Print
'            End If
'            '��ݼ�Ctrl + W��ӡԤ��
'        ElseIf KeyCode = vbKeyW Then
'            If Me.Toolbar.Buttons(sKey_Preview).Enabled = True Then
'                ButtonClick sKey_Preview
'            End If
'            '��ݼ�Ctrl + F3����
'        ElseIf KeyCode = vbKeyF3 Then
'            If Me.Toolbar.Buttons(sKey_Locate).Enabled = True Then
'                ButtonClick sKey_Locate
'            End If
'            '��ݼ�Ctrl + G����
'        ElseIf KeyCode = vbKeyG Then
'            If Me.Toolbar.Buttons(sKey_ReferVoucher).Enabled = True Then
'                ButtonClick sKey_ReferVoucher
'            End If
'            '��������
'        ElseIf KeyCode = vbKeyS Then
'            ' ButtonClick sKey_ReferVoucher
'            If Me.Toolbar.Buttons(sKey_Save).Enabled = True Then
'                ButtonClick sKey_Save
'                Call g_oBusiness.SelectToolbarButton(Toolbar.Buttons(sKey_Add))
'            End If
'
'            '����
'        ElseIf KeyCode = vbKeyZ Then
'            If Me.Toolbar.Buttons(sKey_Discard).Enabled = True Then
'                ButtonClick sKey_Discard
'            End If
'            '�ύ
'        ElseIf KeyCode = vbKeyJ Then
'            If Me.Toolbar.Buttons(sKey_Submit).Enabled = True Then
'                ButtonClick sKey_Submit
'            End If
'            '���
'        ElseIf KeyCode = vbKeyU Then
'            If Me.Toolbar.Buttons(sKey_Confirm).Enabled = True Then
'                ButtonClick sKey_Confirm
'            End If
'
'            '����
'        ElseIf KeyCode = vbKeyF5 Then
'            If Me.Toolbar.Buttons(sKey_Copy).Enabled = True Then
'                ButtonClick sKey_Copy
'            End If
'            '����
'        ElseIf KeyCode = vbKeyG Then
'            'ButtonClick sKey_ReferVoucher
'        ElseIf KeyCode = vbKeyR Then                       'ˢ��
'            If Me.Toolbar.Buttons(sKey_Refresh).Enabled = True Then
'                ButtonClick sKey_Refresh
'            End If
'
'        ElseIf KeyCode = vbKeyD Then                       'ɾ��
'            If Me.Toolbar.Buttons(sKey_Deleterecord).Enabled = True Then
'                ButtonClick sKey_Deleterecord
'
'            End If
'        ElseIf KeyCode = vbKeyA Then                       'ɾ��
'            If Me.Toolbar.Buttons(sKey_Addrecord).Enabled = True Then
'                ButtonClick sKey_Addrecord
'            End If
'        ElseIf KeyCode = vbKeyF4 Then
'            Call ExitForm(0, 0)
'            '��ݼ�Ctrl+E����Ctrl+B���Զ�ָ������,��ⵥ��
'        ElseIf KeyCode = vbKeyE Or KeyCode = vbKeyB Or KeyCode = vbKeyQ Or KeyCode = vbKeyO Then
'            Call GetBatchInfoFun(Voucher, KeyCode, Shift)
'            KeyCode = 0
'        End If
'        ' End If
'    End If
'
'    Select Case KeyCode
'        Case vbKeyF1                                       '����
'            Call LoadHelpId(Me, "15030910")
'        Case vbKeyF5                                       '����
'
'            If Me.Toolbar.Buttons(sKey_Add).Enabled = True Then
'                ButtonClick sKey_Add
'                'Call g_oBusiness.SelectToolbarButton(Toolbar.Buttons(sKey_Add))
'                '         ElseIf Me.Toolbar.Buttons(sKey_Add1).Enabled = True Then
'                '              Call g_oBusiness.SelectToolbarButton(Toolbar.Buttons(sKey_Add1))
'            End If
'            '  ButtonClick sKey_Add
'        Case vbKeyF6                                       '����
'            If Me.Toolbar.Buttons(sKey_Save).Enabled = True Then
'                ButtonClick sKey_Save
'            End If
'        Case vbKeyF8                                       '�޸�
'            If Me.Toolbar.Buttons(sKey_Modify).Enabled = True Then
'                ButtonClick sKey_Modify
'            End If
'        Case vbKeyDelete                                   'ɾ��
'            If Me.Toolbar.Buttons(sKey_Delete).Enabled = True Then
'                ButtonClick sKey_Delete
'            End If
'        Case vbKeyPageUp
'            If Me.Toolbar.Buttons(sKey_Previous).Enabled = True Then
'                ButtonClick sKey_Previous                  '��һҳ
'            End If
'        Case vbKeyPageDown
'            If Me.Toolbar.Buttons(sKey_Next).Enabled = True Then
'                ButtonClick sKey_Next                      '��һҳ
'            End If
'
'
'    End Select
'
'    strSql = "select * from UFMeta_" + g_oLogin.cAcc_Id + "..AA_CustomerButton  where cFormKey='" & gstrCardNumber & "' and cLocaleID='zh-cn' and cHotkey='" & KeyCode & "'   order by cButtonID"
'    Set rs = g_Conn.Execute(strSql)
'    If Not rs.EOF Then
'        ButtonClick CStr(rs!cButtonkey)
'    End If
'    SetToolbarVisible
End Sub

''���ذ����ļ���ID
'Private Sub LoadHelpId(HelpID As String)
'    App.HelpFile = GetWinSysPath & "ufcomsql\��ҵ���\������ҵ\Ӧ��˵��\������ҵ�������.chm"
'    Me.HelpContextID = HelpID
'End Sub

Private Sub Form_Load()
    bInitForm = True
    '�����ʼ������
    Call InitForm
    ' * ���ص�������

    '����frm�˵� modify by chenliangc ��Me.PopMenu1.Visible("retmenu") = False�ŵ�load�¼���
'    Me.PopMenu1.Visible("retmenu") = False
'    '

    '    MsgBox "LoadData"
    If Not LoadData() Then
        '����ģ������Ȩ�޿���
        
''        MsgBox GetString("U8.DZ.JA.Res130"), vbExclamation, GetString("U8.DZ.JA.Res030")
        If bexitload = False Then
            ExitForm 0, 0
        End If
        bexitload = False
        Exit Sub
    End If
    '��������
    Me.Caption = GetString("U8.DZ.JA.Res140")
    Call RegisterMessage                                   '20110812

    'wangfb 11.0ToobarǨ��2012-03-20 ���ÿɼ���
    SetToolbarVisible
    
    inited = True                                          'chenliangc
    bInitForm = False
    Call InitPopMenuText
End Sub
Private Sub InitPopMenuText()
    PopMenu1.Caption("AddR") = GetString("U8.DZ.JA.btn100")
    PopMenu1.Caption("DelR") = GetString("U8.DZ.JA.btn110")
    PopMenu1.Caption("RsLocate") = GetString("U8.SCM.ST.KCGLSQL.FrmMainST.mnuYJCD.lookrow.Caption")
    PopMenu1.Caption("Incor") = GetString("U8.SCM.ST.KCGLSQL.FrmMainST.mnuYJCD.mnuRowAggre.Caption")
    PopMenu1.Caption("batchModify") = GetString("U8.ST.V870.00163")
    PopMenu1.Caption("copyR") = GetString("U8.SCM.ST.KCGLSQL.FrmMainST.mnuYJCD.mnuCopyLine.Caption")
End Sub
'��ʼ������ĳ�������
'��:����ȫ�ֱ����ĳ�ʼ��,���ݿؼ��ĳ�ʼ��,�������ĳ�ʼ��,�����ֱ�
Private Sub InitForm()
    On Error GoTo Err_Handler
    Dim sSource As String
    '����ģ��id,����ģ��ű�����4λ���ڣ������ܼ��ظ���

    '������Ӧ�����¼�
    Me.KeyPreview = True
    '����ʱ���״̬
    mOpStatus = SHOW_ALL

    '��������
    strwhereVou = ""

    Set VchSrv = New clsVouchServer




    '���õ��ݸ�ʽ
    '    VchSrv.SetVouchStyle g_Conn, Voucher, gstrCardNumber

    pageCount = VchSrv.GetPageCount(g_Conn, gstrCardNumber, HeadPKFld, Replace(sAuth_ALL, "HY_FYSL_Payment.id", "V_HY_FYSL_Payment.id"))
    '���б����
    If CLng(sID) <> 0 Then
        UpdatePageCurrent (sID)
    Else
        PageCurrent = pageCount
    End If

    '���ɱ༭��ɫ
    Voucher.DisabledColor = &HE8E8E8

    '��������
    Voucher.ShowSorter = True

    '��������
    Me.Caption = Me.Voucher.TitleCaption

    '��ʼ������pco����,���ݲ���ʱʹ�� USERPCO.VoucherCO
    Dim errStr As String
    Set clsbill = CreateObject("USERPCO.VoucherCO")        'New USERPCO.VoucherCO
    clsbill.IniLogin g_oLogin, errStr
    Set mologin = clsbill.login
    '*************wangfb 11.0ToobarǨ��2012-03-20 start ****************
    Set UFToolbar.Business = g_oBusiness
    Call InitToolBar(mologin, "HY_FYSL_Payment001", Toolbar, UFToolbar, Me.Voucher)
    Call UFToolbar.InitExternalButton("Payment001", mologin.OldLogin)
    Call UFToolbar.SetFormInfo(Me.Voucher, Me)
    '�ڵ���InitExternalButton��������Ҫ���µ���SetToolbar�����������Զ��尴ť���ز���
    UFToolbar.SetToolbar Toolbar
    
    '11.0ToolbarǨ�Ƴ�ʼ����ʾ���ӡģ�尴ť
    Call InitComTemplate
    Call InitComTemplatePRN
'    If IsObject(Toolbar.Buttons("PrintTemplate").Tag) Then
'        Set Toolbar.Buttons("PrintTemplate").Tag.Tag = Me.ComTemplatePRN
'    End If
'    If IsObject(Toolbar.Buttons("ShowTemplate").Tag) Then
'        Set Toolbar.Buttons("ShowTemplate").Tag.Tag = Me.ComTemplateShow
'    End If
    
    '��������ʼ��
    '11.0toolbarǨ�ƣ�������ҵ�񵥾ݱ�׼��֮��ԭ���Ϳ��İ�ť��ʼ��
    'Call Init_Toolbar(Me.Toolbar)
   
    Call ChangeOneFormTbr(Me, Me.Toolbar, Me.UFToolbar)
'    Call SetWFControlBrns(g_oLogin, g_Conn, Me.Toolbar, Me.UFToolbar, gstrCardNumber)
 
    '*************wangfb 11.0ToobarǨ��2012-03-20 end ****************
'
'    '��ʼ���˵�,�˴�����ʹ��call����
'    Call Me.PopMenu1.SubClassMenu(Me)

    '��ȡU8�汾 -chenliangc
    gU8Version = GetU8Version(g_Conn)


Exit_Label:
    On Error GoTo 0
    Exit Sub



    '�ݴ���
Err_Handler:

    ' * �����������Դ
    If Left(Err.Source, 3) = "***" Then
        sSource = Err.Source
    Else
        sSource = "***Sub InitComTemplate of Form frmVoucher"
    End If

    ' * �׳��쳣
    Err.Raise _
            Number:=Err.Number, _
            Source:=sSource, _
            Description:=Err.Description
End Sub
'��ʼ��������ʾģ��
Private Sub InitComTemplate()

    Dim sql As String
    Dim oRecordset As New Recordset
    
    Dim sMessage As String
    Dim i As Integer
    ' * ����Դ
    Dim sSource As String
    
    Dim strAuth As String
    Dim clsRowAuth As U8RowAuthsvr.clsRowAuth
    Set clsRowAuth = New U8RowAuthsvr.clsRowAuth
    clsRowAuth.Init g_Conn.ConnectionString, g_oLogin.cUserId
    strAuth = clsRowAuth.getAuthString("DJMB", gstrCardNumber)
    
    On Error GoTo Err_Handler
    If strAuth = "" Then
        sql = "SELECT VT_ID ,VT_Name " & vbCrLf _
            & "FROM vouchertemplates " & vbCrLf _
            & "WHERE " & vbCrLf _
            & "     vt_cardnumber = '" & gstrCardNumber & "' AND " & vbCrLf _
            & "     vt_templatemode = '0' " & vbCrLf
    Else
        If strAuth = "1=2" Then
            sql = "SELECT VT_ID ,VT_Name " & vbCrLf _
            & "FROM vouchertemplates " & vbCrLf _
            & "WHERE " & vbCrLf _
            & "     vt_cardnumber = '" & gstrCardNumber & "' AND " & vbCrLf _
            & "     vt_templatemode = '0' " & " and " & strAuth & vbCrLf
        Else
            sql = "SELECT VT_ID ,VT_Name " & vbCrLf _
            & "FROM vouchertemplates " & vbCrLf _
            & "WHERE " & vbCrLf _
            & "     vt_cardnumber = '" & gstrCardNumber & "' AND " & vbCrLf _
            & "     vt_templatemode = '0' " & " and vt_id in (" & strAuth & ")" & vbCrLf
        End If
    End If
'    sql = "SELECT VT_ID ,VT_Name " & vbCrLf _
'            & "FROM vouchertemplates " & vbCrLf _
'            & "WHERE " & vbCrLf _
'            & "     vt_cardnumber = '" & gstrCardNumber & "' AND " & vbCrLf _
'            & "     vt_templatemode = '0' " & vbCrLf

    If oRecordset Is Nothing Then _
            Set oRecordset = CreateObject("ADODB.Recordset")

    If oRecordset.State = adStateOpen Then _
            Call oRecordset.Close

    Call oRecordset.Open( _
            sql, _
            g_Conn, _
            adOpenStatic, _
            adLockReadOnly, _
            adCmdText)

    If Not (oRecordset.BOF And oRecordset.EOF) Then
        oRecordset.MoveFirst
        Me.ComTemplateShow.Clear
        dicTemplate.RemoveAll                              'by liwqa Template
        i = 0
        Do Until oRecordset.EOF
            Me.ComTemplateShow.AddItem oRecordset.Fields(1), i
            Me.ComTemplateShow.ItemData(i) = oRecordset.Fields(0)
            dicTemplate.Add CStr(oRecordset.Fields(0)), i
            i = i + 1
            oRecordset.MoveNext
        Loop
    Else
        Exit Sub
    End If

    '    m_FlagVTID = True   'by liwqa
    '
    '    Me.ComTemplateShow.Text = Me.ComTemplateShow.List(0)
    If Me.ComTemplateShow.Count > 0 Then Me.ComTemplateShow.ListIndex = 0
    'Me.ComTemplatePRN.Visible = False
    'Me.ComTemplateShow.Visible = True
    Me.LblTemplate.Caption = GetString("U8.DZ.JA.Res020")
    Me.LblTemplate.Visible = True

Exit_Label:
    On Error GoTo 0
    If Not oRecordset Is Nothing Then
        If oRecordset.State = adStateOpen Then _
                Call oRecordset.Close
    End If
    Set oRecordset = Nothing

    Exit Sub

    '�ݴ���
Err_Handler:
    If Not oRecordset Is Nothing Then
        If oRecordset.State = adStateOpen Then _
                Call oRecordset.Close
    End If
    Set oRecordset = Nothing

    ' * �����������Դ
    If Left(Err.Source, 3) = "***" Then
        sSource = Err.Source
    Else
        sSource = "***Sub InitComTemplate of Form frmVoucher"
    End If

    ' * �׳��쳣
    Err.Raise _
            Number:=Err.Number, _
            Source:=sSource, _
            Description:=Err.Description

End Sub
'��ʼ�����ݴ�ӡģ�� 'by liwqa Template
Private Sub InitComTemplatePRN()
    On Error Resume Next
    Dim rstTmp As New Recordset
    Dim i As Integer

    Dim strAuth As String
    Dim clsRowAuth As U8RowAuthsvr.clsRowAuth
    Set clsRowAuth = New U8RowAuthsvr.clsRowAuth
    clsRowAuth.Init g_Conn.ConnectionString, g_oLogin.cUserId
    strAuth = clsRowAuth.getAuthString("DJMB", gstrCardNumber)
    
    If strAuth = "1=2" Then
        'rstTmp.Open "select VT_ID ,VT_Name  from vouchertemplates where vt_cardnumber=N'" & gstrCardNumber & "' and vt_templatemode=N'1'" & "and  " & strAuth & "", g_Conn, , adOpenStatic, adLockOptimistic
    Else
        If strAuth = "" Then
            rstTmp.Open "select VT_ID ,VT_Name  from vouchertemplates where vt_cardnumber=N'" & gstrCardNumber & "' and vt_templatemode=N'1'", g_Conn, adOpenStatic, adLockOptimistic
        Else
            rstTmp.Open " select VT_ID ,VT_Name  from vouchertemplates where vt_cardnumber=N'" & gstrCardNumber & "' and vt_templatemode=N'1'" & " and vt_id in (" & strAuth & ")", g_Conn, adOpenStatic, adLockOptimistic
        End If
        'rstTmp.Open "select VT_ID ,VT_Name  from vouchertemplates where vt_cardnumber='" & gstrCardNumber & "' and vt_templatemode='1'", g_Conn
    
        If Not (rstTmp.BOF And rstTmp.EOF) Then
            rstTmp.MoveFirst
            Me.ComTemplatePRN.Clear
            dicTemplatePrint.RemoveAll
            i = 0
            Do Until rstTmp.EOF
                Me.ComTemplatePRN.AddItem rstTmp.Fields(1), i
                Me.ComTemplatePRN.ItemData(i) = rstTmp.Fields(0)
                dicTemplatePrint.Add CStr(rstTmp.Fields(0)), i
                i = i + 1
                rstTmp.MoveNext
            Loop
        Else
            Exit Sub
        End If
    End If
    
    '    Me.ComTemplatePRN.Text = Me.ComTemplatePRN.List(0)
    If Me.ComTemplatePRN.Count > 0 Then Me.ComTemplatePRN.ListIndex = 0
    'Me.ComTemplatePRN.Visible = True
    'Me.ComTemplateShow.Visible = False
    Me.LblTemplate.Caption = GetString("U8.DZ.JA.Res010")
    Me.LblTemplate.Visible = True
End Sub
Private Sub InitVoucher()

    Dim oDataSource As Object
    Dim oRecordset As ADODB.Recordset

    On Error GoTo Err_Handler

    '���ñ���


    ' *******************************************************
    ' * ��ȡ��ǰ����ģ��ID (VT_ID) ֵ
    '
    '    Call LoadVTID

    ' *******************************************************
    ' * �������ݺ�̨�������
    '
    If objVoucherTemplate Is Nothing Then _
            Set objVoucherTemplate = _
            New UFVoucherServer85.clsVoucherTemplate



    '    ' ������������Դ����
    Set oDataSource = CreateObject("IDataSource.DefaultDataSource")

    If oDataSource Is Nothing Then
        MsgBox GetString("U8.DZ.JA.Res160"), vbExclamation, GetString("U8.DZ.JA.Res030")
    End If

    Set oDataSource.SetLogin = g_oLogin

    Set Voucher.SetDataSource = oDataSource

    '��ע��:SetTemplateData  ������� Set oDataSource.SetLogin = g_oLogin ֮��, �������ȸ���������Դ��ʼ��
    Call SetTemplateData

    Voucher.LoginObj = g_oLogin
    Voucher.InitDataSource

Exit_Label:
    On Error GoTo 0
    If Not oRecordset Is Nothing Then
        If oRecordset.State = adStateOpen Then _
                Call oRecordset.Close
    End If
    Set oRecordset = Nothing

    Exit Sub
Err_Handler:
    Call ShowErrorInfo( _
            sHeaderMessage:=GetString("U8.DZ.JA.Res180"), _
            lMessageType:=vbInformation, _
            lErrorLevel:=ufsELOnlyHeader _
            )
    GoTo Exit_Label

End Sub

'���ص���ģ�������
Public Function LoadData() As Boolean

    On Error GoTo Err_Handler
    Dim sql As String
    Dim rs As New ADODB.Recordset


    Call InitVoucher

'    If ComTemplateShow.ListCount = 0 Then
'        Call MsgBox( _
'                Prompt:=GetString("U8.DZ.JA.Res190"), _
'                Buttons:=vbExclamation, Title:=GetString("U8.DZ.JA.Res030"))
'        ButtonClick "Refresh"
'        GoTo Exit_Label
'    End If


    '���б����
    If CLng(sID) <> 0 Then
        lngVoucherID = sID


        'ֱ�ӽ��뵥��
    Else

        sql = "select cValue from AccInformation Where cSysId =N'ST' and  cName= N'VouchViewMode'"

        rs.Open sql, g_Conn
        If Not rs.BOF Or Not rs.EOF Then

            If rs.Fields("cValue") = "Last" Then

                lngVoucherID = GetTheLastID(login:=g_oLogin, _
                        oConnection:=g_Conn, _
                        sTable:=MainTable, _
                        sField:=HeadPKFld, _
                        sDataNumFormat:="0")
            Else
                '                If tmpLinkTbl = "" Then '�������� ʱ ��ť״̬���� by zhangwchb 20110809
                lngVoucherID = 0
                '                End If
            End If
        End If
    End If


    Call LoadVoucherData

    LoadData = True

Exit_Label:
    On Error GoTo 0

    Exit Function
Err_Handler:

    LoadData = False

    If g_blnDEBUG_MODE Then
        Call ShowDebugForm( _
                bErrorMode:=True, _
                sProcedure:="Function LoadData of Form frmVoucher")
    End If
    GoTo Exit_Label

End Function
'
'Private Sub LoadVTID()
'
'    Dim sql As String
'    Dim oRecordset As ADODB.Recordset
'
'    'On Error GoTo Err_Handler
'
'    sql = "SELECT [DEF_ID], " _
     '        & "[DEF_ID_PRN] " _
     '        & "FROM [Vouchers] " _
     '        & "WHERE ([CardNumber] = '" & Trim(gstrCardNumber) & "') "
'
'    If oRecordset Is Nothing Then _
     '       Set oRecordset = New ADODB.Recordset
'
'    Call oRecordset.Open( _
     '         Source:=sql, _
     '         ActiveConnection:=g_Conn, _
     '         CursorType:=adOpenStatic, _
     '         LockType:=adLockReadOnly, _
     '         Options:=adCmdText)
'
'    If oRecordset.RecordCount < 1 Then
'        Call Err.Raise( _
         '             Number:=vbObjectError + 512 + 6002, _
         '             Description:=GetString("U8.DZ.JA.Res200"))
'    End If
'
'    '�������������û�б����ȡĬ��ģ�� by liwq
'    If m_strVT_ID = "" Or m_strVT_ID = "0" Then
'        m_strVT_ID = Null2Something( _
         '                     vTarget:=oRecordset.Fields("DEF_ID").Value, _
         '                     vReplace:=0)
'    End If
'    m_strVT_PRN_ID = Null2Something( _
     '                     vTarget:=oRecordset.Fields("DEF_ID_PRN").Value, _
     '                     vReplace:=0)
'
'Exit_Label:
'    On Error GoTo 0
'    If Not oRecordset Is Nothing Then
'        If oRecordset.State = adStateOpen Then _
         '           Call oRecordset.Close
'    End If
'    Set oRecordset = Nothing
'
'    Exit Sub
'Err_Handler:
'    If Not oRecordset Is Nothing Then
'        If oRecordset.State = adStateOpen Then _
         '           Call oRecordset.Close
'    End If
'    Set oRecordset = Nothing
'
'    Err.Raise _
     '            Number:=Err.Number, _
     '            Source:="Sub LoadVTID of Form frmVoucher", _
     '            Description:=Err.Description
'
'End Sub

Private Sub SetTemplateData()

    Dim oRecordset As ADODB.Recordset                      'ģ�����ݼ�¼��
    Dim sAuth As String                                    '�ֶ�Ȩ���ַ���
    Dim sNumber As String                                  '���ݱ�Ź����ַ���
    Dim lColor1 As Long
    Dim lColor2 As Long

    'On Error GoTo Err_Handler

    ' *******************************************************
    ' * �õ�����ģ������,����ָ���ĵ������ͺ�ģ��IDȡ�õ�������
    '
    Set oRecordset = objVoucherTemplate.GetTemplateData2( _
            conn:=g_Conn, _
            sBillName:=gstrCardNumber, _
            vTemplateID:=m_strVT_ID)


    ' *******************************************************
    ' * ȡ��ָ������Ա�Ե�ǰ���ݵ�Ȩ��,�Ա����Ȩ�޿���
    ' *
    ' * ע:
    ' *     1)  ÿ�λ�ģ���ʱ����ҪӦ��һ��
    sAuth = objVoucherTemplate.getAuthString( _
            ologin:=g_oLogin, _
            nID:=gstrCardNumber)


    ' *******************************************************
    ' * ȡ�� Rule ��ɫ
    '
    Call objVoucherTemplate.GetRuleColor( _
            strConn:=g_Conn, _
            clrDisable:=lColor1, _
            clrNeed:=lColor2)


    ' *******************************************************
    ' * ���õ��ݿؼ����ɼ�
    '
    Voucher.Visible = False


    ' *******************************************************
    ' * ע:
    ' *     1)  SetVoucherAuth ���������� SetTemplateData ��
    ' *         ��ǰʹ��
    '
    Call Voucher.SetVoucherAuth(sAuth)
    Call Voucher.SetRuleColor(lColor1, lColor2)

    '    FormatVouchList oRecordset    ' ������ģ�徫������


    Call Voucher.SetTemplateData(oRecordset)

    Voucher.Visible = True

    'by liwq
    If gcCreateType = "�ڳ�����" Then
        Me.LabTitle = GetString("U8.DZ.JA.Res170")
    Else
        Me.LabTitle = Voucher.TitleCaption                 ' GetString("U8.DZ.JA.Res140")
    End If

    Me.UFFrmCaptionMgr.Caption = Voucher.TitleCaption
    'ToolbarǨ�� wangfb 2012-03-30 title�Ƶ�Voucher��
    'Voucher.TitleCaption = ""
    Me.LabTitle = ""


    ' *******************************************************
    ' * ���õ��ݱ�� Rule
    '
    If objBill Is Nothing Then Set objBill = New UFBillComponent.clsBillComponent
    Call objBill.InitBill(g_Conn.ConnectionString, gstrCardNumber)    ' m_bill(cboBill.ListIndex))
    sNumber = objBill.GetBillFormat
    Voucher.SetBillNumberRule sNumber

Exit_Label:
    On Error GoTo 0

    Exit Sub
Err_Handler:
    If g_blnDEBUG_MODE Then
        Call ShowDebugForm( _
                bErrorMode:=True, _
                sProcedure:="Sub SetTemplateData of Form frmVoucher")
    End If
    Err.Raise _
            Number:=Err.Number, _
            Source:="Sub SetTemplateData of Form frmVoucher", _
            Description:=Err.Description

End Sub

'by zhangwchb 20110718 ��չ�ֶ�
Public Sub GetExtendedInfo(ByRef conn As Connection, CardNum As String)

    Set oExtend = CreateObject("VoucherExtendService.ClsExtendServer")
    Call oExtend.GetExtendInfo(conn, CardNum, "T", sExtendField, sExtendJoinSQL)
    Call oExtend.GetExtendInfo(conn, CardNum, "B", sExtendBodyField, sExtendBodyJoinSQL)

End Sub

Private Sub LoadVoucherData()

    Dim oDomHead As New DOMDocument
    Dim oDomBody As New DOMDocument
    Dim oelement As IXMLDOMElement
    Dim tSql As String

    Dim sql As String
    Dim oRecordset As New Recordset

    Dim sMessage As String

    ' * ����Դ
    Dim sSource As String
    On Error GoTo Err_Handler
    
    numappprice = 0

    Screen.MousePointer = vbHourglass

    oRecordset.CursorLocation = adUseClient

    'by zhangwchb 20110718 ��չ�ֶ�
'    Call GetExtendedInfo(g_Conn, gstrCardNumber)

    sql = "SELECT *,'' AS editprop " & vbCrLf _
            & " FROM  " & MainView & vbCrLf _
            & " WHERE  " & HeadPKFld & "= " & lngVoucherID & " "

    If oRecordset Is Nothing Then _
            Set oRecordset = CreateObject("ADODB.Recordset")

    If oRecordset.State = adStateOpen Then _
            Call oRecordset.Close

    Call oRecordset.Open( _
            sql, _
            g_Conn, _
            adOpenStatic, _
            adLockReadOnly, _
            adCmdText)

    If oRecordset.EOF Then
        mOpStatus = SHOW_NOTHING
    Else
       ' gcCreateType = vFieldVal(oRecordset.Fields("cCreateType"))
        Call setTemplate(CStr(Null2Something(oRecordset.Fields("ivtid").Value)))    'by liwqa Template
        Call InitVoucher
        Call SetLayOut
    End If

    ' * ת���� XML ���ݸ�ʽ
    oRecordset.Save oDomHead, adPersistXML

    'by zhangwchb 20110718 ��չ�ֶ�
    sql = "SELECT *,'' AS editprop " & sExtendBodyField & vbCrLf _
            & "FROM " & DetailsView & " " & sExtendBodyJoinSQL & vbCrLf _
            & "WHERE " & HeadPKFld & " = " & lngVoucherID & " "

    If oRecordset Is Nothing Then _
            Set oRecordset = CreateObject("ADODB.Recordset")

    If oRecordset.State = adStateOpen Then _
            Call oRecordset.Close

    Call oRecordset.Open( _
            sql, _
            g_Conn, _
            adOpenStatic, _
            adLockReadOnly, _
            adCmdText)

'
'    ' * ת���� XML ���ݸ�ʽ
'    oRecordset.Save oDomBody, adPersistXML
    Voucher.setVoucherDataXML oDomHead, oDomBody
    SetStamp Voucher

    mOpStatus = SHOW_ALL
    Call SetCtlStyle(Me, Voucher, Me.Toolbar, Me.UFToolbar, mOpStatus)

    '��ȡʱ���
    OldTimeStamp = GetTimeStamp(g_Conn, MainTable, lngVoucherID)

Exit_Label:
    On Error GoTo 0
    Screen.MousePointer = vbDefault

    If Not oRecordset Is Nothing Then
        If oRecordset.State = adStateOpen Then _
                Call oRecordset.Close
    End If
    Set oRecordset = Nothing

    Exit Sub
Err_Handler:
    Screen.MousePointer = vbDefault

    If Not oRecordset Is Nothing Then
        If oRecordset.State = adStateOpen Then _
                Call oRecordset.Close
    End If
    Set oRecordset = Nothing

    ' * �����������Դ
    If Left(Err.Source, 3) = "***" Then
        sSource = Err.Source
    Else
        sSource = "***Sub LoadVoucherData of Form frmVoucher"
    End If

    ' * �׳��쳣
    Err.Raise _
            Number:=Err.Number, _
            Source:=sSource, _
            Description:=Err.Description

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = m_Cancel
'    DropTable sTmpTableName
    sTmpTableName = ""
End Sub

'ÿ�����嶼��Ҫ���������Cancel��UnloadMode�Ĳ����ĺ�����QueryUnload�Ĳ�����ͬ��
Public Sub ExitForm(Cancel As Integer, UnloadMode As Integer)
    If mOpStatus <> SHOW_ALL Then
        If MsgBox(GetString("U8.DZ.JA.Res210"), vbYesNo, GetString("U8.DZ.JA.Res030")) = vbYes Then
            m_Cancel = 1
            Cancel = 1
            Exit Sub
        Else
            m_Cancel = 0
        End If
    Else
        m_Cancel = 0
    End If


    Unload Me
    Cancel = m_Cancel
    UnloadMode = m_UnloadMode

    lngVoucherID = 0
    pageCount = 0
    PageCurrent = 0
    sID = 0
    sAutoId = 0
    '    tmpLinkTbl = "" ''�������� ʱ ��ť״̬���� by zhangwchb 20110809
    Call UnRegisterMessage                                 '20110812

End Sub

Private Sub Form_Resize()
    Call SetLayOut
End Sub

'�Ҽ��˵��¼�����
Private Sub PopMenu1_MenuClick(sMenuKey As String)

    On Error Resume Next

    Select Case LCase(sMenuKey)

        Case "addr"                                        '����
            Voucher.AddLine Voucher.BodyRows + 1

        Case "delr"                                        'ɾ��
            Voucher.DelLine Voucher.row

        Case "rslocate"                                    '��λ��¼
            Voucher.ShowFindDlg

        Case "incor"                                       '�ϲ���ʾ
            Call Execincor

        Case "batchmodify"                                 '����
            Call ExecBathModify

        Case "copyr"                                       '������
            Voucher.DuplicatedLine Voucher.row
            Voucher.bodyText(Voucher.row, "AutoID") = ""
            Voucher.bodyText(Voucher.row, "cbsysbarcode") = "" '�и�����Ҫ����������������
    End Select

End Sub

Public Sub ExecBathModify()    '�����޸�
    On Error GoTo ErrHandle
    '    Screen.MousePointer = vbHourglass
    '
    '    Dim m_oDataSource As Object
    '    Set m_oDataSource = CreateObject("IDataSource.DefaultDataSource")
    '
    '    If m_oDataSource Is Nothing Then
    '        MsgBox getstring("U8.DZ.JA.Res220"), vbExclamation, getstring("U8.DZ.JA.Res030")
    '        Exit Sub
    '    End If
    '
    '    Set m_oDataSource.SetLogin = g_oLogin
    '    Set Voucher.SetDataSource = m_oDataSource
    '
    '    Screen.MousePointer = vbDefault
    
    
    Voucher.ShowBatchModify

    Exit Sub
ErrHandle:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Sub

Public Sub Execincor()    '�ϲ���ʾ

    On Error GoTo ErrHandle

    Screen.MousePointer = vbHourglass
    Voucher.ProtectUnload2

    Dim m_oDataSource As Object
    Set m_oDataSource = CreateObject("IDataSource.DefaultDataSource")

    If m_oDataSource Is Nothing Then
        MsgBox GetString("U8.DZ.JA.Res220"), vbExclamation, GetString("U8.DZ.JA.Res030")
        Exit Sub
    End If

    Set m_oDataSource.SetLogin = g_oLogin
    Set Voucher.SetDataSource = m_oDataSource

    Screen.MousePointer = vbDefault
    Voucher.SHowAggregateSetupDlg

    Exit Sub

ErrHandle:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Sub
Private Sub SaveVoucherTemplateInfo()
    Voucher.SaveVoucherTemplateInfo

End Sub

Private Sub ShowVoucherDesign()

    If Voucher.ShowVoucherDesign = True Then               '��true������£�false��������

        '        Dim domHead As New DOMDocument
        '
        '        Dim domBody As New DOMDocument

        '        Voucher.getVoucherDataXML domHead, domBody

        '�������õ��ݸ�ʽ

        Call SetTemplateData
        
        Call LoadVoucherData

        '        Voucher.setVoucherDataXML domHead, domBody

        SetLayOut
    End If

    '   Voucher.ShowVoucherDesign
End Sub

'�������еİ�ť�¼�

Public Sub ButtonClick(strbuttonkey As String)

    Dim rs       As New ADODB.Recordset

    Dim strSql   As String

    Dim obj      As Object

    Dim sMessage As String

    Dim oDomHead As New DOMDocument, oDomBody As New DOMDocument

    On Error GoTo Err_Handler

    strSql = "select * from UFMeta_" + g_oLogin.cAcc_Id + "..AA_CustomerButton  where cFormKey='" & gstrCardNumber & "' and cLocaleID='zh-cn' and cButtonKey='" & strbuttonkey & "'  order by cButtonID"
    Set rs = g_Conn.Execute(strSql)

    If Not rs.EOF Then
        Set obj = CreateObject(CStr(rs!cCustomerObjectName))

        If VBA.LCase(CStr(rs!cEnableAsKey)) = VBA.LCase(sKey_Add) Then

            If ZwTaskExec(g_oLogin, AuthAdd, 1) = False Then Exit Sub
            Call ExecSubAdd
            Call obj.ButtonClick(CStr(rs!cButtonkey), CStr(rs!cButtonType), Me, Me.Voucher)
            Call ZwTaskExec(g_oLogin, AuthAdd, 0)
        End If

    Else

        'Ĭ����������
        If LCase(strbuttonkey) = "add" Then strbuttonkey = sKey_Add2

        Select Case strbuttonkey

                '��ӡ
            Case sKey_Print

                If ZwTaskExec(g_oLogin, AuthPrint, 1) = False Then Exit Sub
                Call ExecSubVoucherPrint(oConnection:=g_Conn, oVoucher:=Voucher, sBillNumber:=gstrCardNumber, sTemplateID:=m_strVT_PRN_ID, bPreview:=False)
                ZwTaskExec g_oLogin, AuthPrint, 0

                'Ԥ��
            Case sKey_Preview

                If ZwTaskExec(g_oLogin, AuthPrint, 1) = False Then Exit Sub
                Call ExecSubVoucherPrint(oConnection:=g_Conn, oVoucher:=Voucher, sBillNumber:=gstrCardNumber, sTemplateID:=m_strVT_PRN_ID, bPreview:=True)
                ZwTaskExec g_oLogin, AuthPrint, 0

                '���
            Case sKey_Output

                If ZwTaskExec(g_oLogin, AuthOut, 1) = False Then Exit Sub
                Call ExportVoucherDataToFile(oConnection:=g_Conn, oVoucher:=Voucher, sBillNumber:=gstrCardNumber, sTemplateID:=m_strVT_PRN_ID)
                ZwTaskExec g_oLogin, AuthOut, 0

            Case sKey_VoucherDesign
                ShowVoucherDesign

            Case sKey_SaveVoucherDesign
                Call SaveVoucherTemplateInfo

                '��һҳ
            Case sKey_First
                Call ExecSubPageFirst

                '��һҳ
            Case sKey_Previous
                Call ExecSubPageUp

                '��һҳ
            Case sKey_Next
                Call ExecSubPageDown

                '���һҳ
            Case sKey_Last
                Call ExecSubPageLast

            Case sKey_RefVoucher                         'zhangwchb 20110714 ���ӹ�������
                ShowRefVouchers False

            Case "tlbLinkAllVouch"
                ShowRefVouchers True

                '����
            Case sKey_Add2
                '����ʱ����Ȩ��
                '                If Me.ComTemplateShow.ListCount = 0 Then
                '                    'MsgBox "�Բ�����û��ģ��Ȩ�ޣ��޷����ӣ�", vbInformation, pustrMsgTitle
                '                    MsgBox GetResString("U8.ST.Default.00156"), vbInformation, GetResString("U8.SCM.ST.KCGLSQL.FrmMainST.rop.Caption") 'pustrMsgTitle
                '                    Exit Sub
                '                End If
                
                If ZwTaskExec(g_oLogin, AuthAdd, 1) = False Then Exit Sub
                Call initCustomRelation                    '20110822 by zhangwchb
                Call ExecSubAdd
                Call ZwTaskExec(g_oLogin, AuthAdd, 0)

            Case sKey_Add1                                 'strAdd1
                '����ʱ����Ȩ��
                '                If Me.ComTemplateShow.ListCount = 0 Then
                '                    'MsgBox "�Բ�����û��ģ��Ȩ�ޣ��޷����ӣ�", vbInformation, pustrMsgTitle
                '                    MsgBox GetResString("U8.ST.Default.00156"), vbInformation, GetResString("U8.SCM.ST.KCGLSQL.FrmMainST.rop.Caption") 'pustrMsgTitle
                '                    Exit Sub
                '                End If
                
                If ZwTaskExec(g_oLogin, AuthAdd, 1) = False Then Exit Sub
                Call initCustomRelation                    '20110822 by zhangwchb
                Call ExecSubAdd(0)
                Call ZwTaskExec(g_oLogin, AuthAdd, 0)

                '�޸�
            Case sKey_Modify

                If Check_Auth = False Then Exit Sub

                If ZwTaskExec(g_oLogin, AuthModify, 1) = False Then Exit Sub

                '���������޸ĵ��� 20110817 by zhangwchb
                '                If Not CheckVerModify Then
                '
                '                    Exit Sub
                '
                '                ElseIf Voucher.headerText("iswfcontrolled") = 1 And Voucher.headerText("iverifystate") = 1 Then
                '
                '                    If ZwTaskExec(g_oLogin, AuthModifyWF, 1) = False Then
                '
                '                        Exit Sub
                '
                '                    End If
                '                End If
                ''                If bCheckUser = True Then
                '                    If CheckUserAuth(g_Conn.ConnectionString, g_oLogin.cUserId, Voucher.headerText(StrcMaker), "W") = False Then
                '                        MsgBox GetString("U8.pu.VoucherCommon.00163"), vbInformation, GetString("U8.DZ.JA.Res030")
                '
                '                        Exit Sub
                '
                '                    End If
                '                End If

                If ExecFunCompareUfts = True Then
                    '                    Call setAllDisable
                    'Call initCustomRelation                '20110822 by zhangwchb
                    ExecSubModify
                End If

                Call ZwTaskExec(g_oLogin, AuthModify, 0)

                'ɾ��
            Case sKey_Delete

                If Check_Auth = False Then Exit Sub

                If ZwTaskExec(g_oLogin, AuthDelete, 1) = False Then Exit Sub
              
                '                If bCheckUser = True Then
                '                    If CheckUserAuth(g_Conn.ConnectionString, g_oLogin.cUserId, Voucher.headerText(StrcMaker), "W") = False Then
                '                        MsgBox GetString("U8.OM.VoucherControl.00002"), vbInformation, GetString("U8.DZ.JA.Res030")
                '
                '                        Exit Sub
                '
                '                    End If
                '                End If

                If ExecFunCompareUfts = True Then ExecSubDelete
                Call ZwTaskExec(g_oLogin, AuthDelete, 0)

                '��������(��ͬ)
            Case "Prorefer1"
            
                frmExcelDR.Show vbModal
                frmExcelDR.ZOrder 0

                Call ZwTaskExec(g_oLogin, AuthProrefer, 0)
            '��������(����)
            Case "Prorefer2"
            
                If ZwTaskExec(g_oLogin, AuthProrefer1, 1) = False Then Exit Sub
                Call initCustomRelation                    '20110822 by zhangwchb
                Call ExecSubAdd

                If ReferVoucheng Then
                    '��ʾ���շ��ص�����
                    '        Debug.Print gDomReferHead.xml
                    '        Debug.Print gDomReferBody.xml
                    '��������
                    ProcessDataeng Voucher
                    '�û�������ť
                    Me.Toolbar.Buttons("Prorefer1").Enabled = False
                    Me.Toolbar.Buttons("Prorefer2").Enabled = False
                    'Me.Toolbar.Buttons("Prorefer3").Enabled = False
                    Me.UFToolbar.RefreshEnable
                Else
                    
                    Call ExecSubRefresh
                    mOpStatus = SHOW_ALL

                     Call SetCtlStyle(Me, Voucher, Me.Toolbar, Me.UFToolbar, mOpStatus)
                End If

                Call ZwTaskExec(g_oLogin, AuthProrefer1, 0)
                
                '��������(��Ŀ)
            Case "Prorefer3"
            
                If ZwTaskExec(g_oLogin, AuthProrefer, 1) = False Then Exit Sub
                Call initCustomRelation                    '20110822 by zhangwchb
                Call ExecSubAdd

                If ReferVouchpro Then
                    '��ʾ���շ��ص�����
                    '        Debug.Print gDomReferHead.xml
                    '        Debug.Print gDomReferBody.xml
                    '��������
                    ProcessDatapro Voucher
                    '�û�������ť
                    Me.Toolbar.Buttons("Prorefer1").Enabled = False
                    Me.Toolbar.Buttons("Prorefer2").Enabled = False
                    'Me.Toolbar.Buttons("Prorefer3").Enabled = False
                    Me.UFToolbar.RefreshEnable
                    Else
                    
                    Call ExecSubRefresh
                    mOpStatus = SHOW_ALL

                     Call SetCtlStyle(Me, Voucher, Me.Toolbar, Me.UFToolbar, mOpStatus)
                 
                End If

                Call ZwTaskExec(g_oLogin, AuthProrefer, 0)
                  
 
                '�Ƶ�-���۶���
            Case sKey_CreateSAVoucher

                '�Ƶ�ǰ�Ƚ�ʱ���,��������ʱ���
                If SdFlg(lngVoucherID, OldTimeStamp) <> "" Then
                    ReDim varArgs(0)
                    varArgs(0) = Voucher.headerText("cCODE")
                    MsgBox GetStringPara("U8.DZ.JA.Res230", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    '                MsgBox "���ݺ�Ϊ" & Voucher.headerText("cCODE") & "�ĵ������������ε��ݣ�", vbInformation, GetString("U8.DZ.JA.Res030")
                    Exit Sub

                End If

                Voucher.getVoucherDataXML oDomHead, oDomBody    '�������
                ExecmakeDom oDomHead, oDomBody, g_Conn     '��֯����

                g_Conn.BeginTrans

                If ExecCreateVoucher(oDomHead, oDomBody, g_Conn, g_oLogin, ����, ���۶���) Then    '�Ƶ�����д����״̬
                    g_Conn.CommitTrans
                Else
                    g_Conn.RollbackTrans                   '�Ƶ�ʧ�ܻع�
                End If

                Call ExecSubRefresh

                '�Ƶ�-�ɹ�����
            Case sKey_CreatePUVoucher

                '�Ƶ�ǰ�Ƚ�ʱ���,��������ʱ���
                If SdFlg(lngVoucherID, OldTimeStamp) <> "" Then
                    ReDim varArgs(0)
                    varArgs(0) = Voucher.headerText("cCODE")
                    MsgBox GetStringPara("U8.DZ.JA.Res230", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    '                MsgBox "���ݺ�Ϊ" & Voucher.headerText("cCODE") & "�ĵ������������ε��ݣ�", vbInformation, GetString("U8.DZ.JA.Res030")
                    Exit Sub

                End If

                Voucher.getVoucherDataXML oDomHead, oDomBody    '�������
                ExecmakeDom oDomHead, oDomBody, g_Conn     '��֯����

                g_Conn.BeginTrans

                If ExecCreateVoucher(oDomHead, oDomBody, g_Conn, g_oLogin, �ɹ�, �ɹ�����) Then
                    g_Conn.CommitTrans
                Else
                    g_Conn.RollbackTrans                   '�Ƶ�ʧ�ܻع�
                End If

                Call ExecSubRefresh

                '�Ƶ�-������ⵥ
            Case sKey_CreateSCVoucher

                '�Ƶ�ǰ�Ƚ�ʱ���,��������ʱ���
                If SdFlg(lngVoucherID, OldTimeStamp) <> "" Then
                    ReDim varArgs(0)
                    varArgs(0) = Voucher.headerText("cCODE")
                    MsgBox GetStringPara("U8.DZ.JA.Res230", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    '               MsgBox "���ݺ�Ϊ" & Voucher.headerText("cCODE") & "�ĵ������������ε��ݣ�", vbInformation, GetString("U8.DZ.JA.Res030")
                    Exit Sub

                End If

                Voucher.getVoucherDataXML oDomHead, oDomBody    '�������
                ExecmakeDom oDomHead, oDomBody, g_Conn     '��֯����

                g_Conn.BeginTrans

                If ExecCreateVoucher(oDomHead, oDomBody, g_Conn, g_oLogin, ���, ������ⵥ) Then
                    g_Conn.CommitTrans
                Else
                    g_Conn.RollbackTrans                   '�Ƶ�ʧ�ܻع�
                End If

                Call ExecSubRefresh

                '�Ƶ�-Ӧ����
            Case sKey_CreateAPVoucher

                '�Ƶ�ǰ�Ƚ�ʱ���,��������ʱ���
                If SdFlg(lngVoucherID, OldTimeStamp) <> "" Then
                    ReDim varArgs(0)
                    varArgs(0) = Voucher.headerText("cCODE")
                    MsgBox GetStringPara("U8.DZ.JA.Res230", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    '               MsgBox "���ݺ�Ϊ" & Voucher.headerText("cCODE") & "�ĵ������������ε��ݣ�", vbInformation, GetString("U8.DZ.JA.Res030")
                    Exit Sub

                End If

                Voucher.getVoucherDataXML oDomHead, oDomBody    '�������
                ExecmakeDom oDomHead, oDomBody, g_Conn     '��֯����

                g_Conn.BeginTrans

                If ExecCreateVoucher(oDomHead, oDomBody, g_Conn, g_oLogin, Ӧ��) Then
                    g_Conn.CommitTrans
                Else
                    g_Conn.RollbackTrans                   '�Ƶ�ʧ�ܻع�
                End If

                Call ExecSubRefresh

                '����
            Case sKey_Copy
                '����ʱ����Ȩ��
                '                If Me.ComTemplateShow.ListCount = 0 Then
                '                    'MsgBox "�Բ�����û��ģ��Ȩ�ޣ��޷����ӣ�", vbInformation, pustrMsgTitle
                '                    MsgBox GetResString("U8.ST.Default.00156"), vbInformation, GetResString("U8.SCM.ST.KCGLSQL.FrmMainST.rop.Caption") 'pustrMsgTitle
                '                    Exit Sub
                '                End If
                
                If Check_Auth = False Then Exit Sub

                If ZwTaskExec(g_oLogin, AuthAdd, 1) = False Then Exit Sub
                Call initCustomRelation                    '20110822 by zhangwchb
                Call ExecSubCopy
                Call ZwTaskExec(g_oLogin, AuthAdd, 0)

                '����
            Case sKey_Save

                If Voucher.ProtectUnload2 <> 2 Then
                    Voucher.SetFocus

                    Exit Sub

                End If

                Call ExecSubSave

                '����
            Case sKey_Discard

                Call ExecSubDiscard

                '���
            Case sKey_Confirm
                '                If bCheckUser = True Then
                '                    If CheckUserAuth(g_Conn.ConnectionString, g_oLogin.cUserId, Voucher.headerText(StrcMaker), "V") = False Then
                '                        MsgBox GetStringPara("U8.pu.VoucherCommon.00115", Voucher.headerText("cCODE")), vbInformation, GetString("U8.DZ.JA.Res030")
                '
                '                        Exit Sub
                '
                '                    End If
                '                End If

                '                If CheckSubmit(MainTable, "ID", CStr(lngVoucherID)) Then
                '                    wfcBack = 2
                '                    Call ExecRequestAudit(gsGUIDForVouch)
                '                Else

                If ZwTaskExec(g_oLogin, AuthVerify, 1) = False Then Exit Sub
                If ExecFunCompareUfts = True Then ExecSubConfirm
                Call ZwTaskExec(g_oLogin, AuthVerify, 0)
                '                End If
                
                '����
                'enum by modify
            Case sKey_Cancelconfirm
                '                If bCheckUser = True Then
                '                    If CheckUserAuth(g_Conn.ConnectionString, g_oLogin.cUserId, Voucher.headerText(StrcMaker), "U") = False Then
                '                        MsgBox GetStringPara("U8.pu.VoucherCommon.00129", Voucher.headerText("cCODE")), vbInformation, GetString("U8.DZ.JA.Res030")
                '
                '                        Exit Sub
                '
                '                    End If
                '                End If
                '                If Voucher.headerText("cCreateType") = "ת������" Then
                '                    MsgBox GetString("U8.DZ.JA.Res240"), vbInformation, GetString("U8.DZ.JA.Res030")
                '
                '                    Exit Sub
                '
                '                End If

                '                If gcCreateType = "�ڳ�����" Then
                '                    If VoucherIsCreate2(lngVoucherID) Then
                '                        MsgBox GetString("U8.DZ.JA.Res250"), vbInformation, GetString("U8.DZ.JA.Res030")
                '
                '                        Exit Sub
                '
                '                    End If
                '
                '                Else

                If VoucherIsCreate(lngVoucherID) Then
                    MsgBox GetString("U8.DZ.JA.Res250"), vbInformation, GetString("U8.DZ.JA.Res030")

                    Exit Sub

                End If

                '                End If

                '����
                '                If CheckSubmit(MainTable, "ID", CStr(lngVoucherID)) Then
                '                    wfcBack = 2
                '                    Call ExecCancelAudit(gsGUIDForVouch)
                '                Else

                If ZwTaskExec(g_oLogin, AuthUnVerify, 1) = False Then Exit Sub
                If ExecFunCompareUfts = True Then ExecSubCancelconfirm
                Call ZwTaskExec(g_oLogin, AuthUnVerify, 0)
                '                End If

                '��
            Case sKey_Open

                If ZwTaskExec(g_oLogin, AuthOpen, 1) = False Then Exit Sub
                If ExecFunCompareUfts = True Then ExecSubOpen
                Call ZwTaskExec(g_oLogin, AuthOpen, 0)

                '�ر�
            Case sKey_Close

                If ZwTaskExec(g_oLogin, AuthClose, 1) = False Then Exit Sub
                If bCheckUser = True Then
                    If CheckUserAuth(g_Conn.ConnectionString, g_oLogin.cUserId, Voucher.headerText(StrcMaker), "C") = False Then
                        MsgBox GetString("U8.pu.VoucherCommon.00106"), vbInformation, GetString("U8.DZ.JA.Res030")

                        Exit Sub

                    End If
                End If

                If ExecFunCompareUfts = True Then ExecSubClose
                Call ZwTaskExec(g_oLogin, AuthClose, 0)

                '��λ
            Case sKey_Locate
                Call ExecLocate

                'ˢ��
            Case sKey_Refresh
                Call ExecSubRefresh

                '����
            Case gstrHelpCode
                SendKeys "{F1}"

                '����
            Case sKey_Addrecord
                Call ExecSubAddRecord

                '����
            Case sKey_InsertRecord
                Call ExecSubInsertRecord

                'ɾ��
            Case sKey_Deleterecord
                Call ExecSubDeleterecord

                '?����
            Case sKey_Acc
                Voucher.SelectFile

                '������ chenliangc
            Case sKey_Submit

                'by zhangwchb 20110829 �ύ����
                If Voucher.VoucherStatus <> VSNormalMode Then
                    Call ExecSubSave

                    If isSavedOK = False Then Exit Sub
                End If

                Call ExecSubmit(True, MainTable, "ID", lngVoucherID)
                Call ExecSubRefresh

            Case sKey_Unsubmit

                If bCheckUser = True Then
                    If CheckUserAuth(g_Conn.ConnectionString, g_oLogin.cUserId, Voucher.headerText(StrcMaker), "A") = False Then
                        MsgBox GetStringPara("U8.pu.VoucherCommon.01035", Voucher.headerText("cCODE")), vbInformation, GetString("U8.DZ.JA.Res030")

                        Exit Sub

                    End If
                End If

                Call ExecSubmit(False, MainTable, "ID", lngVoucherID)
                Call ExecSubRefresh

            Case sKey_Resubmit
                wfcBack = 2
                Call ExecRequestAudit(gsGUIDForVouch)

            Case sKey_ViewVerify
                Call ExecViewVerify(gsGUIDForVouch)

                'ȡ��
            Case "rowprice", "allprice"
                Call GetPrice(LCase(strbuttonkey), "97", Voucher)    '97���۶���
                
                '����
            Case "BatchModify"
                Call PopMenu1_MenuClick(strbuttonkey)
                
                '�и���
            Case "mnuCopyLine"
                Call PopMenu1_MenuClick("copyr")
                
                '�����
            Case "mnuSplitLine"
                Call PopMenu1_MenuClick("bsplit")
            
            Case "QueryStockAll"
                Call QueryStockAll(Voucher)
                
            Case "QueryStock"
                Call QueryStock(Voucher)
                
                'ˢ�±����ִ���
            Case "RefreshStock"
                Call ShowBodyStockAll(Voucher)
                
                '�黹
            Case "Return"

                Dim sTmp               As String

                Dim strMsg             As String

                Dim cCode              As String

                Dim IsBackWfcontrolled As Boolean '����黹���Ƿ���������

                If getIsWfControl(g_oLogin, g_Conn, sTmp, "HYJCGH005") Then          '����������
                    IsBackWfcontrolled = True
                Else
                    IsBackWfcontrolled = False
                End If

                cCode = Voucher.headerText("cCODE")
                
                If CheckCanBack(lngVoucherID, cCode, gcCreateType, sTmp) Then
                    If ExecReturn(lngVoucherID, sTmp, IsBackWfcontrolled, Voucher.headerText("ufts")) Then
                        strMsg = strMsg & GetStringPara("U8.ST.V870.00759", cCode) & vbCrLf '

                        'strMsg = strMsg & "���� " & cCode & " �黹�ɹ���" & vbCrLf
                        If sTmp <> "" Then strMsg = strMsg & sTmp & vbCrLf
                    Else
                        strMsg = strMsg & GetStringPara("U8.ST.V870.00760", cCode) & vbCrLf '

                        ' strMsg = strMsg & "���� " & cCode & " �黹ʧ�ܣ�" & vbCrLf
                        If sTmp <> "" Then strMsg = strMsg & sTmp & vbCrLf
                    End If

                Else

                    If sTmp <> "" Then strMsg = strMsg & sTmp & vbCrLf
                End If

                If strMsg <> "" Then
                    Screen.MousePointer = vbDefault
                    Load FrmMsgBox
                    FrmMsgBox.Text1 = strMsg
                    FrmMsgBox.Show 1
                    Screen.MousePointer = vbDefault
                End If
                
                

        End Select

Exit_Label:

        On Error GoTo 0

        Screen.MousePointer = vbDefault
        SetToolbarVisible

        Exit Sub
    
Err_Handler:
        sMessage = "" 'GetString("U8.DZ.HU.Res240")

        ' * ��ʾ�ѺõĴ�����Ϣ
        Call ShowErrorInfo(sHeaderMessage:=sMessage, lMessageType:=vbExclamation, lErrorLevel:=ufsELHeaderAndDescription)
        GoTo Exit_Label

    End If

End Sub

'���������޸ĵ��� 20110817 by zhangwchb
Private Function CheckVerModify() As Boolean
    On Error GoTo lerr
    Dim AuditServiceProxy As Object
    Dim objCalledContext As Object
    Dim strErr As String
    Dim IsChangeableVoucher As Boolean

    If Voucher.headerText("iswfcontrolled") = 1 Then
        If Voucher.headerText("iverifystate") = "0" Or Voucher.headerText("iverifystate") = "" Then
            GoTo ExitOK
            '       Else ' ����������,�Ѿ��ύ,����������ж�
        End If
    Else
        GoTo ExitOK
    End If

    '�ж�����ǰ�Ƿ������޸�**********************************************************************************
    Set objCalledContext = CreateObject("UFSoft.U8.Framework.LoginContext.CalledContext")
    objCalledContext.SubId = g_oLogin.cSub_Id
    objCalledContext.TaskId = g_oLogin.TaskId
    objCalledContext.token = g_oLogin.userToken
    Set AuditServiceProxy = CreateObject("UFIDA.U8.Audit.ServiceProxy.AuditServiceProxy")

    IsChangeableVoucher = AuditServiceProxy.IsChangeableVoucher(Voucher.headerText("ID"), LCase$(gstrCardNumber), Voucher.headerText("cCode"), objCalledContext, strErr)

    '�ж�����ǰ�Ƿ������޸�**********************************************************************************

    If strErr <> "" Then
        MsgBox strErr, vbInformation, GetString("U8.DZ.HU.Res030")
        GoTo lerr
    ElseIf Not IsChangeableVoucher Then
        MsgBox GetString("U8.DZ.HU.Res830"), vbInformation, GetString("U8.DZ.HU.Res030")
        GoTo lerr
    End If
    Set objCalledContext = Nothing
    Set AuditServiceProxy = Nothing

ExitOK:
    CheckVerModify = True
    Exit Function
lerr:
    CheckVerModify = False
    Exit Function
End Function

Private Sub ExecLocate()

    On Error GoTo Err_Handler

    If GetFilter(g_oLogin) = False Then Exit Sub

    'by liwqa MainView ��ΪMainTable


    sAuth_ALL = Replace(sAuth_ALL, "HY_DZ_BorrowOut.id", "V_HY_DZ_BorrowOut.id")


    lngVoucherID = GetTheLastID(login:=g_oLogin, _
            oConnection:=g_Conn, _
            sTable:=MainView, _
            sField:=HeadPKFld, _
            sDataNumFormat:="0", _
            sWhereStatement:=strwhereVou)


    sAuth_ALL = Replace(sAuth_ALL, "V_HY_DZ_BorrowOut.id", "HY_DZ_BorrowOut.id")

    If strwhereVou <> "" And lngVoucherID = 0 Then
        MsgBox GetString("U8.DZ.JA.Res260"), vbInformation, GetString("U8.DZ.JA.Res030")
        Exit Sub
    End If
    Call LoadVoucherData

    '���µ�ǰҳ����PageCurrent
    Call UpdatePageCurrent(lngVoucherID)

    mOpStatus = SHOW_ALL
    Call SetCtlStyle(Me, Voucher, Me.Toolbar, Me.UFToolbar, mOpStatus)

    '��ȡʱ���
    OldTimeStamp = GetTimeStamp(g_Conn, MainTable, lngVoucherID)

    Exit Sub


Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Sub

'��ȡʱ����������ʱ����Ƚ�
Private Function ExecFunCompareUfts() As Boolean

    '��ȡʱ���
    TimeStamp = GetTimeStamp(g_Conn, MainTable, lngVoucherID)

    If TimeStamp = RecordDeleted Then
        MsgBox GetString("U8.DZ.JA.Res270"), vbInformation, GetString("U8.DZ.JA.Res030")
        ExecFunCompareUfts = False
        Exit Function
    ElseIf TimeStamp = RecordError Then
        MsgBox GetString("U8.DZ.JA.Res280"), vbInformation, GetString("U8.DZ.JA.Res030")
        ExecFunCompareUfts = False
        Exit Function
    ElseIf OldTimeStamp <> TimeStamp Then
        MsgBox GetString("U8.DZ.JA.Res290"), vbInformation, GetString("U8.DZ.JA.Res030")
        ExecFunCompareUfts = False
        Exit Function
    Else
        OldTimeStamp = TimeStamp
        ExecFunCompareUfts = True
    End If

End Function

'�ر�
Private Sub ExecSubClose()
    On Error GoTo Err_Handler
    'by liwqa ����
    Dim m_bTrans As Boolean
    Dim lngAct As Long
    Dim strSql As String
    g_Conn.BeginTrans
    m_bTrans = True

    g_Conn.Execute "update " & MainTable & " set " & StrCloseUser & "='" & g_oLogin.cUserName & "' , " & StrdCloseDate & "='" & g_oLogin.CurDate & "' , " & StriStatus & "=4 " & vbCrLf & _
            " where convert(nchar,convert(money,ufts),2)  = " & OldTimeStamp & " and " & HeadPKFld & "=" & lngVoucherID, lngAct

    If lngAct = 0 Then

        If m_bTrans Then
            g_Conn.RollbackTrans
            m_bTrans = False
        End If
        MsgBox GetString("U8.DZ.JA.Res290"), vbInformation, GetString("U8.DZ.JA.Res030")
        GoTo Exit_Label
    End If

    If m_bTrans Then
        g_Conn.CommitTrans
        m_bTrans = False
    End If
    
     
       strSql = " update HY_FYSL_Contract set totalappmoney= " & Null2Something(Voucher.headerText("appprice"), 0) - Null2Something(Voucher.headerText("payamount"), 0) & " where  ccode= '" & Voucher.headerText("concode") & "'"
       g_Conn.Execute strSql
        
    

    Call ExecSubRefresh

Exit_Label:
    On Error GoTo 0
    Exit Sub
Err_Handler:

    If m_bTrans Then
        g_Conn.RollbackTrans
        m_bTrans = False
    End If

End Sub
'��
Private Sub ExecSubOpen()

    '״̬:��������˲�Ϊ��,��Ϊ3,��������;���������Ϊ��:
    '                                                   �������˲�Ϊ��,��Ϊ2,�����;��������Ϊ��,����Ϊ1,���½�

    On Error GoTo Err_Handler
    'by liwqa ����
    Dim m_bTrans As Boolean
    Dim lngAct As Long
    g_Conn.BeginTrans
    m_bTrans = True

    g_Conn.Execute "update " & MainTable & " set " & StrCloseUser & "=null , " & StrdCloseDate & "=null, " & StriStatus & "= case when isnull(" & StrcHandler & ",N'')<>N'' then 2 else 1   end " & vbCrLf & _
            " where convert(nchar,convert(money,ufts),2)  = " & OldTimeStamp & " and " & HeadPKFld & "=" & lngVoucherID, lngAct

    If lngAct = 0 Then

        If m_bTrans Then
            g_Conn.RollbackTrans
            m_bTrans = False
        End If
        MsgBox GetString("U8.DZ.JA.Res290"), vbInformation, GetString("U8.DZ.JA.Res030")
        GoTo Exit_Label
    End If

    If m_bTrans Then
        g_Conn.CommitTrans
        m_bTrans = False
    End If

    Call ExecSubRefresh

Exit_Label:
    On Error GoTo 0
    Exit Sub
Err_Handler:

    If m_bTrans Then
        g_Conn.RollbackTrans
        m_bTrans = False
    End If
End Sub

'����
Private Sub ExecSubKCreateBill()

    g_Conn.Execute "update " & MainTable & " set " & StrIntoUser & "='" & g_oLogin.cUserName & "' , " & StrdIntoDate & "='" & g_oLogin.CurDate & "' , " & StriStatus & "=3 where " & HeadPKFld & "=" & lngVoucherID
    Call ExecSubRefresh

End Sub

'���
Private Sub ExecSubConfirm()
    On Error GoTo Err_Handler
    'by liwqa ����
    Dim m_bTrans As Boolean
    Dim lngAct As Long
    g_Conn.BeginTrans
    m_bTrans = True
    Dim sMsg As String
     Dim sql As String

    g_Conn.Execute "update " & MainTable & " set " & StrcHandler & "='" & g_oLogin.cUserName & "' , " & StrdVeriDate & "='" & g_oLogin.CurDate & "' , " & StriStatus & "=2 " & vbCrLf & _
            " where convert(nchar,convert(money,ufts),2)  = " & OldTimeStamp & " and " & HeadPKFld & "=" & lngVoucherID, lngAct

    If lngAct = 0 Then

        If m_bTrans Then
            g_Conn.RollbackTrans
            m_bTrans = False
        End If
        MsgBox GetString("U8.DZ.JA.Res290"), vbInformation, GetString("U8.DZ.JA.Res030")
        GoTo Exit_Label
    End If
'
'
' sql = "update HY_FYSL_Contract  set conpaymoney=" & Voucher.headerText("engamounts") & " ,conpaytolmoney= " & Voucher.headerText("contolamounts") & " where  ccode ='" & Voucher.headerText("concode") & "'"
'
'        g_Conn.Execute sql, lngAct
'
'       If lngAct = 0 Then
'        If m_bTrans Then
'            g_Conn.RollbackTrans
'            m_bTrans = False
'
'        End If
'           MsgBox GetString("U8.DZ.JA.Res290"), vbInformation, GetString("U8.DZ.JA.Res030")
'        GoTo Exit_Label
'      End If
       
       
          sql = " update HY_FYSL_Contract set totalappmoney=isnull(totalappmoney,0)+" & Voucher.headerText("appprice") & " where  ccode= '" & Voucher.headerText("concode") & "'"
       g_Conn.Execute sql
       
       
    
    If sMsg = "" Then
        sMsg = GetString("U8.DZ.JA.Res1940") & vbCrLf '"������˳ɹ�!"
          
    End If
    
'    '����Զ������������ⵥ
'    If gcCreateType <> "�ڳ�����" Then
'        If LCase(getAccinformation("ST", "bautolendout", g_Conn)) = "true" Then
'            sMsg = sMsg & ExecPushOtherOut(lngVoucherID)
'        End If
'    End If
     g_Conn.CommitTrans
     m_bTrans = False
    Screen.MousePointer = vbDefault
    Load FrmMsgBox
    FrmMsgBox.Text1 = sMsg
    FrmMsgBox.Show 1
    Screen.MousePointer = vbDefault
        
    Call ExecSubRefresh
    
Exit_Label:
    On Error GoTo 0
    Exit Sub
Err_Handler:

    If m_bTrans Then
        g_Conn.RollbackTrans
        m_bTrans = False
    End If
End Sub


'����
Private Sub ExecSubCancelconfirm()
    On Error GoTo Err_Handler
    'by liwqa ����
    Dim m_bTrans As Boolean
    Dim lngAct As Long
    g_Conn.BeginTrans
    m_bTrans = True
    Dim sql As String
    
    
        
      '************************************
     '������Ŀ�������ۼ������ͽ��

     sql = " update HY_FYSL_Contract set totalappmoney= " & Null2Something(Voucher.headerText("totalappmoney"), 0) - Null2Something(Voucher.headerText("appprice"), 0) & " where  ccode= '" & Voucher.headerText("concode") & "'"
       g_Conn.Execute sql

     '***********************************
    

    g_Conn.Execute "update " & MainTable & " set " & StrcHandler & "=null , " & StrdVeriDate & "=null , " & StriStatus & "=1 " & vbCrLf & _
            " where convert(nchar,convert(money,ufts),2)  = " & OldTimeStamp & " and " & HeadPKFld & "=" & lngVoucherID, lngAct
 
       If lngAct = 0 Then
        If m_bTrans Then
            g_Conn.RollbackTrans
            m_bTrans = False
        End If
        MsgBox GetString("U8.DZ.JA.Res290"), vbInformation, GetString("U8.DZ.JA.Res030")
      End If
      
'       sql = "update HY_FYSL_Contract  set conpaymoney=null,conpaytolmoney=null where  ccode ='" & Voucher.headerText("concode") & "'"
'
'      g_Conn.Execute sql, lngAct
''
'       If lngAct = 0 Then
'        If m_bTrans Then
'            g_Conn.RollbackTrans
'            m_bTrans = False
'
'        End If
'           MsgBox GetString("U8.DZ.JA.Res290"), vbInformation, GetString("U8.DZ.JA.Res030")
'             GoTo Exit_Label
'      End If

'    If m_bTrans Then
'        'ҵ��֪ͨ
'        NotifySrvSend "HYJCGH001", "HYJCGH001.UnAudit", CStr(lngVoucherID), g_oLogin
'        g_Conn.CommitTrans
'        m_bTrans = False
'    End If
 g_Conn.CommitTrans
 m_bTrans = False
    Call ExecSubRefresh

Exit_Label:
    On Error GoTo 0
    Exit Sub
Err_Handler:

    If m_bTrans Then
        g_Conn.RollbackTrans
        m_bTrans = False
    End If
End Sub
 Public Function Updateconnactdata(isconfim As Boolean) As Boolean
   On Error GoTo Err_Handler
   Dim sql As String
   Dim rs As New ADODB.Recordset
   
   
     Dim m_bTrans As Boolean
    Dim lngAct As Long
    g_Conn.BeginTrans
    m_bTrans = True
    Updateconnactdata = True
    
    If isconfim = True Then
       sql = "  delete HY_FYSL_Contract where ccode='" & Voucher.headerText("concode") & "'  insert into   HY_FYSL_Contract (id, ccode, cname, ddate, cversion, ecustcode, contype, custcontacta, contacta, custcontactb, contactb, acccode, icode, virtualCon, smtype, smapptype, engcode, proccode, consdate, conedate, coneffedate, " & _
       " mconcode, paytype, contolprice, conmoney, designmoney, conpaytolmoney, conpaymoney, accdesignmoney, concompleted, totalappmoney, totalpaymoney, designunits, consubject, cmemo, condescriptiona, condescriptionb, condescriptionc, cdefine1, " & _
       " cdefine2, cdefine3, cdefine4, cdefine5, cdefine6, cdefine7, cdefine8, cdefine9, cdefine10, cdefine11, cdefine12, cdefine13, cdefine14, cdefine15, cdefine16, CloseUser, dCloseDate, iStatus, VoucherType, iverifystate, iswfcontrolled, cHandler, dVeriDate, " & _
       " iVTID, cMaker, dmDate, payperone, paydate, modperone, moddate)  select   b.id,  b.ccode,  b.cname, b.ddate, a.cversion, a.ecustcode, a.contype, a.custcontacta, a.contacta, a.custcontactb, a.contactb, a.acccode, a.icode, a.virtualCon, a.smtype, a.smapptype, a.engcode, a.proccode, a.consdate, a.conedate, a.coneffedate, " & _
       " a.mconcode, a.paytype, a.contolprice, a.conmoney, a.designmoney, a.conpaytolmoney, a.conpaymoney, a.accdesignmoney, a.concompleted, a.totalappmoney, a.totalpaymoney, a.designunits, a.consubject, a.cmemo, a.condescriptiona, a.condescriptionb, a.condescriptionc, a.cdefine1, " & _
       " a.cdefine2, a.cdefine3, a.cdefine4, a.cdefine5, a.cdefine6, a.cdefine7, a.cdefine8, a.cdefine9, a.cdefine10, a.cdefine11, a.cdefine12, a.cdefine13, a.cdefine14, a.cdefine15, a.cdefine16, b.CloseUser, b.dCloseDate,b.iStatus, b.VoucherType, b.iverifystate, b.iswfcontrolled, b.cHandler, b.dVeriDate," & _
       " b.iVTID , b.cMaker, b.dmDate, b.payperone, b.paydate, b.modperone, b.moddate " & _
       " from   HY_FYSL_Payment  a " & _
       " inner join   HY_FYSL_Contract b  on a.concode =b.ccode   " & _
       " where  a.id=" & lngVoucherID
    
    Else
           
       sql = " delete  HY_FYSL_Contract  where ccode= '" & Voucher.headerText("concode") & "'" & _
       "   insert  into  HY_FYSL_Contract(id, ccode, cname, ddate, cversion, ecustcode, contype, custcontacta, contacta, custcontactb, contactb, acccode, icode, virtualCon, smtype, smapptype, engcode, proccode, consdate, conedate, coneffedate, " & _
       " mconcode, paytype, contolprice, conmoney, designmoney, conpaytolmoney, conpaymoney, accdesignmoney, concompleted, totalappmoney, totalpaymoney, designunits, consubject, cmemo, condescriptiona, condescriptionb, condescriptionc, cdefine1, " & _
       " cdefine2, cdefine3, cdefine4, cdefine5, cdefine6, cdefine7, cdefine8, cdefine9, cdefine10, cdefine11, cdefine12, cdefine13, cdefine14, cdefine15, cdefine16, CloseUser, dCloseDate, iStatus, VoucherType, iverifystate, iswfcontrolled, cHandler, dVeriDate, " & _
       " iVTID, cMaker, dmDate, payperone, paydate, modperone, moddate)  select id, ccode, cname, ddate, cversion, ecustcode, contype, custcontacta, contacta, custcontactb, contactb, acccode, icode, virtualCon, smtype, smapptype, engcode, proccode, consdate, conedate, coneffedate, " & _
       "  mconcode, paytype, contolprice, conmoney, designmoney, conpaytolmoney, conpaymoney, accdesignmoney, concompleted, totalappmoney, totalpaymoney, designunits, consubject, cmemo, condescriptiona, condescriptionb, condescriptionc, cdefine1, " & _
       " cdefine2, cdefine3, cdefine4, cdefine5, cdefine6, cdefine7, cdefine8, cdefine9, cdefine10, cdefine11, cdefine12, cdefine13, cdefine14, cdefine15, cdefine16, CloseUser, dCloseDate, iStatus, VoucherType, iverifystate, iswfcontrolled, cHandler, dVeriDate," & _
       " iVTID, cMaker, dmDate, payperone, paydate, modperone, moddate from  HY_FYSL_Contracthistory where  ccode= '" & Voucher.headerText("concode") & "'"

    End If
    g_Conn.Execute sql, lngAct
    
       If lngAct = 0 Then
        If m_bTrans Then
            g_Conn.RollbackTrans
            m_bTrans = False
            
        End If
          Updateconnactdata = False
             Exit Function
      End If
    
      g_Conn.CommitTrans
      Updateconnactdata = True
Exit_Label:
    On Error GoTo 0
    Exit Function
Err_Handler:

    If m_bTrans Then
        g_Conn.RollbackTrans
        m_bTrans = False
    End If
   
 End Function
 


'����
Private Sub ExecSubAddRecord()
    Voucher.AddLine
End Sub
'����
Private Sub ExecSubInsertRecord()
    Dim iRow As Long
    iRow = Voucher.row
    If iRow = 0 Then
        Exit Sub
    Else
        Voucher.AddLine Voucher.row, , ALSPrevious
    End If
End Sub
'ɾ��
Private Sub ExecSubDeleterecord()
    Voucher.DelLine Voucher.row
End Sub

'����'
Private Sub ExecSubDiscard()

    If MsgBox(GetString("U8.DZ.JA.Res300"), vbYesNo, GetString("U8.DZ.JA.Res030")) = vbYes Then
        mOpStatus = SHOW_ALL
        Call ExecSubRefresh
    End If

End Sub


'��ҳ
Public Sub ExecSubPageFirst()

    '��ҳ����λ�ᵼ��ȡ��id�Ǵ����
    Dim sql As String
    Dim rsid As New ADODB.Recordset

    PageCurrent = 1

    '��ȡ����Ȩ��
    '    Dim sRet As String
    '    sRet = GetRowAuthAlls(g_Conn, g_oLogin.cUserId, "R")
    ''    If sRet <> "" Then
    ''    sRet = " where cCreateType<>'�ڳ�����' and 1=1 " & sRet
    '    sRet = " where  1=1 " & sRet


    '    If tmpLinkTbl <> "" Then  '�������� ʱ ��ť״̬���� by zhangwchb 20110809
    '        sql = "select min(" & MainTable & "." & HeadPKFld & ") id from " & MainTable & _
             '            " inner join " & tmpLinkTbl & " on " & MainTable & "." & HeadPKFld & " = " & tmpLinkTbl & ".id " & _
             '            " where " & sAuth_ALL
    '    Else
    If sTmpTableName <> "" Then
        sql = "select min(" & HeadPKFld & ") id from " & MainTable & " inner join " & sTmpTableName & " as search on " & MainTable & ".id=search.cvoucherid  where " & sAuth_ALL
    Else
        sql = "select min(" & HeadPKFld & ") id from " & MainTable & " where " & sAuth_ALL
    End If
    '    End If

    rsid.Open sql, g_Conn, 1, 1


    If Not rsid.EOF Then
        lngVoucherID = rsid("id")
    End If

    rsid.Close
    Set rsid = Nothing

    Call ExecSubRefresh

End Sub

'ĩҳ
Public Sub ExecSubPageLast()
    '��ҳ����λ�ᵼ��ȡ��id�Ǵ����
    Dim sql As String
    Dim rsid As New ADODB.Recordset

    PageCurrent = pageCount

    '��ȡ����Ȩ��
    '    Dim sRet As String
    '    sRet = GetRowAuthAlls(g_Conn, g_oLogin.cUserId, "R")
    ''    If sRet <> "" Then
    ''    sRet = " where cCreateType<>'�ڳ�����' and  1=1 " & sRet
    '    sRet = " where  1=1 " & sRet

    '    If tmpLinkTbl <> "" Then  '�������� ʱ ��ť״̬���� by zhangwchb 20110809
    '        sql = "select  isnull(max(" & MainTable & "." & HeadPKFld & "),0) id from " & MainTable & _
             '            " inner join " & tmpLinkTbl & " on " & MainTable & "." & HeadPKFld & " = " & tmpLinkTbl & ".id " & _
             '            " where " & sAuth_ALL
    '    Else
    If sTmpTableName <> "" Then
        sql = "select isnull(max(" & HeadPKFld & "),0) id from " & MainTable & " inner join " & sTmpTableName & " as search on " & MainTable & ".id=search.cvoucherid where " & sAuth_ALL
    Else
        sql = "select isnull(max(" & HeadPKFld & "),0) id from " & MainTable & " where " & sAuth_ALL
    End If
    '    End If
    rsid.Open sql, g_Conn, 1, 1
    If Not rsid.EOF Then
        lngVoucherID = rsid("id")
    End If

    rsid.Close
    Set rsid = Nothing

    If lngVoucherID > 0 Then Call ExecSubRefresh
End Sub

'��һҳ
Public Sub ExecSubPageUp()
    '��ҳ����λ�ᵼ��ȡ��id�Ǵ����
    Dim sql As String
    Dim rsid As New ADODB.Recordset

    If PageCurrent > 1 Then
        PageCurrent = PageCurrent - 1

        '��ȡ����Ȩ��
        '        Dim sRet As String
        '        sRet = GetRowAuthAlls(g_Conn, g_oLogin.cUserId, "R")


        '        SQL = "select top 1 " & HeadPKFld & " as id from " & MainTable & " where cCreateType<>'�ڳ�����' and " & HeadPKFld & "<" & lngVoucherID & sRet & " order by " & HeadPKFld & " desc"

        '        If tmpLinkTbl <> "" Then  '�������� ʱ ��ť״̬���� by zhangwchb 20110809
        '            sql = "select top 1 " & MainTable & "." & HeadPKFld & " as id from " & MainTable & _
                     '                " inner join " & tmpLinkTbl & " on " & MainTable & "." & HeadPKFld & " = " & tmpLinkTbl & ".id " & _
                     '                " where " & MainTable & "." & HeadPKFld & " < " & lngVoucherID & " and " & sAuth_ALL & _
                     '                " order by " & MainTable & "." & HeadPKFld & " desc "
        '        Else
        If sTmpTableName <> "" Then
            sql = "select top 1 " & HeadPKFld & " as id from " & MainTable & " inner join " & sTmpTableName & " as search on " & MainTable & ".id=search.cvoucherid where  " & HeadPKFld & "<" & lngVoucherID & " and " & sAuth_ALL & " order by " & HeadPKFld & " desc"
        Else
            sql = "select top 1 " & HeadPKFld & " as id from " & MainTable & " where  " & HeadPKFld & "<" & lngVoucherID & " and " & sAuth_ALL & " order by " & HeadPKFld & " desc"
        End If
        '        End If

        rsid.Open sql, g_Conn, 1, 1
        If Not rsid.EOF Then
            lngVoucherID = rsid("id")
        End If

        rsid.Close
        Set rsid = Nothing

        Call ExecSubRefresh

    End If


End Sub

'��һҳ
Public Sub ExecSubPageDown()
    '��ҳ����λ�ᵼ��ȡ��id�Ǵ����
    Dim sql As String
    Dim rsid As New ADODB.Recordset

    If PageCurrent < pageCount Then

        '��ȡ����Ȩ��
        Dim sRet As String
        '        sRet = GetRowAuthAlls(g_Conn, g_oLogin.cUserId, "R")

        '        SQL = "select top 1 " & HeadPKFld & " as id from " & MainTable & " where  cCreateType<>'�ڳ�����' and " & HeadPKFld & ">" & lngVoucherID & sRet & " order by " & HeadPKFld & " asc "

        '        If tmpLinkTbl <> "" Then  ''�������� ʱ ��ť״̬���� by zhangwchb 20110809
        '            sql = "select top 1 " & MainTable & "." & HeadPKFld & " as id from " & MainTable & _
                     '                " inner join " & tmpLinkTbl & " on " & MainTable & "." & HeadPKFld & " = " & tmpLinkTbl & ".id " & _
                     '                " where " & MainTable & "." & HeadPKFld & " > " & lngVoucherID & " and " & sAuth_ALL & _
                     '                " order by " & MainTable & "." & HeadPKFld & " asc "
        '        Else
        If sTmpTableName <> "" Then
            sql = "select top 1 " & HeadPKFld & " as id from " & MainTable & " inner join " & sTmpTableName & " as search on " & MainTable & ".id=search.cvoucherid where   " & HeadPKFld & ">" & lngVoucherID & " and " & sAuth_ALL & " order by " & HeadPKFld & " asc "
        Else
            sql = "select top 1 " & HeadPKFld & " as id from " & MainTable & " where   " & HeadPKFld & ">" & lngVoucherID & " and " & sAuth_ALL & " order by " & HeadPKFld & " asc "
        End If
        '        End If
        rsid.Open sql, g_Conn, 1, 1
        If Not rsid.EOF Then
            lngVoucherID = rsid("id")
        End If

        PageCurrent = PageCurrent + 1

        rsid.Close
        Set rsid = Nothing

        'mod���ܵ���frm �к���
        Call ExecSubRefresh

    End If


End Sub

'ˢ��
'���ݵ�ǰlngvoucherID����
'����,��ҳ,ɾ��ʱ�������lngvoucherID
Public Sub ExecSubRefresh()
    bAlter = False
    
    Call LoadVoucherData

    If Voucher.headerText("cCODE") = "" Then
        Call ExecSubPageLast
    End If

    'ˢ��ʱ���µõ���̨���� by liwq
    Set VchSrv = New clsVouchServer
    pageCount = VchSrv.GetPageCount(g_Conn, gstrCardNumber, HeadPKFld, Replace(sAuth_ALL, "HY_FYSL_Payment.id", "V_HY_FYSL_Payment.id"))
    UFToolbar.RefreshVisible
    UFToolbar.RefreshEnable
    mdOldSelRow = 0
    mdOldSelCol = 0
'    If CBool(mologin.Account.AutoShowBodyStock) = True Then
'        Call ShowBodyStockAll(Voucher)
'    End If
    SendMessgeToPortal "CurrentDocChanged", gsGUIDForVouch
End Sub

Private Sub UFKeyHookCtrl1_ContainerKeyDown(KeyCode As Integer, ByVal Shift As Integer)
    Dim strSql As String
    Dim rs As New ADODB.Recordset


    If Shift = vbAltMask Then
        '��ݼ�ALT + E���
        If KeyCode = vbKeyE Then
            If Me.Toolbar.Buttons(sKey_Output).Enabled = True Then
                ButtonClick sKey_Output
            End If
            '����
        ElseIf KeyCode = vbKeyU Then
            If Me.Toolbar.Buttons(sKey_Cancelconfirm).Enabled = True Then
                ButtonClick sKey_Cancelconfirm
            End If
            '����
        ElseIf KeyCode = vbKeyJ Then
            If Me.Toolbar.Buttons(sKey_Unsubmit).Enabled = True Then
                ButtonClick sKey_Unsubmit
            End If
        ElseIf KeyCode = vbKeyPageUp Then                  '��ҳ
            If Me.Toolbar.Buttons(sKey_First).Enabled = True Then
                ButtonClick sKey_First
            End If
        ElseIf KeyCode = vbKeyPageDown Then                'ĩҳ
            If Me.Toolbar.Buttons(sKey_Last).Enabled = True Then
                ButtonClick sKey_Last
            End If
        ElseIf KeyCode = vbKeyC Then                       '�ر�
            If Me.Toolbar.Buttons(sKey_Close).Enabled = True Then
                ButtonClick sKey_Close
            End If
        ElseIf KeyCode = vbKeyO Then                       '��
            If Me.Toolbar.Buttons(sKey_Open).Enabled = True Then
                ButtonClick sKey_Open
            End If
        End If
    ElseIf Shift = vbCtrlMask Then
        '��ݼ�Ctrl + P��ӡ
        If KeyCode = vbKeyP Then
            If Me.Toolbar.Buttons(sKey_Print).Enabled = True Then
                ButtonClick sKey_Print
            End If
            '��ݼ�Ctrl + W��ӡԤ��
        ElseIf KeyCode = vbKeyW Then
            If Me.Toolbar.Buttons(sKey_Preview).Enabled = True Then
                ButtonClick sKey_Preview
            End If
            '��ݼ�Ctrl + F3����
        ElseIf KeyCode = vbKeyF3 Then
            If Me.Toolbar.Buttons(sKey_Locate).Enabled = True Then
                ButtonClick sKey_Locate
            End If
            '��ݼ�Ctrl + G����
        ElseIf KeyCode = vbKeyG Then
            If Me.Toolbar.Buttons(sKey_ReferVoucher).Enabled = True Then
                ButtonClick sKey_ReferVoucher
            End If
            '��������
        ElseIf KeyCode = vbKeyS Then
            ' ButtonClick sKey_ReferVoucher
            If Me.Toolbar.Buttons(sKey_Save).Enabled = True Then
                ButtonClick sKey_Save
                Call g_oBusiness.SelectToolbarButton(Toolbar.Buttons(sKey_Add))
            End If

            '����
        ElseIf KeyCode = vbKeyZ Then
            If Me.Toolbar.Buttons(sKey_Discard).Enabled = True Then
                ButtonClick sKey_Discard
            End If
            '�ύ
        ElseIf KeyCode = vbKeyJ Then
            If Me.Toolbar.Buttons(sKey_Submit).Enabled = True Then
                ButtonClick sKey_Submit
            End If
            '���
        ElseIf KeyCode = vbKeyU Then
            If Me.Toolbar.Buttons(sKey_Confirm).Enabled = True Then
                ButtonClick sKey_Confirm
            End If

            '����
        ElseIf KeyCode = vbKeyF5 Then
            If Me.Toolbar.Buttons(sKey_Copy).Enabled = True Then
                ButtonClick sKey_Copy
            End If
            '����
        ElseIf KeyCode = vbKeyG Then
            'ButtonClick sKey_ReferVoucher
        ElseIf KeyCode = vbKeyR Then                       'ˢ��
            If Me.Toolbar.Buttons(sKey_Refresh).Enabled = True Then
                ButtonClick sKey_Refresh
            End If

        ElseIf KeyCode = vbKeyD Then                       'ɾ��
            If Me.Toolbar.Buttons(sKey_Deleterecord).Enabled = True Then
                ButtonClick sKey_Deleterecord

            End If
        ElseIf KeyCode = vbKeyA Then                       'ɾ��
            If Me.Toolbar.Buttons(sKey_Addrecord).Enabled = True Then
                ButtonClick sKey_Addrecord
            End If
        ElseIf KeyCode = vbKeyF4 Then
            Call ExitForm(0, 0)
            '��ݼ�Ctrl+E����Ctrl+B���Զ�ָ������,��ⵥ��
        ElseIf KeyCode = vbKeyE Or KeyCode = vbKeyB Or KeyCode = vbKeyQ Or KeyCode = vbKeyO Then
            'Call GetBatchInfoFun(Voucher, KeyCode, Shift)
            KeyCode = 0
        End If
        ' End If
    End If

    Select Case KeyCode
        Case vbKeyF1                                       '����
            Call LoadHelpId(Me, "15030910")
        Case vbKeyF5                                       '����

            If Me.Toolbar.Buttons(sKey_Add).Enabled = True Then
                ButtonClick sKey_Add
                'Call g_oBusiness.SelectToolbarButton(Toolbar.Buttons(sKey_Add))
                '         ElseIf Me.Toolbar.Buttons(sKey_Add1).Enabled = True Then
                '              Call g_oBusiness.SelectToolbarButton(Toolbar.Buttons(sKey_Add1))
            End If
            '  ButtonClick sKey_Add
        Case vbKeyF6                                       '����
            If Me.Toolbar.Buttons(sKey_Save).Enabled = True Then
                ButtonClick sKey_Save
            End If
        Case vbKeyF8                                       '�޸�
            If Me.Toolbar.Buttons(sKey_Modify).Enabled = True Then
                ButtonClick sKey_Modify
            End If
        Case vbKeyDelete                                   'ɾ��
            If Me.Toolbar.Buttons(sKey_Delete).Enabled = True Then
                ButtonClick sKey_Delete
            End If
        Case vbKeyPageUp
            If Me.Toolbar.Buttons(sKey_Previous).Enabled = True Then
                ButtonClick sKey_Previous                  '��һҳ
            End If
        Case vbKeyPageDown
            If Me.Toolbar.Buttons(sKey_Next).Enabled = True Then
                ButtonClick sKey_Next                      '��һҳ
            End If


    End Select

    strSql = "select * from UFMeta_" + g_oLogin.cAcc_Id + "..AA_CustomerButton  where cFormKey='" & gstrCardNumber & "' and cLocaleID='zh-cn' and cHotkey='" & KeyCode & "'   order by cButtonID"
    Set rs = g_Conn.Execute(strSql)
    If Not rs.EOF Then
        ButtonClick CStr(rs!cButtonkey)
    End If
    'SetToolbarVisible
End Sub

Private Sub UFToolbar_OnCommand(ByVal enumType As UFToolBarCtrl.ENUM_MENU_OR_BUTTON, ByVal cButtonId As String, ByVal cMenuId As String)


    If enumType = enumButton Then
        Call ButtonClick(Toolbar.Buttons(cButtonId).Key)
    Else
        Call ButtonClick(Toolbar.Buttons(cButtonId).ButtonMenus(cMenuId).Key)
    End If
    SetToolbarVisible
End Sub

'����
Public Sub ExecSubAdd(Optional byt1 As Byte = 1)
    On Error GoTo Err_Handler:
    Dim rs As New ADODB.Recordset
    Dim sql, sCache, sMessage, sSource As String
    Dim oDomHead, oDomBody As DOMDocument
    Set oDomHead = New DOMDocument
    Set oDomBody = New DOMDocument
    numappprice = 0
    bAlter = False
    
    '���ݳ�ʼ��
    'enum by modify
    If byt1 = 1 Then
        gcCreateType = "��������"
    Else
        gcCreateType = "�ڳ�����"
    End If

    Call setTemplate("")                                   'by liwqa Template

    Call InitVoucher
    Call SetLayOut

    '��ͷ
    sql = "select * from " & MainView & " where 1=2"
    Call rs.Open(sql, g_Conn, 1, 1)
    rs.Save oDomHead, adPersistXML
    rs.Close
    Set rs = Nothing


    '����
    sql = "select * from " & DetailsView & " where 1=2"
    Call rs.Open(sql, g_Conn, 1, 1)
    rs.Save oDomBody, adPersistXML
    rs.Close
    Set rs = Nothing

    '��������
    'enum by modify
    Voucher.AddNew ANMNormalAdd, oDomHead, oDomBody
    Voucher.SetBillNumberRule sCache
    '�Ƶ���,��������
    Voucher.headerText(StrcMaker) = g_oLogin.cUserName
    Voucher.headerText("dmDate") = Format(g_oLogin.CurDate, "yyyy-mm-dd")
    
'    If byt1 = 1 Then
''        Voucher.headerText(StrdDate) = Format(g_oLogin.CurDate, "yyyy-mm-dd")
'    Else
'        sql = "select dateadd(day,-1,CVALUE) as cvalue from AccInformation where csysid='ST' AND CNAME='dSTFirstDate'"
'        rs.Open sql, g_Conn
'        If Not rs.EOF Then
'            Voucher.headerText(StrdDate) = Format(rs.Fields("cvalue"), "yyyy-mm-dd")
'        End If
'    End If

'    Voucher.headerText("cexch_name") = "�����"
'    Voucher.headerText("iexchrate") = 1#
    symbol = "*"                                           '�������㷽ʽ
    Voucher.headerText("iStatus") = "1"
    '    Voucher.headerText("cCreateType") = "��������"
  
  If isfromcon = True Then
   Voucher.headerText("sourcetype") = "FYSL0004"
 
 End If
     

    '    Voucher.headerText("cAboutVoucher") = "���۳��ⵥ"
'    Voucher.headerText("isengdec") = "��"
    Voucher.EnableHead "supengcode", False
    Voucher.EnableHead "supcname", False
    

    Voucher.headerText("ivtid") = m_strVT_ID               'by liwqa Template

    Voucher.headerText("VoucherType") = gstrCardNumber     '"HY99"
    Dim errMsg As String
    If getIsWfControl(g_oLogin, g_Conn, errMsg, gstrCardNumber) Then
        Voucher.headerText("iswfcontrolled") = 1
    End If

    '��־״̬
    Voucher.VoucherStatus = VSeAddMode
    '���õ��ݱ���Ƿ�ɱ༭
    Dim manual As Boolean                                  ' �Ƿ���ȫ�ֹ����
    Call SetVouchCodeEnable(manual)

    mOpStatus = ADD_MAIN

    Call SetCtlStyle(Me, Voucher, Me.Toolbar, Me.UFToolbar, mOpStatus)

    'by zhangwchb 20110829 �ύ����  If bInWorkFlow = True Then
'        Me.Toolbar.Buttons("Submit").Enabled = True
'        UFToolbar.RefreshEnable
'    Else
'        Me.Toolbar.Buttons("Submit").Enabled = False
'        UFToolbar.RefreshEnable
'    End If
'

    '�����ݱ��嵥Ԫ�����ͼƬ
    '    Dim body As Object
    '    Set body = Voucher.GetBodyObject
    '
    '    body.Cell(flexcpPicture, 1, 1, 1, 1) = Me.ImageList1.ListImages.Item(2).Picture



Exit_Label:
    On Error GoTo 0

    Exit Sub
Err_Handler:
    sMessage = GetString("U8.DZ.JA.Res310")

    ' * ��ʾ�ѺõĴ�����Ϣ
    Call ShowErrorInfo( _
            sHeaderMessage:=sMessage, _
            lMessageType:=vbExclamation, _
            lErrorLevel:=ufsELHeaderAndDescription _
            )

    ' * �����������Դ

    If Left(Err.Source, 3) = "***" Then
        sSource = Err.Source
    Else
        sSource = "***Sub AddNewVoucher of Form frmPMRecord"
    End If

    ' * ����ģʽʱ����ʾ���Դ��ڣ����ڸ��ٴ���
    If g_blnDEBUG_MODE Then
        Call ShowDebugForm( _
                bErrorMode:=True, _
                sProcedure:=sSource)
    End If
    GoTo Exit_Label
End Sub

'����
Private Sub ExecSubCopy()

    Dim oDomHead, oDomBody As New DOMDocument
    Dim sCache As String

    bAlter = False
    
    Voucher.getVoucherDataXML oDomHead, oDomBody

    Voucher.AddNew ANMCopyALL, oDomHead, oDomBody
    Voucher.SetBillNumberRule sCache

    '�Ƶ���,��������
    Voucher.headerText(StrcMaker) = g_oLogin.cUserName
   
    Voucher.headerText(HeadPKFld) = ""                     '�����־id
    Voucher.headerText(StrcHandler) = ""                   '�����
    Voucher.headerText(StrdVeriDate) = ""                  '�������
    Voucher.headerText(StrCloseUser) = ""                  '�ر���
    Voucher.headerText(StrdCloseDate) = ""                 '�ر�����
    Voucher.headerText(StrIntoUser) = ""                   '������
    Voucher.headerText(StrdIntoDate) = ""                  '��������
    Voucher.headerText("iStatus") = "�½�"
    '    Voucher.headerText("cCreateType") = "��������"
    'ֻ��ת��������ʱ�Ÿ�Ϊ��������
  
    '    Voucher.headerText("cType") = "�ͻ�"
    'Voucher.headerText("cmemo") = ""
 

    Voucher.headerText("VoucherType") = gstrCardNumber     '"HY99"
    Dim errMsg As String
'    If getIsWfControl(g_oLogin, g_Conn, errMsg, gstrCardNumber) Then
'        Voucher.headerText("iswfcontrolled") = 1
'    Else
'        Voucher.headerText("iswfcontrolled") = 0
'    End If

    Voucher.headerText("iverifystate") = ""
    Voucher.headerText("ireturncount") = ""

 

    '���õ��ݱ���Ƿ�ɱ༭
    Dim manual As Boolean                                  ' �Ƿ���ȫ�ֹ����
    Call SetVouchCodeEnable(manual)

    If manual Then                                         '��ȫ�ֹ���������ÿ�
        Voucher.headerText(strcCode) = ""
    End If


    '��־״̬
    Voucher.VoucherStatus = VSeAddMode
    mOpStatus = ADD_MAIN

'    Call setAllDisable
    Call SetCtlStyle(Me, Voucher, Me.Toolbar, Me.UFToolbar, mOpStatus)

    '���������û�
    Me.Toolbar.Buttons("Prorefer").Enabled = False
    Me.UFToolbar.RefreshEnable

    'by zhangwchb 20110829 �ύ����
'    If bInWorkFlow = True Then
'        Me.Toolbar.Buttons("Submit").Enabled = True
'        UFToolbar.RefreshEnable
'    Else
'        Me.Toolbar.Buttons("Submit").Enabled = False
'        UFToolbar.RefreshEnable
'    End If

End Sub

Private Sub UFToolbar_OnSelectedIndexChanged(ByVal enumType As UFToolBarCtrl.ENUM_MENU_OR_BUTTON, ByVal cButtonId As String, ByVal cMenuId As String, ByVal iSelectedIndex As Integer)
    If enumType = enumCombItem Then
        Select Case LCase(cButtonId)
        Case "printtemplate"
            Me.ComTemplatePRN.ListIndex = iSelectedIndex
        Case "showtemplate"
            sPreVTID = ComTemplateShow.ListIndex
            Me.ComTemplateShow.ListIndex = iSelectedIndex
        End Select
    End If
End Sub

Private Sub Voucher_BatchModify(sItemXML As String)
    On Error Resume Next
    If sItemXML <> "" Then
        Dim oDom As New DOMDocument30
        oDom.loadXML sItemXML
        Dim ele As IXMLDOMElement
        Dim ndLst As IXMLDOMNodeList
        Set ndLst = oDom.selectSingleNode("//Data").childNodes
        For Each ele In ndLst
            If InStr(1, LCase(ele.nodename), "cfree") <> 0 Or InStr(1, LCase(ele.nodename), "cdefine") <> 0 Then
                ele.setAttribute "reftype", ""
                ele.setAttribute "cRefID", ""
            Else
                Select Case LCase(ele.nodename)
                    Case "cinvcode", "cinvname"
                        ele.setAttribute "reftype", "ref"
                        ele.setAttribute "cRefID", "Inventory_AA"
                        If LCase(getAccinformation("ST", "bInventoryCheck", g_Conn)) = "true" Then
                            ele.setAttribute "bAuth", "1"
                        Else
                            ele.setAttribute "bAuth", "0"
                        End If
                    Case "cwhcode", "cwhname"
                        ele.setAttribute "reftype", "ref"
                        ele.setAttribute "cRefID", "Warehouse_AA"
                        If LCase(getAccinformation("ST", "bWarehouseCheck", g_Conn)) = "true" Then
                            ele.setAttribute "bAuth", "1"
                        Else
                            ele.setAttribute "bAuth", "0"
                        End If
                    Case "cexpname"
                        ele.setAttribute "reftype", "ref"
                        ele.setAttribute "cRefID", "ExpenseItem_AA"
                    Case "cmemo"
                        ele.setAttribute "reftype", "ref"
                        ele.setAttribute "cRefID", "gl_bdigest_st"
                    Case "ccusinvcode", "ccusinvname"      ''�ͻ��������
                        ele.setAttribute "reftype", "ref"
                        ele.setAttribute "cRefID", "CusInvContrapose_AA"
                    Case "cinva_unit"
                        ele.setAttribute "reftype", "ref"
                        ele.setAttribute "cRefID", "ComputationUnit_AA"
                    Case "cvencode"
                        ele.setAttribute "reftype", "ref"
                        ele.setAttribute "cRefID", "Vendor_AA"
                    Case "cposition2", "cposition"
                        ele.setAttribute "reftype", "ref"
                        ele.setAttribute "cRefID", "Position_AA"
                        If LCase(getAccinformation("ST", "bPositionCheck", g_Conn)) = "true" Then
                            ele.setAttribute "bAuth", "1"
                        Else
                            ele.setAttribute "bAuth", "0"
                        End If
                    Case Else
                        If ele.getAttribute("reftype") = "ref" And ele.getAttribute("cRefID") = "" Then
                            ele.setAttribute "reftype", ""
                        End If
                End Select
            End If
        Next
        sItemXML = oDom.xml
        Set oDom = Nothing
    End If
End Sub

'�Զ���ȡ���ݺ�
Private Sub Voucher_BillNumberChecksucceed()
    Dim oDomHead As DOMDocument
    Dim oDomFormat As DOMDocument
    Dim oelement As IXMLDOMElement
    Dim bManualCode As Boolean
    Dim bCanModyCode As Boolean
    Dim strVoucherNo As String

    Dim sKey As String

    Dim sError As String

    Set oDomHead = Voucher.GetHeadDom

    If GetVoucherNO(g_Conn, oDomHead, gstrCardNumber, sError, strVoucherNo, oDomFormat, True, , False) = True Then
        Voucher.SetBillNumberRule oDomFormat.xml

        Set oelement = oDomFormat.selectSingleNode("//���ݱ��")

        '֧����ȫ�ֹ����
        '�����ֹ��޸ĵú���Ϊ ��ȫ�ֹ���ţ� �غ��Զ���ȡ�ĺ���Ϊ �ֹ��޸ģ��غ��Զ���ȡ
        bManualCode = oelement.getAttribute("�����ֹ��޸�")
        bCanModyCode = oelement.getAttribute("�����ֹ��޸�") Or oelement.getAttribute("�غ��Զ���ȡ")
    Else
        MsgBox sError, vbInformation, GetString("U8.DZ.JA.Res030")
    End If

    '֧����ȫ�ֹ���ţ���ʱ��ȡ���ݺ� 2003-07-16 �Ƴ���
    If Not bManualCode Then
        With Me.Voucher
            sKey = strcCode

            Set oDomHead = Voucher.GetHeadDom

            If GetVoucherNO(g_Conn, oDomHead, gstrCardNumber, sError, strVoucherNo, , , , False) = False Then
                MsgBox sError, vbInformation, GetString("U8.DZ.JA.Res030")
            Else
                .headerText(sKey) = strVoucherNo
            End If
        End With
    End If
End Sub

Private Sub SetVouchCodeEnable(manual As Boolean)
    If Voucher.VoucherStatus <> VSNormalMode Then
        Voucher.EnableHead strcCode, GetbCanModifyVCode(manual)
    End If
End Sub

Private Sub Voucher_bodyBrowUser(ByVal row As Long, ByVal Col As Long, sRet As Variant, referpara As UAPVoucherControl85.ReferParameter)
    '
    Call VoucherbodyBrowUser(Voucher, row, Col, sRet, referpara)
    
    '
    '     Dim sMetaItemXML As String
    '     Dim sMetaXML As String
    '    sMetaItemXML = Voucher.ItemState(Col, sibody).sFieldName
    '
    '
    '            sMetaXML = "<Ref><RefSet bAuth='0' bMultiSel= '1' /></Ref>"
    '
    ''            referpara.Cancel = False
    '            referpara.id = "Inventory_AA"
    '            referpara.sSql = " cInvCode like '%" & sRet & "%'"
    '            referpara.ReferMetaXML = sMetaXML
    '
    '    Dim vis As UAPVoucherControl85.clsItemState
    '    'ʹ���µĲ��շ���
    '    Dim objRefer As New U8RefService.IService
    '    Set vis = Voucher.ItemState(Col, 1)                    '��ס�˴����ǹؼ�
    '    Dim sqlstr As String
    '    Dim rstClass As New ADODB.Recordset
    '    Dim rstGrid As New ADODB.Recordset
    '    Dim ErrMsg As String
    '    Dim sBodyItemName As String
    '    Dim sMetaXML As String
    '    Dim i As Integer
    '    Dim rs As String
    '    Dim batchP As Boolean
    '    Dim isBatch As Boolean
    '    Dim sRefFieldName As String
    '    Dim sRefTableName As String
    '    Dim sRefCardNumber As String
    '    Dim oDefPro As U8DefPro.clsDefPro
    '    '�����Ҫ���������¼�����Ϊ������״̬
    '    referpara.Cancel = True
    '
    ''    On Error GoTo errhandle
    '    '���ò����Ƿ��ѡ
    '    sMetaXML = "<Ref><RefSet bAuth='0' bMultiSel= '1' /></Ref>"
    '    Dim moRef As New UFReferC.UFReferClient
    '    moRef.SetLogin g_oLogin
    '    sBodyItemName = LCase(Voucher.ItemState(Col, sibody).sFieldName)
    '    If sBodyItemName = "cinvcode" Then
    '        'objRefer.RefID = "Inventory_AA"
    '        referpara.Cancel = False
    '        referpara.id = "Inventory_AA"
    '        referpara.ReferMetaXML = sMetaXML
    '
    '    End If
End Sub

Private Sub Voucher_bodyCellCheck(retvalue As Variant, bChanged As Long, ByVal r As Long, ByVal c As Long, referpara As UAPVoucherControl85.ReferParameter)

    Call VoucherbodyCellCheck(Voucher, retvalue, bChanged, r, c, referpara)

End Sub


Private Sub Voucher_CopySelect(bAuther As Boolean)
    bAuther = ZwTaskExec(g_oLogin, AuthOut, 1, True)
End Sub

'Private Sub Voucher_FillHeadComboBox(ByVal Index As Long, pCom As Object)
'    Dim i As Integer
'    'enum by modify
'    If LCase(Voucher.ItemState(Index, siheader).sFieldName) = LCase("cType") Then
'        pCom.AddItem "�ͻ�"
'        pCom.AddItem "��Ӧ��"
'        pCom.AddItem "����"
'        '        pCom.AddItem "��Ա"
'    End If
'
'    If LCase(Voucher.ItemState(Index, siheader).sFieldName) = LCase("cAboutVoucher") Then
'        pCom.AddItem "���۳��ⵥ"
'        pCom.AddItem "����"
'        pCom.AddItem "ί����ⵥ"
'    End If
'
'    If LCase(Voucher.ItemState(Index, siheader).sFieldName) = LCase("cfreight") Then
'        pCom.AddItem "��"
'        pCom.AddItem "��"
'    End If
'End Sub


Private Sub Voucher_headBrowUser(ByVal Index As Variant, sRet As Variant, referpara As UAPVoucherControl85.ReferParameter)
    Call VoucherheadBrowUser(Voucher, Index, sRet, referpara)
End Sub


Private Sub Voucher_headCellCheck(Index As Variant, retvalue As String, bChanged As UAPVoucherControl85.CheckRet, referpara As UAPVoucherControl85.ReferParameter)
    Call VoucherheadCellCheckFun(Voucher, Index, retvalue, bChanged, referpara)
End Sub

Private Sub Voucher_MouseUp(ByVal section As UAPVoucherControl85.SectionsConstants, ByVal Index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    '�����Ҽ��˵�
'    If Button = 2 Then
'
'        If Me.Voucher.VoucherStatus = VSNormalMode Then
'
'            Me.PopMenu1.Visible("AddR") = False            '����
'            Me.PopMenu1.Visible("DelR") = False            'ɾ��
'            Me.PopMenu1.Visible("batchModify") = False     '����
'            Me.PopMenu1.Visible("copyR") = False           '������
'            '        Me.PopMenu1.Visible("bsplit") = False
'            Me.PopMenu1.Visible("A") = False               '����
'            Me.PopMenu1.Visible("B") = False               '����
'
'        Else
'
'            Me.PopMenu1.Visible("AddR") = False             '����
'            Me.PopMenu1.Visible("DelR") = True             'ɾ��
'            Me.PopMenu1.Visible("copyR") = True            '������
'            Me.PopMenu1.Visible("batchModify") = True      '����
'            '        Me.PopMenu1.Visible("bsplit") = True
'
'            Me.PopMenu1.Visible("A") = True                '����
'            Me.PopMenu1.Visible("B") = True                '����
'
'        End If
'
'        'retmenu:�˵����ڵ�����
'        Me.PopMenu1.ShowPopupMenu Voucher.VoucherBody, "retmenu", X, Y
'
'    End If

End Sub



Private Sub Voucher_RowChanged(ByVal vtOldRow As Variant, ByVal vtNewRow As Variant)
    Dim sInvCode As String
    Dim sError As String
    Dim tmpstr As String
    Dim nRow As Long
    nRow = Voucher.row
    If nRow <> mdOldSelRow Then
        sInvCode = Voucher.bodyText(nRow, "cInvCode")
        If sInvCode <> "" Then
            Call ShowStock(Voucher, sInvCode, nRow)
            RowChange
        End If
        mdOldSelRow = nRow
    End If
End Sub

Private Sub RowChange()
    Dim i As Long
    Dim rs As New ADODB.Recordset
    i = Voucher.row
    If i > 0 And Voucher.VoucherStatus <> VSNormalMode And Voucher.bodyText(i, "cinvcode") <> "" Then
        Set rs = cInvCodeRefer(Voucher.bodyText(i, "cinvcode"))
        SetBodyControl Voucher, rs, i
    End If
End Sub

Private Sub Voucher_RowColChange()
    Dim nCol As Long
    nCol = Voucher.Col
    If nCol <> mdOldSelCol Then
        RowChange
        mdOldSelCol = nCol
    End If
End Sub

Private Sub Voucher_SaveSettingEvent(ByVal varDevice As Variant)

    '�����ӡ��ʽ����
    Dim VoucherTempDate As Object
    Set VoucherTempDate = CreateObject("ufvoucherserver85.clsVoucherTemplate")
    VoucherTempDate.SaveDeviceCapabilities g_Conn.ConnectionString, m_strVT_PRN_ID, varDevice


    Set VoucherTempDate = Nothing

End Sub



Public Function setAllDisable() As Boolean
    On Error GoTo Err_Handler

    Dim sql As String
    Dim i As Long
    Dim rs As New ADODB.Recordset
    'enum by modify
    If Voucher.headerText("isengdec") = "��" Then
 
    Voucher.EnableHead "supengcode", False
    Voucher.EnableHead "supcname", False
    End If

  

    Set rs = Nothing
    Exit Function

Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Function



'���ݱ�ͷ�ĵ���ģ��
Private Sub setTemplate(vtid As String)

    On Error GoTo ErrHandler:

    Dim rs As ADODB.Recordset
    If vtid <> "" Then
        Set rs = g_Conn.Execute("select * from VoucherTemplates where vt_id=N'" & vtid & "' and localeid=N'" & g_oLogin.LanguageRegion & "'")
        If rs.EOF Then
            vtid = objVoucherTemplate.GetDefaultTemplate(g_Conn, gstrCardNumber, g_oLogin.cUserId)
            m_strVT_PRN_ID = objVoucherTemplate.GetDefaultTemplate(g_Conn, gstrCardNumber, g_oLogin.cUserId, False, vtid)
        End If
    Else
        vtid = objVoucherTemplate.GetDefaultTemplate(g_Conn, gstrCardNumber, g_oLogin.cUserId)
        m_strVT_PRN_ID = objVoucherTemplate.GetDefaultTemplate(g_Conn, gstrCardNumber, g_oLogin.cUserId, False, vtid)
    End If
    '�뵱ǰģ��һ��ֱ���˳�
    'If vtid = Mid(ComTemplateShow.Text, 1, InStr(1, ComTemplateShow.Text, " ") - 1) Then
    If vtid = CStr(ComTemplateShow.ItemData(ComTemplateShow.ListIndex)) Then
        m_strVT_ID = vtid
        Exit Sub
    End If
    If vtid = "" Then
        ComTemplateShow.ListIndex = 0
        vtid = Mid(ComTemplateShow.Text, 1, InStr(1, ComTemplateShow.Text, " ") - 1)
    Else
        ComTemplateShow.ListIndex = dicTemplate.Item(vtid)
    End If
    m_strVT_ID = vtid
    SetTemplateData
    ComTemplateShow.ListIndex = dicTemplate.Item(m_strVT_ID)
    ComTemplatePRN.ListIndex = dicTemplatePrint.Item(m_strVT_PRN_ID)
    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbExclamation
End Sub

'20110812
Private Sub m_mht_ReceiveMessage(ByVal sender As UFPortalProxyMessage.IMessageHandler, _
                                 ByVal sMessage As String, _
                                 vReserved As Variant)

    If (sender.MessageType = "DocAuditComplete") Then

        Dim domMessage As New DOMDocument

        Dim eleMessage As IXMLDOMElement

        domMessage.loadXML sMessage
        Set eleMessage = domMessage.selectSingleNode("//Message/Selection/Element")

        If eleMessage.getAttribute("ExecuteResult") = True And eleMessage.getAttribute("cCardNum") = gstrCardNumber Then
            If eleMessage.getAttribute("cVoucherId") = GetHeadItemValue(Voucher.GetHeadDom, "ID") Then
                Call ButtonClick(sKey_Refresh)
            End If
        End If

        Set domMessage = Nothing
    End If

End Sub

Private Sub RegisterMessage()
    Set m_mht = New UFPortalProxyMessage.IMessageHandler
    m_mht.MessageType = "DocAuditComplete"
    If Not g_oBusiness Is Nothing Then
        Call g_oBusiness.RegisterMessageHandler(m_mht)
    End If
End Sub

Private Sub UnRegisterMessage()
    If m_mht Is Nothing Then Exit Sub
    If Not g_oBusiness Is Nothing Then
        Call g_oBusiness.UnregisterMessageHandler(m_mht)
    End If
End Sub

'��ʼ���Զ���������� 20110822
Private Sub initCustomRelation()

    Voucher.SetCustomRelation mobjSubServ.GetCustomRelationRecord(g_Conn, gstrCardNumber)

End Sub

'���ݿؼ����Զ�������¼� 20110822
Private Sub Voucher_AutoFillBackEvent(vtIndex As Variant, _
                                         ByVal vtCurrentValue As Variant, _
                                         ByVal vtCurrentFieldObject As Variant, _
                                         ByVal vtAutoFieldInfo As Variant)

    Dim strErr As String

    On Error GoTo ErrHandle

    mobjSubServ.AutoFillRelations g_Conn, Voucher, vtCurrentFieldObject, vtAutoFieldInfo, strErr

    Exit Sub

ErrHandle:
    MsgBox strErr, vbInformation, GetString("U8.DZ.JA.Res030")
End Sub

Private Function Check_Auth() As Boolean
    'Ȩ�޼��
    Dim rs As New ADODB.Recordset
    Dim sql As String
    Dim ctype  As String, bObjectCode As String
    Dim i As Long
    Check_Auth = True
'    ctype = Voucher.headerText("cType")
'    bObjectCode = Voucher.headerText("bObjectCode") & ""
'    If Voucher.headerText("bObjectCode") & "" <> "" Then
'        Select Case ctype
'            Case "��Ӧ��"
'                sql = "SELECT distinct bObjectCode  from HY_DZ_BorrowOut a left join vendor b on a.bObjectCode=b.cvencode where a.bObjectCode='" & Voucher.headerText("bObjectCode") & "' "
'                sql = sql & IIf(sAuth_vendorW = "", "", " and b.iid in (" & sAuth_vendorW & ")")
'            Case "����"
'                sql = "SELECT distinct bObjectCode  from HY_DZ_BorrowOut a left join department b on a.bObjectCode=b.cdepcode where a.bObjectCode='" & Voucher.headerText("bObjectCode") & "' "
'                sql = sql & IIf(sAuth_depW = "", "", " and b.cdepcode in (" & sAuth_depW & ")")
'            Case "�ͻ�"
'                sql = "SELECT distinct bObjectCode  from HY_DZ_BorrowOut a left join customer b on a.bObjectCode=b.ccuscode where a.bObjectCode='" & Voucher.headerText("bObjectCode") & "' "
'                sql = sql & IIf(sAuth_CusW = "", "", " and b.iid in (" & sAuth_CusW & ")")
'        End Select
'        If sql <> "" Then
'        rs.Open sql, g_Conn
'        If rs.EOF Or rs.BOF Then
'
'            Select Case ctype
'                Case "��Ӧ��"
'                    MsgBox GetStringPara(("U8.DZ.JA.Res2040"), bObjectCode), vbInformation, GetString("U8.DZ.JA.Res030")
'                    Check_Auth = False
'                    Exit Function
'                Case "����"
'                    MsgBox GetStringPara(("U8.DZ.JA.Res2050"), bObjectCode), vbInformation, GetString("U8.DZ.JA.Res030")
'                    Check_Auth = False
'                    Exit Function
'                Case "�ͻ�"
'                    MsgBox GetStringPara(("U8.DZ.JA.Res2060"), bObjectCode), vbInformation, GetString("U8.DZ.JA.Res030")
'                    Check_Auth = False
'                    Exit Function
'            End Select
'
'            Check_Auth = False
'            Exit Function
'        End If
'        End If
'    End If
    '����
    If Voucher.headerText("chdepartcode") <> "" Then
        sql = "SELECT distinct chdepartcode  from HY_FYSL_Payment a left join department b on a.chdepartcode=b.cdepcode where a.chdepartcode='" & Voucher.headerText("chdepartcode") & "' "
        sql = sql & IIf(sAuth_depW = "", "", " and b.cdepcode in (" & sAuth_depW & ")")

        Set rs = New ADODB.Recordset
        rs.Open sql, g_Conn
        If rs.EOF Or rs.BOF Then
            MsgBox GetStringPara(("U8.DZ.JA.Res2050"), Voucher.headerText("chdepartcode")), vbInformation, GetString("U8.DZ.JA.Res030")
            Check_Auth = False
            Exit Function
        End If
    End If

'    For i = 1 To Voucher.BodyRows
'        '���
'
'        sql = "select a.cinvcode from HY_DZ_BorrowOuts a left join inventory i on a.cinvcode=i.cinvcode where a.cinvcode='" & Voucher.bodyText(i, "cinvcode") & "' " & IIf(sAuth_invW = "", "", " and i.iid in (" & sAuth_invW & ")")
'
'        Set rs = New ADODB.Recordset
'        rs.Open sql, g_Conn
'        If rs.EOF Or rs.BOF Then
'            MsgBox GetStringPara(("U8.DZ.JA.Res2080"), Voucher.bodyText(i, "cinvcode")), vbInformation, GetString("U8.DZ.JA.Res030")
'            Check_Auth = False
'            Exit Function
'        End If
'        '�ֿ�
'        sql = "select a.cinvcode from HY_DZ_BorrowOuts a  left join Warehouse b on a.cwhcode=b.cwhcode where cinvcode='" & Voucher.bodyText(i, "cinvcode") & "'" & IIf(sAuth_WareHouseW = "", "", " and( ISNULL(b.cwhcode,N'')=N'' OR b.cwhcode in (" & sAuth_WareHouseW & "))")
'
'        Set rs = New ADODB.Recordset
'        rs.Open sql, g_Conn
'        If rs.EOF Or rs.BOF Then
'            MsgBox GetStringPara(("U8.DZ.JA.Res2090"), Voucher.bodyText(i, "cwhcode")), vbInformation, GetString("U8.DZ.JA.Res030")
'            Check_Auth = False
'            Exit Function
'        End If
'        '��λ
'        If Voucher.bodyText(i, "cPosition") & "" <> "" Then
'            sql = "select a.cinvcode from HY_DZ_BorrowOuts a  left join Position b on a.cPosition=b.cposcode  where a.cinvcode='" & Voucher.bodyText(i, "cinvcode") & "'" & IIf(sAuth_PositionW = "", "", " and b.cposcode in (" & sAuth_PositionW & ")")
'
'            Set rs = New ADODB.Recordset
'            rs.Open sql, g_Conn
'            If rs.EOF Or rs.BOF Then
'                MsgBox GetStringPara(("U8.DZ.JA.Res2100"), Voucher.bodyText(i, "cPosition")), vbInformation, GetString("U8.DZ.JA.Res030")
'                Check_Auth = False
'                Exit Function
'            End If
'        End If
'    Next i

End Function

'�жϵ��������Ƿ�����������  'by zhangwchb 20110829 �ύ����
Private Function bInWorkFlow() As Boolean
    Dim AuditServiceProxy As Object
    Dim objCalledContext As Object
    Dim strErr As String
    Dim blnSubmit As Boolean

    Set objCalledContext = CreateObject("UFSoft.U8.Framework.LoginContext.CalledContext")
    objCalledContext.SubId = g_oLogin.cSub_Id
    objCalledContext.TaskId = g_oLogin.TaskId
    objCalledContext.token = g_oLogin.userToken
    Set AuditServiceProxy = CreateObject("UFIDA.U8.Audit.ServiceProxy.AuditServiceProxy")
    blnSubmit = AuditServiceProxy.isflowenabled(gstrCardNumber, gstrCardNumber & ".Submit", objCalledContext, strErr)

    If strErr = "" Then
        bInWorkFlow = blnSubmit
    End If

    Set objCalledContext = Nothing
    Set AuditServiceProxy = Nothing
End Function

Private Function SetStamp(ctlVoucher As UAPVoucherControl85.ctlVoucher)
    ctlVoucher.SetStampAll
End Function

'wangfb 11.0ToolbarǨ�� 2012-03-31
Private Function SetToolbarVisible()
    On Error Resume Next
    With Toolbar
        '11.0�¹淶���ù���������ʾ
'        If IsWFControlled() Then
'            .Buttons("Submit").Visible = True
'            .Buttons("Resubmit").Visible = True
'            .Buttons("Unsubmit").Visible = True
'            .Buttons("ViewVerify").Visible = True
'        Else
'            .Buttons("Submit").Visible = False
'            .Buttons("Resubmit").Visible = False
'            .Buttons("Unsubmit").Visible = False
'            .Buttons("ViewVerify").Visible = False
'        End If
        '������Ԥ�����������������Ƶ���ȡ�۰�ť����
 
        .Buttons("Preview").Visible = False
        .Buttons("Help").Visible = False
        .Buttons("Locate").Visible = False
        
        If Voucher.VoucherStatus <> VSNormalMode Then
            '11.0ȥ�����а�ť
            '��ʾ��ʽ�µĺϲ���ʾ��ȡ���ϲ������úϲ���������
     
            Toolbar.Buttons(sKey_RefVoucher).Visible = False
            
            If Voucher.VoucherStatus = VSeAddMode Then
                '����ʱ���ۿ��ã�����ע��֪ͨ������
                Toolbar.Buttons("Notes").Enabled = False
                Toolbar.Buttons("Discuss").Enabled = True
                Toolbar.Buttons("Notify").Enabled = False
            Else
                '�޸�ʱ��ע�����ۿ��ã���֪ͨ������
                Toolbar.Buttons("Notes").Enabled = True
                Toolbar.Buttons("Discuss").Enabled = True
                Toolbar.Buttons("Notify").Enabled = False
            End If
            Toolbar.Buttons("tlbLinkAllVouch").Enabled = False
        Else
             
            
            '�հ׵���ʱ��ע�����ۺ�֪ͨ�������ã����򶼿���
            If Voucher.headerText("ccode") <> "" Then
                Toolbar.Buttons("Notes").Enabled = True
                Toolbar.Buttons("Discuss").Enabled = True
                Toolbar.Buttons("Notify").Enabled = True
                Toolbar.Buttons("tlbLinkAllVouch").Enabled = True
            Else
                Toolbar.Buttons("Notes").Enabled = False
                Toolbar.Buttons("Discuss").Enabled = False
                Toolbar.Buttons("Notify").Enabled = False
                Toolbar.Buttons("tlbLinkAllVouch").Enabled = False
            End If
             
            
        End If
    End With
    Me.ComTemplatePRN.Visible = False
    Me.ComTemplateShow.Visible = False
    UFToolbar.RefreshVisible
    UFToolbar.RefreshEnable
    VBA.Err.Clear
End Function

Private Sub Voucher_SearchClick(ByVal cSearchKey As String)
    Dim tmpid As String
    Dim tmpRst As New ADODB.Recordset
    Dim strSql As String
    Dim oId As String
    If sTmpTableName = "" Then
        sTmpTableName = "tempdb..TEMP_STSearchTableName_" & sGUID
    End If
    DropTable sTmpTableName
          
    strSql = "select ID as cVoucherId,ccode as cVoucherCode,cast(null as nvarchar(1)) as cVoucherName,cast(null as nvarchar(1)) as cCardNum,cast(null as nvarchar(1)) as cMenu_Id,cast(null as nvarchar(1)) as cAuth_Id,cast(null as nvarchar(1)) as cSub_Id into " & sTmpTableName & " from " & MainTable & "  where  (ccode like N'%" & Trim(cSearchKey) & "%')"
    
    If sAuth_ALL <> "" Then
        strSql = strSql + " and  (" + sAuth_ALL + ") "
    End If
    
    mologin.AccountConnection.Execute strSql
    strSql = "select cVoucherId from " & sTmpTableName
    tmpRst.Open strSql, mologin.AccountConnection, adOpenForwardOnly, adLockReadOnly
    If Not tmpRst.EOF Then
        tmpid = tmpRst(0)
        Voucher.SearchTableName = sTmpTableName
        If tmpRst.RecordCount = 1 Then
            sTmpTableName = ""
        End If
    Else
        tmpid = ""
        sTmpTableName = ""
    End If
    tmpRst.Close
    Set tmpRst = Nothing
    If CStr(lngVoucherID) <> tmpid And tmpid <> "" Then
        lngVoucherID = tmpid
        ExecSubRefresh
    End If
End Sub
Private Sub Voucher_ReleaseSearchClick()
    'DropTable IIf(sTmpTableName = "", sTmpTableName, Voucher.SearchTableName)
    sTmpTableName = ""
    Voucher.SearchTableName = ""
End Sub

Private Sub Voucher_GoToVoucher(ByVal cVoucherInfo As String)
    lngVoucherID = cVoucherInfo
    ExecSubRefresh
End Sub
Public Sub SetSearchState4List()
    If sTmpTableName <> "" Then
        Voucher.SearchTableName = sTmpTableName
        Voucher.SearchValueType = ListSearch
        Voucher.ClearButtonVisible = True
    Else
        Voucher.SearchTableName = ""
        Voucher.ClearButtonVisible = False
    End If
End Sub
Private Sub ShowRefVouchers(ByVal bAll As Boolean)
    Dim sRet         As String
    Dim strTable     As String
    Dim strLinkField As String
    Dim strLinkFieldValues As String
    Dim i As Integer

    If Voucher.headerText("ccreatetype") = "ת������" Then
        strTable = "hy_dz_borrowouts as jachangeout"
    Else
        strTable = "hy_dz_borrowouts"
    End If

    strLinkField = "autoid"
    
    If bAll Then
        For i = 1 To Voucher.BodyRows
            If Voucher.bodyText(i, strLinkField) <> "" Then
                If i = 1 Then
                    strLinkFieldValues = Voucher.bodyText(i, strLinkField)
                Else
                    strLinkFieldValues = strLinkFieldValues + "," + Voucher.bodyText(i, strLinkField)
                End If
            End If
        Next
    Else
        strLinkFieldValues = Voucher.bodyText(Voucher.row, strLinkField)
    End If

    Set objRelation = CreateObject("SCMBillRelation.clsBill")
    Set objRelation.Business = g_oBusiness
    objRelation.InitLogin g_oLogin

    If Trim$(Voucher.bodyText(Voucher.row, strLinkField)) <> "" Or bAll Then
        sRet = objRelation.getLinkVouchsMenu(strTable, strLinkFieldValues, "ST")

        If sRet = "" Then Exit Sub

        Call objRelation.showList(strTable, strLinkFieldValues, "ST", sRet)

    End If
End Sub

Public Function processdataforcon() As Boolean

    Dim strSql   As String

    Dim rs       As New ADODB.Recordset

    Dim oDomHead As DOMDocument

    processdataforcon = True
 
    strSql = "select 'Y' as selcol, * from  V_HY_FYSL_Payment_refer  where  id ='" & conid & "'"
    Set rs = New ADODB.Recordset
    rs.Open strSql, g_Conn
    Set oDomHead = New DOMDocument
    If Not rs.EOF Then
        rs.Save oDomHead, adPersistXML
        Set gDomReferHead = oDomHead
        ProcessData Voucher
        processdataforcon = True
    Else
        processdataforcon = False
    End If

End Function
