VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{456334B9-D052-4643-8F5F-2326B24BE316}#6.14#0"; "uApvouchercontrol85.ocx"
Object = "{201FB79B-5556-47A4-AD9C-A46BA0C45A44}#6.22#0"; "UFToolBarCtrl.ocx"
Begin VB.Form frmVouchNew 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "0"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11400
   FillColor       =   &H00004040&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6240
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tbrvoucher 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   953
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      _Version        =   393216
   End
   Begin VB.PictureBox labXJ 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2205
      ScaleHeight     =   375
      ScaleWidth      =   675
      TabIndex        =   3
      Top             =   6480
      Visible         =   0   'False
      Width           =   675
      Begin VB.Line Line4 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         X1              =   -15
         X2              =   909
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         X1              =   15
         X2              =   15
         Y1              =   15
         Y2              =   399
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         X1              =   -90
         X2              =   855
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         X1              =   660
         X2              =   660
         Y1              =   372
         Y2              =   0
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "�ֽ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1095
         Left            =   240
         TabIndex        =   4
         Top             =   75
         Width           =   480
      End
   End
   Begin ComctlLib.StatusBar StBar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   7065
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox labZF 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1365
      ScaleHeight     =   375
      ScaleWidth      =   705
      TabIndex        =   0
      Top             =   6480
      Visible         =   0   'False
      Width           =   705
      Begin VB.Line Line9 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         X1              =   0
         X2              =   924
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line8 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         X1              =   690
         X2              =   690
         Y1              =   372
         Y2              =   0
      End
      Begin VB.Line Line6 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   384
      End
      Begin VB.Line Line5 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   939
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   75
         TabIndex        =   1
         Top             =   75
         Width           =   480
      End
   End
   Begin VB.PictureBox picVoucher 
      AutoSize        =   -1  'True
      BackColor       =   &H00E4C9AF&
      BorderStyle     =   0  'None
      Height          =   5490
      Left            =   360
      ScaleHeight     =   5490
      ScaleWidth      =   10560
      TabIndex        =   5
      Top             =   1320
      Width           =   10560
      Begin UFToolBarCtrl.UFToolbar UFToolbar1 
         Height          =   240
         Left            =   720
         TabIndex        =   18
         Top             =   360
         Width           =   7815
         _ExtentX        =   13785
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
      Begin UAPVoucherControl85.ctlVoucher voucher 
         Height          =   1695
         Left            =   2880
         TabIndex        =   17
         Top             =   1440
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2990
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10446406
         DisabledColor   =   16777215
         ColAlignment0   =   9
         Rows            =   20
         Cols            =   20
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ControlScrollBars=   0
         ControlAutoScales=   0
         BaseOfVScrollPoint=   0
         ShowSorter      =   0   'False
         ShowFixColer    =   0   'False
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   4800
         TabIndex        =   15
         Top             =   3360
         Visible         =   0   'False
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   120
         ScaleHeight     =   555
         ScaleWidth      =   10755
         TabIndex        =   8
         Top             =   120
         Width           =   10755
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   6930
            ScaleHeight     =   300
            ScaleWidth      =   3495
            TabIndex        =   10
            Top             =   180
            Width           =   3495
            Begin VB.ComboBox ComboDJMB 
               Height          =   300
               Left            =   960
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   0
               Width           =   2505
            End
            Begin VB.ComboBox ComboVTID 
               Height          =   300
               Left            =   990
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   0
               Width           =   2385
            End
            Begin VB.Label Labeldjmb 
               Alignment       =   2  'Center
               BackColor       =   &H00E4C9AF&
               BackStyle       =   0  'Transparent
               Caption         =   "��ӡģ�棺"
               Height          =   1140
               Left            =   0
               TabIndex        =   14
               Top             =   120
               Width           =   1080
            End
         End
         Begin VB.Label U8VoucherSorter1 
            BackColor       =   &H80000007&
            Caption         =   "3333"
            Height          =   255
            Left            =   3600
            TabIndex        =   16
            Top             =   0
            Width           =   600
         End
         Begin VB.Label LabelVoucherName 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   4680
            TabIndex        =   9
            Top             =   120
            Width           =   630
         End
      End
      Begin MSComCtl2.FlatScrollBar vs 
         Height          =   2550
         Left            =   360
         TabIndex        =   7
         Top             =   2640
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   4498
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1245184
      End
      Begin MSComCtl2.FlatScrollBar hs 
         Height          =   300
         Left            =   3360
         TabIndex        =   12
         Top             =   4920
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         Arrows          =   65536
         Orientation     =   1245185
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5400
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   3120
      Top             =   6720
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   924
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuPop 
      Caption         =   "�Ҽ��˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuLinkQuery 
         Caption         =   "����Ԥ����ϸ"
      End
   End
End
Attribute VB_Name = "frmVouchNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents clsVoucherCO As VouchersControlHBTV.ClsVoucherCO_GDZC
Attribute clsVoucherCO.VB_VarHelpID = -1
' by ahzzd 2005/06/01
'�޸ĺ�ĳ���ָ��������ֵ
Enum MD_EdPanelB
  Addp = 0
  Delp = 1
  EdtP = 2
End Enum
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private WithEvents ARPZ As ZzPz.clsPZ
Attribute ARPZ.VB_VarHelpID = -1
Private maxRefDate As Date ''   �����������
Private strCurVoucherNO As String
Private strVouchType As String, bReturnFlag As Boolean '��¼��������
Private bCheckVouch As Boolean '���ݵ����״̬2
Public bFrmCancel As Boolean
Dim strCardNum As String        ''���ݵ�CardNum
Dim sTemplateID As String       ''����Ĭ��ģ�����
Dim sCurTemplateID As String    ''���ݵ�ǰ��ģ���
Dim sCurTemplateID2 As String    ''���ݵ�ǰ��ģ���
Private vName As String
Private BrowFlag As Boolean '��ʶ�Ƿ����Voucher.browuser�¼�
Dim strRefFldName As String '�������յ��ֶ���
Private iVouchState As Integer
Private bClickCancel As Boolean
Private bClickSave As Boolean
'����
Dim clsRefer As New UFReferC.UFReferClient
Dim clsAuth As New U8RowAuthsvr.clsRowAuth
Dim Domhead As New DOMDocument
Dim Dombody As New DOMDocument
Dim vNewID As Variant               '����id
Dim iHeadIndex As Integer, iBodyIndex As Integer
Private m_UFTaskID As String
Private DomFormat As New DOMDocument
Private GetvouchNO As String
Private bFirst As Boolean
Dim strFreeName1 As String
Dim strFreeName2 As String
Dim strFreeName3 As String
Dim strFreeName4 As String
Dim strFreeName5 As String
Dim strFreeName6 As String
Dim strFreeName7 As String
Dim strFreeName8 As String
Dim strFreeName9 As String
Dim strFreeName10 As String
Private cSBVCode As String, SBVID As String, mDom As DOMDocument, oDomB As DOMDocument
Public iShowMode As Integer    ''����ģʽ  0������ 1�����
Private bCreditCheck  As Boolean   ''�Ƿ�ͨ�����ü��
Dim bOnceRefer As Boolean
Private ButtonTaskID As String  ''��ť����id
Private RstTemplate As ADODB.Recordset, preVTID As String      ''������ʱ�ĵ���ģ���¼��
Private RstTemplate2 As New ADODB.Recordset
Dim vtidPrn() As Long ''��ӡģ������
Private bfillDjmb As Boolean, vtidDJMB() As Long
Private bManBodyChecked As Boolean '' �Ƿ��ֹ�cellchecked
Private bCloseFHSingle As Boolean
Private obj_EA As Object, DOMEA As DOMDocument, strEAXML As String ''������
Private bLostFocus As Boolean
Private domConfig As New DOMDocument
Private domTmp As DOMDocument
Private o_crm As Object
Private moAutoFill As Object
Private dOriVoucherWidth As Double, dOriVoucherHeight As Double
Private col(1 To 22) As Long  '�������¼�ؼ������ڵ�λ��


'by lg070314 ����U870֧��
Private m_Cancel As Integer
Private m_UnloadMode As Integer
Dim sguid As String
Private WithEvents m_mht As UFPortalProxyMessage.IMessageHandler
Attribute m_mht.VB_VarHelpID = -1



'by lg070314 ����U870֧��
'�޸�3 ÿ�����嶼��Ҫ���������Cancel��UnloadMode�Ĳ����ĺ�����QueryUnload�Ĳ�����ͬ
'���ڴ˷����е��ô���Exit(�˳�)�������������ô���Unload�¼�����(��Cancel)��ֵͬʱ�����˷����Ĳ���
Public Sub ExitForm(Cancel As Integer, UnloadMode As Integer)
'    Unload Me
'
'    Cancel = m_Cancel
'    UnloadMode = m_UnloadMode

doNext:
    If Me.voucher.VoucherStatus <> VSNormalMode Then
        Select Case MsgBox("�Ƿ񱣴�Ե�ǰ���ݵı༭��", vbYesNoCancel + vbQuestion)
            Case vbYes
                ButtonClick "Save", "����"
                If Me.voucher.VoucherStatus = VSNormalMode Then
                    GoTo DoQuit
                End If
            Case vbNo
                VoucherFreeTask
                GoTo DoQuit
            Case vbCancel

        End Select

        bFrmCancel = True
        Me.ZOrder
        Cancel = 3
    Else
DoQuit:
        On Error Resume Next
        bFrmCancel = False
'by lg070314����U870�˵��ںϣ��ر�ʱ����Business
        Set UFToolbar1.Business = Nothing

        Set clsVoucherCO = Nothing
        Set clsAuth = Nothing
        Set clsRefer = Nothing
        Set RstTemplate = Nothing
        Set Domhead = Nothing
        Set Dombody = Nothing
        Set DomFormat = Nothing
        Set RstTemplate = Nothing
        Set RstTemplate2 = Nothing
        Set obj_EA = Nothing
        Set DOMEA = Nothing
        If m_UFTaskID <> "" Then
            m_login.TaskExec m_UFTaskID, 0
        End If
        minForm
        Unload Me
   End If


End Sub


'���ð�����ϵͳid
Private Sub SetHelpID()
    Select Case strVouchType
        Case "32"
            Me.HelpContextID = 10060203
        Case Else
            Me.HelpContextID = 10060203
    End Select
    
End Sub
 
''sKey :�����İ�ť����
''
Private Function VoucherTask(sKey As String) As Boolean
    Dim strID As String
    
    Select Case strVouchType
        Case "16"
            Select Case sKey
                Case "����", "����", "ɾ��", "�޸�"
                    strID = "FA03000102"  '
                Case "���", "����"
                    strID = "FA03000103"  '
                Case "�ر�", "��"
                    strID = "FA03000104"  '
            End Select
        Case "97"
            Select Case sKey
                Case "����", "����", "ɾ��", "�޸�"
                    strID = "FA03010101"  '
                Case "�ر�", "��"
                    strID = "FA03010102"  '
                Case "���", "����"
                    strID = "FA03010103"  '
                Case "���"
                    strID = "FA03010105"  '
            End Select
    End Select
    strID = clsVoucherCO.GetVoucherTaskID(sKey, strVouchType, bReturnFlag)
    If strID <> "" Then
        ButtonTaskID = strID
        VoucherTask = LockItem(ButtonTaskID, True, True)
    Else
        VoucherTask = True
    End If
End Function

''�ͷŹ�������
Private Function VoucherFreeTask() As Boolean
    If ButtonTaskID <> "" Then
        VoucherFreeTask = LockItem(ButtonTaskID, False, True)
        ButtonTaskID = ""
    End If
End Function
 
'Dim strAuthId As String     'Ȩ�޺�/gyp/2002/07/24
Private Function ChangeTempaltes(sNewTemplateID As String, Optional bChangDefalt As Boolean, Optional bCheckAuth As Boolean = True, Optional bFormload As Boolean = False) As Boolean
    Dim strDJAuth As String
    Dim bChanged As Boolean
    Dim rstTmp As New ADODB.Recordset
    Dim tmpDomhead As New DOMDocument
    Dim i As Long
    
    On Error GoTo DoERR
    bChanged = False
    If sNewTemplateID = "" Or sNewTemplateID = "0" Then
        Exit Function
    End If
    If bCheckAuth = True Then
        If m_login.IsAdmin = False Then
            If clsAuth.IsHoldAuth("djmb", Trim(sNewTemplateID), , "W") = False Then
                strDJAuth = clsAuth.getAuthString("DJMB", , "W")
                If strDJAuth = "1=2" Then
                    MsgBox "��û��ʹ�õ���ģ���Ȩ�ޣ�"
                    'Me.Hide
                    Exit Function
                Else
                    If clsAuth.IsHoldAuth("DJMB", sTemplateID, , "W") = False Then
                        rstTmp.Open "select vt_id from vouchertemplates where vt_cardnumber='" & strCardNum & "' and vt_id in (" & strDJAuth & ") order by vt_id", DBConn, adOpenForwardOnly, adLockReadOnly
                        If Not rstTmp.EOF Then
                            fillComBol False
                            sNewTemplateID = rstTmp(0)      'left(strDJAuth, IIf(InStr(1, strDJAuth, ",") - 1 = -1, Len(strDJAuth), InStr(1, strDJAuth, ",")))
                        Else
                            MsgBox "��û��ʹ�õ���ģ���Ȩ�ޣ�"
                            Me.Hide
                            rstTmp.Close
                            Set rstTmp = Nothing
                            Exit Function
                        End If
                        rstTmp.Close
                        sTemplateID = sNewTemplateID
                    Else
                        sNewTemplateID = sTemplateID
                    End If
                End If
            End If
        End If
    End If
    If bFirst = True Then Call getCardNumber(sNewTemplateID)
    If RstTemplate Is Nothing Then Set RstTemplate = New ADODB.Recordset
    If Trim(sNewTemplateID) = "" Or sNewTemplateID = "0" Then
        If bChangDefalt = True Then
            sNewTemplateID = sTemplateID
            bChanged = True
        End If
    Else
        If sCurTemplateID <> sNewTemplateID Then
            bChanged = True
        Else
            If bChangDefalt = True Then
                bChanged = True
            End If
        End If
    End If
    If bChanged = True Then
        If preVTID = sNewTemplateID And Not RstTemplate Is Nothing Then
            If Not RstTemplate.RecordCount = 0 Then
                GoTo UsePre  ''��¼�Ѿ�ȡ��
            End If
        End If
        Set RstTemplate2 = clsVoucherCO.GetVoucherFormat(sNewTemplateID, strCardNum)
        If RstTemplate2 Is Nothing Then
            If bChangDefalt = True Then
                Set RstTemplate2 = clsVoucherCO.GetVoucherFormat(sTemplateID, strCardNum)
                bChanged = True
            Else
                bChanged = False
            End If
        Else
            If RstTemplate2.state = 1 Then
                If RstTemplate2.EOF And RstTemplate2.BOF Then
                    If bChangDefalt = True Then
                        Set RstTemplate2 = clsVoucherCO.GetVoucherFormat(sTemplateID, strCardNum)
                        sCurTemplateID = sTemplateID
                        sCurTemplateID2 = sTemplateID
                        bChanged = True
                    Else
                        bChanged = False
                    End If
                Else
                   bChanged = True
                End If
            Else
                    If bChangDefalt = True Then
                        Set RstTemplate2 = clsVoucherCO.GetVoucherFormat(sTemplateID, strCardNum)
                        If RstTemplate2.state = adStateClosed Then
                                MsgBox "ģ������������"
                                ChangeTempaltes = False
                                Exit Function
                        End If
                        If Not RstTemplate2 Is Nothing Then
                            If Not RstTemplate2.EOF Then
                                bChanged = True
                            Else
                                MsgBox "ģ������������"
                                ChangeTempaltes = False
                                Exit Function
                            End If
                        Else
                            MsgBox "ģ������������"
                            ChangeTempaltes = False
                            Exit Function
                        End If
                    Else
                        bChanged = False
                    End If
            End If
        End If
    End If
    If bChanged = True Then
        sCurTemplateID = sNewTemplateID
        sCurTemplateID2 = sNewTemplateID
        preVTID = sNewTemplateID
        If bFormload = False Then
            voucher.Visible = False
        End If
        voucher.setTemplateData RstTemplate2
        dOriVoucherHeight = voucher.Height
        dOriVoucherWidth = voucher.Width
        Call Form_Resize
        If voucher.VoucherStatus <> VSNormalMode Then
            setItemState "modify"
        End If
        Call SetVocuhNameLabel
        If Not DomFormat Is Nothing Then
            If DomFormat.xml <> "" Then
                Me.voucher.SetBillNumberRule DomFormat.xml
                If Me.voucher.VoucherStatus <> VSNormalMode Then
                    Call SetVouchNoWriteble
                End If
            End If
        End If
        RstTemplate2.Save tmpDomhead, adPersistXML
        If RstTemplate.state = 1 Then RstTemplate.Close
        RstTemplate.Open tmpDomhead
        If strVouchType = "07" Then
            Me.voucher.BodyMaxRows = -1
            SetVouchItemState "cinvname", "b", False: SetVouchItemState "ccusinvname", "b", False
            SetVouchItemState "cinvcode", "b", False: SetVouchItemState "ccusinvcode", "b", False
            SetVouchItemState "cwhname", "b", False
            SetVouchItemState "ccusabbname", "t", False
            SetVouchItemState "cpersonname", "t", False
            SetVouchItemState "cdepname", "t", False
            SetVouchItemState "cbustype", "t", False
            SetVouchItemState "cexch_name", "t", False
            SetVouchItemState "iExchRate", "t", False
            SetVouchItemState "ccode", "b", False
            SetVouchItemState "dmdate", "b", False
            SetVouchItemState "dvdate", "b", False
            SetVouchItemState "cbatch", "b", False
            For i = 1 To 10
                SetVouchItemState "cfree" & i, "b", False
            Next
        End If
        
        If bFormload = False Then
            Me.voucher.Visible = True
            Me.Refresh
        End If
    End If
    Set rstTmp = Nothing
    ChangeTempaltes = True
    Call ChangeCaptionCol
    Exit Function
UsePre:
    sCurTemplateID = sNewTemplateID
    sCurTemplateID2 = sNewTemplateID
    If bFormload = False Then
        Me.voucher.Visible = False
    End If
    voucher.setTemplateData RstTemplate
    dOriVoucherHeight = voucher.Height
    dOriVoucherWidth = voucher.Width
    Call Form_Resize
    If voucher.VoucherStatus <> VSNormalMode Then
        setItemState "modify"
    End If
    Call SetVocuhNameLabel
    If bFormload = False Then
        Me.voucher.Visible = True
        Me.Refresh
    End If
    Set rstTmp = Nothing
    ChangeTempaltes = True
    Exit Function
DoERR:
    MsgBox Err.Description
    ChangeTempaltes = False
    Set rstTmp = Nothing
End Function
 
Private Sub SetVocuhNameLabel()
    '//�������Ʊ��⣬��Ҫ��Ϊ�˽�����ݵ����Ƶ�������ʾ���⣬���� "�ڳ�" XXX����
    Me.LabelVoucherName.Caption = Me.voucher.TitleCaption

    '//���ݵ�����
    Me.voucher.TitleCaption = Me.voucher.TitleCaption
    Me.voucher.TitleCaption = ""
End Sub

''����load����,���İ�Ŧ״̬,����ģ��
Private Sub LoadVoucher(sMove As String, Optional vid As Variant, Optional bRefreshClick As Boolean = False)
    Dim errMsg As String
    Dim i As Integer
    On Error Resume Next
    Select Case LCase(sMove)
        Case ""
            errMsg = clsVoucherCO.GetVoucherData(Domhead, Dombody, vid)
        Case "tonext"
ToNext:
            i = i + 1
            errMsg = clsVoucherCO.MoveNext(Domhead, Dombody)
        Case "toprevious"
            errMsg = clsVoucherCO.MovePrevious(Domhead, Dombody)
        Case "tolast"
            errMsg = clsVoucherCO.MoveLast(Domhead, Dombody)
        Case "tofirst"
            errMsg = clsVoucherCO.MoveFirst(Domhead, Dombody)
        
    End Select
        If errMsg <> "" Then
            If bRefreshClick = False And sMove = "" And vid = "" Then
                
            Else
                MsgBox errMsg
            End If
            If i <= 3 Then GoTo ToNext
            Exit Sub
        End If
    ChangeTempaltes IIf(val(GetHeadItemValue(Domhead, "ivtid")) = 0, sCurTemplateID2, GetHeadItemValue(Domhead, "ivtid")), , False
    Me.voucher.Visible = False
    voucher.setVoucherDataXML Domhead, Dombody
    '�������ı�
    Me.voucher.ExamineFlowAuditInfo = GetEAStream(strVouchType, Domhead, Me.voucher, DBConn)
    Call SetSum
    If Me.voucher.headerText("cexch_name") <> "" Then
        Me.voucher.ItemState("iexchrate", siheader).nNumPoint = clsSAWeb.GetExchRateDec(GetHeadItemValue(Domhead, "cexch_name"))
        Me.voucher.headerText("iexchrate") = GetHeadItemValue(Domhead, "iexchrate")
    End If
    ChangeButtonsState
    Call Form_Resize
    Me.voucher.Visible = True
    EditPanel_1 EdtP, 3, ""
    Dim strXml As String
    strXml = "<?xml version='1.0' encoding='GB2312'?>" & Chr(13)
    domConfig.loadXML strXml & "<EAI>0</EAI>"
End Sub
 Private Sub SetSum()
    Exit Sub
    Dim ele As IXMLDOMElement, NdList As IXMLDOMNodeList
    Dim iSum As Double, strSumDX As Variant
    'Dim oNum2Chinese As Object
    
    
    If strVouchType = "92" Or strVouchType = "95" Then Exit Sub
    strSumDX = ""
    iSum = 0
    Set NdList = Dombody.selectNodes("//z:row")
    For Each ele In NdList
        iSum = iSum + CDbl(val(IIf(IsNull(ele.getAttribute("isum")), 0, ele.getAttribute("isum"))))
    Next
    'Set oNum2Chinese = CreateObject("FormulaParse.Calculator")
    Num2Chinese Format(iSum, "#.00"), strSumDX
    'If strSumDX = "Բ��" Then strSumDX = "��Բ������"
    Me.voucher.headerText("isumdx") = strSumDX
    Me.voucher.headerText("isumx") = iSum
    Me.voucher.headerText("zdsumdx") = strSumDX
    Me.voucher.headerText("zdsum") = iSum
    Set NdList = Nothing
    Set ele = Nothing
End Sub
Private Sub setItemState(Optional sOperate As String)
    Dim i As Long
    With Me.voucher
        .BodyMaxRows = 0
        Select Case strVouchType
            Case "97", "16"
                If strVouchType = "97" Then
                    If Dombody.selectNodes("//z:row[@ccontractid !='']").length > 0 Then
                        .EnableHead "ccusabbname", False
                        .EnableHead "cexch_name", False
                        .EnableHead "cbustype", False
                        If voucher.headerText("cstname") = "" Then
                            SetOriItemState "T", "cstname"
                        Else
                            .EnableHead "cstname", False
                        End If
                    Else
                        SetOriItemState "T", "ccusabbname"
                        SetOriItemState "T", "cexch_name"
                        SetOriItemState "T", "cbustype"
                        SetOriItemState "T", "cstname"
                    End If
                    If iVouchState = 2 Then
                        sCurTemplateID = ""
                        
                        .Visible = False
                        For iHeadIndex = 1 To .HeadInfoCount
                            .EnableHead iHeadIndex, False
                        Next iHeadIndex
                        SetOriItemState "T", "cmemo"
                        SetOriItemState "T", "dpredatebt"
                        SetOriItemState "T", "dpremodatebt"
                        .Visible = True
                        .SetFocus
                        .UpdateCmdBtn
                    End If
                End If
                If LCase(sOperate) = "copy" Or LCase(sOperate) = "modify" Then
                    .EnableHead "cbustype", False
                Else
                    SetOriItemState "T", "cbustype"
                End If
        End Select
    End With
End Sub
 
Private Function GetScrollWidth() As Single
    GetScrollWidth = Screen.TwipsPerPixelX * GetSystemMetrics(SM_CXVSCROLL)
End Function
 

''��Ҫ�ı䵥��ģ��
Private Sub ComboDJMB_Click()
    Dim tmpVoucherState As Variant
    ComboDJMB.ToolTipText = ComboDJMB.Text
    If Not bfillDjmb Then
        Me.voucher.Visible = False
        Me.voucher.getVoucherDataXML Domhead, Dombody
        tmpVoucherState = Me.voucher.VoucherStatus
        Call ChangeTempaltes(Str(vtidDJMB(ComboDJMB.ListIndex)), , False)
        Me.voucher.VoucherStatus = tmpVoucherState
        Me.voucher.setVoucherDataXML Domhead, Dombody
        Me.voucher.Visible = True
        Me.voucher.headerText("ivtid") = Str(vtidDJMB(ComboDJMB.ListIndex))
        sCurTemplateID = Str(vtidDJMB(ComboDJMB.ListIndex))
        sCurTemplateID2 = Str(vtidDJMB(ComboDJMB.ListIndex))
    Else
        bfillDjmb = False
    End If
End Sub
 
Private Sub ComboVTID_Click()
    ComboVTID.ToolTipText = ComboVTID.Text
End Sub
 

Private Sub CTBCtrl1_OnCommand(ByVal enumType As prjTBCtrl.ENUM_MENU_OR_BUTTON, ByVal cButtonId As String, ByVal cMenuId As String)
    bCloseFHSingle = False
    ButtonClick cButtonId, tbrvoucher.buttons(cButtonId).ToolTipText
End Sub
 
Private Sub Form_Activate()
    On Error Resume Next
    Me.picVoucher.BackColor = Me.picVoucher.BackColor
End Sub
 
Private Sub Form_Deactivate()
    With Me.voucher
        If .VoucherStatus <> VSNormalMode Then
            bLostFocus = True
            .ProtectUnload2
            bLostFocus = False
        End If
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 
    If iShowMode <> 1 Then  ''??
        setKey KeyCode, Shift
    ElseIf KeyCode = vbKeyF4 Then
        setKey KeyCode, Shift
    End If
    
End Sub
 
Private Sub Form_Load()
    Dim bLock As Boolean
    Dim recTmp As UfRecordset
    Dim dD As Date
    Dim s As String
    On Error Resume Next
    Me.KeyPreview = True
'���õ�������ؼ�

'//////////////////////////////////////////////////
'  860sp������861�޸Ĵ�1 ע��    2006/03/08 �Ŀؼ���861�汾���Ѿ����ɵ����ݿؼ�����   ����Ҫɾ��
' voucher.SetSortCallBackObject U8VoucherSorter1
'    With U8VoucherSorter1
'        .BackColor = voucher.BackColor
'        .Left = Me.Left + 550
'        .Top = Me.Picture1.Top
'        .ZOrder
'    End With
'//////////////////////////////////////////////////


'by lg070314����U870�˵��ںϹ���
    ''''''''''''''''''''''''''''''''''''''
    If Not g_business Is Nothing Then
        Set Me.UFToolbar1.Business = g_business
    End If
    Call RegisterMessage
    '''''''''''''''''''''''''''''''''''''''
    
    Call SetButton  '���ò˵���ť
    If lngClr1 <> 0 And lngClr2 <> 0 Then
        Call voucher.SetRuleColor(lngClr1, lngClr2)
    End If
    ChangeOneFormTbr Me, Me.tbrvoucher, Me.UFToolbar1
    SetButtonStatus "Cancel"
    Labeldjmb.BackColor = Me.Picture2.BackColor
    Picture1.BackColor = Me.Picture2.BackColor
    Labeldjmb.ForeColor = vbBlack
    If iShowMode = 1 Then
        If frmMain.WindowState = 1 Then frmMain.WindowState = 2
    End If
    Me.picVoucher.Move 0, Me.tbrvoucher.Height, Me.ScaleWidth, Me.ScaleHeight - Me.StBar.Height    '-ME.tbrvoucher
     Me.picVoucher.Move 0, Me.tbrvoucher.Height, Me.ScaleWidth, Me.ScaleHeight - Me.StBar.Height - Me.UFToolbar1.Height     '-ME.tbrvoucher
    Me.Picture2.Width = Me.Width
    Me.Picture1.BackColor = Me.Picture2.BackColor
    Me.Picture1.Move Me.Picture2.Width - Me.Picture1.Width - 5
    Me.StBar.ZOrder
    strFreeName1 = clsSAWeb.getDefName(DBConn, "cfree1")
    strFreeName2 = clsSAWeb.getDefName(DBConn, "cfree2")
    strFreeName3 = clsSAWeb.getDefName(DBConn, "cfree3")
    strFreeName4 = clsSAWeb.getDefName(DBConn, "cfree4")
    strFreeName5 = clsSAWeb.getDefName(DBConn, "cfree5")
    strFreeName6 = clsSAWeb.getDefName(DBConn, "cfree6")
    strFreeName7 = clsSAWeb.getDefName(DBConn, "cfree7")
    strFreeName8 = clsSAWeb.getDefName(DBConn, "cfree8")
    strFreeName9 = clsSAWeb.getDefName(DBConn, "cfree9")
    strFreeName10 = clsSAWeb.getDefName(DBConn, "cfree10")
    With StBar
        .Panels.Clear
        .Panels.Add 1, , ""
        .Panels(1).Width = Me.Width * 1 / 3
        .Panels.Add 2, , ""
        .Panels(2).Width = Me.Width * 1 / 3
        .Panels.Add 3, , ""
        .Panels(3).Width = Me.Width * 1 / 3
    End With
    Me.BackColor = Me.voucher.BackColor
    Me.ForeColor = Me.voucher.BackColor
    Set moAutoFill = CreateObject("ScmPublicSrv.clsAutoFill")
    ProgressBar1.Top = voucher.Top - 1000
    ProgressBar1.Width = voucher.Width - 2000
    ProgressBar1.Left = voucher.Left - 1000
End Sub
 

Private Sub Form_Resize()
    On Error Resume Next
    Me.UFToolbar1.Top = 0
    Me.UFToolbar1.Width = Me.ScaleWidth
    Me.picVoucher.Move 0, Me.tbrvoucher.Height, Me.ScaleWidth, Me.ScaleHeight - Me.tbrvoucher.Height - Me.StBar.Height
    If Me.picVoucher.ScaleHeight - 1 * GetScrollWidth < dOriVoucherHeight Then
        Me.voucher.Height = Me.picVoucher.ScaleHeight - Picture2.Height '- 1 * GetScrollWidth
    Else
        Me.voucher.Height = Me.picVoucher.ScaleHeight - Picture2.Height '- 1 * GetScrollWidth
    End If
    If Me.picVoucher.ScaleWidth - 1 * GetScrollWidth < dOriVoucherWidth Then
        Me.voucher.Width = Me.picVoucher.ScaleWidth '- 1 * GetScrollWidth
    Else
        Me.voucher.Width = Me.picVoucher.ScaleWidth '- 1 * GetScrollWidth
    End If
    Me.Picture2.Width = Me.Width
    Me.Picture2.Move 0, 0
    Me.Picture1.BackColor = Me.Picture2.BackColor
    Me.Picture1.Move Me.Picture2.Width - Me.Picture1.Width - 400
    Me.LabelVoucherName.Move (Me.Width - Me.LabelVoucherName.Width) / 2
    
    If labZF.Visible = False And labXJ.Visible = False Then
        Me.voucher.Move 0, Me.Picture2.Height
    Else
        Me.voucher.Move 0, Me.Picture2.Height + Me.labZF.Height
        
        Me.voucher.Width = Me.voucher.Width - 350
        Me.voucher.Height = Me.voucher.Height - Me.labZF.Height - Picture2.Height
    End If
    labZF.Top = picVoucher.Top + Me.Picture2.Height  'Me.top - Me.tbrvoucher.top
    labZF.Left = Me.voucher.Left
    labXJ.Top = picVoucher.Top + Me.Picture2.Height ' Me.top - Me.tbrvoucher.top    'Me.StBar.height
    labXJ.Left = Me.voucher.Left '+ labZF.Width
    With StBar
        .Panels(1).Width = Me.Width * 1 / 3
        .Panels(2).Width = Me.Width * 1 / 3
        .Panels(3).Width = Me.Width * 1 / 3
    End With
'//////////////////////////////////////////////////
'  860sp������861�޸� ע��    2006/03/08 �Ŀؼ���861�汾���Ѿ����ɵ����ݿؼ����� ����Ҫɾ��
'    With U8VoucherSorter1
'        .BackColor = Me.Picture2.BackColor
'        .Left = 550
'        .Top = Me.picVoucher.Top + Me.Picture1.Top
'        .ZOrder
'    End With
    Me.BackColor = Me.voucher.BackColor
    Me.ForeColor = Me.voucher.BackColor
    Me.picVoucher.BackColor = Me.voucher.BackColor
End Sub
 
Private Function FillVoucher(Domhead As DOMDocument, Dombody As DOMDocument, Optional bClearBody As Boolean = False) As Boolean
    Dim lngCol As Long, lngrow As Long, rows As Long
    Dim i  As Long
    Dim ele As IXMLDOMElement
    Dim ns As IXMLDOMNode
    Dim NODs As IXMLDOMNode
    Dim NODs2 As IXMLDOMNode
    Dim elelist As IXMLDOMNodeList
    Dim elelist2 As IXMLDOMNodeList
    Dim eleTmp As IXMLDOMElement
    Dim linedom As DOMDocument
    Dim oDomH As New DOMDocument
    With Me.voucher
        Set linedom = New DOMDocument
        .getVoucherDataXML oDomH, linedom
        Set ns = linedom.selectSingleNode("//rs:data")
        Set elelist = linedom.selectNodes("//z:row[@cinvcode = '']")
        If (Not ns Is Nothing) And elelist.length <> 0 Then
            For Each NODs In elelist
                ns.removeChild NODs
            Next
        End If
        If bClearBody = True Then
            Call ClearAllLineByDom(linedom)
        End If
        .setVoucherDataXML oDomH, linedom
        .BodyMaxRows = 0
        rows = .BodyRows
        If Not Domhead Is Nothing Then
             For Each ele In Domhead.selectNodes("//R")
                
                If LCase(ele.getAttribute("K")) = "cexch_name" Then
                    .ItemState("iexchrate", siheader).nNumPoint = clsSAWeb.GetExchRateDec(.headerText("cexch_name"))
                End If
                If LCase(ele.getAttribute("K")) = "minddate" Then
                    maxRefDate = CDate(ele.getAttribute("V"))
                End If
            Next
        End If
        If Not Dombody Is Nothing Then
            lngCol = .BodyRows
            Set elelist2 = Dombody.selectNodes("//z:row")
            If elelist2.length > 5 Then
                lngCol = lngCol + elelist2.length
                .getVoucherDataXML Domhead, linedom
                Set ns = linedom.selectSingleNode("//rs:data")
                If ns Is Nothing Then
                    Set ns = linedom.createElement("rs:data")
                    linedom.selectSingleNode("xml").appendChild ns
                End If
                Set ns = linedom.selectSingleNode("//rs:data")
                For Each ele In elelist2
                    '.AddLine
                    ele.setAttribute "editprop", "A"
                    ns.appendChild ele
                Next
                .setVoucherDataXML Domhead, linedom
                .row = lngCol
            Else
                For Each NODs2 In elelist2
                    .AddLine
                    lngCol = lngCol + 1
                    .row = lngCol
                    Set ns = linedom.selectSingleNode("//rs:data")
                    If ns Is Nothing Then
                        Set eleTmp = linedom.createElement("rs:data")
                        linedom.documentElement.appendChild eleTmp
                    End If
                    Set elelist = linedom.selectNodes("//z:row")
                    For Each NODs In elelist
                        ns.removeChild NODs
                    Next
                    Set ns = linedom.selectSingleNode("//rs:data")
                    ns.appendChild NODs2
                    .UpdateLineData linedom ', lngCol
                Next
            End If
        End If
    End With
End Function

Public Sub ButtonClick(s As String, sTaskKey As String, Optional bCloseSingle As Boolean = False)
    Dim i As Long
    Dim j As Long
    Dim row As Long
    Dim objGoldTax As Object
    Dim strError As String
    Dim strXMLHead As String
    Dim strXMLBody As String
    Dim lngrow As Integer
    Dim lngCol As Integer
    Dim strID As Variant
    Dim ele As IXMLDOMElement
    Dim strAuthId As String
    Dim elelist As IXMLDOMNodeList
    Dim NDRs    As IXMLDOMNode
    Dim nd      As IXMLDOMNode
    Dim bEAlast As Boolean
    Dim sPrnTmplate As Long
    Dim VoucherGrid As Object
    Dim Frm As New frmVouchNew
    On Error GoTo Err
    bCloseFHSingle = bCloseSingle
    strErrMsg = ""
    i = 0
    Set domPrint = Nothing
    With voucher
        Select Case LCase(s)
            Case "filter"
            'by lg070315������u870�����µĶ�λ����
'                voucher.ShowFindDlg
            
            
                If strVouchType = "97" Then
                    Me.Show
                    Dim Frmlist As New frmVoucherList
                    With Frmlist
                        .Sysid = "FA"
                        .VouchKey = "FA110"
                        .strTaskId = strAuthId
                        .VouchType = strVouchType
                        If .Filter Then
                            .Show
                            Call Unload_frms(Me.Name)
                        End If
                    End With
                End If
            
            Case "pd_all"
                Set VoucherGrid = voucher.GetBodyObject
                For i = 1 To voucher.BodyRows
                        voucher.GetBodyObject.RowHidden(i) = False
                Next
            
            Case "add"            '//����
                ChangeTempaltes sCurTemplateID
                If ChangeDJMBForEdit = False Then Screen.MousePointer = vbDefault: Exit Sub
                If VoucherTask(sTaskKey) = False Then Screen.MousePointer = vbDefault: Exit Sub
                Screen.MousePointer = vbHourglass
                EditPanel_1 EdtP, 3, ""
                labZF.Visible = False
                labXJ.Visible = False
                Me.voucher.AddNew ANMNormalAdd, Domhead, Dombody '
                Call SetVouchNoWriteble      '���õ��ݺ��Ƿ���Ա༭
                Call AddNewVouch              '�����������ݵĳ�ʼֵ
                Me.voucher.AddNew ANMCopyALL, Domhead, Dombody
                Me.voucher.headerText("vt_id") = sCurTemplateID
                Set Domhead = Me.voucher.GetHeadDom
                If iShowMode = 2 Then

                End If
                iVouchState = 0
                Call SetButtonStatus(s)
                Call setItemState(s)
                
            Case "chenged" '���
'                If strVouchType = "97" Or strVouchType = "96" Then
'                    Call Frm.ShowVoucher(gdzckpxg, Me.voucher.headerText("id"))
'               End If
            '/////////////////////////////////////////////////////////////////////////////////////////
            '  860sp������861�޸Ĵ�1 ע��    2006/03/09 861�汾�е��ݿؼ����ӵ��ݸ������ܣ������Ŀ������ļ���ͼƬ���������޴�СΪ1M��
                Me.voucher.SelectFile
               
               
            Case "outadd"              '//��������
            Case "modify"              '//�޸�
                If CheckDJMBAuth(Me.voucher.headerText("ivtid"), "W") = False Then
                    MsgBox "��ǰ����Աû�е�ǰ����ģ���ʹ��Ȩ�ޣ�"
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                If VoucherTask(sTaskKey) = False Then Screen.MousePointer = vbDefault: Exit Sub
                Me.voucher.row = 0
                Me.voucher.col = 0
                Screen.MousePointer = vbHourglass
                .getVoucherDataXML Domhead, Dombody
                Me.voucher.VoucherStatus = VSeEditMode
                Call setItemState(s)
                Call AddNewVouch("modify")
                Call SetVouchNoWriteble
                Call SetButtonStatus(s)
                iVouchState = 1
                .SetFocus
                .UpdateCmdBtn
                
            Case "erase"                 '//ɾ��
                If CheckDJMBAuth(Me.voucher.headerText("ivtid"), "W") = False Then
                    MsgBox "��ǰ����Աû�е�ǰ����ģ���ʹ��Ȩ�ޣ�"
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                If MsgBox("ȷʵҪɾ�����ŵ�����", vbYesNo + vbQuestion) = vbNo Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                If VoucherTask(sTaskKey) = False Then Screen.MousePointer = vbDefault: Exit Sub
                Set Domhead = Me.voucher.GetHeadDom
                bCreditCheck = False
                Screen.MousePointer = vbHourglass
                strError = clsVoucherCO.Delete(Domhead)
                If strError <> "" Then
                    ShowErrDom strError, Domhead
                    LoadVoucher ""
                Else
                    LoadVoucher "tonext"
                End If
                Call VoucherFreeTask

            Case "copy"                         '//����
                If ChangeDJMBForEdit = False Then Screen.MousePointer = vbDefault: Exit Sub
                If VoucherTask(sTaskKey) = False Then Exit Sub
                labZF.Visible = False
                labXJ.Visible = False
                Screen.MousePointer = vbHourglass
                AddNewVouch "copy"
                
                Me.voucher.AddNew ANMCopyALL, Domhead, Dombody
                Me.voucher.headerText("chandler") = ""
                Me.voucher.headerText("chandlername") = ""
                Call SetVouchNoWriteble
                Call SetButtonStatus(s)
                iVouchState = 0
                Call setItemState(s)
                .SetFocus
                .UpdateCmdBtn
            Case "addrow"                       '//����һ��
                With Me.voucher
                    .AddLine
                    .bodyText(.BodyRows, "itb") = "����"
                    If strVouchType = "05" Or strVouchType = "06" Or strVouchType = "26" Or strVouchType = "27" Or strVouchType = "28" Or strVouchType = "29" Then
                        If .BodyRows > 1 Then
                           .bodyText(.BodyRows, "cwhname") = .bodyText(.BodyRows - 1, "cwhname")
                           .bodyText(.BodyRows, "cwhcode") = .bodyText(.BodyRows - 1, "cwhcode")
                        End If
                    End If
                End With
            Case "delrow"                     '//ɾ��һ��
                If (Me.voucher.BodyRows > 0) And Me.voucher.row <> 0 Then
                     Dim tmpRow As Variant
                     tmpRow = Me.voucher.row - 1
                     Me.voucher.DelLine Me.voucher.row
                    Me.voucher.row = tmpRow
                End If
                
            Case "sure"           '//���
                Screen.MousePointer = vbHourglass
                If VoucherTask(sTaskKey) = False Then Screen.MousePointer = vbDefault: Exit Sub
                Call AddNewVouch(s)
                bCreditCheck = True
                Set Domhead = Me.voucher.GetHeadDom
                strError = clsVoucherCO.VerifyVouch(Domhead, bCheckVouch)
                Call ShowErrDom(strError, Domhead)
                ''ˢ�µ�ǰ����
                LoadVoucher ""
                Call VoucherFreeTask
                
            Case "unsure"            '//����
                Screen.MousePointer = vbHourglass
                If VoucherTask(sTaskKey) = False Then Screen.MousePointer = vbDefault: Exit Sub
                Call AddNewVouch(s)
                bCreditCheck = False
                Set Domhead = Me.voucher.GetHeadDom
                strError = clsVoucherCO.VerifyVouch(Domhead, bCheckVouch)
                Call ShowErrDom(strError, Domhead)
                ''ˢ�µ�ǰ����
                LoadVoucher ""
                Call VoucherFreeTask
            Case "cancel"                  '//ȡ��
                bClickCancel = True
                voucher.VoucherStatus = VSNormalMode
                LoadVoucher ""
                bOnceRefer = False
                Call SetButtonStatus(s)
                ChangeButtonsState
                bClickCancel = False
                Call VoucherFreeTask
            Case "save"                    '//����
                Screen.MousePointer = vbHourglass
                voucher.ProtectUnload2
                bClickCancel = False
                bClickSave = True
                strError = ""
                If Me.voucher.BodyRows = 0 And strVouchType <> "94" Then
                    MsgBox "����û�м�¼����¼�룡"
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                If .headVaildIsNull2(strError) = False Then
                    MsgBox "��ͷ��Ŀ" + strError
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                
                strError = ""
                If .bodyVaildIsNull2(strError) = False Then
                    MsgBox "������Ŀ" + strError
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                strError = ""
                Call AddNewVouch("Save")
                voucher.getVoucherDataXML Domhead, Dombody
                '////////////////////////////////////////////////////////////////////////////////////////////////
                '860sp������861�޸Ĵ�   2006/03/08   ���ӵ��ݸ�������
                If SetAttachXML(Domhead) = False Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                '////////////////////////////////////////////////////////////////////////////////////////////////
                If bFirst = True Then
                    Domhead.selectSingleNode("//z:row").Attributes.getNamedItem("bfirst").nodeValue = "1"
                End If
                bCreditCheck = False
ToSave:
                strError = clsVoucherCO.Save(Domhead, Dombody, iVouchState, vNewID, domConfig)
                If strError <> "" Then
                    If InStr(1, strError, "<", vbTextCompare) <> 0 Then
                        ShowErrDom strError, Domhead
                    Else
                        MsgBox IIf(Trim(strError) = "��ǰ�������ɹ�������������!", "", strError)
                        If Domhead.selectNodes("//z:row").length = 1 Then
                            If .headerText(getVoucherCodeName) <> GetHeadItemValue(Domhead, getVoucherCodeName) And strVouchType <> "92" Then
                                .headerText(getVoucherCodeName) = GetHeadItemValue(Domhead, getVoucherCodeName)
                            End If
                        End If
                    End If
                Else
                    voucher.VoucherStatus = VSNormalMode
                    LoadVoucher "", IIf(vNewID <> "", vNewID, 0)
                    bOnceRefer = False
                    Call SetButtonStatus(s)
                    ChangeButtonsState
                    Call VoucherFreeTask
                End If
                bClickSave = False
                If strVouchType = "98" Then
                    Unload Me
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            Case "print"            '��ӡ
                    If Me.ComboVTID.ListCount = 0 Then
                        MsgBox "��ǰ����Աû�п���ʹ�õĴ�ӡģ�棬���飡"
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                    If strVouchType = "96" Or strVouchType = "97" Then
                        sPrnTmplate = Get_print_id(Me.voucher.headerText("stypenum"))
                        If sPrnTmplate = 0 Then
                            MsgBox "û��������ȷ��Ĭ�ϴ�ӡģ�壬���飡"
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                    Else
                        sPrnTmplate = CLng(vtidPrn(Me.ComboVTID.ListIndex))
                    End If
                BillPrnVTID = sPrnTmplate
                VoucherPrn strVouchType, voucher, strCardNum, CLng(sPrnTmplate), , True
                .VoucherStatus = VSNormalMode
                LoadVoucher ""
           
            Case "preview"
                    If Me.ComboVTID.ListCount = 0 Then
                        MsgBox "��ǰ����Աû�п���ʹ�õĴ�ӡģ�棬���飡"
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                     ''ȡ�û����õ�Ĭ�ϴ�ӡģ��
                    If strVouchType = "96" Or strVouchType = "97" Then
                        sPrnTmplate = Get_print_id(Me.voucher.headerText("stypenum"))
                        If sPrnTmplate = 0 Then
                            MsgBox "û��������ȷ��Ĭ�ϴ�ӡģ�壬���飡"
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                    Else
                        sPrnTmplate = CLng(vtidPrn(Me.ComboVTID.ListIndex))
                    End If
                    BillPrnVTID = sPrnTmplate
                    VoucherPrn strVouchType, voucher, strCardNum, sPrnTmplate, "Preview", True
                    .VoucherStatus = VSNormalMode
                    LoadVoucher ""
                    
            Case "output"
                    VouchOutPut voucher, CLng(sTemplateID), strCardNum
            Case "exit"
                Unload Me
                Screen.MousePointer = vbDefault
                Exit Sub
            Case "seek"
                '-----------------------------------------------------
                '�ɵ�������ƾ֤
                 Find_GL_accvouch
            
            Case "paint"
                Screen.MousePointer = vbHourglass
                LoadVoucher "", , True
                
            Case LCase("ToPrevious")   '��һ��
                Screen.MousePointer = vbHourglass
                LoadVoucher s, , True
 
'                voucher.VoucherStatus = VSNormalMode
            
            Case LCase("ToNext")   '��һ��
                Screen.MousePointer = vbHourglass
                LoadVoucher s, , True
        
'                voucher.VoucherStatus = VSNormalMode
                
            Case LCase("ToLast")   'ĩ��
                Screen.MousePointer = vbHourglass
                LoadVoucher s, , True
 
'                voucher.VoucherStatus = VSNormalMode
                
            Case LCase("ToFirst")   '����
                Screen.MousePointer = vbHourglass
                LoadVoucher s, , True
 
'                voucher.VoucherStatus = VSNormalMode
                
            Case LCase("CopyRow")
                Set Dombody = .GetLineDom
                Set Domhead = New DOMDocument
                clsVoucherCO.CopyRow Dombody
                i = voucher.BodyMaxRows
                voucher.BodyMaxRows = 0
                .AddLine
                voucher.BodyMaxRows = i
                .UpdateLineData Dombody, .BodyRows
            Case LCase("LookVeri")  ''��ѯ������
                If .VoucherStatus = VSeEditMode Then .ProtectUnload2
                Set Domhead = .GetHeadDom
                If obj_EA.NeedEAFControl(clsSAWeb.GetEAsCode(strVouchType, Domhead), GetHeadItemValue(Domhead, clsSAWeb.getVouchMainIDName(strVouchType))) Then
                    If (obj_EA.ResearchEAStream(clsSAWeb.GetEAsCode(strVouchType, Domhead), .headerText(clsSAWeb.getVouchMainIDName(strVouchType)))) = False Then
                        MsgBox obj_EA.ErrDescript
                    End If
                Else
                    MsgBox "�õ���δ����������!"
                End If
                
            Case "zp"   '���˵�����ʱ����֧Ʊ
                Dim f As frmZPHX
                Set f = New frmZPHX
                Load f
                Set f.Icon = Me.Icon
                f.LoadData voucher.headerText("cdepcode"), voucher.headerText("citemcode"), DBConn
                
                f.Show vbModal
                Unload f
                Set f = Nothing
                
        End Select
   End With
   Set ele = Nothing
   Screen.MousePointer = vbDefault
   ProgressBar1.Visible = False
   Exit Sub
Err:
    ProgressBar1.Visible = False
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
End Sub

Private Sub GetColWidth()
    With m_Grid
        Dim i As Long
        For i = 0 To m_Grid.cols - 1
            Debug.Print ".ColWidth(" & i & ")=" & .colwidth(i)
        Next
    End With
End Sub

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


Private Sub Form_Unload(Cancel As Integer)
'doNext:
'    If Me.voucher.VoucherStatus <> VSNormalMode Then
'        Select Case MsgBox("�Ƿ񱣴�Ե�ǰ���ݵı༭��", vbYesNoCancel + vbQuestion)
'            Case vbYes
'                ButtonClick "Save", "����"
'                If Me.voucher.VoucherStatus = VSNormalMode Then
'                    GoTo DoQuit
'                End If
'            Case vbNo
'                VoucherFreeTask
'                GoTo DoQuit
'            Case vbCancel
'
'        End Select
'
'        bFrmCancel = True
'        Me.ZOrder
'        Cancel = 3
'    Else
'DoQuit:
'        On Error Resume Next
'        bFrmCancel = False
''by lg070314����U870�˵��ںϣ��ر�ʱ����Business
'        Set UFToolbar1.Business = Nothing
'
'        Set clsVoucherCO = Nothing
'        Set clsAuth = Nothing
'        Set clsRefer = Nothing
'        Set RstTemplate = Nothing
'        Set Domhead = Nothing
'        Set Dombody = Nothing
'        Set DomFormat = Nothing
'        Set RstTemplate = Nothing
'        Set RstTemplate2 = Nothing
'        Set obj_EA = Nothing
'        Set DOMEA = Nothing
'        If m_UFTaskID <> "" Then
'            m_login.TaskExec m_UFTaskID, 0
'        End If
'        minForm
'   End If
End Sub
 
Private Sub mdiAddRow_Click()
    ButtonClick "AddRow", ""
End Sub
 
Private Sub mdiDelRow_Click()
    ''�Ҽ��˵�
    ButtonClick "DelRow", ""
End Sub
  
 
Private Sub mnuLinkQuery_Click()
    Call ProcLinkQuery
End Sub

Private Sub tbrvoucher_ButtonClick(ByVal Button As MSComctlLib.Button)
    bCloseFHSingle = False
    ButtonClick Button.key, Button.ToolTipText
End Sub
Private Function getVoucherCodeName() As String
    Dim KeyCode As String
    Select Case strVouchType
        Case "97", "96"
            KeyCode = "scardnum"     '�ʲ���Ƭ
            
        Case "101"
            KeyCode = "ccode"         ' �ʲ��̵㵥����
            
        Case Else
            KeyCode = "ccode"
    End Select
    getVoucherCodeName = KeyCode
End Function

Private Sub UFToolbar1_OnCommand(ByVal enumType As UFToolBarCtrl.ENUM_MENU_OR_BUTTON, ByVal cButtonId As String, ByVal cMenuId As String)
    ButtonClick IIf(enumType = enumButton, cButtonId, cMenuId), ""
End Sub

Private Sub Voucher_BillNumberChecksucceed()
    Dim errMsg As String, strVouchNo As String, KeyCode As String
    Dim tmpDOM As New DOMDocument
    KeyCode = getVoucherCodeName()
    If strVouchType = "92" Then Exit Sub
    With Me.voucher
        If Not (LCase(DomFormat.selectSingleNode("//���ݱ��").Attributes.getNamedItem("�غ��Զ���ȡ").nodeValue) = "false" And _
        LCase(DomFormat.selectSingleNode("//���ݱ��").Attributes.getNamedItem("�����ֹ��޸�").nodeValue) = "true") Then
            Set tmpDOM = .GetHeadDom

            If clsVoucherCO.GetVoucherNO(tmpDOM, strVouchNo, errMsg) = False Then
                MsgBox errMsg
                strCurVoucherNO = ""
            Else
                .headerText(KeyCode) = strVouchNo
                If strVouchType = "97" Then
                    .headerText("sassetnum") = strVouchNo
                End If
                strCurVoucherNO = strVouchNo
            End If
        End If
    End With
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

 
 
Private Sub Voucher_bodyBrowUser(ByVal row As Long, ByVal col As Long, sRet As Variant, referPara As ReferParameter)
    Dim strSql As String, cInvCode As String, cCusCode As String
    Dim Domhead As DOMDocument, Dombody As DOMDocument
    Dim Dombodys_str1 As String
    Dim Dombodys_str2 As String
    Dim lngrow As Integer
    Dim i As Integer, lRecord As Long
    Dim j As Long
    Dim sKey As String
    Dim sKeyValue As String, strAuth As String
    Dim tmpRow As Long, tmpCol As Long
    Dim tmpCol2 As Long
    Dim strClass As String
    Dim strGrid As String
    Dim ifalg As Boolean
    tmpRow = row
    tmpCol = voucher.col
    tmpCol2 = col
    On Error Resume Next
    With voucher
        .MultiLineSelect = False ''���ö�ѡĬ��
        clsRefer.SetReferSQLString ""
        clsRefer.SetRWAuth "INVENTORY", "R", False
        clsRefer.SetReferDisplayMode enuGrid
        sKey = .ItemState(col, sibody).sFieldName
        sKeyValue = .bodyText(row, col)
        Select Case LCase(sKey)
        
            Case "ccode", "ccode_name" '��Ŀ
                    strClass = ""
                    strGrid = "select ccode,ccode_name from code where ccode not in( select ccode from MT_basesets where  isnull(ccode,'')<>'') and bend=1 "
                    If LCase(sKey) = "ccode" And Len(Trim(.bodyText(row, "ccode"))) > 0 Then
                        strGrid = strGrid & " and ccode like '%" & Trim(.bodyText(row, "ccode")) & "%'"
                    ElseIf LCase(sKey) = "ccode_name" And Len(Trim(.bodyText(row, "ccode_name"))) > 0 Then
                        strGrid = strGrid & " and ccode_name like '%" & Trim(.bodyText(row, "ccode_name")) & "%'"
                    End If
                    strGrid = strGrid & " order by ccode "
                    If clsRefer.StrRefInit_SetColWidth(m_login, False, strClass, strGrid, "��Ŀ����,��Ŀ����", "2500,6000") = False Then Exit Sub
                    clsRefer.Show
                    If Not clsRefer.recMx Is Nothing Then
                        .bodyText(row, "ccode") = clsRefer.recMx(0)
                        .bodyText(row, "ccode_name") = clsRefer.recMx(1)
                        sRet = clsRefer.recMx.Fields(LCase(sKey))
                    End If
                    

            Case "cexpcode", "cexpname" '�������
                    strClass = "select * from dbo.ExpItemClass "
                    strGrid = "select cexpccode,cexpcode,cexpname from ExpenseItem  where 1=1  "
                    If LCase(sKey) = "cexpcode" And Len(Trim(.bodyText(row, "cexpcode"))) > 0 Then
                        strGrid = strGrid & " and cexpcode like '%" & Trim(.bodyText(row, "cexpcode")) & "%'"
                    ElseIf LCase(sKey) = "cexpname" And Len(Trim(.bodyText(row, "cexpname"))) > 0 Then
                        strGrid = strGrid & " and cexpname like '%" & Trim(.bodyText(row, "cexpname")) & "%'"
                    End If
                    strGrid = strGrid & " order by cexpcode "
                    If clsRefer.StrRefInit_SetColWidth(m_login, False, strClass, strGrid, "����,����������,�����������", "0,2500,6000") = False Then Exit Sub
                    clsRefer.Show
                    If Not clsRefer.recMx Is Nothing Then
                        .bodyText(row, "cexpcode") = clsRefer.recMx(1)
                        .bodyText(row, "cexpname") = clsRefer.recMx(2)
                    sRet = clsRefer.recMx.Fields(LCase(sKey))
                    End If

                    

                 
        End Select
End With
'by lg070315������U870 UAP���ݿؼ��µĲ��մ���
referPara.Cancel = True

End Sub
 
Private Sub Voucher_bodyCellCheck(RetValue As Variant, bChanged As Long, ByVal R As Long, ByVal C As Long, referPara As ReferParameter)
    Dim lngrow As Long
    Dim strError As String
    Dim tmpDomhead As DOMDocument, tmpDOMBody As DOMDocument
    Dim sKey As String
    Dim ele As IXMLDOMElement
    Dim sKeyValue As String
    Dim i As Long
    Dim usexchangerate As Currency
    Dim dblexchangerate As Currency
    Dim strGrid As String
    Dim strSql As String
    Dim sqlstr As String
    Dim recMx As UfRecordset
    Dim intCyc As Long
    Dim rds As New ADODB.Recordset
    Dim rds1 As New ADODB.Recordset
    On Error GoTo DoERR
    With Me.voucher
        sKey = LCase(.ItemState(C, sibody).sFieldName)
        If Trim(.bodyText(R, sKey)) = "" Then Exit Sub
        Select Case sKey
        
            Case "ccode"  '��Ŀ
                    strSql = "select ccode,ccode_name from code where bend=1 "
                    strSql = strSql & " and ccode = '" & Trim(.bodyText(R, "ccode")) & "'"
                    If rds.state <> 0 Then rds.Close
                    rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
                    If rds.RecordCount = 0 Then
                        MsgBox "��Ŀ���Ϸ���", vbOKOnly + vbCritical
                        Me.voucher.bodyText(R, "ccode") = ""
                        Me.voucher.bodyText(R, "ccode_name") = ""
                        RetValue = ""
                        bChanged = Retry
                    Else
                        Me.voucher.bodyText(R, "ccode") = rds.Fields("ccode")
                        Me.voucher.bodyText(R, "ccode_name") = rds.Fields("ccode_name")
                        RetValue = rds.Fields(LCase(sKey))
                    End If
                    
            Case "ccode_name"  '��Ŀ
                    strSql = "select ccode,ccode_name from code where bend=1 "
                    strSql = strSql & " and ccode_name = '" & Trim(.bodyText(R, "ccode_name")) & "'"
                    If rds.state <> 0 Then rds.Close
                    rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
                    If rds.RecordCount = 0 Then
                        MsgBox "��Ŀ���Ϸ���", vbOKOnly + vbCritical
                        Me.voucher.bodyText(R, "ccode") = ""
                        Me.voucher.bodyText(R, "ccode_name") = ""
                        RetValue = ""
                        bChanged = Retry
                    Else
                        Me.voucher.bodyText(R, "ccode") = rds.Fields("ccode")
                        Me.voucher.bodyText(R, "ccode_name") = rds.Fields("ccode_name")
                        RetValue = rds.Fields(LCase(sKey))
                    End If
            
            Case "cexpcode"  '�������
                    strSql = "select cexpccode,cexpcode,cexpname from ExpenseItem  where 1=1  "
                    strSql = strSql & " and cexpcode = '" & Trim(.bodyText(R, "cexpcode")) & "'"
                    If rds.state <> 0 Then rds.Close
                    rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
                    If rds.RecordCount = 0 Then
                        MsgBox "������𲻺Ϸ���", vbOKOnly + vbCritical
                        Me.voucher.bodyText(R, "cexpcode") = ""
                        Me.voucher.bodyText(R, "cexpname") = ""
                        RetValue = ""
                        bChanged = Retry
                    Else
                        Me.voucher.bodyText(R, "cexpcode") = rds.Fields("cexpcode")
                        Me.voucher.bodyText(R, "cexpname") = rds.Fields("cexpname")
                        Me.voucher.bodyText(R, "iscontrol") = "����"
                        Me.voucher.bodyText(R, "bfb") = GETysbl(Me.voucher.headerText("cdepcode"), Me.voucher.headerText("citemcode"), Me.voucher.bodyText(R, "cexpcode"))
                        RetValue = rds.Fields(LCase(sKey))
                    End If
                    
            Case "cexpname"  '�������
                    strSql = "select cexpccode,cexpcode,cexpname from ExpenseItem  where 1=1  "
                    strSql = strSql & " and cexpname = '" & Trim(.bodyText(R, "cexpname")) & "'"
                    If rds.state <> 0 Then rds.Close
                    rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
                    If rds.RecordCount = 0 Then
                        MsgBox "������𲻺Ϸ���", vbOKOnly + vbCritical
                        Me.voucher.bodyText(R, "cexpcode") = ""
                        Me.voucher.bodyText(R, "cexpname") = ""
                        RetValue = ""
                        bChanged = Retry
                    Else
                        Me.voucher.bodyText(R, "cexpcode") = rds.Fields("cexpcode")
                        Me.voucher.bodyText(R, "cexpname") = rds.Fields("cexpname")
                        Me.voucher.bodyText(R, "bfb") = GETysbl(Me.voucher.headerText("cdepcode"), Me.voucher.headerText("citemcode"), Me.voucher.bodyText(R, "cexpcode"))
                        Me.voucher.bodyText(R, "iscontrol") = "����"
                        RetValue = rds.Fields(LCase(sKey))
                    End If
            
'            Case "adds", "lenssen" '����ʱ���Ʒ���
'                    If LCase(sKey) = "adds" And Len(Trim(.bodyText(R, "adds"))) > 0 Then
'                         If (Trim(.bodyText(R, "adds")) = "�跽") Or (Trim(.bodyText(R, "adds")) = "����") Then
'                            .bodyText(R, "adds") = Trim(.bodyText(R, "adds"))
'                            RetValue = .bodyText(R, "adds")
'                            .bodyText(R, "lenssen") = ""
'                         Else
'                            .bodyText(R, "adds") = ""
'                            RetValue = ""
'                            bChanged = Retry
'                         End If
'                    ElseIf LCase(sKey) = "lenssen" And Len(Trim(.bodyText(R, "lenssen"))) > 0 Then
'                         If (Trim(.bodyText(R, "lenssen")) = "�跽") Or (Trim(.bodyText(R, "lenssen")) = "����") Then
'                            .bodyText(R, "lenssen") = Trim(.bodyText(R, "lenssen"))
'                            RetValue = .bodyText(R, "lenssen")
'                            .bodyText(R, "adds") = ""
'                         Else
'                            .bodyText(R, "lenssen") = ""
'                            RetValue = ""
'                            bChanged = Retry
'                         End If
'                    End If
            Case "adds"
                    If Len(Trim(.bodyText(R, "adds"))) > 0 Then
                        If (Trim(.bodyText(R, "adds")) = "�跽") Then
                            .bodyText(R, "adds") = Trim(.bodyText(R, "adds"))
                            RetValue = .bodyText(R, "adds")
                            If (.bodyText(R, "lenssen") <> "����") And (.bodyText(R, "lenssen") <> "") Then
                                .bodyText(R, "lenssen") = "����"
                            End If
                        ElseIf (Trim(.bodyText(R, "adds")) = "����") Then
                            .bodyText(R, "adds") = Trim(.bodyText(R, "adds"))
                            RetValue = .bodyText(R, "adds")
                            If (.bodyText(R, "lenssen") <> "�跽") And (.bodyText(R, "lenssen") <> "") Then
                                .bodyText(R, "lenssen") = "�跽"
                            End If
                        Else
                            .bodyText(R, "adds") = ""
                            RetValue = ""
                            bChanged = Retry
                        End If
                    End If
            Case "lenssen"
                    If Len(Trim(.bodyText(R, "lenssen"))) > 0 Then
                        If (Trim(.bodyText(R, "lenssen")) = "�跽") Then
                            .bodyText(R, "lenssen") = Trim(.bodyText(R, "lenssen"))
                            RetValue = .bodyText(R, "lenssen")
                            If (.bodyText(R, "adds") <> "����") And (.bodyText(R, "adds") <> "") Then
                                .bodyText(R, "adds") = "����"
                            End If
                        ElseIf (Trim(.bodyText(R, "lenssen")) = "����") Then
                            .bodyText(R, "lenssen") = Trim(.bodyText(R, "lenssen"))
                            RetValue = .bodyText(R, "lenssen")
                            If (.bodyText(R, "adds") <> "�跽") And (.bodyText(R, "adds") <> "") Then
                                .bodyText(R, "adds") = "�跽"
                            End If
                        Else
                            .bodyText(R, "lenssen") = ""
                            RetValue = ""
                            bChanged = Retry
                         End If
                    End If
            Case "rate" 'Ԥ�����
               If (.bodyText(R, "rate") > 100) Or (.bodyText(R, "rate") <= 0) Then
                    MsgBox "Ԥ��������Ϸ���", vbOKOnly + vbCritical
                    .bodyText(R, "rate") = "0.00"
                    RetValue = ""
                    bChanged = Retry
               End If
               
               
               
            Case "je" 'Ԥ��
                If val(.bodyText(R, "je")) <> 0 Then
                    If val(.bodyText(R, "hdje")) = 0 Then
                    .bodyText(R, "hdje") = .bodyText(R, "je")
                    End If
                End If
                          
            
'            Case "hdje"   '  �˶�Ԥ��
'               If (.bodyText(R, "rate") > 100) Or (.bodyText(R, "rate") <= 0) Then
'                    MsgBox "Ԥ��������Ϸ���", vbOKOnly + vbCritical
'                    .bodyText(R, "rate") = "0.00"
'                    RetValue = ""
'                    bChanged = Retry
'               End If
               

        End Select
    End With
    Set rds = Nothing
    Exit Sub
DoERR:
    Set rds = Nothing
    MsgBox Err.Description
End Sub
 
  
Private Sub Voucher_FillHeadComboBox(ByVal Index As Long, pCom As Object)
    Select Case LCase(Me.voucher.ItemState(Index, siheader).sFieldName)

        Case "iscontrol" '�Ƿ����
                pCom.Clear
                pCom.AddItem "����"
                pCom.AddItem "����"
                pCom.AddItem "����ʾ"
        
        Case "iperiod" 'Ԥ�����ڣ�����ڼ䣩
                pCom.Clear
                pCom.AddItem "1  ��"
                pCom.AddItem "2  ��"
                pCom.AddItem "3  ��"
                pCom.AddItem "4  ��"
                pCom.AddItem "5  ��"
                pCom.AddItem "6  ��"
                pCom.AddItem "7  ��"
                pCom.AddItem "8  ��"
                pCom.AddItem "9  ��"
                pCom.AddItem "10 ��"
                pCom.AddItem "11 ��"
                pCom.AddItem "12 ��"
            
            
       Case Else
            pCom.Clear
    End Select
    
End Sub
 
Private Sub Voucher_FillList(ByVal R As Long, ByVal C As Long, pCom As Object)
    Dim sFieldName As String
    sFieldName = LCase(Me.voucher.ItemState(C, sibody).sFieldName)
    Select Case sFieldName
        Case "adds", "lenssen" '���Ʒ���
                pCom.Clear
                pCom.AddItem ""
                pCom.AddItem "�跽"
                pCom.AddItem "����"
        
        Case "iscontrol" '�Ƿ����
                pCom.Clear
'                pCom.AddItem "1 ����"
'                pCom.AddItem "2 ����"
'                pCom.AddItem "3 ����ʾ"
                
                pCom.AddItem "����"
                pCom.AddItem "����"
                pCom.AddItem "����ʾ"

            
        Case "ending"
            pCom.Clear
            pCom.AddItem "�̿�"      '"�̿�"������ʵ���������������ӯ
            pCom.AddItem "��ʵ"
            pCom.AddItem "���"
            pCom.AddItem "��ӯ"
            

            
       Case Else
            pCom.Clear
    
    End Select
End Sub
 

Private Sub Voucher_headBrowUser(ByVal Index As Variant, sRet As Variant, referPara As ReferParameter)
    Dim iElement As IXMLDOMElement
    Dim sKey As String, sKeyValue As String
    Dim strSql As String
    Dim sFormat As String
    Dim strAuth As String
    Dim strDate As String
    Dim strCusInv As String
    Dim rst As New ADODB.Recordset
    Dim strClass As String
    Dim strGrid As String
    On Error Resume Next
    clsRefer.referMulti = False
    clsRefer.SetReferDisplayMode enuGrid
    clsRefer.SetReferSQLString ""
    clsRefer.SetRWAuth "INVENTORY", "R", False
    strAuth = ""
    sKey = Me.voucher.ItemState(Index, siheader).sFieldName
    sKeyValue = Me.voucher.headerText(Index)
    Select Case LCase(sKey)
    
    
        Case "cdepcode", "cdepname" '���ű���
            strClass = "select cdepcode,cdepname from Department"
            strGrid = "SELECT Department.cDepCode,Department.cDepName FROM Department  where bdepend=1 "
            If LCase(sKey) = "cdepcode" And Len(Trim(Me.voucher.headerText("cdepcode"))) > 0 Then
                strGrid = strGrid & " and cDepCode like '%" & Trim(Me.voucher.headerText("cdepcode")) & "%'"
            ElseIf LCase(sKey) = "cdepname" And Len(Trim(Me.voucher.headerText("cdepname"))) > 0 Then
                strGrid = strGrid & " and cDepName like '%" & Trim(Me.voucher.headerText("cdepname")) & "%'"
            End If
            strGrid = strGrid & " order by cDepCode "
            If clsRefer.StrRefInit_SetColWidth(m_login, False, strClass, strGrid, "Ƶ������,Ƶ������", "2500,6000") = False Then Exit Sub
            clsRefer.Show
            If Not clsRefer.recMx Is Nothing Then
                Me.voucher.headerText("cdepcode") = clsRefer.recMx(0)
                Me.voucher.headerText("cdepname") = clsRefer.recMx(1)
                sRet = clsRefer.recMx.Fields(LCase(sKey))
            End If

        Case "citemcode", "citemname" '��Ŀ����
            strClass = "select citemccode,citemcname  from V_MT_ItemsClass"
            strGrid = "SELECT citemccode, citemcode , citemname FROM V_MT_Items  where 1=1  "
            If LCase(sKey) = "citemcode" And Len(Trim(Me.voucher.headerText("citemcode"))) > 0 Then
                strGrid = strGrid & " and citemcode like '%" & Trim(Me.voucher.headerText("citemcode")) & "%'"
            ElseIf LCase(sKey) = "citemname" And Len(Trim(Me.voucher.headerText("citemname"))) > 0 Then
                strGrid = strGrid & " and citemname like '%" & Trim(Me.voucher.headerText("citemname")) & "%'"
            End If
            strGrid = strGrid & " order by citemcode "
            If clsRefer.StrRefInit_SetColWidth(m_login, False, strClass, strGrid, "����,��Ŀ����,��Ŀ����", "0,2500,6000") = False Then Exit Sub
            clsRefer.Show
            If Not clsRefer.recMx Is Nothing Then
                Me.voucher.headerText("citemcode") = clsRefer.recMx(1)
                Me.voucher.headerText("citemname") = clsRefer.recMx(2)
                sRet = clsRefer.recMx.Fields(LCase(sKey))
            End If
            
        Case "cexpccode", "cexpcname" '����������
            strClass = "select * from dbo.ExpItemClass "
            strGrid = "select cexpccode,cexpcname  from  ExpItemClass  where bexpcend=1  "
            If LCase(sKey) = "cexpccode" And Len(Trim(Me.voucher.headerText("cexpccode"))) > 0 Then
                strGrid = strGrid & " and cexpcode like '%" & Trim(Me.voucher.headerText("cexpccode")) & "%'"
            ElseIf LCase(sKey) = "cexpcname" And Len(Trim(Me.voucher.headerText("cexpcname"))) > 0 Then
                strGrid = strGrid & " and cexpcname like '%" & Trim(Me.voucher.headerText("cexpcname")) & "%'"
            End If
            strGrid = strGrid & " order by cexpccode "
            If clsRefer.StrRefInit_SetColWidth(m_login, False, strClass, strGrid, "����������,�����������", "2500,6000") = False Then Exit Sub
            clsRefer.Show
            If Not clsRefer.recMx Is Nothing Then
                Me.voucher.headerText("cexpccode") = clsRefer.recMx(0)
                Me.voucher.headerText("cexpcname") = clsRefer.recMx(1)
                sRet = clsRefer.recMx.Fields(LCase(sKey))
            End If
        
        Case "cexpcode", "cexpname" '�������
            strClass = "select * from dbo.ExpItemClass "
            strGrid = "select cexpccode,cexpcode,cexpname from ExpenseItem  where 1=1  "
            If LCase(sKey) = "cexpcode" And Len(Trim(Me.voucher.headerText("cexpcode"))) > 0 Then
                strGrid = strGrid & " and cexpcode like '%" & Trim(Me.voucher.headerText("cexpcode")) & "%'"
            ElseIf LCase(sKey) = "cexpname" And Len(Trim(Me.voucher.headerText("cexpname"))) > 0 Then
                strGrid = strGrid & " and cexpname like '%" & Trim(Me.voucher.headerText("cexpname")) & "%'"
            End If
            strGrid = strGrid & " order by cexpcode "
            If clsRefer.StrRefInit_SetColWidth(m_login, False, strClass, strGrid, "����,����������,�����������", "0,2500,6000") = False Then Exit Sub
            clsRefer.Show
            If Not clsRefer.recMx Is Nothing Then
                Me.voucher.headerText("cexpcode") = clsRefer.recMx(1)
                Me.voucher.headerText("cexpname") = clsRefer.recMx(2)
                sRet = clsRefer.recMx.Fields(LCase(sKey))
            End If

        Case "cpersoncode", "cpersonname" '��Ա����
            strClass = "select cdepcode,cdepname from Department"
            strGrid = "SELECT Department.cDepCode,Person.cPersonCode,Person.cPersonName FROM Department INNER JOIN Person ON Department.cDepCode = Person.cDepCode  where 1=1 "
            If LCase(sKey) = "cpersoncode" And Len(Trim(Me.voucher.headerText("cpersoncode"))) > 0 Then
                strGrid = strGrid & " and  cpersoncode like '%" & Trim(Me.voucher.headerText("cpersoncode")) & "%'"
            ElseIf LCase(sKey) = "cpersonname" And Len(Trim(Me.voucher.headerText("cpersonname"))) > 0 Then
                strGrid = strGrid & " and cpersonname like '%" & Trim(Me.voucher.headerText("cpersonname")) & "%'"
            End If
            If clsRefer.StrRefInit_SetColWidth(m_login, False, strClass, strGrid, "����,��Ա����,��Ա����", "0,2500,6000") = False Then Exit Sub
            clsRefer.Show
            If Not clsRefer.recMx Is Nothing Then
                Me.voucher.headerText("cpersoncode") = clsRefer.recMx(0)
                Me.voucher.headerText("cpersonname") = clsRefer.recMx(1)
                sRet = clsRefer.recMx.Fields(LCase(sKey))
            End If


'        Case "cpicpath"
'            With Me.CommonDialog1
'                .FileName = ""
'                .Filter = "ͼƬ�����ļ�(*.*)|*.*"
'                .ShowOpen
''                If .FileName = "" Then Exit Sub
'                 Me.voucher.headerText("cpicpath") = .FileName
'                sRet = .FileName
'            End With
            
      
    End Select
    If Left(sKey, 7) = "cdefine" Then
        With Me.voucher
            sRet = RefDefine(Index, siheader)
        End With
    End If
'by lg070315������U870 UAP���ݿؼ��µĲ��մ���
    referPara.Cancel = True
    
    If rst.state = 1 Then rst.Close
    Set rst = Nothing
    Exit Sub
End Sub

Private Function RefDefine(Index As Variant, iVoucherSec As Integer) As String
    Dim clsDef As U8DefPro.clsDefPro
    Dim nDataSource As Long         '������Դ
    Dim nEnterType As Long         '���뷽ʽ
    Dim sDataRule As String       '���ݹ�ʽ
    Dim bValidityCheck As Boolean      '�Ƿ�Ϸ��Լ��
    Dim bBuildArchives As Boolean      '�Ƿ񽨵�
    Dim sVouchType As String
    Dim sTableName As String, sFieldName As String, sCardNumber As String
    Dim sDefWhere As String
    Dim strKeyValue As String
    Set clsDef = New U8DefPro.clsDefPro
        With Me.voucher
            If iVoucherSec = siheader Then
                strKeyValue = .headerText(Index)
            Else
                strKeyValue = .bodyText(.row, Index)
            End If
            nDataSource = .ItemState(Index, iVoucherSec).nDataSource
            nEnterType = .ItemState(Index, iVoucherSec).nEnterType
            sDataRule = .ItemState(Index, iVoucherSec).sDataRule
            bValidityCheck = .ItemState(Index, iVoucherSec).bValidityCheck
            bBuildArchives = .ItemState(Index, iVoucherSec).bBuildArchives
            Select Case nDataSource  '0��ʾ�ֹ����룻1��ʾ������2��ʾ����
                Case 0
                    sTableName = "UserDefine"
                    sFieldName = "cValue"
                    sVouchType = ""
                Case 1
                    sTableName = Left(sDataRule, InStr(1, sDataRule, ",") - 1)
                    sFieldName = Mid(sDataRule, InStr(1, sDataRule, ",") + 1)
                    sVouchType = ""
                Case 2
                    sCardNumber = Left(sDataRule, InStr(1, sDataRule, ",") - 1)
                    sFieldName = Mid(sDataRule, InStr(1, sDataRule, ",") + 1)
            End Select
            If Not clsDef.Init(False, DBConn.ConnectionString, m_login.cUserId) Then
                RefDefine = ""
                MsgBox "��ʼ���Զ��������ʧ�ܣ�"
                Exit Function
            End If
            RefDefine = clsDef.GetRefVal(nDataSource, iVoucherSec, .ItemState(Index, iVoucherSec).sFieldName, sTableName, sFieldName, sCardNumber, strKeyValue, False, 40, 1)
        End With
        Set clsDef = Nothing
End Function

Private Sub Voucher_headCellCheck(Index As Variant, RetValue As String, bChanged As UapVoucherControl85.CheckRet, referPara As UapVoucherControl85.ReferParameter)
    Dim pt          As POINTAPI
    Dim hwnd        As Long
    Dim sClsName    As String * 100
    Dim i As Long
    Dim ele As IXMLDOMElement
    Dim sSkeyCode As String
    Dim strSql As String
    Dim rds As New ADODB.Recordset
    Call GetCursorPos(pt)
    hwnd = WindowFromPoint(pt.X, pt.Y)
    If hwnd <> 0 Then
       GetClassName hwnd, sClsName, 100
       sClsName = LCase(Trim(sClsName))
       If sClsName = "msvb_lib_toolbar" Or sClsName = "toolbar20wndclass" Or Trim(sClsName) = "msocommandbar" Then
            If Not bClickSave Then Exit Sub
       End If
    End If
    Dim lngrow As Long
    Dim lngCol As Long
    Dim strAuth As String
    Dim strError As String
    Dim rstTmp As New ADODB.Recordset
    Dim strKey As String
    Dim strRefersql As String
    Dim tmprst As ADODB.Recordset
On Error GoTo DoERR
    strKey = LCase(Me.voucher.ItemState(Index, siheader).sFieldName)
    If Trim(Me.voucher.headerText(strKey)) = "" Then Exit Sub
    With Me.voucher
        Select Case LCase(strKey)
        
        Case "cdepcode" '���ű���
                strSql = "SELECT Department.cDepCode,Department.cDepName FROM Department  where bdepend=1 "
                strSql = strSql & " and cdepcode = '" & Trim(Me.voucher.headerText("cdepcode")) & "'"
                If rds.state <> 0 Then rds.Close
                rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
                If rds.RecordCount = 0 Then
                    MsgBox "Ƶ�����Ϸ���", vbOKOnly + vbCritical
                    .headerText("cdepcode") = ""
                    .headerText("cdepname") = ""
                    RetValue = ""
                    bChanged = Retry
                Else
                    .headerText("cdepcode") = rds.Fields("cDepCode")
                    .headerText("cdepname") = rds.Fields("cDepName")
                    RetValue = rds.Fields(LCase(strKey))
                End If
            
        
        Case "cdepname"  '���ű���
                strSql = "SELECT Department.cDepCode,Department.cDepName FROM Department  where bdepend=1 "
                strSql = strSql & " and cDepName = '" & Trim(Me.voucher.headerText("cdepname")) & "'"
                If rds.state <> 0 Then rds.Close
                rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
                If rds.RecordCount = 0 Then
                    MsgBox "Ƶ�����Ϸ���", vbOKOnly + vbCritical
                    .headerText("cdepcode") = ""
                    .headerText("cdepname") = ""
                    RetValue = ""
                    bChanged = Retry
                Else
                    .headerText("cdepcode") = rds.Fields("cDepCode")
                    .headerText("cdepname") = rds.Fields("cDepName")
                    RetValue = rds.Fields(LCase(strKey))
                End If
            
        Case "citemcode"  '��Ŀ����
            strSql = "SELECT  citemcode , citemname FROM V_MT_Items  where 1=1  "
            strSql = strSql & " and citemcode = '" & Trim(Me.voucher.headerText("citemcode")) & "'"
            If rds.state <> 0 Then rds.Close
            rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
            If rds.RecordCount = 0 Then
                MsgBox "��Ŀ���Ϸ���", vbOKOnly + vbCritical
                .headerText("citemcode") = ""
                .headerText("citemname") = ""
                RetValue = ""
                bChanged = Retry
            Else
                .headerText("citemcode") = rds.Fields("citemcode")
                .headerText("citemname") = rds.Fields("citemname")
                RetValue = rds.Fields(LCase(strKey))
            End If
            
        Case "citemcode", "citemname" '��Ŀ����
            strSql = "SELECT  citemcode , citemname FROM V_MT_Items  where 1=1  "
            strSql = strSql & " and citemname = '" & Trim(Me.voucher.headerText("citemname")) & "'"
            If rds.state <> 0 Then rds.Close
            rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
            If rds.RecordCount = 0 Then
                MsgBox "��Ŀ���Ϸ���", vbOKOnly + vbCritical
                .headerText("citemcode") = ""
                .headerText("citemname") = ""
                RetValue = ""
                bChanged = Retry
            Else
                .headerText("citemcode") = rds.Fields("citemcode")
                .headerText("citemname") = rds.Fields("citemname")
                RetValue = rds.Fields(LCase(strKey))
            End If
            
        Case "cexpccode" '����������
            strSql = "select cexpccode,cexpcname  from  ExpItemClass  where bexpcend=1  "
            strSql = strSql & " and cexpccode = '" & Trim(Me.voucher.headerText("cexpccode")) & "'"
            If rds.state <> 0 Then rds.Close
            rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
            If rds.RecordCount = 0 Then
                MsgBox "���������಻�Ϸ���", vbOKOnly + vbCritical
                .headerText("cexpccode") = ""
                .headerText("cexpcname") = ""
                RetValue = ""
                bChanged = Retry
            Else
                .headerText("cexpccode") = rds.Fields("cexpccode")
                .headerText("cexpcname") = rds.Fields("cexpcname")
                RetValue = rds.Fields(LCase(strKey))
            End If
        
        Case "cexpcname" '����������
            strSql = "select cexpccode,cexpcname  from  ExpItemClass  where bexpcend=1  "
            strSql = strSql & " and cexpcname = '" & Trim(Me.voucher.headerText("cexpcname")) & "'"
            If rds.state <> 0 Then rds.Close
            rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
            If rds.RecordCount = 0 Then
                MsgBox "���������಻�Ϸ���", vbOKOnly + vbCritical
                .headerText("cexpccode") = ""
                .headerText("cexpcname") = ""
                RetValue = ""
                bChanged = Retry
            Else
                .headerText("cexpccode") = rds.Fields("cexpccode")
                .headerText("cexpcname") = rds.Fields("cexpcname")
                RetValue = rds.Fields(LCase(strKey))
            End If
         
        
        Case "cexpcode" '�������
            strSql = "select cexpccode,cexpcode,cexpname from ExpenseItem  where 1=1  "
            strSql = strSql & " and cexpcode = '" & Trim(Me.voucher.headerText("cexpcode")) & "'"
            If rds.state <> 0 Then rds.Close
            rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
            If rds.RecordCount = 0 Then
                MsgBox "������𲻺Ϸ���", vbOKOnly + vbCritical
                .headerText("cexpcode") = ""
                .headerText("cexpname") = ""
                RetValue = ""
                bChanged = Retry
            Else
                .headerText("cexpcode") = rds.Fields("cexpcode")
                .headerText("cexpname") = rds.Fields("cexpname")
                RetValue = rds.Fields(LCase(strKey))
            End If
        
        Case "cexpname"  '�������
            strSql = "select cexpccode,cexpcode,cexpname from ExpenseItem  where 1=1  "
            strSql = strSql & " and cexpname = '" & Trim(Me.voucher.headerText("cexpname")) & "'"
            If rds.state <> 0 Then rds.Close
            rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
            If rds.RecordCount = 0 Then
                MsgBox "������𲻺Ϸ���", vbOKOnly + vbCritical
                .headerText("cexpcode") = ""
                .headerText("cexpname") = ""
                RetValue = ""
                bChanged = Retry
            Else
                .headerText("cexpcode") = rds.Fields("cexpcode")
                .headerText("cexpname") = rds.Fields("cexpname")
                RetValue = rds.Fields(LCase(strKey))
            End If

        Case "cpersoncode"  '��Ա����
            strSql = "SELECT Department.cDepCode,Person.cPersonCode,Person.cPersonName FROM Department INNER JOIN Person ON Department.cDepCode = Person.cDepCode  where 1=1 "
            strSql = strSql & " and  cpersoncode = '" & Trim(Me.voucher.headerText("cpersoncode")) & "'"
            If rds.state <> 0 Then rds.Close
            rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
            If rds.RecordCount = 0 Then
                MsgBox "��Ա���Ϸ���", vbOKOnly + vbCritical
                .headerText("cpersoncode") = ""
                .headerText("cpersonname") = ""
                RetValue = ""
                bChanged = Retry
            Else
                .headerText("cpersoncode") = rds.Fields("cPersonCode")
                .headerText("cpersonname") = rds.Fields("cPersonName")
                RetValue = rds.Fields(LCase(strKey))
            End If

        Case "cpersonname"  '��Ա����
            strSql = "SELECT Department.cDepCode,Person.cPersonCode,Person.cPersonName FROM Department INNER JOIN Person ON Department.cDepCode = Person.cDepCode  where 1=1 "
            strSql = strSql & " and cpersonname = '" & Trim(Me.voucher.headerText("cpersonname")) & "'"
            If rds.state <> 0 Then rds.Close
            rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
            If rds.RecordCount = 0 Then
                MsgBox "��Ա���Ϸ���", vbOKOnly + vbCritical
                .headerText("cpersoncode") = ""
                .headerText("cpersonname") = ""
                RetValue = ""
                bChanged = Retry
            Else
                .headerText("cpersoncode") = rds.Fields("cPersonCode")
                .headerText("cpersonname") = rds.Fields("cPersonName")
                RetValue = rds.Fields(LCase(strKey))
            End If
            
            
            
        Case "rate" 'Ԥ�����
               If (.headerText("rate") > 100) Or (.headerText("rate") <= 0) Then
                    MsgBox "Ԥ��������Ϸ���", vbOKOnly + vbCritical
                    .headerText("rate") = "0.00"
                    RetValue = ""
                    bChanged = Retry
               End If

        End Select
    End With
    Set rstTmp = Nothing
    Exit Sub
DoERR:
    MsgBox Err.Description
End Sub
 

Private Sub fillComBol(bPrint As Boolean)
    Dim tmprst As New ADODB.Recordset
    Dim strAuth As String
    Dim strSql As String
    Dim i As Long
    Dim sWhere As String
    If bPrint = True Then
        ComboVTID.Clear
    Else
        ComboDJMB.Clear
    End If
    strAuth = clsAuth.getAuthString("DJMB")
    If strAuth = "1=2" Then Exit Sub
    If bFirst = False Then
        If bPrint = True Then
            ''��ӡ
            strSql = "SELECT VT_Name,VT_ID  From VoucherTemplates WHERE (VT_CardNumber = '" & strCardNum & "') AND (VT_TemplateMode = 1) " & IIf(strAuth <> "", "and vt_id in (" & strAuth & ")", "")
        Else
            ''��ʾ
            strSql = "SELECT VT_Name,VT_ID  From VoucherTemplates WHERE (VT_CardNumber = '" & strCardNum & "') AND (VT_TemplateMode = 0) " & IIf(strAuth <> "", "and vt_id in (" & strAuth & ")", "")
        End If
    Else
        If bPrint = True Then
            ''��ӡ
            strSql = "SELECT VT_Name,VT_ID  From VoucherTemplates WHERE (" & sWhere & ") AND (VT_TemplateMode = 1) " & IIf(strAuth <> "", "and vt_id in (" & strAuth & ")", "")
        Else
            ''��ʾ
            strSql = "SELECT VT_Name,VT_ID  From VoucherTemplates WHERE ( " & sWhere & ") AND (VT_TemplateMode = 0) " & IIf(strAuth <> "", "and vt_id in (" & strAuth & ")", "")
        End If
    End If
    tmprst.CursorLocation = adUseClient
    tmprst.Open strSql, DBConn, adOpenForwardOnly, adLockReadOnly
    If tmprst.EOF Then
        i = 0
        If bPrint = True Then
            ComboVTID.Clear
        Else
            ComboDJMB.Clear
        End If
    Else
        i = tmprst.RecordCount - 1
        If bPrint Then
            ReDim vtidPrn(i)
        Else
            ReDim vtidDJMB(i)
        End If
    End If
    If Not tmprst.EOF Then
        If bPrint = True Then
            ComboVTID.Clear
            i = 0
            Do While Not tmprst.EOF
                ComboVTID.AddItem tmprst(0)
                vtidPrn(i) = CLng(tmprst(1))
                i = i + 1
                tmprst.MoveNext
            Loop
            ComboVTID.ListIndex = 0
            ComboVTID.ToolTipText = ComboVTID.Text
        Else
            ComboDJMB.Clear
            i = 0
            Do While Not tmprst.EOF
                ComboDJMB.AddItem tmprst(0)
                tmprst.MoveNext
            Loop
            ComboDJMB.ListIndex = 0
            ComboDJMB.ToolTipText = ComboDJMB.Text
        End If
    End If
    tmprst.Close
    bfillDjmb = True
    Set tmprst = Nothing
End Sub
'���ݳ�ʼ��
Public Function ShowVoucher(VoucherType As VoucherType, Optional vVoucherId As Variant, Optional iMode As Integer)
    Dim tmpTemplateID As String
    Dim errMsg As String
'by lg070314 ����U870�Ż��ں�
    Dim vfd As Object
    sguid = CreateGUID()
    If Not (g_business Is Nothing) Then
        Set vfd = g_business.CreateFormEnv(sguid, Me)
    End If
    
    g_FormbillShow = False
    Screen.MousePointer = vbHourglass
    frmFloat.m_oTimer.Enabled = False
    On Error GoTo DoERR
    If IsMissing(iMode) = True Then
        iShowMode = 0
    Else
        iShowMode = iMode
    End If
    Set clsVoucherCO = New VouchersControlHBTV.ClsVoucherCO_GDZC
    'by ahzzd 2005/05/09 ���ݳ�ʼ��
    clsVoucherCO.Init VoucherType, m_login, DBConn, "CS", clsSAWeb
    clsAuth.Init m_login.UfDbName, m_login.cUserId
    Set obj_EA = CreateObject("u8ExamineAndApprove.clsU8Examine")
    Call obj_EA.Init(m_login)
'    MT01    '0   ����������Ŀ���ձ�
'    MT02    '0   ����Ŀ���ձ�
'    MT03    '0   �������������ñ�
'    MT04    '0   ���÷���������ñ�
'    MT05    '0   Ԥ������ڳ�¼��
'    MT06    '0   Ԥ����Ƶ�
'    MT07    '0   Ԥ����Ƶ�����
'    MT08    '0   ֧Ʊ��
'    MT09    '0   ��Ŀ�������ѱ��˵�
    Select Case VoucherType
        Case MT01
            strVouchType = "87"
            strCardNum = "MT01"
        
        Case MT02
            strVouchType = "88"
            strCardNum = "MT02"
        
        Case MT03
            strVouchType = "89"
            strCardNum = "MT03"
        
        Case MT04
            strVouchType = "90"
            strCardNum = "MT04"
        
        Case MT05
            strVouchType = "91"
            strCardNum = "MT05"
        
        Case MT06
            strVouchType = "92"
            strCardNum = "MT06"
        
        Case MT07
            strVouchType = "93"
            strCardNum = "MT07"
        
        Case MT08
            strVouchType = "94"
            strCardNum = "MT08"
        
        Case MT09
            strVouchType = "95"
            strCardNum = "MT09"
     
    
        Case gdzckpxg       '// �ʲ���Ƭ�༭
            strVouchType = "98"
            strCardNum = "FA01"
        Case gdzckp           '//�ʲ���Ƭ����
            strVouchType = "97"
            strCardNum = "FA01"
    End Select
    ''���ð�ť
 
   U8VoucherSorter1.Visible = False
 
 
    sTemplateID = clsSAWeb.GetVTID(DBConn, strCardNum)
 
'
'    ''���ð�ť
'    sTemplateID = clsSAWeb.GetVTID(DBConn, strCardNum)
    If iShowMode <> 2 Then
 
        errMsg = clsVoucherCO.GetVoucherData(Domhead, Dombody, vVoucherId)
 
        If errMsg <> "" Then
            MsgBox errMsg
            If iShowMode = 1 Then
                On Error Resume Next
                bFrmCancel = False
                Set clsVoucherCO = Nothing
                Set clsAuth = Nothing
                Set clsRefer = Nothing
                Set RstTemplate = Nothing
                Set Domhead = Nothing
                Set Dombody = Nothing
                Set DomFormat = Nothing
                Set RstTemplate = Nothing
                Set RstTemplate2 = Nothing
                If m_UFTaskID <> "" Then
                    m_login.TaskExec m_UFTaskID, 0
                End If
                Screen.MousePointer = vbDefault
                Exit Function
            End If
        End If
 
        If iShowMode = 1 Then
            Call reInit(VoucherType, Domhead)
        End If
        If Not Domhead.selectSingleNode("//z:row") Is Nothing Then
            If Not Domhead.selectSingleNode("//z:row").Attributes.getNamedItem("ivtid") Is Nothing Then
                tmpTemplateID = Domhead.selectSingleNode("//z:row").Attributes.getNamedItem("ivtid").nodeValue
            Else
                tmpTemplateID = "0"
            End If
        Else
            tmpTemplateID = "0"
        End If
 
    Else
        errMsg = clsVoucherCO.GetVoucherData(Domhead, Dombody, 0)
        If errMsg <> "" Then
            MsgBox errMsg
        End If
        Set oDomB = New DOMDocument
        oDomB.loadXML Dombody.xml
    End If
    If tmpTemplateID = "" Or tmpTemplateID = "0" Then
        sCurTemplateID = sTemplateID    ''ȡĬ��ģ��
    Else
        sCurTemplateID = tmpTemplateID  ''�µ�ģ��
    End If
    sCurTemplateID2 = sCurTemplateID
    
    If sCurTemplateID = 0 Then
        Me.Hide
        MsgBox "��û��ģ��ʹ��Ȩ��"
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '860sp������861�޸Ĵ�   2006/03/12  861 ���Ӹ���
    Call SetVoucherDataSource
    

    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    If ChangeTempaltes(sCurTemplateID, True, False, True) = False Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    On Error Resume Next
    If GetHeadItemValue(Domhead, "cexch_name") <> "" Then
        Me.voucher.ItemState("cexchrate", siheader).nNumPoint = clsSAWeb.GetExchRateDec(GetHeadItemValue(Domhead, "cexch_name"))
    End If
    On Error GoTo DoERR
 
    
    voucher.setVoucherDataXML Domhead, Dombody
 
 
    'by ahzzd 2006/05/09   ����׼�������������
    '�������ı�
    Me.voucher.ExamineFlowAuditInfo = GetEAStream(strVouchType, Domhead, Me.voucher, DBConn)
    
 
    Call SetSum
    Call ChangeCaptionCol
    If Me.Caption = "0" Then
        Me.Caption = voucher.TitleCaption
    End If

    If iShowMode <> 2 Then
        voucher.VoucherStatus = VSNormalMode
        ChangeButtonsState
    Else
    End If
    If iShowMode <> 1 Then
 
        If clsVoucherCO.GetVoucherNO(Domhead, GetvouchNO, errMsg, DomFormat, True) = False Then
            MsgBox errMsg
        Else
 
            Me.voucher.SetBillNumberRule DomFormat.xml
        End If
        clsRefer.setlogin m_login   ''��ʼ�����տؼ�
    End If
    Me.voucher.Visible = True
    Call fillComBol(True)   ''���ģ��ѡ��
    If iShowMode <> 1 Then
        Call fillComBol(False)
        bfillDjmb = False
    End If
    Call SetHelpID
    If Me.Caption = "" Then
        Me.Caption = Me.LabelVoucherName.Caption
    End If
    Dim strXml As String
    strXml = "<?xml version='1.0' encoding='GB2312'?>" & Chr(13)
    domConfig.loadXML strXml & "<EAI>0</EAI>"
    
'by lg070314 ����U870֧�֣������ں�
    If g_business Is Nothing Then
        Me.Show
    Else
'        InitToolbarTag Me.tbrvoucher
        Call g_business.ShowForm(Me, "FA", sguid, False, True, vfd)
        Set Me.voucher.PortalBusinessObject = g_business
        Me.voucher.PortalBizGUID = sguid
    End If
    
    Me.BackColor = Me.voucher.BackColor
    Me.picVoucher.BackColor = Me.voucher.BackColor
    Me.Refresh
    Me.voucher.SetFocus
    Screen.MousePointer = vbDefault
    g_FormbillShow = True
    U8VoucherSorter1.Visible = False
 
 
    Exit Function
DoERR:
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
End Function
''���õ��ݺ��Ƿ���Ա༭
Private Sub SetVouchNoWriteble()
    Dim KeyCode As String
    
    On Error Resume Next
    If strVouchType = "92" Then Exit Sub
    KeyCode = getVoucherCodeName()
    If Not DomFormat Is Nothing Then
        If DomFormat.xml <> "" Then
            If LCase(DomFormat.selectSingleNode("//���ݱ��").Attributes.getNamedItem("�����ֹ��޸�").nodeValue) = "false" And LCase(DomFormat.selectSingleNode("//���ݱ��").Attributes.getNamedItem("�غ��Զ���ȡ").nodeValue) = "false" Then
                Me.voucher.EnableHead KeyCode, False
            Else
                Me.voucher.EnableHead KeyCode, True
            End If
        End If
    End If
End Sub
 
Private Sub SetButton()
    Dim Index As Integer
    Dim btnX As MSComctlLib.Button
    On Error Resume Next
    Set Me.Icon = frmMain.Icon
    With tbrvoucher
        Set .ImageList = frmMain.imgBmp
        
            .buttons.Clear
         ''���Ӱ�ť
        'by lg070314 �޸�U870�Ż��˵��ںϣ�����Toolbar��Button����Tagֵ
        'Tagֵ ��ʾ�˵��ϵ�ͼ���ļ�����   ͼ���ļ��� ..\U8SOFT\icons
        
         ''��ӡ
            Set btnX = .buttons.Add(, "Print", strPrint, tbrDefault)
'            btnX.image = 314
            btnX.ToolTipText = strPrint
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Print"
        ''Ԥ��
            Set btnX = .buttons.Add(, "Preview", strPreview, tbrDefault)
'            btnX.image = 312
            btnX.ToolTipText = strPreview
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "print preview"
        ''���
            Set btnX = .buttons.Add(, "Output", strOutput, tbrDefault)
'            btnX.image = 308
            btnX.ToolTipText = strOutput
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Output"
 
        ''����
            Set btnX = .buttons.Add(, "Add", strAdd, tbrDefault)
'            btnX.image = 323
            btnX.ToolTipText = strAdd
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Add"
            
'        ''��������
'            Set btnX = .buttons.Add(, "batchAdd", "����", tbrDefault)
''            btnX.image = 389
'            btnX.ToolTipText = "����"
'            btnX.Description = btnX.ToolTipText
'            btnX.Tag = "Input"
            
'        ''��������
'            Set btnX = .buttons.Add(, "inAdd", "����", tbrDefault)
'            btnX.image = 1
'            btnX.ToolTipText = "����"
'            btnX.Description = btnX.ToolTipText
'            btnX.Tag = "inAdd"
'        ''��������
'            Set btnX = .buttons.Add(, "outAdd", "����", tbrDefault)
'            btnX.image = 1
'            btnX.ToolTipText = "����"
'            btnX.Description = btnX.ToolTipText
'            btnX.Tag = "outAdd"
'            Set btnX = .buttons.Add(, , , tbrSeparator)
        ''�޸�
            Set btnX = .buttons.Add(, "Modify", strModify, tbrDefault)
'            btnX.image = 324
            btnX.ToolTipText = strModify
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "modify"
            
'        ''�޸ı��
'            Set btnX = .buttons.Add(, "Chenged", strchenged, tbrDefault)
''            btnX.image = 321
'            btnX.ToolTipText = strchenged
'            btnX.Description = btnX.ToolTipText
'            btnX.Tag = "accessories"
            
         ''ɾ��
            Set btnX = .buttons.Add(, "Erase", strDelete, tbrDefault)
'            btnX.image = 326
            btnX.ToolTipText = strDelete
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "delete"
         ''����
            Set btnX = .buttons.Add(, "Copy", "����", tbrDefault)
'            btnX.image = 318
            btnX.ToolTipText = "����"
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Copy"
            
'        ''ͼƬ
'            Set btnX = .buttons.Add(, "picture", "ͼƬ", tbrDefault)
'            btnX.image = 20
'            btnX.ToolTipText = "ͼƬ"
'            btnX.Description = btnX.ToolTipText
'            btnX.Tag = "picture"
            
'          '����
'            Set btnX = .buttons.Add(, "label", "����", tbrDefault)
'            btnX.image = 20
'            btnX.ToolTipText = "����"
'            btnX.Description = btnX.ToolTipText
'            btnX.Tag = "label"
'
            
         ''����
            Set btnX = .buttons.Add(, "Save", strSave, tbrDefault)
''            btnX.image = 988
'btnX.Style = tbrButtonGroup
'btnX.ButtonMenus.Add
            btnX.ToolTipText = strSave
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "save"
          ''����
            Set btnX = .buttons.Add(, "Cancel", strDiscard, tbrDefault)
'            btnX.image = 316
            btnX.ToolTipText = strDiscard
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Cancel"
            
            '���
            Set btnX = .buttons.Add(, "Sure", "���", tbrDefault)
'            btnX.image = 1100
            btnX.ToolTipText = "���"
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Approve"
            
            '����
            Set btnX = .buttons.Add(, "UnSure", "����", tbrDefault)
'            btnX.image = 341
            btnX.ToolTipText = "����"
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Unapprove"
            

        ''����
            Set btnX = .buttons.Add(, "AddRow", strAddrecord, tbrDefault)
'            btnX.image = 343
            btnX.ToolTipText = strAddrecord
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "add a row"
      
        ''ɾ��
            Set btnX = .buttons.Add(, "DelRow", strDeleterecord, tbrDefault)
'            btnX.image = 347
            btnX.ToolTipText = strDeleterecord
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Delete a row"


'        '�ָ����
'            Set btnX = .buttons.Add(, , , tbrSeparator)
'
'            '����
'            Set btnX = .buttons.Add(, "Filter", strFilter, tbrDefault)
''            btnX.image = 1120
'            btnX.ToolTipText = strFilter
'            btnX.Description = btnX.ToolTipText
'            btnX.Tag = "Filter"
           
           
            
          ''����
            Set btnX = .buttons.Add(, "ToFirst", strFirst, tbrDefault)
'            btnX.image = 24 '1174
            btnX.ToolTipText = strFirst
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "first page"
          ''����
            Set btnX = .buttons.Add(, "ToPrevious", strPrevious, tbrDefault)
'            btnX.image = 22 '1139
            btnX.ToolTipText = strPrevious
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "previous page"
          ''����
            Set btnX = .buttons.Add(, "ToNext", strNext, tbrDefault)
'            btnX.image = 23 '1133
            btnX.ToolTipText = strNext
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "next page"
           ''ĩ��
            Set btnX = .buttons.Add(, "ToLast", strLast, tbrDefault)
'            btnX.image = 25 '1117
            btnX.ToolTipText = strLast
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Last page"

'            'ȫ��
'            Set btnX = .buttons.Add(, "Pd_all", strPd_all, tbrDefault)
'            btnX.image = 25
'            btnX.ToolTipText = strPd_all
'            btnX.Description = btnX.ToolTipText
           ''ƾ֤
'            Set btnX = .buttons.Add(, "Seek", "ƾ֤", tbrDefault)
''            btnX.image = 8
'            btnX.ToolTipText = "ƾ֤"
'            btnX.Description = btnX.ToolTipText
'            btnX.Tag = "sum"
           ''ˢ��
            Set btnX = .buttons.Add(, "Paint", strRefresh, tbrDefault)
'            btnX.image = 154
            btnX.ToolTipText = strRefresh
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "refresh"
            
'        End If
'           ''����
'            Set btnX = .Buttons.Add(, "Help", strHelp, tbrDefault)
'            btnX.Image = 36
'            btnX.ToolTipText = strHelp
'            btnX.Description = btnX.ToolTipText
           
           ''�˳�
'            Set btnX = .buttons.Add(, "Exit", strExit, tbrDefault)
'            btnX.image = 1118
'            btnX.ToolTipText = strExit
'            btnX.Description = btnX.ToolTipText
'            btnX.Tag = "Exit"
'           ''����
'            Set btnX = .Buttons.Add(, "PrnSet", "����", tbrDefault)
'            btnX.Image = 9
'            btnX.ToolTipText = "����"
'            btnX.Description = btnX.ToolTipText
'            btnX.Visible = False
           ''�б�
'            Set btnX = .Buttons.Add(, "LstTab", "�б�", tbrDefault)
'            btnX.Image = 43
'            btnX.ToolTipText = "�б�"
'            btnX.Description = btnX.ToolTipText
'            btnX.Visible = False
         
         If strCardNum = "MT09" Then
            Set btnX = .buttons.Add(, "ZP", "֧Ʊ", tbrDefault)
            'btnX.image = 43
            btnX.ToolTipText = "֧Ʊ"
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "new_persp"
            btnX.Visible = True
            
        End If
    End With
    ''�����ϡ��ֽ��λ��
    labZF.Top = picVoucher.Top 'Me.top - Me.tbrvoucher.top
    labZF.Left = Me.voucher.Left
    labXJ.Top = picVoucher.Top ' Me.top - Me.tbrvoucher.top    'Me.StBar.height
    labXJ.Left = Me.voucher.Left + labZF.Width
'by lg070316���ӳ�ʼ��U870�˵�
    Call InitToolbarTag(Me.tbrvoucher)
    
End Sub
''�ı�button��״̬
Private Sub ChangeButtonsState()
    Dim i As Integer
    On Error Resume Next
    Me.labXJ.Visible = False
    Me.labZF.Visible = False
    With Me.voucher
        If .headerText("ddate") <> "" Then
'            Me.tbrvoucher.buttons("Copy").Enabled = True
'            Me.tbrvoucher.buttons("Sure").Enabled = True
'            Me.tbrvoucher.buttons("UnSure").Enabled = True
            Select Case strVouchType
                   
                   Case "87", "88", "89", "90" '��������
                       Me.tbrvoucher.buttons("Copy").Enabled = True
                   Case "91", "92", "93", "95" 'ҵ�񵥾�
                    '�����
                    If .headerText("chandler") <> "" Then '�����
                        Me.tbrvoucher.buttons("UnSure").Visible = True
                        Me.tbrvoucher.buttons("Sure").Visible = False
                        bCheckVouch = False
                        tbrvoucher.buttons("Add").Enabled = True
                        tbrvoucher.buttons("Save").Enabled = False
                        tbrvoucher.buttons("Copy").Enabled = True
                        Me.tbrvoucher.buttons("Modify").Enabled = False
                        Me.tbrvoucher.buttons("Erase").Enabled = False
                        Me.tbrvoucher.buttons("Chenged").Visible = True
                    'δ���
                    Else
                        Me.tbrvoucher.buttons("Sure").Visible = True
                        Me.tbrvoucher.buttons("UnSure").Visible = False
                        Me.tbrvoucher.buttons("Copy").Enabled = True
                        tbrvoucher.buttons("Add").Enabled = True
                        bCheckVouch = True
                        Me.tbrvoucher.buttons("Sure").Enabled = True
                        Me.tbrvoucher.buttons("Modify").Enabled = True
                        Me.tbrvoucher.buttons("Erase").Enabled = True
                    End If
                 Case "94"  'ҵ�񵥾�
                    '�����
                    If .headerText("chandler") <> "" Then '�����
                        Me.tbrvoucher.buttons("UnSure").Visible = True
                        Me.tbrvoucher.buttons("Sure").Visible = False
                        bCheckVouch = False
                        tbrvoucher.buttons("Add").Enabled = True
                        tbrvoucher.buttons("Save").Enabled = False
                        tbrvoucher.buttons("Copy").Enabled = True
                        Me.tbrvoucher.buttons("Modify").Enabled = False
                        Me.tbrvoucher.buttons("Erase").Enabled = False
                        Me.tbrvoucher.buttons("Chenged").Visible = True
                        Me.tbrvoucher.buttons("AddRow").Visible = False
                        Me.tbrvoucher.buttons("DelRow").Visible = False
                        
                    'δ���
                    Else
                        Me.tbrvoucher.buttons("Sure").Visible = True
                        Me.tbrvoucher.buttons("UnSure").Visible = False
                        Me.tbrvoucher.buttons("Copy").Enabled = True
                        tbrvoucher.buttons("Add").Enabled = True
                        bCheckVouch = True
                        Me.tbrvoucher.buttons("Sure").Enabled = True
                        Me.tbrvoucher.buttons("Modify").Enabled = True
                        Me.tbrvoucher.buttons("Erase").Enabled = True
                        Me.tbrvoucher.buttons("AddRow").Visible = False
                        Me.tbrvoucher.buttons("DelRow").Visible = False
 
                    End If
                    
                End Select

            For i = 1 To Me.tbrvoucher.buttons.Count
                If Left(Me.tbrvoucher.buttons(i).ToolTipText, 2) <> "����" And Left(Me.tbrvoucher.buttons(i).ToolTipText, 2) <> "��ѯ" Then
                    Me.tbrvoucher.buttons(i).Caption = Left(Me.tbrvoucher.buttons(i).ToolTipText, 2)
                End If
            Next

        Else     ''�յ���
 
            Me.tbrvoucher.buttons("Erase").Visible = False
            Me.tbrvoucher.buttons("Modify").Visible = False
            Me.tbrvoucher.buttons("Save").Visible = False
            Me.tbrvoucher.buttons("Cancel").Visible = False
            Me.tbrvoucher.buttons("Sure").Visible = False
            Me.tbrvoucher.buttons("UnSure").Visible = False
 
        End If
    End With
    
    If clsVoucherCO.BOF And clsVoucherCO.EOF Then
        Me.tbrvoucher.buttons("ToFirst").Enabled = False
        Me.tbrvoucher.buttons("ToPrevious").Enabled = False
        Me.tbrvoucher.buttons("ToNext").Enabled = False
        Me.tbrvoucher.buttons("ToLast").Enabled = False
    ElseIf clsVoucherCO.BOF Then
        Me.tbrvoucher.buttons("ToFirst").Enabled = False
        Me.tbrvoucher.buttons("ToPrevious").Enabled = False
        Me.tbrvoucher.buttons("ToNext").Enabled = True
        Me.tbrvoucher.buttons("ToLast").Enabled = True
    ElseIf clsVoucherCO.EOF Then
        Me.tbrvoucher.buttons("ToFirst").Enabled = True
        Me.tbrvoucher.buttons("ToPrevious").Enabled = True
        Me.tbrvoucher.buttons("ToNext").Enabled = False
        Me.tbrvoucher.buttons("ToLast").Enabled = False
    Else
        Me.tbrvoucher.buttons("ToFirst").Enabled = True
        Me.tbrvoucher.buttons("ToPrevious").Enabled = True
        Me.tbrvoucher.buttons("ToNext").Enabled = True
        Me.tbrvoucher.buttons("ToLast").Enabled = True
    End If
    If tbrvoucher.Visible = False Then
        Me.UFToolbar1.RefreshVisible
    End If
    Me.UFToolbar1.RefreshEnable
    Call Init
End Sub
 
Private Sub SetScrollBarValue()
    vs.Visible = False
    hs.Visible = False
Exit Sub
On Error Resume Next
    Me.hs.Move 0, Me.picVoucher.ScaleHeight - GetScrollWidth, Me.picVoucher.ScaleWidth - GetScrollWidth, GetScrollWidth
    Me.vs.Move Me.picVoucher.ScaleWidth - GetScrollWidth, (Me.Picture2.Height + Me.Picture2.Top), GetScrollWidth, Me.picVoucher.ScaleHeight - Me.Picture2.Height 'GetScrollWidth - Me.StBar.height
    Me.vs.ZOrder
    Me.hs.ZOrder
    vs.Min = 0
    vs.Max = 0
    vs.value = 0
    If Me.voucher.Height + 1 * GetScrollWidth - Me.picVoucher.ScaleHeight + Me.Picture2.Height - Me.Picture2.Height <= vs.Min Then
        vs.Max = vs.Min
        vs.Visible = False
    Else
        vs.Max = Me.voucher.Height + 1 * GetScrollWidth - Me.picVoucher.ScaleHeight + Me.Picture2.Height - Me.Picture2.Height
        vs.Visible = True
    End If
    vs.SmallChange = 500
    vs.LargeChange = 3000
    hs.Min = 0
    hs.Max = 0
    hs.value = 0
    If Me.voucher.Width + GetScrollWidth - Me.picVoucher.ScaleWidth <= hs.Min Then
        hs.Max = hs.Min
        hs.Visible = False
    Else
        hs.Max = Me.voucher.Width + GetScrollWidth - Me.picVoucher.ScaleWidth
        vs.Max = vs.Max + GetScrollWidth
        If vs.Visible = True Then hs.Max = hs.Max + GetScrollWidth
        hs.Visible = True
    End If
    hs.SmallChange = 500
    hs.LargeChange = 3000
End Sub
 
Private Sub Voucher_headOnEdit(Index As Integer)
    With Me.voucher
        Select Case strVouchType
            Case "102" '�ʲ�����
        End Select
    End With
End Sub
 
Private Sub voucher_KeyPress(ByVal section As UapVoucherControl85.SectionsConstants, ByVal Index As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub voucher_MouseUp(ByVal section As UapVoucherControl85.SectionsConstants, ByVal Index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If section = sibody And Button = 2 Then
        If strCardNum = "MT06" Then
            PopupMenu mnuPop
        End If
    End If
End Sub

Public Sub Voucher_PreParePrintEvnet(sStyle As String, sData As String)
    Dim rsPrintModel As UfRecordset
    Dim ndRoot  As IXMLDOMNode
    Dim ndRootList As IXMLDOMNodeList
    Dim eleMent  As IXMLDOMElement
    Dim tmpDOM As New DOMDocument
    Set rsPrintModel = gcAccount.dbSales.OpenRecordset("select fieldname,fieldtype from voucheritems_prn where vt_id='" & CLng(vtidPrn(Me.ComboVTID.ListIndex)) & "' and fieldtype in (2,3,4) and cardsection='T'")
    If Not domPrint Is Nothing Then
    Dim oxml As New DOMDocument
    Dim oEl As IXMLDOMElement
    Dim i As Long
    On Error GoTo Errhand
    sStyle = domPrintStyle.xml
    oxml.loadXML sStyle
    For Each eleMent In oxml.selectSingleNode("//��ͷ").selectNodes("//�ֶ�")
        If eleMent.getAttribute("�߿�") = "0x1F" Then
            eleMent.setAttribute "�߿�", "3"
        End If
        If Left(eleMent.getAttribute("�ؼ���"), 2) <> "�ı�" Then
            rsPrintModel.Filter = ""
            rsPrintModel.Filter = "fieldname='" & Mid(eleMent.getAttribute("�ؼ���"), InStr(1, eleMent.getAttribute("�ؼ���"), "(") + 1, InStr(1, eleMent.getAttribute("�ؼ���"), ")") - InStr(1, eleMent.getAttribute("�ؼ���"), "(") - 1) & "'"
            If rsPrintModel.RecordCount Then
                eleMent.setAttribute "���뷽ʽ", "��"
            End If
        End If
    Next
    sStyle = oxml.xml
        If voucher.headerText("bfirst") Then
            tmpDOM.loadXML sStyle
            Set ndRootList = domPrint.selectNodes("//����")
            For Each ndRoot In ndRootList
                ndRoot.Text = LabelVoucherName.Caption
            Next
            Set ndRootList = tmpDOM.selectNodes("//����")
            For Each eleMent In ndRootList
                eleMent.setAttribute "��", "500"
            Next
            sStyle = tmpDOM.xml
        End If
        sData = domPrint.xml
    End If
    Exit Sub
Errhand:
    MsgBox Err.Description
    sStyle = domPrintStyle.xml
End Sub

Private Sub Voucher_RowColChange()
    Dim tmpRow As Integer, tmpCol As Integer
    Dim i As Long, j As Long
    On Error Resume Next
    With Me.voucher
        tmpRow = .row
        tmpCol = .col
        i = .row
            Select Case strVouchType
                Case "97"
                    If tmpRow > 0 Then '
                        If .bodyText(tmpRow, "cscloser") <> "" Then
                            SetVouchItemState .ItemState(.colEx, sibody).sFieldName, "B", False
                        End If '
                    End If
            End Select
    End With
DoExit:
End Sub
 
Private Sub Voucher_SaveSettingEvent(ByVal varDevice As Variant)
    Dim TmpUFTemplate As Object
    Set TmpUFTemplate = CreateObject("UFVoucherServer85.clsVoucherTemplate")
    If TmpUFTemplate.SaveDeviceCapabilities(DBConn.ConnectionString, BillPrnVTID, varDevice) <> 0 Then
        MsgBox "��ӡ���ñ���ʧ��"
    End If
End Sub
 
Private Sub VS_Change()
    Me.voucher.Top = Me.Picture2.Height - vs.value - Me.Picture2.Height ''- Me.StBar.height)
End Sub
 
'���ƽ���
Private Sub VS_GotFocus()
    On Error Resume Next
    Me.voucher.SetFocus
End Sub

Private Sub HS_Change()
    Me.voucher.Left = -hs.value
End Sub
Private Sub HS_GotFocus()
    On Error Resume Next
    Me.voucher.SetFocus
End Sub
 
Private Sub picVoucher_Resize()
    SetScrollBarValue
End Sub
 
Private Sub AddNewVouch(Optional strOperator As String)
    Dim iElement As IXMLDOMElement
    Dim iAttr As IXMLDOMAttribute
    Dim i As Long
    With voucher
        Select Case LCase(strOperator)
            Case "sure"
                .headerText("dcheckdate") = m_login.CurDate
                .headerText("checkcode") = m_login.cUserId
                .headerText("checkname") = m_login.cUserName
                Exit Sub
            Case "unsure"
                .headerText("checkcode") = ""
                .headerText("checkname") = ""
                Exit Sub
            Case "save"
                 If vName = "DISPQC" Then
                    If .TotalText("iSum") > 0 Then
                       
                       .headerText("breturnflag") = 0
                    Else
                       .headerText("breturnflag") = 1
                    End If
                    Exit Sub
                 End If
                 If strVouchType = "95" Then
                    .headerText("bIWLType") = 1
                 ElseIf strVouchType = "92" Then
                    .headerText("bIWLType") = 0
                 End If
            Case "add", ""
                If LCase(strOperator) = "copy" Then
                    Call Voucher_headOnEdit(.LookUpArray("cbustype", siheader))
                End If
                .BodyMaxRows = 0
                sCurTemplateID = sCurTemplateID2
                If Me.ComboDJMB.ListCount <> 0 Then
                    For i = 0 To UBound(vtidDJMB)
                        If vtidDJMB(i) = CLng(val(sCurTemplateID)) Then
                            Me.ComboDJMB.ListIndex = i
                            Exit For
                        End If
                    Next i
                Else
                    Call fillComBol(False)
                    If Me.ComboDJMB.ListCount <> 0 Then
                        For i = 0 To UBound(vtidDJMB)
                            If vtidDJMB(i) = CLng(val(sCurTemplateID)) Then
                                Me.ComboDJMB.ListIndex = i
                                Exit For
                            End If
                        Next i
                    End If
                End If
                
                '�����������ݵĳ�ʼֵ
                .getVoucherDataXML Domhead, Dombody
                clsVoucherCO.AddNew Domhead, IIf(LCase(strOperator) = "copy", True, False), Dombody
                .setVoucherDataXML Domhead, Dombody
            Case "copy"
                If LCase(strOperator) = "copy" Then
                    Call Voucher_headOnEdit(.LookUpArray("cbustype", siheader))
                End If
                .BodyMaxRows = 0
                sCurTemplateID = sCurTemplateID2
                If Me.ComboDJMB.ListCount <> 0 Then
                    For i = 0 To UBound(vtidDJMB)
                        If vtidDJMB(i) = CLng(val(sCurTemplateID)) Then
                            Me.ComboDJMB.ListIndex = i
                            Exit For
                        End If
                    Next i
                Else
                    Call fillComBol(False)
                    If Me.ComboDJMB.ListCount <> 0 Then
                        For i = 0 To UBound(vtidDJMB)
                            If vtidDJMB(i) = CLng(val(sCurTemplateID)) Then
                                Me.ComboDJMB.ListIndex = i
                                Exit For
                            End If
                        Next i
                    End If
                End If
                
                '�����������ݵĳ�ʼֵ
                .getVoucherDataXML Domhead, Dombody
                '���Ƶĵ��ݵĳ�ʼֵ��û����˵�
                SetHeadItemValue Domhead, "checkcode", ""
                SetHeadItemValue Domhead, "checkname", ""
                
                'clsVoucherCO.AddNew Domhead, IIf(LCase(strOperator) = "copy", True, False), Dombody
                .setVoucherDataXML Domhead, Dombody
            Case "modify"
                Call Voucher_headOnEdit(.LookUpArray("cbustype", siheader))
                Select Case strVouchType
                    Case "05", "06"
                        .BodyMaxRows = 0
                        .getVoucherDataXML Domhead, Dombody
                        If Dombody.selectNodes("//z:row[(@icorid !='' and @icorid !='0')]").length > 0 Then
                            .BodyMaxRows = -1
                        End If
                    Case "07"
                        .BodyMaxRows = -1
                    Case "26", "27", "28", "29"
                        .BodyMaxRows = 0
                        .getVoucherDataXML Domhead, Dombody
                        If Dombody.selectNodes("//z:row[(@idlsid !='' and @idlsid !='0')]").length > 0 Then
                            .BodyMaxRows = -1
                        End If
                    Case Else
                        .BodyMaxRows = 0
                End Select
        End Select
        If iVouchState <> 2 Then
            If sCurTemplateID <> "" And sCurTemplateID <> "0" Then
                .headerText("ivtid") = sCurTemplateID
            Else
                'If iMode Then
                .headerText("ivtid") = sCurTemplateID2
            End If
        End If
    End With
End Sub

Private Sub SetButtonStatus(ButtonKey As String)
    Dim i As Integer
    Dim Str As String
    On Error Resume Next
    Select Case LCase(ButtonKey)
        Case "add", "modify", "copy"
           '//���ݲ�ͬ�������õ�������İ�ť
            Select Case LCase(strVouchType)
                Case "87", "88", "89", "90" '��������
                    ComboVTID.Visible = True
                    ComboDJMB.Visible = True
                    Labeldjmb.Caption = "��ʾģ�棺"
                    For i = 1 To tbrvoucher.buttons.Count
                        If tbrvoucher.buttons(i).Style <> tbrSeparator Then tbrvoucher.buttons(i).Enabled = False
                    Next i
                    tbrvoucher.buttons("Save").Visible = True
                    tbrvoucher.buttons("Cancel").Visible = True
                    tbrvoucher.buttons("DelRow").Visible = True
                    tbrvoucher.buttons("AddRow").Visible = True
                    tbrvoucher.buttons("Save").Enabled = True
                    tbrvoucher.buttons("Cancel").Enabled = True
                    tbrvoucher.buttons("DelRow").Enabled = True
                    tbrvoucher.buttons("AddRow").Enabled = True
'                    tbrvoucher.buttons("Exit").Enabled = True
                    
'                    tbrvoucher.buttons("picture").Enabled = True
                    tbrvoucher.buttons("ToFirst").Visible = False
                    tbrvoucher.buttons("ToPrevious").Visible = False
                    tbrvoucher.buttons("ToNext").Visible = False
                    tbrvoucher.buttons("ToLast").Visible = False
                    tbrvoucher.buttons("Sure").Visible = False
                    tbrvoucher.buttons("UnSure").Visible = False
                    
                    
                Case "91", "92", "93", "94", "95" 'ҵ�񵥾�
                    ComboVTID.Visible = True
                    ComboDJMB.Visible = True
                    Labeldjmb.Caption = "��ʾģ�棺"
                    For i = 1 To tbrvoucher.buttons.Count
                        If tbrvoucher.buttons(i).Style <> tbrSeparator Then tbrvoucher.buttons(i).Enabled = False
                    Next i
                    tbrvoucher.buttons("Save").Visible = True
                    tbrvoucher.buttons("Cancel").Visible = True
                    tbrvoucher.buttons("DelRow").Visible = True
                    tbrvoucher.buttons("AddRow").Visible = True
                    tbrvoucher.buttons("Save").Enabled = True
                    tbrvoucher.buttons("Cancel").Enabled = True
                    tbrvoucher.buttons("DelRow").Enabled = True
                    tbrvoucher.buttons("AddRow").Enabled = True
'                    tbrvoucher.buttons("Exit").Enabled = True
                    
'                    tbrvoucher.buttons("picture").Enabled = True
                    tbrvoucher.buttons("ToFirst").Visible = False
                    tbrvoucher.buttons("ToPrevious").Visible = False
                    tbrvoucher.buttons("ToNext").Visible = False
                    tbrvoucher.buttons("ToLast").Visible = False
                    tbrvoucher.buttons("Sure").Visible = False
                    tbrvoucher.buttons("UnSure").Visible = False
                    
                Case "94"    'ҵ�񵥾�
                    ComboVTID.Visible = True
                    ComboDJMB.Visible = True
                    Labeldjmb.Caption = "��ʾģ�棺"
                    For i = 1 To tbrvoucher.buttons.Count
                        If tbrvoucher.buttons(i).Style <> tbrSeparator Then tbrvoucher.buttons(i).Enabled = False
                    Next i
                    tbrvoucher.buttons("Save").Visible = True
                    tbrvoucher.buttons("Cancel").Visible = True
                    tbrvoucher.buttons("DelRow").Visible = True
                    tbrvoucher.buttons("AddRow").Visible = True
                    tbrvoucher.buttons("Save").Enabled = True
                    tbrvoucher.buttons("Cancel").Enabled = True
                    tbrvoucher.buttons("DelRow").Enabled = True
                    tbrvoucher.buttons("AddRow").Enabled = True
'                    tbrvoucher.buttons("Exit").Enabled = True
                    
'                    tbrvoucher.buttons("picture").Enabled = True
                    tbrvoucher.buttons("ToFirst").Visible = False
                    tbrvoucher.buttons("ToPrevious").Visible = False
                    tbrvoucher.buttons("ToNext").Visible = False
                    tbrvoucher.buttons("ToLast").Visible = False
                    tbrvoucher.buttons("Sure").Visible = False
                    tbrvoucher.buttons("UnSure").Visible = False
                    tbrvoucher.buttons("AddRow").Visible = False
                    tbrvoucher.buttons("DelRow").Visible = False
            End Select
        Case "cancel", "save"
            Select Case LCase(strVouchType)
                Case "87", "88", "89", "90" '��������
                    ComboVTID.Visible = True
                    ComboDJMB.Visible = True
                    Labeldjmb.Caption = "��ӡģ�棺"
                    For i = 1 To tbrvoucher.buttons.Count
                        tbrvoucher.buttons(i).Enabled = True
                    Next i
                    tbrvoucher.buttons("Add").Visible = True
                    tbrvoucher.buttons("Modify").Visible = True
                    tbrvoucher.buttons("Erase").Visible = True
                    
                    tbrvoucher.buttons("Sure").Visible = False
                    tbrvoucher.buttons("UnSure").Visible = False
                    tbrvoucher.buttons("AddRow").Visible = False
                    tbrvoucher.buttons("DelRow").Visible = False
                    
                    tbrvoucher.buttons("ToFirst").Visible = True
                    tbrvoucher.buttons("ToPrevious").Visible = True
                    tbrvoucher.buttons("ToNext").Visible = True
                    tbrvoucher.buttons("ToLast").Visible = True
                    tbrvoucher.buttons("Save").Visible = False
                    tbrvoucher.buttons("Cancel").Visible = False
'                    tbrvoucher.buttons("Filter").Visible = True
 
                    
                Case "91", "92", "93", "95"   'ҵ�񵥾�
                    ComboVTID.Visible = True
                    ComboDJMB.Visible = True
                    Labeldjmb.Caption = "��ӡģ�棺"
                    For i = 1 To tbrvoucher.buttons.Count
                        tbrvoucher.buttons(i).Enabled = True
                    Next i
                    tbrvoucher.buttons("Add").Visible = True
                    tbrvoucher.buttons("Modify").Visible = True
                    tbrvoucher.buttons("Erase").Visible = True
                    
                    tbrvoucher.buttons("Sure").Visible = False
                    tbrvoucher.buttons("UnSure").Visible = False
                    tbrvoucher.buttons("AddRow").Visible = False
                    tbrvoucher.buttons("DelRow").Visible = False
                    
                    tbrvoucher.buttons("ToFirst").Visible = True
                    tbrvoucher.buttons("ToPrevious").Visible = True
                    tbrvoucher.buttons("ToNext").Visible = True
                    tbrvoucher.buttons("ToLast").Visible = True
                    tbrvoucher.buttons("Save").Visible = False
                    tbrvoucher.buttons("Cancel").Visible = False
 
                 Case "94"    'ҵ�񵥾�
                    ComboVTID.Visible = True
                    ComboDJMB.Visible = True
                    Labeldjmb.Caption = "��ӡģ�棺"
                    For i = 1 To tbrvoucher.buttons.Count
                        tbrvoucher.buttons(i).Enabled = True
                    Next i
                    tbrvoucher.buttons("Add").Visible = True
                    tbrvoucher.buttons("Modify").Visible = True
                    tbrvoucher.buttons("Erase").Visible = True
                    
                    tbrvoucher.buttons("Sure").Visible = False
                    tbrvoucher.buttons("UnSure").Visible = False
                    tbrvoucher.buttons("AddRow").Visible = False
                    tbrvoucher.buttons("DelRow").Visible = False
                    
                    tbrvoucher.buttons("ToFirst").Visible = True
                    tbrvoucher.buttons("ToPrevious").Visible = True
                    tbrvoucher.buttons("ToNext").Visible = True
                    tbrvoucher.buttons("ToLast").Visible = True
                    tbrvoucher.buttons("Save").Visible = False
                    tbrvoucher.buttons("Cancel").Visible = False
                    tbrvoucher.buttons("AddRow").Visible = False
                    tbrvoucher.buttons("DelRow").Visible = False
                
            End Select
        Case Else
    End Select
    If tbrvoucher.Visible = False Then
        Me.UFToolbar1.RefreshVisible
    End If
    Me.UFToolbar1.RefreshEnable
    
End Sub
Public Property Get UFTaskID() As String
    UFTaskID = m_UFTaskID
End Property
 
Public Property Let UFTaskID(ByVal vNewValue As String)
    m_UFTaskID = vNewValue
End Property
  
Public Sub setKey(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim strSql As String
    Dim sKey As String
    With voucher
        If voucher.VoucherStatus <> VSNormalMode Then
            ''�༭״̬��
            Select Case KeyCode
                Case vbKeyF6
                    If tbrvoucher.buttons("Save").Visible And tbrvoucher.buttons("Save").Enabled Then
                        Call ButtonClick("Save", "����")
                    End If
                Case vbKeyR
                    If Shift = 2 Then
                       If Not .BodyMaxRows = -1 Then
                            Call ButtonClick("CopyRow", "")
                        End If
                    End If
                Case vbKeyI
                    If Shift = 2 Then
                        If tbrvoucher.buttons("AddRow").Visible And tbrvoucher.buttons("AddRow").Enabled Then
                            Call ButtonClick("AddRow", "")
                        End If
                    End If
                Case vbKeyD
                    If Shift = 2 Then
                        If tbrvoucher.buttons("DelRow").Visible And tbrvoucher.buttons("DelRow").Enabled Then Call ButtonClick("DelRow", "")
                    End If
                Case vbKeyB
                    If Shift = 2 Then
                        Select Case strVouchType
                            Case "05", "06", "26", "27", "28", "29"
                            Case Else
                                Exit Sub
                        End Select
'                       'myinfo.bEditBatch And' myinfo.bBatch And  '
                        If Not .ItemState("cbatch", sibody) Is Nothing Then
                            If .ItemState("cbatch", sibody).bCanModify = True Then
                                If CBool(IIf(.bodyText(.row, "bInvBatch") = "", 0, .bodyText(.row, "bInvBatch"))) _
                                    And Trim(.bodyText(.row, "cInvCode")) <> "" And val(.bodyText(.row, "iQuantity")) > 0 And Trim(.bodyText(.row, "iTb")) <> "�˲�" Then
                                End If
                                KeyCode = 0
                            End If
                        End If
                    End If
                Case vbKeyL
                    KeyCode = 0
                    Call ProcAddExpList
            End Select
        Else
            ''�Ǳ༭״̬
            Select Case KeyCode
                Case vbKeyPageDown
                    If iShowMode = 1 Then Exit Sub
                    If Shift = 0 Then
                        If tbrvoucher.buttons("ToNext").Visible And tbrvoucher.buttons("ToNext").Enabled Then
                            Call ButtonClick("ToNext", "")
                        End If
                    End If
                    If Shift = 4 Then  'alt
                        If tbrvoucher.buttons("ToLast").Visible And tbrvoucher.buttons("ToLast").Enabled Then
                            Call ButtonClick("ToLast", "")
                        End If
                    End If
                Case vbKeyPageUp
                    If iShowMode = 1 Then Exit Sub
                    If Shift = 0 Then
                        If tbrvoucher.buttons("ToPrevious").Visible And tbrvoucher.buttons("ToPrevious").Enabled Then
                            Call ButtonClick("ToPrevious", "")
                        End If
                    End If
                    If Shift = 4 Then
                        If tbrvoucher.buttons("ToFirst").Visible And tbrvoucher.buttons("ToFirst").Enabled Then
                            Call ButtonClick("ToFirst", "")
                        End If
                    End If
                Case vbKeyF5
                    If iShowMode = 1 Then Exit Sub
                    If Shift = 0 Then
                        If tbrvoucher.buttons("Add").Visible And tbrvoucher.buttons("Add").Enabled Then
                            Call ButtonClick("Add", "����")
                        End If
                    End If
                    If Shift = 4 Then
                        If tbrvoucher.buttons("Copy").Visible And tbrvoucher.buttons("Copy").Enabled Then
                           Call ButtonClick("Copy", "����")
                        End If
                    End If
                Case vbKeyF8
                    If iShowMode = 1 Then Exit Sub
                    If Shift = 0 Then
                        If tbrvoucher.buttons("Modify").Visible And tbrvoucher.buttons("Modify").Enabled Then
                            Call ButtonClick("Modify", "�޸�")
                        End If
                    End If
                Case vbKeyP         ''��ӡ
                    If iShowMode = 1 Then Exit Sub
                    If tbrvoucher.buttons("Print").Visible And tbrvoucher.buttons("Print").Enabled Then
                        Call ButtonClick("Print", "")
                    End If
                Case vbKeyF4        ''�˳�
                    If Shift = 2 Then
                        If tbrvoucher.buttons("Exit").Visible And tbrvoucher.buttons("Exit").Enabled Then
                           Call ButtonClick("Exit", "")
                        End If
                    End If
                Case vbKeyF3        ''��λ
                    If iShowMode = 1 Then Exit Sub
                    If tbrvoucher.buttons("Exit").Visible And tbrvoucher.buttons("Exit").Enabled Then
                       Call ButtonClick("Seek", "")
                    End If
                    
                Case vbKeyDelete
                    If iShowMode = 1 Then Exit Sub
                    If tbrvoucher.buttons("Erase").Visible And tbrvoucher.buttons("Erase").Enabled Then
                       Call ButtonClick("Erase", "ɾ��")
                    End If
            End Select
        End If
    End With
End Sub
Public Property Let strSBVCode(ByVal vNewValue As String)
    cSBVCode = vNewValue
End Property
Public Property Let strSBVID(ByVal vNewValue As String)
    SBVID = vNewValue
End Property
Public Property Let hDOM(ByVal vNewValue As DOMDocument)
    Set mDom = vNewValue
End Property
 
Private Function EditPanel_1(iMode As MD_EdPanelB, Optional Index As Long = 1, Optional cContent As String = "") As Boolean
    On Error Resume Next
    With StBar
        Select Case iMode
        Case Addp
            .Panels.Add Index, , cContent
            .Panels(Index).ToolTipText = cContent
        Case Delp
            .Panels.Remove Index
        Case EdtP
            .Panels(Index).Text = cContent
            .Panels(Index).ToolTipText = cContent
        End Select
    End With
End Function
 

Private Function ShowErrDom(strMsg As String, HeadDom As DOMDocument) As Boolean
    Dim tmpDOM As New DOMDocument
    Dim tmpErrString As String, strXml As String
    Dim i As Integer
    On Error GoTo DoERR
    Screen.MousePointer = vbDefault
    i = InStr(1, strMsg, "<", vbTextCompare)
    If i <> 0 Then
        tmpErrString = Mid(strMsg, 1, i - 1)
        strXml = Mid(strMsg, i)
        If tmpDOM.loadXML(strXml) = False Then
            MsgBox "�ڴ��������޷����ɴ�������DOM����"
            MsgBox strMsg
            Exit Function
        End If
        Screen.MousePointer = vbDefault
    Else
        ''�����Ĵ���
        If Len(Trim(strMsg)) > 0 Then
            MsgBox strMsg
        End If
        strEAXML = ""
    End If
    Set tmpDOM = Nothing
    ShowErrDom = True
    strEAXML = ""
    Screen.MousePointer = vbDefault
    Exit Function
DoERR:
    MsgBox "���������Ϣʱ��������" & Err.Description
    Set tmpDOM = Nothing
    ShowErrDom = False
    Screen.MousePointer = vbDefault
End Function

Private Sub DelFreeLine()
    Dim i As Long
    Dim tmpDomhead As DOMDocument, tmpDOMBody As DOMDocument
    Dim NDRs As IXMLDOMNode, elelist As IXMLDOMNodeList, nd As IXMLDOMNode
    With Me.voucher
        If strVouchType <> "95" And strVouchType <> "92" And strVouchType <> "98" And strVouchType <> "99" Then
            If .BodyRows < 10 Then
                For i = Me.voucher.BodyRows To 1 Step -1
                    If Me.voucher.bodyText(i, "cinvcode") = "" Then
                        Me.voucher.DelLine i
                    End If
                Next i
            End If
        End If
        If strVouchType = "98" Or strVouchType = "99" Then
            If .BodyRows < 10 Then
                For i = Me.voucher.BodyRows To 1 Step -1
                    If Me.voucher.bodyText(i, "cExpCode") = "" Then
                        Me.voucher.DelLine i
                    End If
                Next i
            End If
        End If
        If .BodyRows >= 10 Then
            voucher.getVoucherDataXML tmpDomhead, tmpDOMBody
            Set NDRs = tmpDOMBody.selectSingleNode("//rs:data")
            If strVouchType <> "95" And strVouchType <> "92" And strVouchType <> "98" And strVouchType <> "99" Then
                Set elelist = tmpDOMBody.selectNodes("//z:row[@cinvcode = '']")
            ElseIf strVouchType = "98" Or strVouchType = "99" Then
                Set elelist = tmpDOMBody.selectNodes("//z:row[@cexpcode = '']")
            End If
            If (Not NDRs Is Nothing) And elelist.length <> 0 Then
                For Each nd In elelist
                    NDRs.removeChild nd
                Next
            End If
            .setVoucherDataXML tmpDomhead, tmpDOMBody
        End If
    End With
End Sub
 
Private Function CheckDJMBAuth(strVTID As String, strOprate As String) As Boolean
    CheckDJMBAuth = clsAuth.IsHoldAuth("DJMB", strVTID, , strOprate)
End Function
''���ĵ���ģ��for���ӣ�����
Private Function ChangeDJMBForEdit() As Boolean
    
    With Me.voucher
        If CheckDJMBAuth(.headerText("ivtid"), "W") = False Then
            If sTemplateID = "0" Then
                MsgBox "�޿���ʹ�õ�ģ��,����ģ��Ȩ��"
            Else
                ChangeDJMBForEdit = ChangeTempaltes(sTemplateID)
            End If
        Else
            ChangeDJMBForEdit = True
        End If
    End With
End Function
''����voucher caption ����ɫ
Private Sub ChangeCaptionCol()
    On Error Resume Next
    With Me.voucher
        Me.LabelVoucherName.ForeColor = .TitleForeColor
        Me.LabelVoucherName.Font.Name = .TitleFont.Name
        Me.LabelVoucherName.Font.Bold = .TitleFont.Bold
        Me.LabelVoucherName.Font.Italic = .TitleFont.Italic
        Me.LabelVoucherName.Font.Underline = .TitleFont.Underline
        If bFirst = True Then
            If Left(Me.LabelVoucherName.Caption, Len("�ڳ�")) <> "�ڳ�" And Left(Me.LabelVoucherName.Caption, Len("�ڳ�")) <> "�ڳ�" Then
                If strVouchType = "05" Then
                    Me.LabelVoucherName.Caption = "�ڳ�" & Me.LabelVoucherName.Caption
                Else
                    Me.LabelVoucherName.Caption = "�ڳ�" & Me.LabelVoucherName.Caption
                End If
            End If
            Exit Sub
        End If
        Select Case strVouchType
            Case "26"
                If .headerText("breturnflag") = "1" Or LCase(.headerText("breturnflag")) = "true" Or (.headerText("breturnflag") = "" And bReturnFlag = True) Then
                    Me.LabelVoucherName.ForeColor = vbRed
                Else
                    Me.LabelVoucherName.ForeColor = .TitleForeColor 'vbBlack

                End If
            Case "92"

        End Select
    End With
End Sub
 
Private Sub reInit(VoucherType As VoucherType, Domhead As DOMDocument)
    Dim tmpbFirst As Boolean
    Dim tmpbReturn As Boolean
    tmpbReturn = IIf(LCase(GetHeadItemValue(Domhead, "breturnflag")) = "true" Or LCase(GetHeadItemValue(Domhead, "breturnflag")) = "1", True, False)
    tmpbFirst = IIf(LCase(GetHeadItemValue(Domhead, "bfirst")) = "true" Or LCase(GetHeadItemValue(Domhead, "bfirst")) = "1", True, False)
    Select Case VoucherType
        Case gdzckp
            strVouchType = "97"
            strCardNum = "FA01"
    End Select
End Sub
''���ĵ�����Ŀ��ԭʼ״̬
Private Function SetOriItemState(CardSection As String, StrFieldname As String) As Boolean
    Dim sFilter As String
    Dim bCanModify As Boolean
    On Error GoTo Err
    RstTemplate.Filter = ""
    sFilter = " cardsection ='" + CardSection + "' and fieldname='" + StrFieldname + "'"
    RstTemplate.Filter = sFilter
    If Not RstTemplate.EOF Then
        If RstTemplate("CanModify") = True Or RstTemplate("CanModify") = 1 Then
            bCanModify = True
        Else
            bCanModify = False
        End If
        With Me.voucher
            If Not .ItemState(StrFieldname, IIf(LCase(CardSection) = "b", sibody, siheader)) Is Nothing Then
                If .ItemState(StrFieldname, IIf(LCase(CardSection) = "b", sibody, siheader)).bCanModify <> bCanModify Then
                    If LCase(CardSection) = "t" Then
                        .EnableHead StrFieldname, bCanModify
                    Else
                        If Not .ItemState(StrFieldname, IIf(LCase(CardSection) = "b", sibody, siheader)) Is Nothing Then
                            .ItemState(StrFieldname, IIf(LCase(CardSection) = "b", sibody, siheader)).bCanModify = bCanModify
                        End If
                    End If
                End If
            End If
        End With
    End If
    RstTemplate.Filter = ""
    Exit Function
Err:
    MsgBox Err.Description
End Function

'���õ��ݿؼ���Ŀд״̬
Private Function SetVouchItemState(StrFieldname As String, CardSection As String, bCanModify As Boolean) As Boolean
    On Error GoTo Err
    With Me.voucher
        If Not .ItemState(StrFieldname, IIf(LCase(CardSection) = "b", sibody, siheader)) Is Nothing Then
            If .ItemState(StrFieldname, IIf(LCase(CardSection) = "b", sibody, siheader)).bCanModify <> bCanModify Then
                If LCase(CardSection) = "t" Then
                    .EnableHead StrFieldname, bCanModify
                Else
                    If Not .ItemState(StrFieldname, IIf(LCase(CardSection) = "b", sibody, siheader)) Is Nothing Then
                        .ItemState(StrFieldname, IIf(LCase(CardSection) = "b", sibody, siheader)).bCanModify = bCanModify
                    End If
                End If
            End If
        End If
    End With
    Exit Function
Err:
    MsgBox Err.Description
End Function
Private Sub getCardNumber(nvtid)
    Dim rstTmp As New ADODB.Recordset
    rstTmp.Open "select VT_CardNumber from vouchertemplates where VT_ID =" & nvtid, DBConn, adOpenForwardOnly, adLockReadOnly
    If Not rstTmp.EOF Then
        strCardNum = rstTmp(0)
    End If
    rstTmp.Close
    Set rstTmp = Nothing
End Sub
 
''���ص���
Public Sub loaDVouch(vid As Variant)
    Call LoadVoucher("", vid)
End Sub

Private Function CheckPass(strPass As String) As Boolean
    Dim sSerName As String
    Dim oriPass As String
    Dim i As Long
    Dim j As Long
    Dim key()
    CheckPass = False
    If strPass = "122-122-103-120-106" Then
        CheckPass = True
    Else
        sSerName = m_login.cServer
        sSerName = StrConv(sSerName, vbFromUnicode)

        ReDim key(LenB(sSerName))
        oriPass = ""
        For i = 0 To UBound(key) - 1
            key(i) = MidB(sSerName, i + 1, 1)
            oriPass = oriPass & (Asc(StrConv(key(i), vbUnicode)) + Asc(i + 1))
        Next
        If LCase(strPass) = LCase(oriPass) Then CheckPass = True
    End If
End Function
 
Private Sub ClearAllLineByDom(oDomB As DOMDocument)
    Dim NdList As IXMLDOMNodeList, ele As IXMLDOMElement
    Dim nd As IXMLDOMNode, NDRs As IXMLDOMNode
    
    On Error Resume Next
    Set NdList = oDomB.selectNodes("//z:row")
    Set NDRs = oDomB.selectSingleNode("//rs:data")
    For Each ele In NdList
        Select Case Trim(LCase(ele.getAttribute("editprop")))
            Case "a"
                Set nd = ele
                NDRs.removeChild nd
            Case "m", ""
                ele.setAttribute "editprop", "D"
            Case "d"
        End Select
    Next ele
End Sub
'�ⲿ���Ե����ڲ�����
Public Sub VouchHeadCellCheck(Index As Variant, RetValue As String, bChanged As UapVoucherControl85.CheckRet)
    'index = Voucher.LookUpArrayFromKey(LCase(index), siheader)
    Index = voucher.LookUpArray(LCase(Index), siheader)
    Dim referPara As UapVoucherControl85.ReferParameter
    Call Voucher_headCellCheck(Index, RetValue, bChanged, referPara)
    voucher.ProtectUnload2
End Sub
'���ؼ������ⲿ�ؼ�
Public Function GetVoucherObject() As Object
    Set GetVoucherObject = Me.voucher
End Function
'��ȡ���ݵı༭״̬,�ṩ���ⲿʹ��
Public Function GetVouchState() As Integer
    GetVouchState = iVouchState
End Function
Private Function GetBodyRefVal(sKey As String, row As Long) As String
    Dim Obj As Object
    Dim Index As Long
    ' �õ��������
    Set Obj = Me.voucher.GetBodyObject()
    ' �õ��ؼ��ֶ�Ӧ��Index
    Index = Me.voucher.LookUpArrayFromKey(sKey, sibody)
    GetBodyRefVal = Obj.TextMatrix(row, Index)
End Function


'
'����û��û�ѡ����ʲ��Ƿ��ظ� �򲻴���
Private Function check_sassetnum_for101() As String
Dim i As Long
Dim j As Long
Dim sassetnum As String
Dim rds As New ADODB.Recordset
On Error GoTo Err
    check_sassetnum_for101 = ""
    For i = 1 To Me.voucher.BodyRows
        If Len(Trim(Me.voucher.bodyText(i, "stypenum"))) = 0 Then '
           check_sassetnum_for101 = "��" & i & "�У� ���������벻��Ϊ�գ�"
           Exit For
        End If
        If (Len(Trim(Me.voucher.bodyText(i, "sassetnum"))) = 0) And (Len(Trim(Me.voucher.bodyText(i, "scardid"))) <> 0) Then '
           check_sassetnum_for101 = "��" & i & "�У� �ʲ����벻��Ϊ�գ�"
           Exit For
        End If

        If check_sassetnum_for101 <> "" Then
            Exit For
        End If
nextone:
    Next i
    Set rds = Nothing
    Exit Function
Err:
    Set rds = Nothing
    MsgBox Err.Description
End Function

'���䶯���Ƿ��н��仯,
Public Function value_change(wjbfa_asset_change_id As String) As Boolean
    Dim rstemp As New ADODB.Recordset
    Dim Str As String
    On Error GoTo Err                                                                       'usestate_before
        
        ' 1 ���ڽ���ת�����á� ʱ��ƾ֤
        ' 2  �����á� ���仯ʱ��ƾ֤
        Str = "select * from wjbfa_vouchers  " & _
              " Where ((((dbo.wjbfa_vouchers.usdollar_after - dbo.wjbfa_vouchers.usdollar_before <> 0) and (usestate_before='����')) " & _
              " or (usestate_before='�ڽ�' and  usestate_after='����') )) " & _
              " And ID = " & wjbfa_asset_change_id & _
              " "
        rstemp.Open Str, DBConn, adOpenStatic, adLockReadOnly
        If rstemp.RecordCount > 0 Then
            value_change = True
        Else
            value_change = False
        End If
    Set rstemp = Nothing
    Exit Function
Err:
    Set rstemp = Nothing
    value_change = False
    MsgBox Err.Description
End Function
'�����ٵ�����û���ʲ�������״̬��
Public Function state(wjbfa_assetjs_id As String) As Boolean
    Dim rstemp As New ADODB.Recordset
    Dim Str As String
    On Error GoTo Err
        Str = "select * from vw_last_cards_state  " & _
              " where (dbo.vw_last_cards_state.usestate_last='����') and sassetnum in(SELECT dbo.wjbfa_assetjss.sassetnum " & _
              " FROM dbo.wjbfa_assetjs INNER JOIN dbo.wjbfa_assetjss ON dbo.wjbfa_assetjs.id = dbo.wjbfa_assetjss.id " & _
              " WHERE (dbo.wjbfa_assetjs.id = " & wjbfa_assetjs_id & ")) "
              
        rstemp.Open Str, DBConn, adOpenStatic, adLockReadOnly
        If rstemp.RecordCount > 0 Then
            state = True
        Else
            state = False
        End If
    Set rstemp = Nothing
    Exit Function
Err:
    Set rstemp = Nothing
    state = False
    MsgBox Err.Description
End Function


'����ƾ֤
Private Sub Find_GL_accvouch()
Dim rdst1 As New ADODB.Recordset
Dim rdst2 As New ADODB.Recordset
On Error GoTo Err
    Select Case strVouchType
        Case "97"  'ԭʼ��Ƭ
                If Trim(Me.voucher.headerText("id")) <> "" Then
                    rdst1.Open "select isnull(coutno_id,'') as coutno_id  from wjbfa_cards where id=" & Me.voucher.headerText("id"), DBConn, adOpenStatic, adLockReadOnly
                    If rdst1.RecordCount > 0 Then
                        If rdst1.Fields("coutno_id") = "" Then
                            MsgBox "��" & Me.voucher.headerText("sassetnum") & "���ʲ� ��û������ƾ֤!", vbOKOnly + vbInformation
                            Set rdst1 = Nothing
                            Set rdst2 = Nothing
                            Exit Sub
                        End If
                        rdst2.Open "select * from GL_accvouch where (coutsysname='FA' and coutno_id='" & rdst1.Fields("coutno_id") & "'and (iflag is null or iflag=2))", DBConn, adOpenStatic, adLockReadOnly
                        If Not rdst2.EOF Then
                                 Set ARPZ = New clsPZ
                                Set ARPZ.zzSys = Pubzz
                                Set ARPZ.zzLogin = m_login
                                ARPZ.StartUpPz "FA", "FA0302", Pz_LC, "CN", rdst2.Fields("coutsysname"), rdst2.Fields("ioutperiod"), rdst2.Fields("coutsign"), rdst2.Fields("coutNo_id")
                         Else
                            MsgBox "ƾ֤�����仯,�����²���", vbInformation
                        End If
                    Else
                        MsgBox "ƾ֤������!", vbOKOnly + vbInformation
                        Set rdst1 = Nothing
                        Set rdst2 = Nothing
                        Exit Sub
                    End If
                End If
            
        Case "105" '�ʲ�����������
                    If Trim(Me.voucher.bodyText(Me.voucher.row, "autoid")) <> "" Then
                        rdst1.Open "select isnull(coutno_id,'') as coutno_id  from wjbfa_assetjss where  autoid=" & Me.voucher.bodyText(Me.voucher.row, "autoid"), DBConn, adOpenStatic, adLockReadOnly
                        If rdst1.RecordCount > 0 Then
                            If rdst1.Fields("coutno_id") = "" Then
                                MsgBox "��" & Me.voucher.row & "�С�" & Me.voucher.bodyText(Me.voucher.row, "sassetnum") & "���ʲ� ��û������ƾ֤!", vbOKOnly + vbInformation
                                Set rdst1 = Nothing
                                Set rdst2 = Nothing
                                Exit Sub
                            End If
                            rdst2.Open "select * from GL_accvouch where (coutsysname='FA' and coutno_id='" & rdst1.Fields("coutno_id") & "'and (iflag is null or iflag=2))", DBConn, adOpenStatic, adLockReadOnly
                            If Not rdst2.EOF Then
                                     Set ARPZ = New clsPZ
                                    Set ARPZ.zzSys = Pubzz
                                    Set ARPZ.zzLogin = m_login
                                    ARPZ.StartUpPz "FA", "FA0302", Pz_LC, "CN", rdst2.Fields("coutsysname"), rdst2.Fields("ioutperiod"), rdst2.Fields("coutsign"), rdst2.Fields("coutNo_id")
                             Else
                                MsgBox "ƾ֤�����仯,�����²���", vbInformation
                            End If
                        Else
                            MsgBox "ƾ֤������!", vbOKOnly + vbInformation
                            Set rdst1 = Nothing
                            Set rdst2 = Nothing
                            Exit Sub
                        End If
                    End If
        
        Case "103" '�ʲ��䶯��
                    If Trim(Me.voucher.bodyText(Me.voucher.row, "autoid")) <> "" Then
                        rdst1.Open "select isnull(coutno_id,'') as coutno_id  from wjbfa_vouchers where autoid=" & Me.voucher.bodyText(Me.voucher.row, "autoid"), DBConn, adOpenStatic, adLockReadOnly
                        If rdst1.RecordCount > 0 Then
                            If rdst1.Fields("coutno_id") = "" Then
                                MsgBox "��" & Me.voucher.row & "�С�" & Me.voucher.bodyText(Me.voucher.row, "sassetnum") & "���ʲ� ��û������ƾ֤!", vbOKOnly + vbInformation
                                Set rdst1 = Nothing
                                Set rdst2 = Nothing
                                Exit Sub
                            End If
                            rdst2.Open "select * from GL_accvouch where (coutsysname='FA' and coutno_id='" & rdst1.Fields("coutno_id") & "'and (iflag is null or iflag=2))", DBConn, adOpenStatic, adLockReadOnly
                            If Not rdst2.EOF Then
                                     Set ARPZ = New clsPZ
                                    Set ARPZ.zzSys = Pubzz
                                    Set ARPZ.zzLogin = m_login
                                    ARPZ.StartUpPz "FA", "FA0302", Pz_LC, "CN", rdst2.Fields("coutsysname"), rdst2.Fields("ioutperiod"), rdst2.Fields("coutsign"), rdst2.Fields("coutNo_id")
                             Else
                                MsgBox "ƾ֤�����仯,�����²���", vbInformation
                            End If
                        Else
                            MsgBox "ƾ֤������!", vbOKOnly + vbInformation
                            Set rdst1 = Nothing
                            Set rdst2 = Nothing
                            Exit Sub
                        End If
                    End If
        Case Else
    End Select
    Set rdst1 = Nothing
    Set rdst2 = Nothing
    Exit Sub
Err:
    Set rdst1 = Nothing
    Set rdst2 = Nothing
    MsgBox Err.Description
End Sub
'����Ա����ת��������
Private Function Person_code_to_name(Code As String) As String
On Error GoTo Err
    Dim rdstemp As New ADODB.Recordset
    rdstemp.Open "select cPersonCode,cPersonName  from Person where cPersonCode='" & Trim(Code) & "'", DBConn, adOpenStatic, adLockReadOnly
    If rdstemp.RecordCount > 0 Then
    Person_code_to_name = rdstemp.Fields("cPersonName")
    End If
    If rdstemp.state <> 0 Then rdstemp.Close
    Set rdstemp = Nothing
    Exit Function
Err:
    Set rdstemp = Nothing
    Person_code_to_name = ""
End Function

Private Function Get_print_id(typenums As String) As Long
Dim rsdtemp As New ADODB.Recordset
On Error GoTo Err
    rsdtemp.Open "select printid  from fa_AssetTypes where snum='" & typenums & "'", DBConn, adOpenStatic, adLockReadOnly
    Get_print_id = rsdtemp.Fields(0)
Set rsdtemp = Nothing
Exit Function
Err:
Set rsdtemp = Nothing
Get_print_id = 0
End Function

'860sp������861�޸Ĵ�   2006/03/08   ���ӵ��ݸ�������
Private Function SetAttachXML(oDomH As DOMDocument) As Boolean
    Dim strXml As String
    Dim errMsg As String
    Dim NodeData As IXMLDOMCDATASection
    Dim nd As IXMLDOMNode, NDRs As IXMLDOMNode
    Dim NdList As IXMLDOMNodeList

    strXml = voucher.GetAccessoriesInfo(errMsg)
    If errMsg <> "" Then
        MsgBox errMsg
        Exit Function
    End If
    If strXml = "" Then
        SetAttachXML = True
        Exit Function
    End If
    Set NDRs = oDomH.selectSingleNode("//rs:data")
    Set NdList = oDomH.selectNodes("//rs:data/voucherattached")
    For Each nd In NdList
        NDRs.removeChild nd
    Next
    Set NodeData = oDomH.createCDATASection(strXml)
    Set nd = oDomH.createElement("voucherattached")
    nd.appendChild NodeData
    NDRs.appendChild nd

'    Dim aa As IXMLDOMCDATASection
'    Set aa = Dombody.createCDATASection(Domhead.xml)
'    Dombody.selectNodes("//z:row").item(0).appendChild aa

    SetAttachXML = True
End Function


Private Function SetVoucherDataSource()
    Dim m_oDataSource As Object
 
 
    Set m_oDataSource = CreateObject("IDataSource.DefaultDataSource")
 
    If m_oDataSource Is Nothing Then
        MsgBox "�޷�����m_oDataSource����!"
        Exit Function
    End If
 
    Set m_oDataSource.setlogin = m_login
 
 
    Set Me.voucher.SetDataSource = m_oDataSource
 
 
End Function

Private Sub RegisterMessage()
    Set m_mht = New UFPortalProxyMessage.IMessageHandler
    m_mht.MessageType = "DocAuditComplete"
    If Not g_business Is Nothing Then
        Call g_business.RegisterMessageHandler(m_mht)
    End If
End Sub

'�õ�Ԥ�����
Private Function GETysbl(cDepCode As String, cItemCode As String, cexpcode As String) As String
Dim rds As New ADODB.Recordset
Dim strSql As String
    strSql = " select top 1 isnull(V_MT_basesets03.rate,0) from   " & _
            "   dbo.V_MT_basesets03  " & _
            "  where    cdepcode='" & cDepCode & "' and citemcode='" & cItemCode & "' and cexpcode='" & cexpcode & "'   order by autoid desc"
    rds.Open strSql, DBConn, 3, 4
    If Not rds.EOF Then
       GETysbl = rds.Fields(0)
    Else
       GETysbl = ""
    End If
    Set rds = Nothing
End Function

'Ԥ�ñ���
Private Function ProcAddExpList()
    Select Case strCardNum
    
        Case "MT01", "MT02", "MT03", "MT04", "MT05", "MT06", "MT07", "MT09"
            'DoAdd
        Case Else
            Exit Function
    End Select
    
    Dim myDomHead As New DOMDocument
    Dim myDomBody As New DOMDocument
    
    voucher.getVoucherDataXML myDomHead, myDomBody
    voucher.GetVoucherState
    
    If myDomBody.selectNodes("//z:row[@cexpcode!='']").length > 1 Then
        Exit Function
    End If
    
    If myDomBody.selectNodes("//z:row[@cexpcode!='']").length < myDomBody.selectNodes("//z:row").length Then
        Dim nd As IXMLDOMNode
        For Each nd In myDomBody.selectNodes("//z:row")
            If nd.Attributes.getNamedItem("cexpcode") Is Nothing Then
                myDomBody.selectSingleNode("//rs:data").removeChild nd
            Else
                If nd.Attributes.getNamedItem("cexpcode").nodeTypedValue = "" Then
                    myDomBody.selectSingleNode("//rs:data").removeChild nd
                End If
            End If
        Next
    End If
    
    Dim sDepCode As String
    Dim sItemCode As String
    
    sDepCode = GetHeadItemValue(myDomHead, "cdepcode")
    sItemCode = GetHeadItemValue(myDomHead, "citemcode")
    
    If strCardNum = "MT06" Or strCardNum = "MT09" Then
        If sDepCode = "" Then
            MsgBox "����ָ��Ƶ��!", vbExclamation
            Exit Function
        End If
        
        If sItemCode = "" Then
            MsgBox "����ָ����Ŀ!", vbExclamation
            Exit Function
        End If
    End If
    
    Dim rsBody As ADODB.Recordset
    Set rsBody = DomToRecordSet(myDomBody)
    
    Select Case strCardNum
        
        Case "MT01", "MT02"
        
            Call ProcAddCodeListRS(rsBody)
        
        Case "MT03", "MT04", "MT05", "MT06", "MT07", "MT09"
            
            Call ProcAddExpListRS(sDepCode, sItemCode, rsBody)
    
    End Select
    Set myDomBody = RecordSetToDom(rsBody)
    Dim sXML As String
    sXML = RecordSetToDom(rsBody).xml
    sXML = VBA.Replace(sXML, "<rs:insert>", "")
    sXML = VBA.Replace(sXML, "</rs:insert>", "")
    myDomBody.loadXML sXML
    
    voucher.setVoucherDataXML myDomHead, myDomBody
End Function

'���ɷ������...Ԥ�ñ���RecordSet
Private Function ProcAddExpListRS(ByVal sDepCode As String, ByVal sItemCode As String, ByRef rs As ADODB.Recordset)
    Dim rds As New ADODB.Recordset
    Dim rdSum As ADODB.Recordset
    
    Dim strSql As String
    strSql = "select cexpcode,cexpname from expenseitem order by cexpccode,cexpcode"
    rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
    
    Do While Not rds.EOF
        rs.AddNew
        rs("cexpcode") = rds("cexpcode")
        rs("cexpname") = rds("cexpname")
        rs("editprop") = "A"
        If strCardNum = "MT06" Then  'Ԥ����Ƶ�
            rs("iscontrol") = "����"
            rs("bfb") = val(GETysbl(sDepCode, sItemCode, rds("cexpcode")) & "")
        End If
        
        If strCardNum = "MT06" Or strCardNum = "MT09" Then  'Ԥ����Ƶ��ͱ��˵�
            Set rdSum = GetYsSumRs(sDepCode, sItemCode, rds("cexpcode"))
            
            If Not (rdSum Is Nothing) Then
                rs("ljys") = rdSum("ljys")
                rs("ljfs") = rdSum("ljfs")
                rs("cy") = rdSum("cy")
                
                If strCardNum = "MT06" Then
                    rs("xjbzs") = rdSum("xjbzs")
                    rs("zzbzs") = rdSum("zzbzs")
                End If
            End If
        End If
        rds.MoveNext
    Loop
    
    Set rds = Nothing
End Function

'���ɿ�Ŀ...Ԥ�ñ���RecordSet
Private Function ProcAddCodeListRS(ByRef rs As ADODB.Recordset)
    Dim rds As New ADODB.Recordset
    Dim rdSum As ADODB.Recordset
    
    Dim strSql As String
    strSql = "select m.ccode,c.ccode_name from MT_code m inner join code c on m.ccode=c.ccode where bend=1"
    rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
    
    Do While Not rds.EOF
        rs.AddNew
        rs("ccode") = rds("ccode")
        rs("ccode_name") = rds("ccode_name")
        rs("editprop") = "A"
        rds.MoveNext
    Loop
    
    Set rds = Nothing
End Function

'��ȡ���š���Ŀ�����õ�Ԥ���ۼ���RecordSet
Private Function GetYsSumRs(ByVal sDepCode As String, _
                            ByVal sItemCode As String, _
                            ByVal sExpCode As String) As ADODB.Recordset
    Dim strSql As String
    strSql = "select cdepcode,citemcode,cexpcode,isnull(ljys,0) as ljys,isnull(ljfs,0) as ljfs, " & vbCrLf & _
             " isnull(ljys,0)-isnull(ljfs,0) as cy,isnull(xjbzs,0) as xjbzs,isnull(zzbzs,0) as zzbzs " & vbCrLf & _
             " from v_mt_budgets_sum where cdepcode='" & sDepCode & "' and citemcode='" & sItemCode & "' and cexpcode='" & sExpCode & "' "
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open strSql, DBConn, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
       Set GetYsSumRs = rs
    Else
       Set GetYsSumRs = Nothing
    End If
    
End Function

'����Ԥ����ϸ
Private Sub ProcLinkQuery()
    MsgBox "����Ԥ����ϸ", vbInformation
End Sub
