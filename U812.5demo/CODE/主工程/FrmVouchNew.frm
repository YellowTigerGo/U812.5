VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{456334B9-D052-4643-8F5F-2326B24BE316}#6.96#0"; "UAPvouchercontrol85.ocx"
Object = "{201FB79B-5556-47A4-AD9C-A46BA0C45A44}#6.42#0"; "UFToolBarCtrl.ocx"
Object = "{005151D4-33F6-471B-B651-222D4E411832}#4.5#0"; "UFFormPartner.ocx"
Object = "{4C2F9AC0-6D40-468A-8389-518BB4F8C67D}#1.0#0"; "UFComboBox.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{D2B3369D-2E6C-45DE-A705-14481242A2BE}#1.12#0"; "UFMenu6U.ocx"
Begin VB.Form frmVouchNew 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "0"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   14115
   FillColor       =   &H00004040&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "FrmVouchNew.frx":0000
   ScaleHeight     =   5730
   ScaleWidth      =   14115
   WindowState     =   2  'Maximized
   Begin UAPVoucherControl85.ctlVoucher Voucher 
      Height          =   3735
      Left            =   1800
      TabIndex        =   9
      Top             =   1440
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   6588
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      TitleForecolor  =   -2147483642
      DisabledColor   =   16777215
      ColAlignment0   =   9
      Rows            =   20
      Cols            =   20
      TitleCaption    =   "单据名称"
      TitleCaption    =   "单据名称"
      TitleForecolor  =   -2147483642
      ControlScrollBars=   0
      ControlAutoScales=   0
      BaseOfVScrollPoint=   0
      ShowSorter      =   0   'False
      ShowFixColer    =   0   'False
   End
   Begin MSComDlg.CommonDialog dlgFileOpen 
      Left            =   960
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picSure 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Enabled         =   0   'False
      FillColor       =   &H000000FF&
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   0
      Picture         =   "FrmVouchNew.frx":3612
      ScaleHeight     =   525
      ScaleWidth      =   1275
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSComctlLib.Toolbar tbrvoucher 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   14115
      _ExtentX        =   24897
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      Begin UFToolBarCtrl.UFToolbar UFToolbar1 
         Height          =   240
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   423
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin UFFormPartner.UFFrmCaption UFFrmCaptionMgr 
      Left            =   3480
      Top             =   840
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "0"
      DebugFlag       =   0   'False
      SkinStyle       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   1200
      ScaleHeight     =   585
      ScaleWidth      =   10395
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   10395
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   7230
         ScaleHeight     =   300
         ScaleWidth      =   3495
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   3495
         Begin UFCOMBOBOXLib.UFComboBox ComboDJMB 
            Height          =   330
            Left            =   930
            TabIndex        =   5
            Top             =   0
            Width           =   2025
            _Version        =   65536
            _ExtentX        =   3572
            _ExtentY        =   582
            _StockProps     =   196
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            Text            =   ""
            Style           =   2
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
         End
         Begin UFCOMBOBOXLib.UFComboBox ComboVTID 
            Height          =   330
            Left            =   930
            TabIndex        =   6
            Top             =   0
            Width           =   2025
            _Version        =   65536
            _ExtentX        =   3572
            _ExtentY        =   582
            _StockProps     =   196
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            Text            =   ""
            Style           =   2
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
         End
         Begin VB.Label Labeldjmb 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00E4C9AF&
            BackStyle       =   0  'Transparent
            Caption         =   "打印模版："
            Height          =   180
            Left            =   60
            TabIndex        =   4
            Top             =   60
            Width           =   900
         End
      End
      Begin VB.Label LabelVoucherName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "单据"
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   2
         Top             =   120
         Width           =   630
      End
   End
   Begin cPopMenu6.PopMenu PopMenu1 
      Left            =   480
      Top             =   1680
      _ExtentX        =   1058
      _ExtentY        =   1058
      HighlightCheckedItems=   0   'False
      TickIconIndex   =   0
      HighlightStyle  =   2
      ActiveMenuForeColor=   -2147483641
      MenuBackgroundColor=   16119285
   End
   Begin VB.Menu mnuPop 
      Caption         =   "右键菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuLinkQuery 
         Caption         =   "联查预算分析表"
      End
   End
End
Attribute VB_Name = "frmVouchNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//警告：以下为API函数,尽量不要改动/////////////////////////////////////////////////////////////////////////////////
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'/////////////////////////////////////////////////////////////////////////////////////////////////////////

'//提醒：以下为窗体视图层所用通用变量申明,请不要随便改动////////////////////////////////////////////////////////
Private sGuid As String               '记载每创建一个窗体的GUID
Public bFrmCancel As Boolean          '记载窗体状态
Public iShowMode As Integer           '记录窗体模式  0：正常 1：浏览 2：变更状态
Private m_UFTaskID As String          '记录当前操作任务号
Private m_FormVisible  As Boolean     '记录窗体显示否
Private m_strHelpId As String         '记录窗体的帮助ID

Private bClickCancel As Boolean       '单据按钮取消操作控制变量
 
Private ButtonTaskID As String        '单据按钮的任务id
Private m_strToolBarName As String    '单据工具栏标识
Private clsTbl As New clsAutoToolBar  '单据工具栏格式化变量

Private m_strVouchType As String      '记录单据类型
Private m_bReturnFlag As Boolean      '记录红蓝单据
Private m_strCardNum As String        '记录单据CardNum
Private m_bFirst As Boolean           '记录是否期初单据
Private bCheckVouch As Boolean        '记录单据的审核状态2
Private bLostFocus As Boolean         '记录单据是否失去焦点

Private sTemplateID As String                '单据默认模板号码
Private sCurTemplateID As String             '单据当前的模板号
Private sCurTemplateID2 As String            '单据当前的模板号
Private preVTID As String                    '单据VTID改变前ID
Private RstTemplate As ADODB.Recordset       '单据模版临时记录集
Private RstTemplate2 As New ADODB.Recordset  '单据模版记录集
Private bfillDjmb As Boolean                 '单据模板ID是否填充标记
Private vtidDJMB() As Long                   '单据显示模板ID数组
Private vtidPrn() As Long                    '单据打印模版ID数组
Private intBodyMaxRows As Integer            '单据表体最大行数

Private iVouchState As Integer        '单据当前的状态(0表示新增、1表示修改、2表示变更等),主要是为提供CO使用。
Private moAutoFill As Object          '单据模板数据自动带入服务对象
Private blnBlank As Boolean           '单据初始界面是否为空单据

Private vNewID As Variant             '单据标识id
Private DomFormat As New DOMDocument  '单据编码规则
Private domHead As New DOMDocument    '单据表头数据
Private domBody As New DOMDocument    '单据表体数据
Private domConfig As New DOMDocument  '单据参数对象

Private WithEvents m_oHelper As Helper  '单据助手对象 草稿
Attribute m_oHelper.VB_VarHelpID = -1
Private m_sCurrentDraftID As String     '草稿
Private bnewDraft As Boolean            '打开草稿时新增处理不参照生单
Private m_MakeVoucherRuleID As String   'UAP生单规则ID

Private WithEvents m_mht As UFPortalProxyMessage.IMessageHandler  '窗体与门户交流通道对象
Attribute m_mht.VB_VarHelpID = -1
Private clsRefer As New UFReferC.UFReferClient                    '旧参照对象
Private clsVoucher As New clsAutoVoucher                          '新单据管理对象
Private clsVoucherRefer As New clsAutoRefer                       '新单据参照管理对象
'///////////////////////////////////////////////////////////////////////////////////////////////////////////

'//提示：以下为业务处理专用全局变量申明，请根据业务需求进行相应改动///////////////////////////////////////////////
Private WithEvents clsVoucherCO As EFVoucherCo.clsVoucherCO    '控制中心对象
Attribute clsVoucherCO.VB_VarHelpID = -1
Private clsVouchModel As New EFVoucherMo.clsVouchLoad     '业务处理中心对象
Private WithEvents ARPZ As ZzPz.clsPZ                                      '凭证对象
Attribute ARPZ.VB_VarHelpID = -1
                              


Public strVoucherUFTS As String  ''参照生单来源单据的时间戳
Public Userdll_UI As New UserDefineDll_UI

Private strUserErr As String '用户插件错误信息
Private UserbSuc As Boolean   '插件执行状态   =true 表示成功  =false 表示失败
Private SA_VoucherListConfigDom As New DOMDocument


'//提醒：注意这个方法不允话删除
'//      门户对与门户融合的窗体的一强制性要求。
'//      Cancel与UnloadMode的参数的含义与QueryUnload的参数相同
Public Sub ExitForm(Cancel As Integer, UnloadMode As Integer)
doNext:
    If Me.Voucher.VoucherStatus <> VSNormalMode Then
        Select Case MsgBox(GetString("U8.SA.xsglsql.01.frmbillvouch.00095"), vbYesNoCancel + vbQuestion) 'zh-CN：是否保存对当前单据的编辑？
            Case vbYes
                ButtonClick "Save", "保存"
                If Me.Voucher.VoucherStatus = VSNormalMode Then
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
        Set UFToolbar1.Business = Nothing

        Set clsVoucherCO = Nothing
        Set clsRefer = Nothing
        Set RstTemplate = Nothing
        Set domHead = Nothing
        Set domBody = Nothing
        Set DomFormat = Nothing
        Set RstTemplate = Nothing
        Set RstTemplate2 = Nothing
        If m_UFTaskID <> "" Then
            m_login.TaskExec m_UFTaskID, 0
        End If
        Unload Me
   End If
End Sub

 

'警告：获取按钮功能ID,并进行权限校验。
'      sKey——操作的按钮名称
Private Function VoucherTask(Skey As String) As Boolean
    Dim strID As String
    strID = clsVoucherCO.GetVoucherTaskID(Skey, strVouchtype, bReturnFlag)
    If strID <> "" Then
        ButtonTaskID = strID
        VoucherTask = LockItem(ButtonTaskID, True, True)
    Else
        VoucherTask = True
    End If
End Function

'警告：释放功能申请
Private Function VoucherFreeTask() As Boolean
    If ButtonTaskID <> "" Then
        VoucherFreeTask = LockItem(ButtonTaskID, False, True)
        ButtonTaskID = ""
    End If
End Function
 
'警告：单据模板选择，请不要对这个函数进行任何改动
Private Function ChangeTempaltes(sNewTemplateID As String, Optional bChangDefalt As Boolean, Optional bCheckAuth As Boolean = True, Optional bFormload As Boolean = False) As Boolean
    Dim RstTemplate2 As New ADODB.Recordset
    Dim rstTmp As New ADODB.Recordset
    Dim strDJAuth As String
    Dim bChanged As Boolean
    Dim tmpDomhead As New DOMDocument
    Dim i As Long

    RstTemplate2.CursorLocation = adUseClient
    
    On Error GoTo DoERR
    bChanged = False
    If sNewTemplateID = "" Or sNewTemplateID = "0" Then
        Exit Function
    End If
    For i = 0 To UBound(vtidDJMB)
        If sNewTemplateID = vtidDJMB(i) Then Exit For
    Next
    If i > UBound(vtidDJMB) Then sNewTemplateID = vtidDJMB(0)
    If bFirst = True Then Call getCardNumber(sNewTemplateID)

    If RstTemplate Is Nothing Then Set RstTemplate = New ADODB.Recordset

    If Trim(sNewTemplateID) = "" Or sNewTemplateID = "0" Then
        If bChangDefalt = True Then
            sNewTemplateID = vtidDJMB(0) ' sTemplateID
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
                GoTo UsePre  ''记录已经取回
            End If
        End If
        Set RstTemplate2 = clsVoucherCO.GetVoucherFormat(sNewTemplateID, strCardNum)
        If RstTemplate2 Is Nothing Then
            If bChangDefalt = True Then
                Set RstTemplate2 = clsVoucherCO.GetVoucherFormat(vtidDJMB(0), strCardNum)
                bChanged = True
            Else
                bChanged = False
            End If
        Else
            If RstTemplate2.State = 1 Then
                If RstTemplate2.EOF And RstTemplate2.BOF Then
                    If bChangDefalt = True Then
                        Set RstTemplate2 = clsVoucherCO.GetVoucherFormat(vtidDJMB(0), strCardNum)
                        sCurTemplateID = vtidDJMB(0)
                        sCurTemplateID2 = vtidDJMB(0)
                        bChanged = True
                    Else
                        bChanged = False
                    End If
                Else
                   bChanged = True
                End If
            Else
                If bChangDefalt = True Then
                    Set RstTemplate2 = clsVoucherCO.GetVoucherFormat(vtidDJMB(0), strCardNum)
                    If RstTemplate2.State = adStateClosed Then
                            MsgBox GetString("U8.SA.xsglsql.01.frmbillvouch.00003"), vbExclamation 'zh-CN：模版设置有问题
                            ChangeTempaltes = False
                            Exit Function
                    End If
                    If Not RstTemplate2 Is Nothing Then
                        If Not RstTemplate2.EOF Then
                            bChanged = True
                        Else
                            MsgBox GetString("U8.SA.xsglsql.01.frmbillvouch.00004"), vbExclamation 'zh-CN：模版设置有问题
                            ChangeTempaltes = False
                            Exit Function
                        End If
                    Else
                        MsgBox GetString("U8.SA.xsglsql.01.frmbillvouch.00005"), vbExclamation 'zh-CN：模版设置有问题
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
            Voucher.Visible = False
        End If
'        Debug.Print Timer
        
        If strVouchtype <> "92" Then   '提示：对于不需要进行单据合计行，请参照此种方法
            Voucher.ShowSummaryView = True
        End If
        '如果是调试状态，不处理附件，以放置弹出‘加载附件失败’窗口
        If clsSAWeb.IsDebug Then RstTemplate2.Fields("vchtblprimarykeynames") = ""
        Voucher.setTemplateData RstTemplate2
'        Debug.Print Timer
        Call Form_Resize
        If Voucher.VoucherStatus <> VSNormalMode Then
            If Voucher.VoucherStatus = VSeAddMode Then
                SetItemState "add"
            End If
            If Voucher.VoucherStatus = VSeEditMode Then
                SetItemState "modify"
            End If
        End If
            
        LabelVoucherName.Caption = Voucher.TitleCaption  '//单据名称标题放入单据头的Label上。
'        Voucher.TitleCaption = ""                        '//单据的名称，清空
    
        If Not DomFormat Is Nothing Then
            If DomFormat.xml <> "" Then
                Me.Voucher.SetBillNumberRule DomFormat.xml
                If Me.Voucher.VoucherStatus <> VSNormalMode Then
                    Call SetVouchNoWriteble
                End If
            End If
        End If
        If RstTemplate.State = 1 Then RstTemplate.Close
        Set RstTemplate = RstTemplate2   '效率修改
        RstTemplate2.MoveFirst
        intBodyMaxRows = RstTemplate2.Fields("vt_bodymaxrows")
        If bFormload = False Then
            Me.Voucher.Visible = True
            Me.Refresh
        End If
        
        bfillDjmb = True
        If iShowMode <> 1 Then
            If Me.ComboDJMB.ListCount > 0 Then
                Me.ComboDJMB.ListIndex = GetDispCobIndex(val(sCurTemplateID))
            Else
                GetString ("U8.SA.xsglsql.01.frmbillvouch.00171") 'zh-CN：您没有模版使用权限
                Screen.MousePointer = vbDefault
                Exit Function
            End If
        End If
        bfillDjmb = False
    End If
    Set rstTmp = Nothing
    ChangeTempaltes = True
    If Not bFormload And bChanged Then Call ChangeCaptionCol
    Exit Function
UsePre:
    sCurTemplateID = sNewTemplateID
    sCurTemplateID2 = sNewTemplateID
    If bFormload = False Then
        Me.Voucher.Visible = False
    End If
    If strVouchtype <> "99" Then
        Voucher.ShowSummaryView = True
    End If
    Voucher.setTemplateData RstTemplate
    Call Form_Resize
    If Voucher.VoucherStatus <> VSNormalMode Then
        SetItemState "modify"
    End If
    If bFormload = False Then
        Me.Voucher.Visible = True
        Me.Refresh
    End If
    Set rstTmp = Nothing
    ChangeTempaltes = True
    Exit Function
DoERR:
    MsgBox Err.Description, vbExclamation
    ChangeTempaltes = False
    Set rstTmp = Nothing
End Function

''警告：加载单据数据,更改模板
Private Sub LoadVoucher(sMove As String, Optional vid As Variant, Optional bRefreshClick As Boolean = False, Optional blnAuth As Boolean = True)
    Dim errMsg As String
    Dim i As Integer
    Dim strXml As String
    Dim UserbSuc As Boolean
    On Error Resume Next
    If Trim(CStr(vid)) = "" Then
        vid = vNewID
    End If
    Select Case LCase(sMove)
        Case ""
            errMsg = clsVoucherCO.GetVoucherData(domHead, domBody, vid, blnAuth)
            
        Case "tonext"
            If val(GetHeadItemValue(domHead, "vt_id")) <> 0 Then
                errMsg = clsVoucherCO.MoveNext(domHead, domBody)
            End If
            
        Case "toprevious"
            If val(GetHeadItemValue(domHead, "vt_id")) = 0 Then
                errMsg = clsVoucherCO.MoveLast(domHead, domBody)
            Else
                errMsg = clsVoucherCO.MovePrevious(domHead, domBody)
            End If
            
        Case "tolast"
ToNext:
            If val(GetHeadItemValue(domHead, "vt_id")) <> 0 Or errMsg <> "" Then
                errMsg = clsVoucherCO.MoveLast(domHead, domBody)
            Else
                errMsg = clsVoucherCO.GetVoucherData(domHead, domBody, vid, blnAuth)
            End If
        
        Case "tofirst"
            errMsg = clsVoucherCO.MoveFirst(domHead, domBody)
        
    End Select
    
    If errMsg <> "" Then
        If bRefreshClick = False And sMove = "" And vid = "" Then
            
        Else
            MsgBox errMsg
        End If
        If i <= 3 Then
            sMove = "tolast"
            GoTo ToNext
        End If
        Exit Sub
    End If
    
    ChangeTempaltes IIf(val(GetHeadItemValue(domHead, "vt_id")) = 0, sCurTemplateID2, GetHeadItemValue(domHead, "vt_id")), , False
    Voucher.Visible = False
    
    Voucher.SkipLoadAccessories = False
    Voucher.setVoucherDataXML domHead, domBody
'    Userdll_UI.LoadAfter_VoucherData errMsg, UserbSuc
        '审批流文本
    Voucher.ExamineFlowAuditInfo = GetEAStream(Me.Voucher, strVouchtype)

     
    
    Call DWINfor
    Call Form_Resize
    Voucher.Visible = True
    
    strXml = "<?xml version='1.0' encoding='GB2312'?>" & Chr(13)
    domConfig.loadXML strXml & "<EAI>0</EAI>"
    SendMessgeToPortal "CurrentDocChanged"
    
End Sub


'提示：设置单据上表头或表体的项目的属性的方法示例
Private Sub SetItemState(Optional sOperate As String)
    Dim i As Long
    Dim iHeadIndex As Integer
    With Me.Voucher
        .BodyMaxRows = 0
        Select Case strVouchtype
            Case "97", "16"
                If strVouchtype = "97" Then
                    If domBody.selectNodes("//z:row[@ccontractid !='']").length > 0 Then
                        .EnableHead "ccusabbname", False
                        .EnableHead "cexch_name", False
                        .EnableHead "cbustype", False
                        If Voucher.headerText("cstname") = "" Then
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
 
'警告：需要改变单据模版
Private Sub ComboDJMB_Click()
    Dim tmpVoucherState As Variant
    ComboDJMB.ToolTipText = ComboDJMB.Text
    If Not bfillDjmb Then
        Me.Voucher.Visible = False
        Me.Voucher.getVoucherDataXML domHead, domBody
        tmpVoucherState = Me.Voucher.VoucherStatus
        Call ChangeTempaltes(Str(vtidDJMB(ComboDJMB.ListIndex)), , False)
        Me.Voucher.VoucherStatus = tmpVoucherState
        'LDX   2009-05-22  Add Beg
'        If strVouchType = 98 Then
            domHead.selectSingleNode("//z:row").Attributes.getNamedItem("vt_id").nodeTypedValue = vtidDJMB(ComboDJMB.ListIndex)
'        End If
        'LDX   2009-05-22  Add End
        Me.Voucher.setVoucherDataXML domHead, domBody
        Me.Voucher.Visible = True
        Me.Voucher.headerText("vt_id") = Str(vtidDJMB(ComboDJMB.ListIndex))
        sCurTemplateID = Str(vtidDJMB(ComboDJMB.ListIndex))
        sCurTemplateID2 = Str(vtidDJMB(ComboDJMB.ListIndex))
    Else
        bfillDjmb = False
    End If
End Sub

'警告
Private Sub ComboVTID_Click()
    ComboVTID.ToolTipText = ComboVTID.Text
End Sub
 
'警告
Private Sub Form_Activate()
    SendMessgeToPortal "DocActivated"
End Sub

'警告
Private Sub Form_Deactivate()
    With Me.Voucher
        If .VoucherStatus <> VSNormalMode Then
            bLostFocus = True
            .ProtectUnload2
            bLostFocus = False
        End If
    End With
    SendMessgeToPortal "DocDeactivated"
End Sub

'警告
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If iShowMode <> 1 Then  ''??
        setKey KeyCode, Shift
    ElseIf KeyCode = vbKeyF4 Then
        setKey KeyCode, Shift
    End If
    
End Sub

Private Sub InitFrm()

    Dim sNumber As String           '单据编号规则字符串
    Dim m_oServer As Object
    Dim m_oVNumber As Object
    Dim m_oDataSource As Object
' 创建单据后台服务对象
    Set m_oServer = CreateObject("UFVoucherServer85.clsVoucherTemplate")
    ' 创建单据编号对象
    Set m_oVNumber = CreateObject("UFBillComponent.clsBillComponent")

    Call m_oVNumber.InitBill(m_login.UfDbName, strCardNum)
    sNumber = m_oVNumber.GetBillFormat
    Voucher.SetBillNumberRule sNumber

    Set m_oDataSource = CreateObject("IDataSource.DefaultDataSource")
    Set m_oDataSource.SetLogin = m_login
    Set Voucher.SetDataSource = m_oDataSource
    Voucher.LoginObj = m_login
    Voucher.InitDataSource
End Sub

'警告
Private Sub Form_Load()
    Dim strErr As String
    On Error Resume Next
    
'    SkinSE_Start hwnd
    InitFrm
'    PopMenu1.AddItem "批改", "BatchModify"
    
    '工具栏获取门信息，或解释成菜单融合功能
    If Not g_business Is Nothing Then
        Set Me.UFToolbar1.Business = g_business
    End If
    Call RegisterMessage
    
    Dim oDicTmp As Object
    Set oDicTmp = CreateObject("Scripting.Dictionary")
    Call Me.UFToolbar1.Settoolbarfromdata(Me.tbrvoucher, DBconn, m_login, strCardNum, strCardNum, oDicTmp)
    Me.UFToolbar1.SetToolbar Me.tbrvoucher
    Me.UFToolbar1.Height = 0
    Me.tbrvoucher.Visible = False
    Set Voucher.OToolbar = Me.UFToolbar1 '上页下页显示位置


    '设置工具栏的相关按钮的图片、文字、描述等属性(可以有其他办法:比如将工具栏按钮的设置放入数据表中)
'    Call SetButton   '用下面新的方法代替
'    clsTbl.initToolBar Me.tbrvoucher, Me.strToolBarName, strErr
'
'    '将微软工具栏与UF工具栏融合
'    ChangeOneFormTbr Me, Me.tbrvoucher, Me.UFToolbar1, strCardNum
'
'    'U872新增将单据控件及窗体传递给UF工具栏
    Call UFToolbar1.SetFormInfo(Me.Voucher, Me)
'
'    '设置工具栏的可见状态
''    SetButtonStatus "Cancel"   '用下面新的方法代替
'    clsTbl.ChangeButtonState Voucher, Me.tbrvoucher, Me.UFToolbar1, Voucher.VoucherStatus
    ChangeButtonState Voucher, Me.tbrvoucher, Me.UFToolbar1, Voucher.VoucherStatus
    
    '单据模板标题及下拉框的显示
    SetCboVtidState
    
    '单据模板设置标题背景、及前景颜色
    Labeldjmb.BackColor = Me.Picture2.BackColor
    Labeldjmb.ForeColor = vbBlack
    
    Picture2.Move 0, Me.tbrvoucher.Height, Me.ScaleWidth
    
    'Picture1里面包含单据模板设置的Lable 及Combox
    Picture1.BackColor = Me.Picture2.BackColor
    Me.Picture1.Move Me.Picture2.Width - Me.Picture1.Width - 5
    
    
    
    '单据设置不可编辑及必需输入颜色
    If lngClr1 <> 0 And lngClr2 <> 0 Then
        Call Voucher.SetRuleColor(lngClr1, lngClr2)
    End If
    Me.Voucher.ControlAutoScales = AutoBoth
    Me.Voucher.ControlScrollBars = ScrollBoth
    Me.Voucher.ShowSorter = True
    
    '窗体设置前景及背景色设置
    Me.BackColor = Me.Voucher.BackColor
    Me.ForeColor = Me.Voucher.BackColor
    
    
    '//创建U8助手对象
    Set m_oHelper = New VoucherHelper.Helper
    If m_oHelper Is Nothing Then
        MsgBox "创建U8助手对象Fail!", vbCritical, "错误"
    End If
    Set m_oHelper.Login = m_login
    
    '//设置窗体的帮助ID
    Me.HelpContextID = val(strHelpId)
    
    '//初始化单据相关参照
    clsVoucherRefer.Init strCardNum, strErr
    clsVoucher.SubInit strCardNum
    
    '//给门户发送消息
    SendMessgeToPortal "DocEditorOpened"
    
End Sub
 
'//警告
Private Sub Form_Resize()
    On Error Resume Next
    Me.UFToolbar1.Top = 0
    Me.UFToolbar1.Width = Me.ScaleWidth
    
    Picture2.Move 0, Me.tbrvoucher.Height, Me.ScaleWidth, Picture2.Height
    Picture2.Top = 0
    LabelVoucherName.Move (Me.Width - Me.LabelVoucherName.Width) / 2
    LabelVoucherName.ZOrder
    
    Picture1.BackColor = Me.Picture2.BackColor
    Picture1.Left = Me.ScaleWidth - Picture1.Width
    
'    Voucher.Move 0, Picture2.Top + Picture2.Height, Me.ScaleWidth, Me.ScaleHeight - Picture2.Height - Picture2.Top
    Voucher.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
  
    
    Me.BackColor = Me.Voucher.BackColor
    Me.ForeColor = Me.Voucher.BackColor
    
    If Voucher.VoucherStatus = VSNormalMode Then
        RefeshVoucher
    End If
End Sub

'//提示：单据功能的方法实现
Public Sub ButtonClick(s As String, sTaskKey As String, Optional bCloseSingle As Boolean = False)
    Dim i As Long
    Dim j As Long
    Dim row As Long
    Dim objGoldTax As Object
    Dim strError As String
    Dim strXMLHead As String
    Dim strXMLBody As String
    Dim lngRow As Integer
    Dim lngCol As Integer
    Dim strID As Variant
    Dim ele As IXMLDOMElement
    Dim strAuthID As String
    Dim elelist As IXMLDOMNodeList
    Dim ndRS    As IXMLDOMNode
    Dim nd      As IXMLDOMNode
    Dim bEAlast As Boolean
    Dim sPrnTmplate As Long
    Dim VoucherGrid As Object
    Dim Frm As New frmVouchNew
    Dim domError As New DOMDocument
    
    Dim strTaskId As String
    Dim retDoUndoSubmit As Boolean
    Dim strErrorResId As String '审批流的错误信息870 added
    Dim m_CardNumber As String
    Dim m_Mid As String
    Dim m_mcode As String
    Dim m_MAuthid As String
    Dim m_TablName As String
    
    On Error GoTo Err
    If clsTbl.ButtonKeyDown(m_login, s) Then
        UserbSuc = False
        Userdll_UI.Before_ButtonClick Me.Voucher, s, strUserErr, UserbSuc
        If UserbSuc Then
            Exit Sub
        End If
    
    
    strErrMsg = ""
    i = 0
    Set domPrint = Nothing
    
    If strVouchtype = "" Then strVouchtype = strCardNum
    domError.loadXML "<Data />"
    
    With Voucher
        Select Case LCase(s)
        
            Case LCase("BatchModify")
                Me.Voucher.ShowBatchModify
                
            Case "import" '导入
                Call EXCEL_Importdate
            
            Case "shiftto"                   '//转入
                LoadVoucher ""
            Case "setfixcols"          '是否显示固定行
                Me.Voucher.ShowFixColer = Not Me.Voucher.ShowFixColer
            Case "sumrow"              '合并显示
                Me.Voucher.SHowAggregateSetupDlg
 
            Case "lookrow"             '定位行
                Me.Voucher.ShowFindDlg
            Case "attached"            '增加附件
                Me.Voucher.SelectFile
            Case "output"              '输出
                strTaskId = clsVoucherCO.GetVoucherTaskID(LCase(s), strVouchtype, bReturnFlag)
                VouchOutPut Voucher, CLng(vtidDJMB(ComboDJMB.ListIndex)), strCardNum
            Case LCase("ToFirst")      '首张
                Screen.MousePointer = vbHourglass
                LoadVoucher s, , True
                Voucher.VoucherStatus = VSNormalMode
                
            Case LCase("ToPrevious")   '上一张
                Screen.MousePointer = vbHourglass
                LoadVoucher s, , True
                Voucher.VoucherStatus = VSNormalMode
            
            Case LCase("ToNext")       '下一张
                Screen.MousePointer = vbHourglass
                LoadVoucher s, , True
                Voucher.VoucherStatus = VSNormalMode
                
            Case LCase("ToLast")       '末张
                Screen.MousePointer = vbHourglass
                LoadVoucher s, , True
                Voucher.VoucherStatus = VSNormalMode
            Case "refresh"               '刷新
                Screen.MousePointer = vbHourglass
                If val(GetHeadItemValue(domHead, "vt_id")) <> 0 Then
                    LoadVoucher "", , True
                End If
            Case "locate"             '定位
                strID = SeekOneVoucher(strCardNum, bFirst)
                If val(strID) <> 0 Then
                    LoadVoucher "", strID, , False
                End If
            Case "cancel"             '取消
                bClickCancel = True
                Voucher.VoucherStatus = VSNormalMode
                If Not blnBlank Then
                    LoadVoucher ""
                Else
                    LoadVoucher "", 0
                End If
                
                bClickCancel = False
                Call VoucherFreeTask
                
            Case "submit"      '提交
                Screen.MousePointer = vbHourglass
                Dim mbilltype As String
                If Voucher.headerText("iswfcontrolled") = "1" And (Voucher.headerText("iverifystate") = "0" _
                    Or (Voucher.headerText("iverifystate") = "1" And val(Voucher.headerText("ireturncount")) > 0)) Then
                    Call GetCardNumberMid(m_CardNumber, m_Mid, m_TablName)
                    Set domHead = .GetHeadDom
                    Set ele = domHead.selectSingleNode("//z:row")
                    retDoUndoSubmit = DoUndoSubmit(True, m_CardNumber, m_Mid, m_TablName, ele.getAttribute("ufts"), CBool(Voucher.headerText("iswfcontrolled")), strErrorResId, , mbilltype)
                    If retDoUndoSubmit = False Then
                        MsgBox strErrorResId
                    Else
                        LoadVoucher ""
                        MsgBox GetString("U8.SA.xsglsql_2.saworkflowsrv.011") '"单据提交成功！"
                    End If
                Else
                    MsgBox GetString("U8.SA.xsglsql_2.saworkflowsrv.001") '"该单据已经提交或者未启用审批流！"
                End If
                Call VoucherFreeTask
                
            Case "unsubmit"   '撤销
                If Not IsHoldAuth(Me.Voucher.headerText("cmaker"), "A") Then
                    MsgBox "当前操作员没有权限", vbInformation
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                Screen.MousePointer = vbHourglass
                If Voucher.headerText("iswfcontrolled") = "1" And Voucher.headerText("iverifystate") <> "0" Then
                    Call GetCardNumberMid(m_CardNumber, m_Mid, m_TablName)
                    Set domHead = .GetHeadDom
                    Set ele = domHead.selectSingleNode("//z:row")
                    retDoUndoSubmit = DoUndoSubmit(False, m_CardNumber, m_Mid, m_TablName, ele.getAttribute("ufts"), CBool(Voucher.headerText("iswfcontrolled")), strErrorResId, Voucher.headerText(getVoucherCodeName()))
                    If retDoUndoSubmit = False Then
                        MsgBox strErrorResId
                    Else
                        LoadVoucher ""
                        MsgBox GetString("U8.SA.xsglsql_2.saworkflowsrv.012") '撤销成功！
                    End If
                Else
                    MsgBox GetString("U8.SA.xsglsql_2.saworkflowsrv.002") '"该单据已经撤销或者未启用审批流！"
                End If
                Call VoucherFreeTask
                
            Case LCase("viewverify")  ''查询审批流
                SendShowViewMessage "UFIDA.U8.Audit.AuditHistoryView"
                SendMessgeToPortal "DocQueryAuditHistory"
                
            Case "print"            '打印
                If Me.ComboVTID.ListCount = 0 Then
                    MsgBox GetString("U8.SA.xsglsql.01.frmbillvouch.00074"), vbExclamation 'zh-CN：当前操作员没有可以使用的打印模版，请检查！
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                sPrnTmplate = CLng(vtidPrn(Me.ComboVTID.ListIndex))
                BillPrnVTID = sPrnTmplate
                VoucherPrn strVouchtype, Voucher, strCardNum, CLng(sPrnTmplate), , True
                .VoucherStatus = VSNormalMode
                LoadVoucher ""
           
            Case "preview"
                If Me.ComboVTID.ListCount = 0 Then
                    MsgBox GetString("U8.SA.xsglsql.01.frmbillvouch.00074"), vbExclamation 'zh-CN：当前操作员没有可以使用的打印模版，请检查！
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                sPrnTmplate = CLng(vtidPrn(Me.ComboVTID.ListIndex))
                BillPrnVTID = sPrnTmplate
                VoucherPrn strVouchtype, Voucher, strCardNum, sPrnTmplate, "Preview", True
                .VoucherStatus = VSNormalMode
                LoadVoucher ""
                
            Case "opendraft"    '//打开草稿
                OpenFromDraft DraftMode
            Case "savedraft"    '//保存草稿
                SaveAsDraft DraftMode
            Case "managedraft"  '//管理草稿
                ManagementDraft DraftMode
            Case "opentemplate" '//打开模板
                OpenFromDraft TemplateMode
            Case "savetemplate" '//保存模板
                SaveAsDraft TemplateMode
            Case "managetemplate" '//管理模板
                ManagementDraft TemplateMode
            Case "help"               '帮助
                On Error Resume Next
                SendKeys "{F1}"
                ShowContextHelp Me.hwnd, App.HelpFile, Me.HelpContextID
                
            Case "exit"                '退出
                Unload Me
                Screen.MousePointer = vbDefault
                Exit Sub
                
            '警告：////////////////////以上内容请不要改动//////////////////////////////////
            
            Case "add"            '//新增
                m_MakeVoucherRuleID = ""  '//对于非UAP制单将制单规则清空
                
                If val(GetHeadItemValue(domHead, "vt_id")) <> 0 Then
                    blnBlank = False
                End If
                
                If moAutoFill Is Nothing Then
                    Set moAutoFill = CreateObject("ScmPublicSrv.clsAutoFill")
                    Voucher.SetCustomRelation moAutoFill.GetCustomRelationRecord(DBconn, strCardNum)
                    If Not BeforeEditVoucher() Then GoTo FreeTask
                End If
                ChangeTempaltes sCurTemplateID
                If ChangeDJMBForEdit = False Then
                    Screen.MousePointer = vbDefault
                    GoTo FreeTask
                End If
                Screen.MousePointer = vbDefault
                Call LoadVoucher("")
                
                Screen.MousePointer = vbHourglass
                Voucher.SkipLoadAccessories = False
                Voucher.AddNew ANMNormalAdd, domHead, domBody    '正常新增时所用的方法
                'voucher.AddNew ANMCopyALL, Domhead, Dombody     'Copy时所用方法
                Call SetVouchNoWriteble                          '设置单据号是否可以编辑
                Call AddNewVouch                                 '设置新增单据的初始值
                
                Screen.MousePointer = vbDefault
                Select Case UCase(strCardNum)  ''增加单据时的参照
                    Case "EFBWGL020301" ''选题论证
'                        If UCase(getAccinformation("EF", "bMustED_booksource", "BWGL")) = "TRUE" Then ''选题论证必有稿源登记
'                            Call ReferVouch
'                        End If
                        If ReferVouch = False Then
                            If UCase(getAccinformation("EF", "bMustED_booksource", "BWGL")) = "TRUE" Then ''选题论证必有稿源登记
                                ButtonClick "cancel", "放弃"
                            End If
                        End If
                    Case "EFBWGL020401" ''选题登陆
                        If ReferVouch = False Then
                            If UCase(getAccinformation("EF", "bMustED_seldclare", "BWGL")) = "TRUE" Then ''选题登陆必有选题申报
                                ButtonClick "cancel", "放弃"
                            End If
                        End If

                    
                End Select
                
                Set domHead = Me.Voucher.GetHeadDom
                iVouchState = 0
                'Call SetButtonStatus(s)
                Call SetItemState(s)
                
                On Error Resume Next
                
    
                picSure.Visible = False
                Voucher.SetFocus
                Voucher.UpdateCmdBtn
            
            Case "modify"              '//修改
                If val(GetHeadItemValue(domHead, "vt_id")) <> 0 Then
                    blnBlank = False
                End If
                
                If Not IsHoldAuth(Me.Voucher.headerText("cmaker"), "W") Then
                    MsgBox "当前操作员没有权限", vbInformation
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                If CheckDJMBAuth(Me.Voucher.headerText("vt_id"), "W") = False Then
                    MsgBox GetString("U8.SA.xsglsql.01.frmbillvouch.00021"), vbExclamation 'zh-CN：当前操作员没有当前单据模版的使用权限！
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                '控制审批中修改
                If GetHeadItemValue(domHead, "iswfcontrolled") = "1" Then
                    If Not (GetHeadItemValue(domHead, "iverifystate") = "0") Then
                        m_MAuthid = clsVoucherCO.GetVoucherTaskID("editforverify", strVouchtype, bReturnFlag)
                        Call GetCardNumberMid(m_CardNumber, m_Mid, m_TablName, m_mcode)
                        If bVerifyCanModify(m_CardNumber, m_Mid, m_mcode, m_MAuthid) = False Then
                            Exit Sub
                        End If
                    End If
                End If
                
                iVouchState = 1
                Voucher.row = 0
                Voucher.col = 0
                
                Screen.MousePointer = vbHourglass
                .getVoucherDataXML domHead, domBody
                Voucher.VoucherStatus = VSeEditMode
                If moAutoFill Is Nothing Then
                    Set moAutoFill = CreateObject("ScmPublicSrv.clsAutoFill")
                    Voucher.SetCustomRelation moAutoFill.GetCustomRelationRecord(DBconn, strCardNum)
                    If Not BeforeEditVoucher() Then Exit Sub
                End If
                Call SetItemState(s)
                Call AddNewVouch("modify")
                Call SetVouchNoWriteble
                
                'If strVouchType = "07" Then Me.voucher.BodyMaxRows = -1   对于单表头的单据，需要用这条语句
                Voucher.headerText("ccrechpname") = ""
                
                '//为了支持新的修改要求，需要增加如下字段cmodifier,dmoddate,dmodifysystime
                Voucher.headerText("cmodifier") = m_login.cUserName
                Voucher.headerText("dmoddate") = m_login.CurDate
                Voucher.headerText("dmodifysystime") = Date + Time()
                
                
                'Call SetButtonStatus(s)
               
                .SetFocus
                .UpdateCmdBtn
                .SetAggregateInfoCancelFlag
                 
                
            Case "copy", "copyvoucher"        '//复制
                If tbrvoucher.buttons("copy").ButtonMenus("copyvoucher").Enabled = False Then Exit Sub
                If val(GetHeadItemValue(domHead, "vt_id")) <> 0 Then
                    blnBlank = False
                End If
                If moAutoFill Is Nothing Then
                    Set moAutoFill = CreateObject("ScmPublicSrv.clsAutoFill")
                    Voucher.SetCustomRelation moAutoFill.GetCustomRelationRecord(DBconn, strCardNum)
                    If Not BeforeEditVoucher() Then Exit Sub
                End If
                If ChangeDJMBForEdit = False Then Screen.MousePointer = vbDefault: Exit Sub
                    Screen.MousePointer = vbDefault
                    Call LoadVoucher("")
                
 
                picSure.Visible = False
                Screen.MousePointer = vbHourglass
                iVouchState = 0
                AddNewVouch "copy"
                Me.Voucher.AddNew ANMCopyALL, domHead, domBody
                CheckAuthAfterDraft
                VouchHeadCellCheck Voucher.LookUpArray("ddate", siHeader), Voucher.headerText("ddate"), success
                Call SetVouchNoWriteble
                Call SetItemState("copyvoucher")
                .SetFocus
                .UpdateCmdBtn
                .SetAggregateInfoCancelFlag
                
            Case "copyline"                      '拷贝行
                i = Voucher.BodyMaxRows
                Voucher.BodyMaxRows = 0
                .DuplicatedLine .row
                Dim domBodyLine As DOMDocument
                Set domBodyLine = .GetLineDom(.BodyRows)
                clsVoucherCO.CopyRow domBodyLine
                Voucher.BodyMaxRows = i
                .UpdateLineData domBodyLine, .BodyRows
                Call Voucher_RowColChange
                
            Case "addline"                       '增行
                With Me.Voucher
                    .AddLine
                End With
                ReSetBodyRowNo
            Case "insertline"
                DoInsertLine
            Case "delline"                       '删行
                Dim tmpRow As Variant
                Dim strRows As String
                Dim varRows As Variant
                Dim myTmpStr As String
                If Voucher.row = 0 Or Voucher.Rows = 0 Then Exit Sub
                strRows = Voucher.GetSelectedRows
                varRows = Split(strRows, ",")
                For i = 0 To UBound(varRows)
                    Voucher.DelLine varRows(i)
                    tmpRow = varRows(i) - 1
                Next
                If tmpRow <> 0 Then
                    Me.Voucher.row = tmpRow
                    Call Voucher_RowColChange
                Else
                    Me.Voucher.row = 0
                    Me.Voucher.col = 0
                End If
                Voucher.getVoucherDataXML domHead, domBody
                ReSetBodyRowNo
                SetItemState
                
            Case "savenew"                 '保存新增
                ButtonClick "save", "保存"
                If Voucher.VoucherStatus = VSNormalMode Then
                    ButtonClick "add", tbrvoucher.buttons("add").Caption
                End If
                
            Case "save"                    '保存
'                Debug.Print Voucher.bodyText(34, 1)
'                Debug.Print Voucher.bodyText(35, 1)
'                Debug.Print Voucher.bodyText(36, 1)
                Voucher.getVoucherDataXML domHead, domBody
                Voucher.RemoveEmptyRow
                
                
                If Not Voucher.headVaildIsNull2(strError) Then
                    MsgBox strError, vbOKOnly + vbInformation
                    Voucher.SetFocus
                    Voucher.VoucherSetFocus siHeader
                    Exit Sub
                End If
                If Trim(domBody.xml) <> "" Then
                    If Not Voucher.bodyVaildIsNull2(strError) Then
                        MsgBox strError, vbOKOnly + vbInformation
                        Voucher.VoucherSetFocus sibody
                        Exit Sub
                    End If
                End If
'                Debug.Print Voucher.bodyText(35, 1)
                Screen.MousePointer = vbHourglass
                If Voucher.ProtectUnload2 <> 2 Then
                    Voucher.SetFocus
                    Exit Sub
                End If
                               
                Call AddNewVouch("Save")
                
                
                If strVoucherUFTS <> "" Then
                    If Not IsExistent Then  ''Clin 09-12-15 保存时判断来源单据状态
                        MsgBox "来源单据已被其他人修改或删除，请刷新后再试！"
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
                
                Voucher.getVoucherDataXML domHead, domBody
                
                SetVouchNoFormat domHead
                If SetAttachXML(domHead) = False Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                
'                If bFirst = True Then
'                    domHead.selectSingleNode("//z:row").Attributes.getNamedItem("bfirst").nodeValue = "1"
'                End If
                
'                If .BodyRows >= 10 Then
'                    Set ndRS = domBody.selectSingleNode("//rs:data")
'                    Me.Voucher.SkipLoadAccessories = True
'                    .StopSetDefaultValue = True
'                    .setVoucherDataXML domHead, domBody
'                    Me.Voucher.SkipLoadAccessories = False
'                    .StopSetDefaultValue = False
'                    If domBody.selectNodes("//z:row").length = 0 Then
'                        MsgBox GetString("U8.SA.xsglsql.01.frmbillvouch.00054"), vbInformation 'zh-CN：表体记录为0！
'                        Screen.MousePointer = vbDefault
'                        Exit Sub
'                    End If
'                End If
                
                'domhead增加结点区分 是销售保存还是 web or eai or others
                Dim strAccessories As String
                Dim accessoriesElement As IXMLDOMElement
                Set accessoriesElement = domHead.createElement("ischeck")
                accessoriesElement.Text = "false"
                domHead.selectSingleNode("//rs:data").appendChild accessoriesElement
                
                
ToSave:
                '设置UAP生单规则 begin
'                Call domHead.documentElement.setAttribute("makevoucherruleid", m_MakeVoucherRuleID)
                '设置UAP生单规则 end
                
                'by ahzzd 20100127
                '保存前检查
                clsVoucher.CheckBeforeSave "before", domHead, domBody, domError
                If domError.documentElement.selectNodes("row[@ignore='false']").length > 0 Then
                    strError = domError.xml
                Else
                    clsSAWeb.bManualTrans = False
                    clsVoucherCO.Init strCardNum, m_login, DBconn, "CS", clsSAWeb
                    strError = clsVoucherCO.Save(domHead, domBody, iVouchState, vNewID, domConfig)
                End If
                
                
                If strError <> "" Then
                    If InStr(1, strError, "<", vbTextCompare) <> 0 Then
                        ShowErrDom strError, domHead
                    Else
                        MsgBox strError
                        If domHead.selectNodes("//z:row").length = 1 Then
                            If .headerText(getVoucherCodeName) <> GetHeadItemValue(domHead, getVoucherCodeName) And strVouchtype <> "92" Then
                                .headerText(getVoucherCodeName) = GetHeadItemValue(domHead, getVoucherCodeName)
                            End If
                        End If
                    End If
                Else
                    Voucher.VoucherStatus = VSNormalMode
                    LoadVoucher "", IIf(vNewID <> "", vNewID, 0)
                    'Call SetButtonStatus(s)
                    'ChangeButtonsState
                    
                    DeleteDraft DraftMode '保存成功删除草稿
                    Call VoucherFreeTask
                    SendMessgeToPortal "DocSaved"
                End If
                
            Case "delete"                 '//删除
                
                If Not IsHoldAuth(Me.Voucher.headerText("cmaker"), "W") Then
                    MsgBox "当前操作员没有权限", vbInformation
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                
                If CheckDJMBAuth(Me.Voucher.headerText("vt_id"), "W") = False Then
                    MsgBox GetString("U8.SA.xsglsql.01.frmbillvouch.00022"), vbExclamation 'zh-CN：当前操作员没有当前单据模版的使用权限！
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                
                If MsgBox(GetString("U8.SA.xsglsql.01.frmbillvouch.00023"), vbYesNo + vbQuestion) = vbNo Then 'zh-CN：确实要删除本张单据吗？
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                
                Set domHead = Me.Voucher.GetHeadDom
                Screen.MousePointer = vbHourglass
                clsSAWeb.bManualTrans = False
                clsVoucherCO.Init strCardNum, m_login, DBconn, "CS", clsSAWeb
                strError = clsVoucherCO.Delete(domHead)
                If strError <> "" Then
                    ShowErrDom strError, domHead
                    LoadVoucher ""
                Else
                    LoadVoucher "tonext"
                End If
                Call VoucherFreeTask
                
            Case "verify", "unverify", "resubmit", "sure", "unsure"
                
                If LCase(s) = "verify" Then
                    If Not IsHoldAuth(Me.Voucher.headerText("cmaker"), "V") Then
                        MsgBox "当前操作员没有权限", vbInformation
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                ElseIf LCase(s) = "unverify" Then
                    If Not IsHoldAuth(Me.Voucher.headerText("cmaker"), "U") Then
                        MsgBox "当前操作员没有权限", vbInformation
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
                
                '判断是否启用审批流
                If Voucher.headerText("iswfcontrolled") = "1" Then
                    If LCase(s) = "verify" Or LCase(s) = "resubmit" Or LCase(s) = "sure" Then
                        SendShowViewMessage ("UFIDA.U8.Audit.AuditViews.TreatTaskViewPart")
                        SendMessgeToPortal "DocRequestAudit"
                    Else
                        SendShowViewMessage ("UFIDA.U8.Audit.AuditViews.TreatTaskViewPart")
                        SendMessgeToPortal "DocRequestCancelAudit"
                    End If
                Else
                    Call AddNewVouch(s)
                    Set domHead = Me.Voucher.GetHeadDom
                    If s = "verify" Or s = "sure" Then
                        bCheckVouch = True
                    Else
                        bCheckVouch = False
                    End If
                    
                    clsSAWeb.bManualTrans = False
                    clsVoucherCO.Init strCardNum, m_login, DBconn, "CS", clsSAWeb
                    strError = clsVoucherCO.VerifyVouch(domHead, bCheckVouch)
                    
                    If strError <> "" Then
                        Call ShowErrDom(strError, domHead)
                    End If
    
                    ''刷新当前单据
                    LoadVoucher ""
                    Call VoucherFreeTask
                End If
            Case "closeorder", "openorder"  '关闭 、打开
                
                If s = "closeorder" Then
                    bCheckVouch = True
                Else
                    bCheckVouch = False
                End If
                If bCheckVouch Then
                    If Not IsHoldAuth(Me.Voucher.headerText("cmaker"), "C") Then
                        MsgBox "当前操作员没有权限", vbInformation
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
                strError = clsVoucherCO.CloseVouch(domHead, bCheckVouch)
                If strError <> "" Then
                    Call ShowErrDom(strError, domHead)
                End If

                ''刷新当前单据
                LoadVoucher ""
                Call VoucherFreeTask
            Case "closeline", "openline"  '关闭 、打开
                If s = "closeline" Then
                    bCheckVouch = True
                Else
                    bCheckVouch = False
                End If
                If Voucher.row = 0 Or Voucher.Rows = 0 Then Exit Sub
                Dim autoids As String
                strRows = Voucher.GetSelectedRows
                varRows = Split(strRows, ",")
                For i = 0 To UBound(varRows)
                    If Voucher.bodyText(varRows(i), "autoid") <> "" Then
                        autoids = autoids & "'" & Voucher.bodyText(varRows(i), "autoid") & "',"
                    End If
                    tmpRow = varRows(i) - 1
                Next
                
                If autoids <> "" Then
                    autoids = Left(autoids, Len(autoids) - 1)
                    strError = clsVoucherCO.CloseVouchBodyLine(domHead, bCheckVouch, autoids)
                    If strError <> "" Then
                        Call ShowErrDom(strError, domHead)
                    End If
                End If

                ''刷新当前单据
                LoadVoucher ""
                Call VoucherFreeTask
            
            Case "receive"   '收稿，收稿记录专用 Clin
                If UCase(strCardNum) = "EFBWGL020502" Then
                    Dim strsql As String
                    .headerText("crecievedate") = m_login.CurDate
                    .headerText("crecievercode") = m_login.cUserId
                    .headerText("crecievername") = m_login.cUserName
                    strsql = "Update EFBWGL_DistRecord set crecievedate = '" & m_login.CurDate & "',crecievercode='" & m_login.cUserId & "',crecievername='" & m_login.cUserName & "' where ccode = '" & .headerText("ccode") & "'"
                    DBconn.Execute strsql
                End If
            Case "formatsetup"  '----格式设置
                If Voucher.ShowVoucherDesign = True Then
                    '重新设置单据格式
                    bfillDjmb = False
                    sCurTemplateID = 0
                    preVTID = 0
'                    Dim DomHtmp As New DOMDocument    '单据表头数据
'                    Dim DomBtmp As New DOMDocument    '单据表体数据

'                    Me.Voucher.getVoucherDataXML DomHtmp, DomBtmp
                    ComboDJMB_Click
'                    Me.Voucher.setVoucherDataXML DomHtmp, DomBtmp
                    Exit Sub
                End If
            Case "savelayout"
                Voucher.SaveVoucherTemplateInfo
        End Select
        
  
        
        '设置工具栏的可见状态
        'SetButtonStatus "Cancel"   用下面新的方法代替
'        clsTbl.ChangeButtonState Voucher, Me.tbrvoucher, Me.UFToolbar1, Voucher.VoucherStatus
        ChangeButtonState Voucher, Me.tbrvoucher, Me.UFToolbar1, Voucher.VoucherStatus
        If Voucher.headerText("cverifier") <> "" Then
            picSure.Visible = True
        Else
            picSure.Visible = False
        End If
        
        Userdll_UI.LoadAfter_VoucherData strUserErr, UserbSuc
        '单据模板标题及下拉框的显示
        SetCboVtidState
   End With
   Set ele = Nothing
   Screen.MousePointer = vbDefault
   
'   Exit Sub
'Err:
'    MsgBox Err.Description
'    Screen.MousePointer = vbDefault

FreeTask:
'    UserbSuc = False
   Userdll_UI.After_ButtonClick Me.Voucher, s, strUserErr, UserbSuc
   clsTbl.ButtonKeyUp m_login, s
'     If UserbSuc = False Then
'         MsgBox strUserErr
'     End If

   End If
   Exit Sub
Err:
    MsgBox Err.Description
    clsTbl.ButtonKeyUp m_login, s
    Screen.MousePointer = vbDefault
    
End Sub

'警告
Private Sub Form_Unload(Cancel As Integer)
    SendMessgeToPortal "DocEditorClosed"
    UnRegisterMessage
End Sub
 
'提示：右键菜单——新增行
Private Sub mdiAddRow_Click()
    ButtonClick "AddRow", ""
End Sub

'提示：右键菜单——删除行
Private Sub mdiDelRow_Click()
    ButtonClick "DelRow", ""
End Sub


Private Sub PopMenu1_MenuClick(sMenuKey As String)
 If sMenuKey = "BatchModify" Then
        Me.Voucher.ShowBatchModify
 End If
End Sub

'警告：
Private Sub tbrvoucher_ButtonClick(ByVal Button As MSComctlLib.Button)
    ButtonClick Button.Key, Button.ToolTipText
End Sub

'提示：单据号字段名称
Private Function getVoucherCodeName() As String
    Dim KeyCode As String
    Select Case strVouchtype
        Case ""
            KeyCode = "ccode"
            
        Case Else
            KeyCode = "ccode"
    End Select
    getVoucherCodeName = KeyCode
End Function

'警告：
Private Sub UFToolbar1_OnCommand(ByVal enumType As UFToolBarCtrl.ENUM_MENU_OR_BUTTON, ByVal cButtonId As String, ByVal cmenuid As String)
    On Error Resume Next
    Dim sButton As String
    Dim sButtonCaption As String
    sButton = IIf(enumType = enumButton, cButtonId, cmenuid)
    sButtonCaption = tbrvoucher.buttons(sButton)
    ButtonClick sButton, sButtonCaption
    
End Sub

Private Sub UFToolbar1_OnSelectedIndexChanged(ByVal enumType As UFToolBarCtrl.ENUM_MENU_OR_BUTTON, ByVal cButtonId As String, ByVal cmenuid As String, ByVal iSelectedIndex As Integer)
    If enumType = enumCombItem Then
        Select Case LCase(cButtonId)
        Case "printtemplate"
            If Voucher.VoucherStatus = VSNormalMode Then
                Me.ComboVTID.ListIndex = iSelectedIndex
            Else
                Me.ComboDJMB.ListIndex = iSelectedIndex
            End If
        Case "displayFormat"
            Me.ComboVTID.ListIndex = iSelectedIndex
        End Select
    End If
End Sub

Private Function SetPrintShowTemplate()
    If IsObject(tbrvoucher.buttons("PrintTemplate").Tag) Then
        If Voucher.VoucherStatus = VSNormalMode Then
            Set tbrvoucher.buttons("PrintTemplate").Tag.Tag = Me.ComboVTID
        Else
            Set tbrvoucher.buttons("PrintTemplate").Tag.Tag = Me.ComboDJMB
        End If
    End If
End Function

Private Sub Voucher_AddNewLineEvent(ByVal nRow As Long)
    Voucher.bodyText(nRow, "irowno") = Voucher.BodyRows
End Sub

Private Sub Voucher_AutoFillBackEvent(vtIndex As Variant, ByVal vtCurrentValue As Variant, ByVal vtCurrentFieldObject As Variant, ByVal vtAutoFieldInfo As Variant)
    Dim sErrMsg As String
    On Error GoTo Errhandle
    moAutoFill.AutoFillRelations DBconn, Voucher, vtCurrentFieldObject, vtAutoFieldInfo, sErrMsg
    Exit Sub
Errhandle:

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
            If InStr(1, LCase(ele.nodeName), "cfree") <> 0 Or InStr(1, LCase(ele.nodeName), "cdefine") <> 0 Then
                ele.setAttribute "reftype", ""
                ele.setAttribute "cRefID", ""
            Else
                Dim strReferName As String
                strReferName = clsVoucherRefer.GetReferName(LCase(ele.nodeName))
                If strReferName = "" And ele.getAttribute("cRefID") = "" Then
                    ele.setAttribute "reftype", ""
                    ele.setAttribute "cRefID", ""
                ElseIf ele.getAttribute("reftype") = "ref" And strReferName = "" Then
                    ele.setAttribute "reftype", ""
                Else
                    If strReferName <> "" Then
                        ele.setAttribute "cRefID", strReferName
                    End If
                End If
            End If
        Next
        sItemXML = oDom.xml
        Set oDom = Nothing
    End If
End Sub

'提示：单据新增时对单据号进行赋值
Private Sub Voucher_BillNumberChecksucceed()
    Dim errMsg As String, strVouchNo As String, KeyCode As String
    Dim tmpDOM As New DOMDocument
    KeyCode = getVoucherCodeName()
    'LDX    2009-05-21  注释    Beg
'    If strVouchType = "92" Then Exit Sub
    'LDX    2009-05-21  注释    End
    With Me.Voucher
        If Not (LCase(DomFormat.selectSingleNode("//单据编号").Attributes.getNamedItem("重号自动重取").nodeValue) = "false" And _
        LCase(DomFormat.selectSingleNode("//单据编号").Attributes.getNamedItem("允许手工修改").nodeValue) = "true") Then
            Set tmpDOM = .GetHeadDom
            If clsVoucherCO.GetVoucherNO(tmpDOM, strVouchNo, errMsg) = False Then
                MsgBox errMsg
            Else
                .headerText(KeyCode) = strVouchNo
            End If
        End If
    End With
End Sub

'提示：单据表体参照事件
Private Sub Voucher_bodyBrowUser(ByVal row As Long, ByVal col As Long, sRet As Variant, referPara As ReferParameter)
    Dim strsql As String, cInvCode As String, cCusCode As String
    Dim domHead As DOMDocument, domBody As DOMDocument
    Dim Dombodys_str1 As String
    Dim Dombodys_str2 As String
    Dim i As Integer, lRecord As Long
    Dim j As Long
    Dim Skey As String
    Dim sKeyValue As String, strAuth As String
    Dim tmpRow As Long, tmpCol As Long
    Dim tmpCol2 As Long
    Dim strClass As String
    Dim strGrid As String
    Dim ifalg As Boolean
    Dim dlgFileOpen As clsFileOperate
    tmpRow = row
    tmpCol = Voucher.col
    tmpCol2 = col
    On Error Resume Next
    With Voucher
        .MultiLineSelect = False ''设置多选默认
        clsRefer.SetReferSQLString ""
        clsRefer.SetRWAuth "INVENTORY", "R", False
        clsRefer.SetReferDisplayMode enuGrid
        Skey = .ItemState(col, sibody).sFieldName
        sKeyValue = .bodyText(row, col)
'        UserbSuc = False
        Userdll_UI.Voucher_bodyBrowUser Me.Voucher, Skey, row, sRet, strUserErr, UserbSuc
        If UserbSuc Then
            referPara.Cancel = True
            Exit Sub
        End If
        
        Select Case LCase(Skey)
            Case "filepath"                   'Clin 2009-11-21
                   '增加文件或修改已经选择的文件
                    If Me.Voucher.bodyText(Voucher.row, Skey) = "" Then
                     Me.dlgFileOpen.CancelError = False
                     Me.dlgFileOpen.Filter = "All Files|*.*"
                     Me.dlgFileOpen.ShowOpen
                    End If
                    
                    '如果取到，则设置文件名
                    If Trim("" & Me.dlgFileOpen.FileName) <> "" Then sRet = Trim("" & Me.dlgFileOpen.FileName)
                
            Case "cexpccode", "cexpcname" '预算费用类别
                    strClass = ""
                    strGrid = "select cexpccode,cexpcname from dbo.ExpItemClass  where 1=1  "
                    If clsRefer.StrRefInit_SetColWidth(m_login, False, strClass, strGrid, "费用类别编码,费用类别名称", "2500,6000") = False Then Exit Sub
                    clsRefer.Show
                    If Not clsRefer.recmx Is Nothing Then
                        .bodyText(row, "cexpccode") = clsRefer.recmx(0)
                        .bodyText(row, "cexpcname") = clsRefer.recmx(1)
                    sRet = clsRefer.recmx.Fields(LCase(Skey))
                    End If
 
           
        End Select
    End With
'    referPara.Cancel = True
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strReferString As String
    lngRow = row
    lngCol = col
    
    clsVoucherRefer.Init strCardNum, ""
    strReferString = clsVoucherRefer.ShowReferCtl(clsVoucher, Voucher, sibody, col, referPara, row)
'    If referPara.Cancel Then
'        sRet = Voucher.bodyText(lngRow, lngCol)
'    End If
    If strCardNum <> "EFBWGL020201" Then    'Clin 2009-11-21
        If referPara.Cancel Then
            sRet = Voucher.bodyText(lngRow, lngCol)
        End If
    End If
End Sub

'提示：单据表体校验事件
Private Sub Voucher_bodyCellCheck(RetValue As Variant, bChanged As Long, ByVal R As Long, ByVal C As Long, referPara As ReferParameter)
    'Dim lngRow As Long
    Dim intNumPoint As Integer
    Dim strFieldName As String
    Dim lngOldRow As Long
    Dim lngOldRows As Long
    Dim i As Long
    Dim strReferString As String
    Dim strsqltemp As String
    Dim rsttemp As New ADODB.Recordset
    Dim cls As New clsFileOperate
    Dim strFileOnServer As String
    Dim Skey As String
    
 
    Skey = LCase(Voucher.ItemState(C, sibody).sFieldName)
'    UserbSuc = False
    Userdll_UI.Voucher_bodyCellCheck Me.Voucher, RetValue, bChanged, Skey, R, strUserErr, UserbSuc
    If UserbSuc Then
        referPara.Cancel = True
        Exit Sub
    End If
          
    
    referPara.bValid = True
    strFieldName = Voucher.ItemState(C, sibody).sFieldName
    lngOldRow = R
    lngOldRows = Voucher.BodyRows
    If Not referPara.rstGrid Is Nothing Then
        strReferString = clsVoucherRefer.FillItemsAfterBrowse(clsVoucher, Voucher, "B", strFieldName, referPara.rstGrid, RetValue, R)
        If Voucher.BodyRows > lngOldRows Then
            For i = lngOldRows + 1 To Voucher.BodyRows
                clsVoucher.CellCheck Voucher, "B", Voucher.ItemState(C, sibody).sFieldName, bChanged, clsVoucherRefer, i
            Next
        End If
        R = lngOldRow
    Else
        strFieldName = clsVoucherRefer.CellCheck("", Voucher, "B", strFieldName, RetValue, R)
        If strFieldName <> "" Then
            If strFieldName = "cancel" Then
                bChanged = Cancel
            Else
                RetValue = ""
'                bChanged = retry
            End If
            Exit Sub
        End If
    End If
    
'    clsVoucherRefer.CellCheck strReferString, voucher, "B", voucher.ItemState(C, sibody).sFieldName, R
    clsVoucher.CellCheck Voucher, "B", Voucher.ItemState(C, sibody).sFieldName, bChanged, clsVoucherRefer, R
    RetValue = Voucher.bodyText(lngOldRow, C)
    If Voucher.ItemState(C, sibody).nFieldType = 4 Then
        intNumPoint = Voucher.ItemState(C, sibody).nNumPoint
        RetValue = Format(RetValue, IIf(intNumPoint = 0, "###0", "###0." & String(intNumPoint, "0")))
    End If
    If Voucher.ItemState(C, sibody).nReferType = 3 Then
        RetValue = Format(RetValue, "YYYY-MM-DD")
    End If
    
    Dim allBodySum As Double
    With Me.Voucher
        Skey = LCase(.ItemState(C, sibody).sFieldName)
        Select Case UCase(strCardNum)
            Case "EFFYGL040201", "EFFYGL040301" '费用预估单、费用结算单
                Select Case LCase(Skey)
                    Case "inumber", "iunitprice", "imoney"
                        allBodySum = 0
                        For i = 1 To Voucher.BodyRows
                            allBodySum = allBodySum + val(Format(Voucher.bodyText(i, "imoney") & "", "#0.00"))
                        Next
                        Voucher.headerText("je") = allBodySum
                End Select
        End Select
    End With
    
    '//以下为预算管理功能特殊处理部分/////////////////////////////////
    ReSetBodyRowNo
    


End Sub
 
Private Sub Voucher_Click(section As UAPVoucherControl85.SectionsConstants, ByVal Index As Long)
Dim UserbSuc As Boolean
Dim strUserErr As String
    If Index > 0 Then
        Userdll_UI.Voucher_Click Me.Voucher, section, Me.Voucher.ItemState(Index, section).sFieldName, Me.Voucher.row, strUserErr, UserbSuc
        If UserbSuc Then
'            referPara.Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub Voucher_DblClick(section As UAPVoucherControl85.SectionsConstants, ByVal Index As Long)
    Dim lngRet As Long
    Dim cls As New clsFileOperate
    Dim UserbSuc As Boolean
    Dim strUserErr As String
    If Index > 0 Then
        Userdll_UI.Voucher_DblClick Me.Voucher, section, Me.Voucher.ItemState(Index, section).sFieldName, Me.Voucher.row, strUserErr, UserbSuc
'        Userdll_UI.Voucher_DblClick Me.Voucher, section, "", Me.Voucher.row, strUserErr, UserbSuc
        If UserbSuc Then
'            referPara.Cancel = True
            Exit Sub
        End If
    End If
    Select Case UCase(m_strCardNum)
        Case "EFBWGL020201"   ''稿源登记
           '修改下句为不只是在双击路径，而是该行
            If section = sibody And Voucher.bodyText(Voucher.RowSel, "filepath") <> "" And Voucher.ColSel = Voucher.ItemState("filepath", 1).Reserve6 Then
                If MsgBox(" 是否打开 " & Voucher.bodyText(Voucher.RowSel, "filename") & "?", vbYesNo + vbQuestion) = vbYes Then
                    cls.SetParam m_login
                    cls.OpenFile "" & Voucher.bodyText(Voucher.RowSel, "filename"), "" & Voucher.bodyText(Voucher.RowSel, "filepath")
                   Set cls = Nothing
                End If
            End If
'        Case "EFZZ0612"
'            If section = sibody And Voucher.bodyText(Voucher.RowSel, "filename") <> "" And Voucher.ColSel = Voucher.ItemState("filename", 1).Reserve6 Then
'               'If ShowMsgBox("是否要打开文件【" & vch.bodyText(vch.RowSel, "strPath") & "】?", vbYesNo + vbQuestion) = vbYes Then
'                If MsgBox(" 是否打开 " & Voucher.bodyText(Voucher.RowSel, "filename") & "?", vbYesNo + vbQuestion) = vbYes Then
'                    'lngRet = ShellExecute(GetDesktopWindow, "open", vch.bodyText(vch.RowSel, "strPath"), vbNullString, vbNullString, vbNormalFocus)
'                    cls.SetParam m_Login
'        '            cls.OpenFile "" & Voucher.bodyText(Voucher.RowSel, "filepath"), "" & Voucher.bodyText(Voucher.RowSel, "filetype")
'                    cls.OpenFile "" & Voucher.bodyText(Voucher.RowSel, "filename"), "" & Voucher.bodyText(Voucher.RowSel, "servername")
'                   Set cls = Nothing
'                End If
'            End If
    End Select
End Sub

'提示：单据表头列表式项目的初始化事件
Private Sub Voucher_FillHeadComboBox(ByVal Index As Long, pCom As Object)
    Dim i As Integer
    Dim rds As New ADODB.Recordset
    Dim Skey As String
'        UserbSuc = False
        Skey = Me.Voucher.ItemState(Index, siHeader).sFieldName
        Userdll_UI.Voucher_FillHeadComboBox Me.Voucher, Skey, pCom, strUserErr, UserbSuc
        If UserbSuc Then
            Exit Sub
        End If
         
        Select Case Skey
            Case "iyear" '编制年度
                    pCom.Clear
                    For i = 0 To 4
                      pCom.AddItem CStr(val(m_login.cIYear) + i)
                    Next
                  
           Case Else
                pCom.Clear
                clsVoucherRefer.FillHeadComboBox Voucher, Index, pCom
        End Select
End Sub

'提示：单据表体列表式项目的初始化事件
Private Sub Voucher_FillList(ByVal R As Long, ByVal C As Long, pCom As Object)
    Dim sFieldName As String
    sFieldName = LCase(Me.Voucher.ItemState(C, sibody).sFieldName)
'    UserbSuc = False
    Userdll_UI.Voucher_FillList Me.Voucher, sFieldName, R, pCom, strUserErr, UserbSuc
    If UserbSuc Then
        Exit Sub
    End If
    
    pCom.Clear
    clsVoucherRefer.FillList Voucher, R, C, pCom
    
End Sub
 
'提示：单据表头参照事件
Private Sub Voucher_headBrowUser(ByVal Index As Variant, sRet As Variant, referPara As ReferParameter)
    'Dim iElement As IXMLDOMElement
    Dim Skey As String, sKeyValue As String
    Dim Str As String
    
    Skey = Me.Voucher.ItemState(Index, siHeader).sFieldName
    sKeyValue = Me.Voucher.headerText(Index)
    Dim strClass As String
    Dim strGrid As String
    Dim strsql As String
    Dim Rst As New ADODB.Recordset
    
'    UserbSuc = False
    Userdll_UI.Voucher_headBrowUser Me.Voucher, Skey, sRet, strUserErr, UserbSuc
    If UserbSuc Then
        referPara.Cancel = True
    Else
        clsRefer.referMulti = False
        clsRefer.SetReferDisplayMode enuGrid
        clsRefer.SetReferSQLString ""
        
        Dim strReferString As String
        strReferString = clsVoucherRefer.ShowReferCtl(clsVoucher, Voucher, siHeader, CLng(Index), referPara)
        sRet = Voucher.headerText(Index)
    End If
    
End Sub

'警告：自定义项参照方法
Private Function RefDefine(Index As Variant, iVoucherSec As Integer) As String
    Dim clsDef As U8DefPro.clsDefPro
    Dim nDataSource As Long         '数据来源
    Dim nEnterType As Long         '输入方式
    Dim sDataRule As String       '数据公式
    Dim bValidityCheck As Boolean      '是否合法性检测
    Dim bBuildArchives As Boolean      '是否建档
    Dim sVouchType As String
    Dim sTableName As String, sFieldName As String, sCardNumber As String
    Dim sDefWhere As String
    Dim strKeyValue As String
    Set clsDef = New U8DefPro.clsDefPro
    With Me.Voucher
        If iVoucherSec = siHeader Then
            strKeyValue = .headerText(Index)
        Else
            strKeyValue = .bodyText(.row, Index)
        End If
        nDataSource = .ItemState(Index, iVoucherSec).nDataSource
        nEnterType = .ItemState(Index, iVoucherSec).nEnterType
        sDataRule = .ItemState(Index, iVoucherSec).sDataRule
        bValidityCheck = .ItemState(Index, iVoucherSec).bValidityCheck
        bBuildArchives = .ItemState(Index, iVoucherSec).bBuildArchives
        Select Case nDataSource  '0表示手工输入；1表示档案；2表示单据
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
        If Not clsDef.Init(False, DBconn.ConnectionString, m_login.cUserId) Then
            RefDefine = ""
            MsgBox "初始化自定义项组件失败！"
            Exit Function
        End If
        RefDefine = clsDef.GetRefVal(nDataSource, iVoucherSec, .ItemState(Index, iVoucherSec).sFieldName, sTableName, sFieldName, sCardNumber, strKeyValue, False, 40, 1)
    End With
    Set clsDef = Nothing
End Function

'提示：单据表头校验事件
Private Sub Voucher_headCellCheck(Index As Variant, RetValue As String, bChanged As UAPVoucherControl85.CheckRet, referPara As UAPVoucherControl85.ReferParameter)
    Dim strCellCheckType As String
    Dim blnTrue As Boolean
    Dim domTmp As New DOMDocument
    Dim strFieldName As String
    Dim strCellCheck As String
    Dim strReferString As String
    
    referPara.bValid = True
    strFieldName = Voucher.ItemState(Index, siHeader).sFieldName
        
'    UserbSuc = False
'    Userdll_UI.Voucher_headBrowUser Me.Voucher, strFieldName, RetValue, strUserErr, UserbSuc
    Userdll_UI.Voucher_headCellCheck Me.Voucher, strFieldName, RetValue, bChanged, strUserErr, UserbSuc
    If UserbSuc Then
        referPara.Cancel = True
    Else
        If Not referPara.rstGrid Is Nothing Then
            strReferString = clsVoucherRefer.FillItemsAfterBrowse(clsVoucher, Voucher, "T", strFieldName, referPara.rstGrid, RetValue)
        Else
            strCellCheck = clsVoucherRefer.CellCheck("", Voucher, "T", strFieldName, RetValue)
            If strCellCheck <> "" Then
                RetValue = ""
                'bChanged = retry
                blnOnEdit = False
                Exit Sub
            End If
        End If
                
        RetValue = Voucher.headerText(Index)
        strCellCheckType = clsVoucher.CellCheck(Voucher, "T", strFieldName, bChanged, clsVoucherRefer)
        If strCellCheckType <> "" Then
            blnTrue = ReferVoucherByInput(strCellCheckType)
            If Not blnTrue Then
                bChanged = Cancel
                domTmp.loadXML strCellCheckType
                If Not domTmp.documentElement.Attributes.getNamedItem("errresid") Is Nothing Then
                    MsgBox GetString(domTmp.documentElement.Attributes.getNamedItem("errresid").Text), vbExclamation
                End If
                Set domTmp = Nothing
            End If
        End If
    End If
    

End Sub

Private Function ReferVoucherByInput(strNodXml As String) As Boolean
    Dim domSet As New DOMDocument
    Dim clsReferVoucher As New clsAutoReferVouch
    Dim strReferKey As String
    Dim strFilter As String
    Dim domHead As New DOMDocument
    Dim domBody As New DOMDocument
    
    domSet.loadXML strNodXml
    ReferVoucherByInput = False
    If Not domSet Is Nothing Then
        strReferKey = domSet.documentElement.Attributes.getNamedItem("cellcheck").Text
        strFilter = ReplaceVoucherItems(domSet.documentElement.Attributes.getNamedItem("desfldname").Text, Voucher)
        clsReferVoucher.Init strCardNum, strReferKey
        clsReferVoucher.InitReferVoucher Nothing
        clsReferVoucher.strFilter = strFilter
        clsReferVoucher.GetOneVoucher domHead, domBody
        If domHead.xml <> "" Then
            Set clsReferVoucher.domSourceHead = domHead
            Set clsReferVoucher.domSourceBody = domBody
            ReferFillVoucher clsReferVoucher
            ReferVoucherByInput = True
        End If
    End If
    Set domSet = Nothing
    Set clsReferVoucher = Nothing
    Set domHead = Nothing
    Set domBody = Nothing
End Function
Private Function ReferFillVoucher(clsReferVoucher As clsAutoReferVouch)
    Dim domHead As DOMDocument
    Dim domBody As DOMDocument
    Dim lngRows As Long
    Dim i As Long
    Dim j As Long
    Dim varFields As Variant
    Dim strFillType As String
    Dim tmpMaxRows As Integer
'    Dim strOldExchName As String
    
    lngRows = clsReferVoucher.domSourceBody.selectNodes("//z:row").length
    strFillType = clsReferVoucher.strFillType
    If strFillType = "removedetails" Then
    For i = Voucher.Rows To 1 Step -1
        Voucher.DelLine i
    Next
    End If
    tmpMaxRows = Voucher.BodyMaxRows
    Voucher.BodyMaxRows = 0
    For i = 0 To lngRows - 1
        Voucher.AddLine
    Next
    Voucher.BodyMaxRows = tmpMaxRows
    Voucher.getVoucherDataXML domHead, domBody
    clsReferVoucher.FillVoucherItems domHead, domBody, True
'    clsReferVoucher.FillCellCheckItems domHead, Dombody
    Dim strExchName As String
    Voucher.setVoucherDataXML domHead, domBody
    Dim referPara As UAPVoucherControl85.ReferParameter
    If clsReferVoucher.strHeadCellCheckFields <> "" Then
        varFields = Split(clsReferVoucher.strHeadCellCheckFields, ",")
        For i = 0 To UBound(varFields)
            Voucher_headCellCheck Voucher.LookUpArray(varFields(i), siHeader), "", Cancel, referPara
        Next
    End If
    If clsReferVoucher.strBodyCellCheckFields <> "" Then
        varFields = Split(clsReferVoucher.strBodyCellCheckFields, ",")
        For i = Voucher.BodyRows To Voucher.BodyRows - lngRows + 1 Step -1
            For j = 0 To UBound(varFields)
                If Voucher.bodyText(i, varFields(j)) <> "" Then
                    Voucher_bodyCellCheck Voucher.bodyText(i, varFields(j)), 1, i, Voucher.LookUpArray(varFields(j), sibody), referPara
'                    Voucher.CallAutoFillBackEvent varFields(j), i
                End If
            Next
        Next
    End If
    Voucher.getVoucherDataXML domHead, domBody
    clsReferVoucher.FillVoucherItems domHead, domBody, False
    Voucher.setVoucherDataXML domHead, domBody
    strExchName = GetHeadItemValue(domHead, "cexch_name")
    If strExchName <> "" Then
        Voucher.headerText("iexchrate") = clsSAWeb.GetExchRate(strExchName, m_login.CurDate, m_login)
        Call Voucher_headCellCheck(Voucher.LookUpArray("iexchrate", siHeader), Voucher.headerText("iexchrate"), success, referPara)
        Voucher.ItemState("iexchrate", siHeader).nNumPoint = clsSAWeb.GetExchRateDec(strExchName)
    End If
    Voucher.ExecCalcExpression
End Function

'警告：填充单据上打印、显示模板框
Private Sub fillComBol(bPrint As Boolean, tmprst As ADODB.Recordset)
    Dim i As Long
    
    If bPrint = True Then
        ComboVTID.Clear
        tmprst.Filter = "VT_TemplateMode=1"
    Else
        ComboDJMB.Clear
        tmprst.Filter = "VT_TemplateMode=0"
    End If
    
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
            bfillDjmb = True
            ComboVTID.ListIndex = 0
            
            bfillDjmb = False
            ComboVTID.UTooltipText = ComboVTID.Text
        Else
            ComboDJMB.Clear
            i = 0
            Do While Not tmprst.EOF
                ComboDJMB.AddItem tmprst(0)
                vtidDJMB(i) = CLng(tmprst(1))
                ComboDJMB.ItemData(ComboDJMB.NewIndex) = tmprst.Fields("printid")
                i = i + 1
                tmprst.MoveNext
            Loop
            bfillDjmb = True
            ComboDJMB.ListIndex = 0
            bfillDjmb = False
            ComboDJMB.UTooltipText = ComboDJMB.Text
        End If
    End If
End Sub


'提醒：单据窗体加载主方法
Public Function ShowVoucher(CardNumbers As String, Optional vVoucherId As Variant, Optional imode As Integer, Optional strCurrentRow As String)
    Dim tmpTemplateID As String
    Dim errMsg As String
    Dim vfd As Object
    Dim strEnterVoucherMode As String   '单据进入的模式(True,空、False最后一张)
    Dim errStr As String
    Dim bSuc As Boolean
    
    On Error GoTo DoERR
    
    
    sGuid = CreateGUID()
    If Not (g_business Is Nothing) Then
        Set vfd = g_business.CreateFormEnv(sGuid, Me)
    End If
    g_FormbillShow = False
    
    '插件初始化
    Get_SA_VoucherListConfig
 
'    UserbSuc = True
    cls_Public.WrtDBlog DBconn, m_login.cUserId, "EFmain", "UI接口装载开始！"
    Userdll_UI.Userdll_Init g_business, m_login, DBconn, Me, CardNumbers, strUserErr, UserbSuc
    If UserbSuc = False Then
        cls_Public.WrtDBlog DBconn, m_login.cUserId, "EFmain", "UI接口装载失败！"
    Else
        cls_Public.WrtDBlog DBconn, m_login.cUserId, "EFmain", "UI接口装载成功！"
    End If
    
    Screen.MousePointer = vbHourglass
    If IsMissing(imode) = True Then
        iShowMode = 0
    Else
        iShowMode = imode
    End If
    
    Set clsVoucherCO = New EFVoucherCo.clsVoucherCO    '提示：请结合在程序设计时所取的控制层类名进行修改
    clsVoucherCO.Init CardNumbers, m_login, DBconn, "CS", clsSAWeb
    cls_Public.WrtDBlog DBconn, m_login.cUserId, "EFmain", CardNumbers & "单据, clsVoucherCO初始化结束！"
    
    ''获到单据初始进入时的选项
    strEnterVoucherMode = "" 'clsSAWeb.GetSysDicOption("SA", "EnterVoucherMode")
    If strEnterVoucherMode = "" Then strEnterVoucherMode = ""
    blnBlank = IIf(strEnterVoucherMode = "Blank", True, False)
    
    
    If iShowMode <> 2 Then
    
        If (IsMissing(vVoucherId) Or IsNull(vVoucherId)) And blnBlank Then
            errMsg = clsVoucherCO.GetVoucherData(domHead, domBody, 0)
        Else
            errMsg = clsVoucherCO.GetVoucherData(domHead, domBody, vVoucherId)
        End If
 
        If errMsg <> "" Then
            MsgBox errMsg
            If iShowMode = 1 Then
                On Error Resume Next
                bFrmCancel = False
                Set clsVoucherCO = Nothing
                Set clsRefer = Nothing
                Set RstTemplate = Nothing
                Set domHead = Nothing
                Set domBody = Nothing
                Set DomFormat = Nothing
                
                If m_UFTaskID <> "" Then
                    m_login.TaskExec m_UFTaskID, 0
                End If
                Screen.MousePointer = vbDefault
                Exit Function
            End If
        End If
 
 
        
        ''获取单据模板
        If Not domHead.selectSingleNode("//z:row") Is Nothing Then
            'LDX    2009-05-22  Modify  Beg
            '增加了当VoucherType = MT66时 ,模版字段的取值
           
                If Not domHead.selectSingleNode("//z:row").Attributes.getNamedItem("vt_id") Is Nothing Then
                    tmpTemplateID = domHead.selectSingleNode("//z:row").Attributes.getNamedItem("vt_id").nodeValue
                ElseIf Not domHead.selectSingleNode("//z:row").Attributes.getNamedItem("vt_id") Is Nothing Then
                    tmpTemplateID = domHead.selectSingleNode("//z:row").Attributes.getNamedItem("vt_id").nodeValue
                Else
                    tmpTemplateID = "0"
                End If
        Else
            tmpTemplateID = "0"
        End If
 
    Else
        errMsg = clsVoucherCO.GetVoucherData(domHead, domBody, 0)
        If errMsg <> "" Then
            MsgBox errMsg
        End If
    End If
    
    '增加附件
    Call SetVoucherDataSource
    
    GetDomVtid
    If Me.ComboDJMB.ListCount <= 0 Then
        MsgBox GetString("U8.SA.xsglsql.01.frmbillvouch.00171"), vbExclamation 'zh-CN：您没有模版使用权限
        Exit Function
    End If
    
    If tmpTemplateID = "" Or tmpTemplateID = "0" Then
        sCurTemplateID = vtidDJMB(0)     ''取默认模板
    Else
        sCurTemplateID = tmpTemplateID  ''新的模板
    End If
    sCurTemplateID2 = sCurTemplateID
    
    If sCurTemplateID = "0" Then
        Me.Hide
        MsgBox GetString("U8.SA.xsglsql.01.frmbillvouch.00171"), vbExclamation 'zh-CN：您没有模版使用权限
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    If ChangeTempaltes(sCurTemplateID, True, False, True) = False Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    On Error Resume Next
    If GetHeadItemValue(domHead, "cexch_name") <> "" Then
        Me.Voucher.ItemState("iexchrate", siHeader).nNumPoint = clsSAWeb.GetExchRateDec(GetHeadItemValue(domHead, "cexch_name"))
    End If
    
    On Error GoTo DoERR
    
    '改变标题颜色
    Call ChangeCaptionCol
    
    'U872窗体标题颜色
    If UFFrmCaptionMgr.Caption = "" Or UFFrmCaptionMgr.Caption = "0" Then
        UFFrmCaptionMgr.Caption = Me.LabelVoucherName.Caption
    End If
    FormVisible = True

    If g_business Is Nothing Then
        Me.Show
    Else
        Call g_business.ShowForm(Me, "EF", sGuid, False, True, vfd)
        Set Me.Voucher.PortalBusinessObject = g_business
        Me.Voucher.PortalBizGUID = sGuid
        SendMessgeToPortal "CurrentDocChanged"
    End If
 
    Me.Voucher.SkipLoadAccessories = False
    
    
    Voucher.setVoucherDataXML domHead, domBody
    vNewID = Voucher.headerText("id")
    If Voucher.headerText("cverifier") <> "" Then
        picSure.Visible = True
    Else
        picSure.Visible = False
    End If
    Userdll_UI.LoadAfter_VoucherData strUserErr, UserbSuc
     
    If strCurrentRow <> "" Then
        Voucher.SetCurrentRow strCurrentRow
    End If
    
    '审批流文本
    Me.Voucher.ExamineFlowAuditInfo = GetEAStream(Me.Voucher, strVouchtype)
    
    If iShowMode <> 2 Then
        Voucher.VoucherStatus = VSNormalMode
'        clsTbl.ChangeButtonState Voucher, Me.tbrvoucher, Me.UFToolbar1, Voucher.VoucherStatus
        ChangeButtonState Voucher, Me.tbrvoucher, Me.UFToolbar1, Voucher.VoucherStatus
    End If
    Me.Voucher.Visible = True
    Call DWINfor
    
    
    Dim strXml As String
    strXml = "<?xml version='1.0' encoding='UTF-8'?>" & Chr(13)
    domConfig.loadXML strXml & "<ATO>" & Chr(13) & " </ATO>"

    Me.BackColor = Me.Voucher.BackColor
    Screen.MousePointer = vbDefault
    g_FormbillShow = True
    Err.Clear
    Exit Function
DoERR:
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
End Function

'提醒：根据单据编号设置规则，设置单据号字段是否可以编辑
Private Sub SetVouchNoWriteble()
    Dim KeyCode As String
    On Error Resume Next
    If strVouchtype = "92" Then Exit Sub
    KeyCode = getVoucherCodeName()
    If Not DomFormat Is Nothing Then
        If DomFormat.xml <> "" Then
            If LCase(DomFormat.selectSingleNode("//单据编号").Attributes.getNamedItem("允许手工修改").nodeValue) = "false" And LCase(DomFormat.selectSingleNode("//单据编号").Attributes.getNamedItem("重号自动重取").nodeValue) = "false" Then
                Me.Voucher.EnableHead KeyCode, False
            Else
                Me.Voucher.EnableHead KeyCode, True
            End If
        End If
    End If
End Sub

'提示：单据表头编辑时发生的事件
Private Sub Voucher_headOnEdit(Index As Integer)
'    Dim tmprst As New ADODB.Recordset
'    Dim sWhere As String
'    Dim intCyc As Integer
'    Dim strsql As String
'    With Me.Voucher
'        'If .LookUpArrayFromKey("citemcode", siheader) = Index Or .LookUpArrayFromKey("citemname", siheader) = Index Then
'            If strCardNum = "MT66" And .headerText("citemcode") <> "" Then
'                If blnOnEdit = True Then
'                    sWhere = " VT_CardNumber = '" + strCardNum + "' and vt_ID in (select left(b.ccode,6) from mt_baseset a left join mt_basesets b on a.id=b.id where a.cvouchtype='11' and a.citemcode='" & Me.Voucher.headerText("citemcode") & "') "
'                    strsql = "SELECT VT_Name,VT_ID,isnull(VT_PrintTemplID,DEF_ID_PRN) as printid,isnull(VT_TemplateMode,0) as VT_TemplateMode From vouchertemplates inner join vouchers_base on vouchertemplates.VT_CardNumber=vouchers_base.cardnumber WHERE (" & sWhere & ")  "
'                    tmprst.CursorLocation = adUseClient
'                    tmprst.Open ConvertSQLString(strsql), DBconn, adOpenForwardOnly, adLockReadOnly
'                    If tmprst.RecordCount > 0 Then
'                        For intCyc = 0 To ComboDJMB.ListCount - 1
'                            If ComboDJMB.List(intCyc) = tmprst.Fields(0) Then
'                                If ComboDJMB.ListIndex <> intCyc Then
'                                    ComboDJMB.ListIndex = intCyc
'                                End If
'                                Exit For
'                            End If
'                        Next
'                    End If
'
'                End If
'
'            End If
'        'End If
'    End With
'    blnOnEdit = False
End Sub
 
Private Sub Voucher_IsAllowBatchModify(bCanModify As Boolean, ByVal row As Long, ByVal colkey As String)
    bCanModify = Voucher.ItemState(colkey, sibody).bCanModify
End Sub

'提示：单据KeyPress事件
Private Sub voucher_KeyPress(ByVal section As UAPVoucherControl85.SectionsConstants, ByVal Index As Long, KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        KeyAscii = 0
'    End If
End Sub

Private Sub Voucher_LoadFromTemplate(ByVal enumType As UAPVoucherControl85.TemplateModes, ByVal TemplateID As String, oDOmHead As Variant, oDomBody As Variant, oOtherDom As Variant)
    Dim ele  As IXMLDOMElement
    
    If enumType = DraftMode Then
        m_sCurrentDraftID = TemplateID
    End If
    If Me.Voucher.VoucherStatus = VSNormalMode And enumType = DraftMode Then
        bnewDraft = True
        ButtonClick "add", ""
        bnewDraft = False
    End If
    Voucher.StopSetDefaultValue = True
    Voucher.SkipLoadAccessories = True
    
    For Each ele In oDomBody.selectNodes("//z:row")    ''Clin 09-12-15 置模版表体标识，否则不能保存表体数据
        ele.setAttribute "editprop", "A"
    Next
    
    Voucher.setVoucherDataXML oDOmHead, oDomBody
    Voucher.StopSetDefaultValue = False
    Voucher.SkipLoadAccessories = False
    CheckAuthAfterDraft
    If enumType = TemplateMode Then
        AddNewVouch "copy"
        Call Voucher_BillNumberChecksucceed
        VouchHeadCellCheck Voucher.LookUpArray("ddate", siHeader), Voucher.headerText("ddate"), success
    End If
End Sub

'提示：单据点击右键事件
Private Sub voucher_MouseUp(ByVal section As UAPVoucherControl85.SectionsConstants, ByVal Index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If section = sibody And Button = 2 Then
'        If strCardNum = "MT06" Or strCardNum = "MT07" Then
'            PopupMenu mnuPop
'        End If
'    End If
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
    For Each eleMent In oxml.selectSingleNode("//表头").selectNodes("//字段")
        If eleMent.getAttribute("边框") = "0x1F" Then
            eleMent.setAttribute "边框", "3"
        End If
        If Left(eleMent.getAttribute("关键字"), 2) <> "文本" Then
            rsPrintModel.Filter = ""
            rsPrintModel.Filter = "fieldname='" & Mid(eleMent.getAttribute("关键字"), InStr(1, eleMent.getAttribute("关键字"), "(") + 1, InStr(1, eleMent.getAttribute("关键字"), ")") - InStr(1, eleMent.getAttribute("关键字"), "(") - 1) & "'"
            If rsPrintModel.RecordCount Then
                eleMent.setAttribute "对齐方式", "右"
            End If
        End If
    Next
    sStyle = oxml.xml
    
        If Voucher.headerText("bfirst") Then
            tmpDOM.loadXML sStyle
            Set ndRootList = domPrint.selectNodes("//标题")
            For Each ndRoot In ndRootList
                ndRoot.Text = LabelVoucherName.Caption
            Next
            Set ndRootList = tmpDOM.selectNodes("//标题")
            For Each eleMent In ndRootList
                eleMent.setAttribute "宽", "500"
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
'    Dim tmpRow As Integer, tmpCol As Integer
'    Dim i As Long, j As Long
'    On Error Resume Next
'    With Me.Voucher
'        tmpRow = .row
'        tmpCol = .col
'        i = .row
'            Select Case strVouchType
'                Case "97"
'                    If tmpRow > 0 Then '
'                        If .bodyText(tmpRow, "cscloser") <> "" Then
'                            SetVouchItemState .ItemState(.colEx, sibody).sFieldName, "B", False
'                        End If '
'                    End If
'                Case "96"
'                    If .row > 1 Then
'                          SetVouchItemState .ItemState(.colEx, sibody).sFieldName, "B", False
'                    Else
'                          SetVouchItemState .ItemState(.colEx, sibody).sFieldName, "B", True
'                    End If
'            End Select
'    End With
'DoExit:
End Sub
 
Private Sub Voucher_SaveSettingEvent(ByVal varDevice As Variant)
    Dim TmpUFTemplate As Object
    Set TmpUFTemplate = CreateObject("UFVoucherServer85.clsVoucherTemplate")
    If TmpUFTemplate.SaveDeviceCapabilities(DBconn.ConnectionString, BillPrnVTID, varDevice) <> 0 Then
        MsgBox GetString("U8.SA.xsglsql.01.frmbillvouch.00361"), vbExclamation 'zh-CN：打印设置保存失败
    End If
    Set TmpUFTemplate = Nothing
End Sub
 

 
'控制界面
Private Sub VS_GotFocus()
    On Error Resume Next
    Me.Voucher.SetFocus
End Sub

Private Sub HS_GotFocus()
    On Error Resume Next
    Me.Voucher.SetFocus
End Sub
 

 
Private Sub AddNewVouch(Optional strOperator As String)
    Dim iElement As IXMLDOMElement
    Dim iAttr As IXMLDOMAttribute
    Dim i As Long
    Dim tmpdate As String
    With Voucher
        Select Case LCase(strOperator)
            Case "verify"
                .headerText("dverifydate") = m_login.CurDate
'                .headerText("checkcode") = m_Login.cUserId
                .headerText("cverifier") = m_login.cUserName
                Exit Sub
            Case "unverify"
                .headerText("dverifydate") = ""
                .headerText("cverifier") = ""
                Exit Sub
            Case "save"
'                 If strVouchType = "95" Then
'                    .headerText("bIWLType") = 1
'                 ElseIf strVouchType = "92" Then
'                    .headerText("bIWLType") = 0
'                 End If
            Case "add", "copy", ""
                If LCase(strOperator) = "copy" Then
                    Call Voucher_headOnEdit(.LookUpArray("cbustype", siHeader))
                End If
                .BodyMaxRows = intBodyMaxRows
                sCurTemplateID = sCurTemplateID2
                .getVoucherDataXML domHead, domBody
                
                clsVoucherCO.AddNew domHead, IIf(LCase(strOperator) = "copy", True, False), domBody
                Me.Voucher.SkipLoadAccessories = False
                .setVoucherDataXML domHead, domBody
                .headerText("cmakerddate") = m_login.CurDate
                .headerText("cvouchtype") = strCardNum
                
                If LCase(strOperator) = "add" Or LCase(strOperator) = "copy" Or LCase(strOperator) = "" Then
                    .headerText("vt_id") = sCurTemplateID       '单据vt_id
                    .headerText("chandler") = ""
                    .headerText("chandlername") = ""
                    tmpdate = .headerText("cdate")
                    .headerText("cdate") = ""
                    .headerText("cdate") = tmpdate
                    tmpdate = .headerText("cmaker")             '制单人
                    .headerText("cmaker") = ""
                    .headerText("cmaker") = tmpdate
                    tmpdate = .headerText("breturnflag")        '红蓝字
                    .headerText("breturnflag") = ""
                    .headerText("breturnflag") = tmpdate
                    
                End If
                
            Case "modify"
                Call Voucher_headOnEdit(.LookUpArray("cbustype", siHeader))
                Select Case strVouchtype
                    Case "05", "06"
                        .BodyMaxRows = 0
                        .getVoucherDataXML domHead, domBody
                        If domBody.selectNodes("//z:row[(@icorid !='' and @icorid !='0')]").length > 0 Then
                            .BodyMaxRows = -1
                        End If

                    Case Else
                        .BodyMaxRows = 0
                End Select
        End Select
        If iVouchState <> 2 Then
            If sCurTemplateID <> "" And sCurTemplateID <> "0" Then
                .headerText("vt_id") = sCurTemplateID
            Else
                .headerText("vt_id") = sCurTemplateID2
            End If
        End If
    End With
End Sub

Private Sub SetCboVtidState()
    If Me.Voucher.VoucherStatus = VSNormalMode Then
        ComboVTID.Visible = True
        ComboDJMB.Visible = False
        Labeldjmb.Caption = "打印模版" 'zh-CN：打印模版
'        Labeldjmb.Caption = GetString("U8.SA.xsglsql.01.frmbillvouch.00398") 'zh-CN：打印模版
    Else
        ComboVTID.Visible = False
        ComboDJMB.Visible = True
        Labeldjmb.Caption = "显示模版" 'zh-CN：显示模版：
'        Labeldjmb.Caption = GetString("U8.SA.xsglsql.01.frmbillvouch.00395") 'zh-CN：显示模版：
    End If
End Sub

Public Property Get UFTaskID() As String
    UFTaskID = m_UFTaskID
End Property
 
Public Property Let UFTaskID(ByVal vNewValue As String)
    m_UFTaskID = vNewValue
End Property

'提示：快捷键所调用的方法
Public Sub setKey(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim strsql As String
    Dim Skey As String
    Dim referPara As UAPVoucherControl85.ReferParameter
    On Error Resume Next
    With Voucher
        If Voucher.VoucherStatus <> VSNormalMode Then
            ''编辑状态下
            Select Case KeyCode
                Case vbKeyF7
                    If tbrvoucher.buttons("fetchprice").Visible And tbrvoucher.buttons("fetchprice").Enabled Then
                        Call g_business.SelectToolbarButton(Me.tbrvoucher.buttons("fetchprice"))
                    End If
                    If Shift = 2 Then
                        If tbrvoucher.buttons("fetchprice").ButtonMenus("rowprice").Visible And tbrvoucher.buttons("fetchprice").ButtonMenus("rowprice").Enabled Then
                            ButtonClick "rowprice", tbrvoucher.buttons("fetchprice").ButtonMenus("rowprice").Tag
                        End If
                    End If
                    If Shift = 4 Then
                        If tbrvoucher.buttons("fetchprice").ButtonMenus("allprice").Visible And tbrvoucher.buttons("fetchprice").ButtonMenus("allprice").Enabled Then
                            ButtonClick "allprice", tbrvoucher.buttons("fetchprice").ButtonMenus("allprice").Tag
                        End If
                    End If
                Case vbKeyG
                    If Shift = 2 Then
                        If tbrvoucher.buttons("refervouch").Visible And tbrvoucher.buttons("refervouch").Enabled Then
                            If tbrvoucher.buttons("refervouch").Style = tbrDropdown Then
                                Call g_business.SelectToolbarButton(Me.tbrvoucher.buttons("refervouch"))
'                            Else
'                                Call ButtonClick("refervouch", "")
                            End If
                        End If
                    End If
                Case vbKeyZ
                    If Shift = 2 Then
                        Call ButtonClick("cancel", tbrvoucher.buttons("cancel").Caption)
                    End If
                Case vbKeyK
                    If Shift = 2 Then
                        If tbrvoucher.buttons("copy").ButtonMenus("copyline").Visible And tbrvoucher.buttons("copy").ButtonMenus("copyline").Enabled Then
                            Call ButtonClick("copyline", tbrvoucher.buttons("copy").ButtonMenus("copyline").Tag)
                        End If
                        If tbrvoucher.buttons("copyline").Visible And tbrvoucher.buttons("copyline").Enabled Then
                            Call ButtonClick("copyline", tbrvoucher.buttons("copyline").Caption)
                        End If
                    End If
                Case vbKeyF6
                    If tbrvoucher.buttons("save").Visible And tbrvoucher.buttons("save").Enabled Then
                        Call ButtonClick("save", tbrvoucher.buttons("save").Caption)
                    End If
                Case vbKeyS
                    If Shift = 2 Then
                        If tbrvoucher.buttons("save").Visible And tbrvoucher.buttons("save").Enabled Then
                            Call ButtonClick("savenew", tbrvoucher.buttons("save").Caption)
                        End If
                    End If
                Case vbKeyR
                    If Shift = 2 Then
                       If Not .BodyMaxRows = -1 Then
                            Call ButtonClick("copyline", tbrvoucher.buttons("copyline").Caption)
                        End If
                    End If
                Case vbKeyN
                    If Shift = 2 Then
                        If tbrvoucher.buttons("addline").Visible And tbrvoucher.buttons("addline").Enabled Then
                            Call ButtonClick("addline", tbrvoucher.buttons("addline").Caption)
                        End If
                    End If
                Case vbKeyD
                    If Shift = 2 Then
                        If tbrvoucher.buttons("delline").Visible And tbrvoucher.buttons("delline").Enabled Then Call ButtonClick("delline", tbrvoucher.buttons("delline").Caption)
                    End If
                Case vbKeyA
                    Select Case strVouchtype
                        Case "05", "06", "26", "27", "28", "29"
                        Case Else
                            Exit Sub
                    End Select
                    If bReturnFlag = True Then Exit Sub
                    If Shift = 2 Then
                        KeyCode = 0
                    End If
                Case vbKeyB
                    If Shift = 2 Then
                        KeyCode = 0
                    End If
                Case vbKeyQ
                    If Shift = 2 Then
                        KeyCode = 0
                    End If
                Case vbKeyO
                    If Shift = 2 Then
                        KeyCode = 0
                    End If
                Case vbKeyE
     
                Case vbKeyF2
                    Dim strDate As String
                    If Shift = 2 Then
                    
                    End If
                    
            End Select

        Else
            ''非编辑状态
            Select Case KeyCode
                Case vbKeyF5
                    If Shift = 0 Then '
                        If tbrvoucher.buttons("add").Visible And tbrvoucher.buttons("add").Enabled Then
                            Call ButtonClick("add", tbrvoucher.buttons("add").Caption)
                        End If
                    End If
                    If Shift = 2 Then '
                        If tbrvoucher.buttons("copy").Enabled And tbrvoucher.buttons("copy").ButtonMenus("copyvoucher").Visible And tbrvoucher.buttons("copy").ButtonMenus("copyvoucher").Enabled Then
                            Call ButtonClick("copyvoucher", tbrvoucher.buttons("copy").ButtonMenus("copyvoucher").Tag)
                        End If
                    End If
                Case vbKeyO
                    If Shift = 4 Then '
                        If tbrvoucher.buttons("openorder").Visible And tbrvoucher.buttons("openorder").Enabled Then
                            Call ButtonClick("openorder", tbrvoucher.buttons("openorder").Caption)
                        End If
                    End If
'                Case vbKeyK
'                    If Shift = 2 Then
'                        If tbrvoucher.Buttons("viewverify").Visible And tbrvoucher.Buttons("viewverify").Enabled Then
'                            Call ButtonClick("viewverify", tbrvoucher.Buttons("viewverify").Caption)
'                        End If
'                    End If
                Case vbKeyR
                    If Shift = 2 Then
                        If tbrvoucher.buttons("refresh").Visible And tbrvoucher.buttons("refresh").Enabled Then
                            Call ButtonClick("refresh", tbrvoucher.buttons("refresh").Caption)
                        End If
                    End If
                Case vbKeyL
                    If strVouchtype = "97" Then
                        If Shift = 2 Then
                            If tbrvoucher.buttons("lock").Visible And tbrvoucher.buttons("lock").Enabled Then
                                Call ButtonClick("lock", tbrvoucher.buttons("lock").Caption)
                            End If
                        End If
                        If Shift = 4 Then
                            If tbrvoucher.buttons("unlock").Visible And tbrvoucher.buttons("unlock").Enabled Then
                                Call ButtonClick("unlock", tbrvoucher.buttons("unlock").Caption)
                            End If
                        End If
                    End If
                Case vbKeyJ
                    If Shift = 2 Then
                        If tbrvoucher.buttons("submit").Visible And tbrvoucher.buttons("submit").Enabled Then
                            Call ButtonClick("submit", tbrvoucher.buttons("submit").Caption)
                        End If
                    End If
                    If Shift = 4 Then
                        If tbrvoucher.buttons("unsubmit").Visible And tbrvoucher.buttons("unsubmit").Enabled Then
                            Call ButtonClick("unsubmit", tbrvoucher.buttons("unsubmit").Caption)
                        End If
                    End If
                Case vbKeyU
                    If Shift = 2 Then
                        If tbrvoucher.buttons("verify").Visible And tbrvoucher.buttons("verify").Enabled Then
                            Call ButtonClick("verify", tbrvoucher.buttons("verify").Caption)
                        End If
                    End If
                    If Shift = 4 Then
                        If tbrvoucher.buttons("unverify").Visible And tbrvoucher.buttons("unverify").Enabled Then
                            Call ButtonClick("unverify", tbrvoucher.buttons("unverify").Caption)
                        End If
                    End If
                Case vbKeyC
                    If Shift = 4 Then
                        If tbrvoucher.buttons("closeorder").Visible And tbrvoucher.buttons("closeorder").Enabled Then
                           Call ButtonClick("closeorder", tbrvoucher.buttons("closeorder").Caption)
                        End If
                    End If
                Case vbKeyF10
                    If tbrvoucher.buttons("copy").Visible And tbrvoucher.buttons("copy").Enabled Then
                           Call ButtonClick("copy", tbrvoucher.buttons("copy").Caption)
                    End If
                Case vbKeyN
                    If Shift = 2 And (strVouchtype = "26" Or strVouchtype = "27" Or strVouchtype = "28" Or strVouchtype = "29") Then
                        If tbrvoucher.buttons("nowpay").Visible And tbrvoucher.buttons("nowpay").Enabled Then
                            Call ButtonClick("nowpay", tbrvoucher.buttons("nowpay").Caption)
                        End If
                    End If
                Case vbKeyPageDown
'                    If iShowMode = 1 Then Exit Sub
                    If Shift = 0 Then
                        If tbrvoucher.buttons("tonext").Visible And tbrvoucher.buttons("tonext").Enabled Then
                            Call ButtonClick("tonext", tbrvoucher.buttons("tonext").Caption)
                        End If
                    End If
                    If Shift = 4 Then  'alt
                        If tbrvoucher.buttons("tolast").Visible And tbrvoucher.buttons("tolast").Enabled Then
                            Call ButtonClick("tolast", tbrvoucher.buttons("tolast").Caption)
                        End If
                    End If
                Case vbKeyPageUp
'                    If iShowMode = 1 Then Exit Sub
                    If Shift = 0 Then
                        If tbrvoucher.buttons("toprevious").Visible And tbrvoucher.buttons("toprevious").Enabled Then
                            Call ButtonClick("toprevious", tbrvoucher.buttons("toprevious").Caption)
                        End If
                    End If
                    If Shift = 4 Then
                        If tbrvoucher.buttons("tofirst").Visible And tbrvoucher.buttons("tofirst").Enabled Then
                            Call ButtonClick("tofirst", tbrvoucher.buttons("tofirst").Caption)
                        End If
                    End If
                Case vbKeyF5
'                    If iShowMode = 1 Then Exit Sub
                    If Shift = 0 Then
                        If tbrvoucher.buttons("add").Visible And tbrvoucher.buttons("add").Enabled Then
                            Call ButtonClick("add", tbrvoucher.buttons("add").Caption)
                        End If
                    End If
                Case vbKeyF8
'                    If iShowMode = 1 Then Exit Sub
                    If Shift = 0 Then
                        If tbrvoucher.buttons("modify").Visible And tbrvoucher.buttons("modify").Enabled Then
                            Call ButtonClick("modify", tbrvoucher.buttons("modify").Caption)
                        End If
                    End If
                Case vbKeyP         ''打印
'                    If iShowMode = 1 Then Exit Sub
                    If Shift = 2 Then
                        If tbrvoucher.buttons("print").Visible And tbrvoucher.buttons("print").Enabled Then
                            Call ButtonClick("print", tbrvoucher.buttons("print").Caption)
                        End If
                    End If
                Case vbKeyW         ''打印
'                    If iShowMode = 1 Then Exit Sub
                    If Shift = 2 Then
                        If tbrvoucher.buttons("preview").Visible And tbrvoucher.buttons("preview").Enabled Then
                            Call ButtonClick("preview", tbrvoucher.buttons("preview").Caption)
                        End If
                    End If
                Case vbKeyE         ''打印
'                    If iShowMode = 1 Then Exit Sub
                    If Shift = 4 Then
                        If tbrvoucher.buttons("output").Visible And tbrvoucher.buttons("output").Enabled Then
                            Call ButtonClick("output", tbrvoucher.buttons("output").Caption)
                        End If
                    End If
'                Case vbKeyF4        ''退出
'                    If Shift = 2 Then
'                        If tbrvoucher.Buttons("Exit").Visible And tbrvoucher.Buttons("Exit").Enabled Then
'                           Call ButtonClick("Exit", "")
'                        End If
'                    End If
                Case vbKeyF3        ''定位
'                    If iShowMode = 1 Then Exit Sub
                    If tbrvoucher.buttons("locate").Visible And tbrvoucher.buttons("locate").Enabled Then
                       Call ButtonClick("locate", tbrvoucher.buttons("locate").Caption)
                    End If

                Case vbKeyDelete
'                    If iShowMode = 1 Then Exit Sub
                    If tbrvoucher.buttons("delete").Visible And tbrvoucher.buttons("delete").Enabled Then
                       Call ButtonClick("delete", tbrvoucher.buttons("delete").Caption)
                    End If

            End Select
        End If
        If KeyCode = vbKeyF Then
            If Shift = 2 Then
                ButtonClick "lookrow", "行定位"
            End If
        End If
    End With
End Sub

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
            MsgBox "在错误处理中无法生成错误生成DOM对象！"
            MsgBox strMsg
            Exit Function
        End If
        Screen.MousePointer = vbDefault
    Else
        ''正常的错误
        If Len(Trim(strMsg)) > 0 Then
            MsgBox strMsg
        End If
        
    End If
    Set tmpDOM = Nothing
    ShowErrDom = True
    
    Screen.MousePointer = vbDefault
    Exit Function
DoERR:
    MsgBox "处理错误信息时发生错误：" & Err.Description
    Set tmpDOM = Nothing
    ShowErrDom = False
    Screen.MousePointer = vbDefault
End Function

 
 
Private Function CheckDJMBAuth(strVTID As String, strOprate As String) As Boolean
'    CheckDJMBAuth = clsSAWeb.clsAuth.IsHoldAuth("DJMB", strVTID, , strOprate)
    CheckDJMBAuth = clsSAWeb.CheckDJMBAuth(strVTID, strOprate)
End Function

Private Function IsHoldAuth(cmaker As String, strOprate As String) As Boolean
    IsHoldAuth = True
    If IsAuthControl(DBconn, "voucher", strCardNum, "bmaker") Then
        IsHoldAuth = clsSAWeb.clsAuth.IsHoldAuth("user", cmaker, , strOprate)
    End If
End Function

Private Function IsAuthControl(DBconn As ADODB.Connection, strFormType As String, strKey As String, authName As String) As Boolean
    Dim Rst As New ADODB.Recordset
    On Error GoTo hErr
    Rst.Open "select * from sa_authconfig where formtype=N'" + strFormType + "' and ckey=N'" + strKey + "' and authname='" & authName & "' and bcontrol=1", DBconn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not Rst.EOF And Not Rst.BOF Then
        IsAuthControl = True
    End If
    Rst.Close
    Set Rst = Nothing
    Exit Function
hErr:

End Function

''更改单据模版for增加，复制
Private Function ChangeDJMBForEdit() As Boolean
    
    With Me.Voucher
        If CheckDJMBAuth(.headerText("vt_id"), "W") = False Then
            If sTemplateID = "0" Then
                MsgBox "无可以使用的模版,请检查模版权限"
            Else
                ChangeDJMBForEdit = ChangeTempaltes(sTemplateID)
            End If
        Else
            ChangeDJMBForEdit = True
        End If
    End With
End Function

'提示：如何更改voucher caption 的颜色及显示名称的方法，可以根据项目需要进行修改
Private Sub ChangeCaptionCol()
    On Error Resume Next
    
    With Me.Voucher
        Me.LabelVoucherName.ForeColor = .TitleForeColor
        Me.LabelVoucherName.Font.Name = .TitleFont.Name
        Me.LabelVoucherName.Font.Bold = .TitleFont.Bold
        Me.LabelVoucherName.Font.Italic = .TitleFont.Italic
        Me.LabelVoucherName.Font.Underline = .TitleFont.Underline
        Me.LabelVoucherName.Font.Size = 15 ' .TitleFont.Size
        If bFirst = True Then
            If Left(Me.LabelVoucherName.Caption, Len(GetString("U8.SA.xsglsql.frmBillVouchNew.02941"))) <> GetString("U8.SA.xsglsql.frmBillVouchNew.02941") And Left(Me.LabelVoucherName.Caption, Len(GetString("U8.SA.xsglsql.frmBillVouchNew.02943"))) <> GetString("U8.SA.xsglsql.frmBillVouchNew.02943") Then
                    Me.LabelVoucherName.Caption = GetString("U8.SA.xsglsql.frmBillVouchNew.02941") & Me.LabelVoucherName.Caption
            End If
            Exit Sub
        End If
        Select Case strVouchtype
            Case "26", "27", "28", "29"
                If .headerText("breturnflag") = "1" Or LCase(.headerText("breturnflag")) = "true" Or (.headerText("breturnflag") = "" And bReturnFlag = True) Then
                    Me.LabelVoucherName.ForeColor = vbRed
                Else
                    Me.LabelVoucherName.ForeColor = .TitleForeColor 'vbBlack
                End If
            Case "92"
        End Select
    End With
End Sub
 
 
''更改单据项目到原始状态
Private Function SetOriItemState(cardsection As String, strFieldName As String) As Boolean
    Dim sFilter As String
    Dim bCanModify As Boolean
    On Error GoTo Err
    RstTemplate.Filter = ""
    sFilter = " cardsection ='" + cardsection + "' and fieldname='" + strFieldName + "'"
    RstTemplate.Filter = sFilter
    If Not RstTemplate.EOF Then
        If RstTemplate("CanModify") = True Or RstTemplate("CanModify") = 1 Then
            bCanModify = True
        Else
            bCanModify = False
        End If
        With Me.Voucher
            If Not .ItemState(strFieldName, IIf(LCase(cardsection) = "b", sibody, siHeader)) Is Nothing Then
                If .ItemState(strFieldName, IIf(LCase(cardsection) = "b", sibody, siHeader)).bCanModify <> bCanModify Then
                    If LCase(cardsection) = "t" Then
                        .EnableHead strFieldName, bCanModify
                    Else
                        If Not .ItemState(strFieldName, IIf(LCase(cardsection) = "b", sibody, siHeader)) Is Nothing Then
                            .ItemState(strFieldName, IIf(LCase(cardsection) = "b", sibody, siHeader)).bCanModify = bCanModify
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

'设置单据控件项目写状态
Private Function SetVouchItemState(strFieldName As String, cardsection As String, bCanModify As Boolean) As Boolean
    On Error GoTo Err
    With Me.Voucher
        If Not .ItemState(strFieldName, IIf(LCase(cardsection) = "b", sibody, siHeader)) Is Nothing Then
            If .ItemState(strFieldName, IIf(LCase(cardsection) = "b", sibody, siHeader)).bCanModify <> bCanModify Then
                If LCase(cardsection) = "t" Then
                    .EnableHead strFieldName, bCanModify
                Else
                    If Not .ItemState(strFieldName, IIf(LCase(cardsection) = "b", sibody, siHeader)) Is Nothing Then
                        .ItemState(strFieldName, IIf(LCase(cardsection) = "b", sibody, siHeader)).bCanModify = bCanModify
                    End If
                End If
            End If
        End If
    End With
    Exit Function
Err:
    MsgBox Err.Description
End Function

'警告：根据单据模板ID,取得单据标识(cardNumber)
Private Sub getCardNumber(nvtid)
    Dim rstTmp As New ADODB.Recordset
    rstTmp.Open "select VT_CardNumber from vouchertemplates where VT_ID =" & nvtid, DBconn, adOpenForwardOnly, adLockReadOnly
    If Not rstTmp.EOF Then
        strCardNum = rstTmp(0)
    End If
    rstTmp.Close
    Set rstTmp = Nothing
End Sub
 
'警告：外部可以调用内部函数——加载单据
Public Sub loaDVouch(vid As Variant)
    Call LoadVoucher("", vid)
End Sub


'警告：外部可以调用内部函数——表头参照
Public Sub VouchHeadCellCheck(Index As Variant, RetValue As String, bChanged As UAPVoucherControl85.CheckRet)
    Index = Voucher.LookUpArray(LCase(Index), siHeader)
    Dim referPara As UAPVoucherControl85.ReferParameter
    Call Voucher_headCellCheck(Index, RetValue, bChanged, referPara)
    Voucher.ProtectUnload2
End Sub

'警告：外部可以调用内部函数——将控件传给外部控件
Public Function GetVoucherObject() As Object
    Set GetVoucherObject = Me.Voucher
End Function

'警告：外部可以调用内部函数——获取单据的编辑状态,提供给外部使用
Public Function GetVouchState() As Integer
    GetVouchState = iVouchState
End Function

'警告：外部可以调用内部函数——获取单据表体某个单元的值
Private Function GetBodyRefVal(Skey As String, row As Long) As String
    Dim Obj As Object
    Dim Index As Long
    ' 得到表体对象
    Set Obj = Me.Voucher.GetBodyObject()
    ' 得到关键字对应的Index
    Index = Me.Voucher.LookUpArrayFromKey(Skey, sibody)
    GetBodyRefVal = Obj.TextMatrix(row, Index)
End Function


'提示：联查凭证示例,可以根据项目情况进行取舍
Private Sub Find_GL_accvouch()
    Dim rdst1 As New ADODB.Recordset
    Dim rdst2 As New ADODB.Recordset
    On Error GoTo Err
    Select Case strVouchtype

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

' 2006/03/08   增加单据附件功能
Private Function SetAttachXML(oDomH As DOMDocument) As Boolean
    Dim strXml As String
    Dim errMsg As String
    Dim NodeData As IXMLDOMCDATASection
    Dim nd As IXMLDOMNode, ndRS As IXMLDOMNode
    Dim NdList As IXMLDOMNodeList

    strXml = Voucher.GetAccessoriesInfo(errMsg)
    If errMsg <> "" Then
        MsgBox errMsg
        Exit Function
    End If
    If strXml = "" Then
        SetAttachXML = True
        Exit Function
    End If
    Set ndRS = oDomH.selectSingleNode("//rs:data")
    Set NdList = oDomH.selectNodes("//rs:data/voucherattached")
    For Each nd In NdList
        ndRS.removeChild nd
    Next
    Set NodeData = oDomH.createCDATASection(strXml)
    Set nd = oDomH.createElement("voucherattached")
    nd.appendChild NodeData
    ndRS.appendChild nd



    SetAttachXML = True
End Function


Private Function SetVoucherDataSource()
    Dim m_oDataSource As Object
    Set m_oDataSource = CreateObject("IDataSource.DefaultDataSource")
    If m_oDataSource Is Nothing Then
        MsgBox "无法创建m_oDataSource对象!"
        Exit Function
    End If
    Set m_oDataSource.SetLogin = m_login
    Set Me.Voucher.SetDataSource = m_oDataSource
End Function

Private Sub RegisterMessage()
    Set m_mht = New UFPortalProxyMessage.IMessageHandler
    m_mht.MessageType = "DocAuditComplete"
    If Not g_business Is Nothing Then
        Call g_business.RegisterMessageHandler(m_mht)
    End If
End Sub






'生成科目...预置表体RecordSet
Private Function ProcAddCodeListRS(ByRef rs As ADODB.Recordset)
    Dim rds As New ADODB.Recordset
    Dim rdSum As ADODB.Recordset
    
    Dim strsql As String
    strsql = "select m.ccode,c.ccode_name from MT_code m inner join code c on m.ccode=c.ccode where bend=1"
    rds.Open strsql, DBconn, adOpenStatic, adLockReadOnly
    
    Do While Not rds.EOF
        rs.AddNew
        rs("ccode") = rds("ccode")
        rs("ccode_name") = rds("ccode_name")
        rs("editprop") = "A"
        rds.MoveNext
    Loop
    
    Set rds = Nothing
End Function









 

 



'警告: 向门户状态栏上传递信息
Private Sub SendMessgeToPortal(strMessageType As String)
    Dim strID As String
    Dim strMaker As String
    Dim strKey As String
    Dim strCode As String
    
    On Error GoTo Errhandle
    strKey = strCardNum
    Select Case strCardNum
        Case Else
            strID = "id"
            strCode = "ccode"
    End Select

    If strCardNum = "13" Then
        If bReturnFlag Then
            strKey = strCardNum & "Red"
        End If
    End If
    strMaker = Voucher.headerText("cmaker")
    If strMaker = "" Then strMaker = m_login.cUserName
    SendPortalMessage sGuid, strKey, Voucher.headerText(strID), strMessageType, strMaker, Voucher.headerText("ufts"), Voucher.headerText(strCode), strVouchtype, bReturnFlag   'Voucher.headerText("ufts")
    Exit Sub
Errhandle:
End Sub

Public Property Get FormVisible() As Boolean
    FormVisible = m_FormVisible
End Property

Public Property Let FormVisible(ByVal vNewValue As Boolean)
    m_FormVisible = vNewValue
End Property

Public Property Get strToolBarName() As String
    strToolBarName = m_strToolBarName
End Property

Public Property Let strToolBarName(ByVal vNewValue As String)
    m_strToolBarName = vNewValue
End Property

Public Property Get strCardNum() As String
    strCardNum = m_strCardNum
End Property

Public Property Let strCardNum(ByVal vNewValue As String)
    m_strCardNum = vNewValue
End Property

Public Property Get strVouchtype() As String
    strVouchtype = m_strVouchType
End Property

Public Property Let strVouchtype(ByVal vNewValue As String)
    m_strVouchType = vNewValue
End Property
Public Property Get bFirst() As Boolean
    bFirst = m_bFirst
End Property

Public Property Let bFirst(ByVal vNewValue As Boolean)
    m_bFirst = vNewValue
End Property

Public Property Get bReturnFlag() As Boolean
    bReturnFlag = m_bReturnFlag
End Property

Public Property Let bReturnFlag(ByVal vNewValue As Boolean)
    m_bReturnFlag = vNewValue
End Property

Public Property Get strHelpId() As String
    strHelpId = m_strHelpId
End Property

Public Property Let strHelpId(ByVal vNewValue As String)
    m_strHelpId = vNewValue
End Property


'注销消息函数
Private Sub UnRegisterMessage()
    If m_mht Is Nothing Then Exit Sub
    If Not g_business Is Nothing Then
        Call g_business.UnregisterMessageHandler(m_mht)
    End If
    Set m_mht = Nothing
End Sub

Private Sub GetDomVtid()
    Dim tmprst As New ADODB.Recordset
    Dim sWhere As String
    Dim strAuth As String
    Dim strsql As String
    strAuth = clsSAWeb.clsAuth.getAuthString("DJMB")
    If strAuth = "1=2" Then Exit Sub
    'LDX    2009-06-04  Add Beg
'    If strVouchType = 98 Then
'        sWhere = " VT_CardNumber = '" + strCardNum + "' and vt_ID in (select left(b.ccode,6) from mt_baseset a left join mt_basesets b on a.id=b.id where a.cvouchtype='11') "
'    Else
        sWhere = " VT_CardNumber = '" + strCardNum + "'"
'    End If
    'LDX    2009-06-04  Add End
    strsql = "SELECT VT_Name,VT_ID,isnull(VT_PrintTemplID,DEF_ID_PRN) as printid,isnull(VT_TemplateMode,0) as VT_TemplateMode From vouchertemplates inner join vouchers_base on vouchertemplates.VT_CardNumber=vouchers_base.cardnumber WHERE (" & sWhere & ")  " & IIf(strAuth <> "", "and vt_id in (" & strAuth & ")", "") + " order by VT_CardNumber,case when vt_id=def_id then vt_id else null end desc"    'AND (VT_TemplateMode = 0)
    tmprst.CursorLocation = adUseClient
    tmprst.Open ConvertSQLString(strsql), DBconn, adOpenForwardOnly, adLockReadOnly
    fillComBol True, tmprst
    fillComBol False, tmprst
    tmprst.Close
    Set tmprst = Nothing
End Sub

Private Function GetDispCobIndex(lngDispId As Long) As Integer
    Dim i As Integer
    If Me.ComboDJMB.ListCount <= 0 Then Exit Function
    GetDispCobIndex = 0
    For i = 0 To UBound(vtidDJMB)
        If vtidDJMB(i) = lngDispId Then
            GetDispCobIndex = i
            Exit Function
        End If
    Next
End Function
Private Sub DWINfor()
    
End Sub

Private Function BeforeEditVoucher() As Boolean
    Dim errMsg As String
    Dim GetvouchNO As String
    BeforeEditVoucher = clsVoucherCO.GetVoucherNO(domHead, GetvouchNO, errMsg, DomFormat, True)
    If BeforeEditVoucher = False Then
        MsgBox errMsg, vbExclamation
    Else
        Me.Voucher.SetBillNumberRule DomFormat.xml
        ''初始化参照控件
        clsRefer.SetLogin m_login
    End If
End Function

 
'by 20100120 ahzzd 取得审批流需要的需要的相关信息
'cardNum  单据编号
'Mid    单据主表ID值
'Tblname  单据主表名
'MCode    单据编码
Private Sub GetCardNumberMid(CardNum As String, Mid As String, Tblname As String, Optional MCode As String)
    Select Case strVouchtype
    Case "26", "27", "28", "29" '发票
        CardNum = "07": Mid = Voucher.headerText("sbvid"): Tblname = "SaleBillVouch": MCode = Voucher.headerText("csbvcode")
    Case "05", "06", "00" '发货
        CardNum = "01": Mid = Voucher.headerText("dlid"): Tblname = "DispatchList": MCode = Voucher.headerText("cdlcode")
    Case "97" '订单
        CardNum = "17": Mid = Voucher.headerText("id"): Tblname = "SO_SOMain": MCode = Voucher.headerText("csocode")
    Case "16" '报价单
        CardNum = "16": Mid = Voucher.headerText("id"): Tblname = "SA_QuoMain": MCode = Voucher.headerText("ccode")
    Case "98" '代垫
        CardNum = "08": Mid = Voucher.headerText("id"): Tblname = "ExpenseVouch": MCode = Voucher.headerText("cevcode")
    Case "99" '费用支出
        CardNum = "09": Mid = Voucher.headerText("id"): Tblname = "SalePayVouch": MCode = Voucher.headerText("cspvcode")
    Case "07" '结算
        CardNum = "02": Mid = Voucher.headerText("id"): Tblname = "SA_SettleVouch": MCode = Voucher.headerText("ccode")
    Case "00"
        CardNum = "28": Mid = Voucher.headerText("dlid"): Tblname = "DispatchList": MCode = Voucher.headerText("cdlcode")
'    Case "95", "92" '包装物
'        cardNum = "10": Mid = "autoid"
    Case "EFBWGL020301"
        CardNum = "EFBWGL020301": Mid = Voucher.headerText("id"): Tblname = "efbwgl_SelDeclare": MCode = Voucher.headerText("ccode")
    Case Else
        CardNum = strVouchtype
        Mid = Voucher.headerText("id")
        MCode = Voucher.headerText("ccode")
    End Select
End Sub






'定位一张单据
'by  ahzzd 20100118
Private Function SeekOneVoucher(strVouchtype As String, Optional strFirst As Boolean = True) As String
'ahzzd 实现思路
'1 、取得单据列表过滤的关键字，并显示过滤界面
'2 、根据过滤对象得到过滤字符串 查询单据列表视图
'3 、取得查询结果的第一个主表ID 值
Dim List_ckey As String
Dim Listfrom As String
Dim Filtername As String
Dim objFilter As New UFGeneralFilter.FilterSrv
Dim Rst As New ADODB.Recordset
Dim strsql As String
On Error GoTo Err0
    SeekOneVoucher = "0"

'    List_ckey = SA_VoucherListConfigDom.selectNodes("//z:row[@ckey='" + strCardNum + "']")
'    List_ckey = SA_VoucherListConfigDom.selectSingleNode("//z:row").Attributes.getNamedItem("listfrom").nodeValue
    Listfrom = SA_VoucherListConfigDom.selectSingleNode("//z:row").Attributes.getNamedItem("listfrom").nodeValue
    Filtername = SA_VoucherListConfigDom.selectSingleNode("//z:row").Attributes.getNamedItem("filtername").nodeValue
    objFilter.OpenFilter m_login, Filtername
  
  
  

    strsql = " select * from " & Listfrom
    If Len(Trim(objFilter.GetSQLWhere)) > 0 Then
        strsql = strsql & " where  cvouchtype ='" & strCardNum & "' and " & objFilter.GetSQLWhere
    End If
    
    Rst.Open strsql, DBconn, 3, 4
    
    
    If Not Rst.EOF Then
        SeekOneVoucher = Rst.Fields("id").value
    End If
 Exit Function
Err0:
    SeekOneVoucher = "0"


End Function

Private Sub CheckAuthAfterDraft()
    Dim blnChange As UAPVoucherControl85.CheckRet
    Dim par As UAPVoucherControl85.ReferParameter
    If Voucher.headerText("ccuscode") <> "" And clsSAWeb.bAuth_Cus And Not m_login.isAdmin Then
        If Not clsSAWeb.clsAuth.IsHoldAuth("Customer", Voucher.headerText("ccuscode"), , "W") Then
            Voucher.headerText("ccusabbname") = ""
            Voucher_headCellCheck "ccusabbname", "", blnChange, par
        End If
    End If
    If Voucher.headerText("cdepcode") <> "" And clsSAWeb.bAuth_dep And Not m_login.isAdmin Then
        If Not clsSAWeb.clsAuth.IsHoldAuth("department", Voucher.headerText("cdepcode"), , "W") Then
            Voucher.headerText("cdepname") = ""
            Voucher_headCellCheck "cdepname", "", blnChange, par
        End If
    End If
    If strVouchtype = "92" Or strVouchtype = "95" Then
        If Voucher.headerText("chandler") <> "" And clsSAWeb.bAuth_Per And Not m_login.isAdmin Then
            If Not clsSAWeb.clsAuth.IsHoldAuth("person", Voucher.headerText("chandler"), , "W") Then
                Voucher.headerText("cpersonname") = ""
                Voucher_headCellCheck "cpersonname", "", blnChange, par
            End If
        End If
        If Voucher.headerText("cinvcode") <> "" And clsSAWeb.bAuth_Inv And Not m_login.isAdmin Then
            If Not clsSAWeb.clsAuth.IsHoldAuth("inventory", Voucher.headerText("cinvcode"), , "W") Then
                Voucher.headerText("cinvname") = ""
                Voucher_headCellCheck "cinvname", "", blnChange, par
            End If
        End If
    Else
        If Voucher.headerText("cpersoncode") <> "" And clsSAWeb.bAuth_Per And Not m_login.isAdmin Then
            If Not clsSAWeb.clsAuth.IsHoldAuth("person", Voucher.headerText("cpersoncode"), , "W") Then
                Voucher.headerText("cpersonname") = ""
                Voucher_headCellCheck "cpersonname", "", blnChange, par
            End If
        End If
        Dim i As Long
        For i = 1 To Voucher.BodyRows
            If Voucher.bodyText(i, "cwhcode") <> "" And clsSAWeb.bAuth_Wh And Not m_login.isAdmin Then
                If Not clsSAWeb.clsAuth.IsHoldAuth("warehouse", Voucher.bodyText(i, "cwhcode"), , "W") Then
                    Voucher.bodyText(i, "cwhname") = ""
                    Voucher_bodyCellCheck "", blnChange, i, Voucher.LookUpArray("cwhname", sibody), par
                End If
            End If
            If Voucher.bodyText(i, "cinvcode") <> "" And clsSAWeb.bAuth_Inv And Not m_login.isAdmin Then
                If Not clsSAWeb.clsAuth.IsHoldAuth("inventory", Voucher.bodyText(i, "cinvcode"), , "W") Then
                    Voucher.bodyText(i, "cinvcode") = ""
                    Voucher_bodyCellCheck "", blnChange, i, Voucher.LookUpArray("cinvname", sibody), par
                End If
            End If
            Voucher.bodyText(i, "natostatus") = ""  ''重新选配
            Voucher.bodyText(i, "cconfigstatus") = "未选配"  ''重新选配
            
        Next
    End If
End Sub

Private Sub SendShowViewMessage(sViewID As String, Optional ByVal strMessageType As String = "SHOWVIEW")
    'sViewID:="UFIDA.U8.Audit.AuditViews.TreatTaskViewPart",审批视图,
    'sViewID:="UFIDA.U8.Audit.AuditHistoryView",审批进程表查,审时调用
    'SHOWVIEW显示视图，HIDEVIEW隐藏视图
    Dim MyStrCardNum As String
    If strCardNum = "01" Then
        MyStrCardNum = "01"
    Else
        MyStrCardNum = strCardNum
    End If
    ShowWorkFlowView sGuid, MyStrCardNum, sViewID, strMessageType
End Sub

Public Sub ShowWorkFlowView(m_strFormGuid As String, strCardNumber As String, sViewID As String, Optional ByVal strMessageType As String = "SHOWVIEW")
    'sViewID:="UFIDA.U8.Audit.AuditViews.TreatTaskViewPart",审批视图,
    'sViewID:="UFIDA.U8.Audit.AuditHistoryView",审批进程表查,审时调用
    'SHOWVIEW显示视图，HIDEVIEW隐藏视图
    Dim tsb As Object
    Dim strXml As String
    If Not (g_business Is Nothing) Then
        Set tsb = g_business.GetToolbarSubjectEx(m_strFormGuid)
    End If
    Debug.Print m_strFormGuid
    strXml = ""
    strXml = strXml & "<Message type='" & strMessageType & "'>"
    strXml = strXml & "   <Selection context= '1K:" + strCardNumber + " '>"
    strXml = strXml & "      <Element typeName = 'ViewPart' viewID = '" & sViewID & "'  isFirstElement = 'true'/> "
    strXml = strXml & "   </Selection>"
    strXml = strXml & "</Message>"
    If Not (tsb Is Nothing) Then
           Call tsb.TransMessage(m_strFormGuid, strXml)
    End If
 
    If Not tsb Is Nothing Then Set tsb = Nothing
   
End Sub

'草稿
Private Sub m_oHelper_LoadFromTemplate(ByVal enumType As VoucherHelper.TemplateModes, ByVal TemplateID As String, oDOmHead As Variant, oDomBody As Variant, oOtherDom As Variant)
    Dim ele  As IXMLDOMElement
    
    If enumType = DraftMode Then
        m_sCurrentDraftID = TemplateID
    End If
    If Me.Voucher.VoucherStatus = VSNormalMode And enumType = DraftMode Then
        bnewDraft = True
        ButtonClick "add", ""
        bnewDraft = False
    End If
    Voucher.StopSetDefaultValue = True
    Voucher.SkipLoadAccessories = True
    
    For Each ele In oDomBody.selectNodes("//z:row")    ''Clin 09-12-15 置模版表体标识，否则不能保存表体数据
        ele.setAttribute "editprop", "A"
    Next
    
    Voucher.setVoucherDataXML oDOmHead, oDomBody
    Voucher.StopSetDefaultValue = False
    Voucher.SkipLoadAccessories = False
    CheckAuthAfterDraft
    If enumType = TemplateMode Then
        AddNewVouch "copy"
        Call Voucher_BillNumberChecksucceed
        VouchHeadCellCheck Voucher.LookUpArray("ddate", siHeader), Voucher.headerText("ddate"), success
    End If
End Sub

'打开草稿模板
Private Sub OpenFromDraft(ByVal nMode As TemplateModes)
    Dim StrDraft As String
    Select Case strVouchtype
        Case "26" 'zp
            StrDraft = IIf(bReturnFlag, GetString("U8.SA.xsglsql.frmMain.00518"), GetString("U8.SA.xsglsql.frmMain.00516"))
        Case "27" 'pp
            StrDraft = IIf(bReturnFlag, GetString("U8.SA.xsglsql.frmMain.00519"), GetString("U8.SA.xsglsql.frmMain.00517"))
        Case "28" 'dbd
            StrDraft = IIf(bReturnFlag, GetString("U8.SA.xsglsql.01.frmbillvouch.00379"), GetString("U8.SA.xsglsql.01.frmbillvouch.00378"))
        Case "29" 'rb
            StrDraft = IIf(bReturnFlag, GetString("U8.SA.xsglsql.frmMain.00528"), GetString("U8.SA.xsglsql.frmMain.00527"))
        Case Else
            StrDraft = Voucher.TitleCaption
    End Select
    
    Call m_oHelper.GetDraftList(nMode, m_strCardNum, StrDraft)
    
End Sub
'存储草稿模板
Private Sub SaveAsDraft(ByVal nMode As TemplateModes)
    
    Dim StrDraft As String
    Dim temp As String
    Dim domHead As DOMDocument
    Dim domBody As DOMDocument
    Call Voucher.getVoucherDataXML(domHead, domBody)
    Select Case strVouchtype
        Case "26" 'zp
            StrDraft = IIf(bReturnFlag, GetString("U8.SA.xsglsql.frmMain.00518"), GetString("U8.SA.xsglsql.frmMain.00516"))
        Case "27" 'pp
            StrDraft = IIf(bReturnFlag, GetString("U8.SA.xsglsql.frmMain.00519"), GetString("U8.SA.xsglsql.frmMain.00517"))
        Case "28" 'dbd
            StrDraft = IIf(bReturnFlag, GetString("U8.SA.xsglsql.01.frmbillvouch.00379"), GetString("U8.SA.xsglsql.01.frmbillvouch.00378"))
        Case "29" 'rb
            StrDraft = IIf(bReturnFlag, GetString("U8.SA.xsglsql.frmMain.00528"), GetString("U8.SA.xsglsql.frmMain.00527"))
        Case Else
            StrDraft = Voucher.TitleCaption
    End Select
    If nMode = DraftMode Then
        temp = m_oHelper.SaveAsDraft(nMode, m_strCardNum, Voucher.GetVoucherState().sTitle, domHead, domBody, , StrDraft, m_sCurrentDraftID)
    Else
        temp = m_oHelper.SaveAsDraft(nMode, m_strCardNum, Voucher.GetVoucherState().sTitle, domHead, domBody, , StrDraft)
    End If
    If Trim(temp) <> "" Then
        If nMode = DraftMode Then
            m_sCurrentDraftID = temp
        End If
    End If
'缺省两个参数
'多表体请将其他的表体合并成一个后放在参数otherdom中
'此处用来唯一标识草稿文件 如果还存在重复则 使用unikey参数加以区分
End Sub

'管理草稿模板
Private Sub ManagementDraft(ByVal nMode As TemplateModes)
    Dim StrDraft As String
    Select Case strVouchtype
        Case "26" 'zp
            StrDraft = IIf(bReturnFlag, GetString("U8.SA.xsglsql.frmMain.00518"), GetString("U8.SA.xsglsql.frmMain.00516"))
        Case "27" 'pp
            StrDraft = IIf(bReturnFlag, GetString("U8.SA.xsglsql.frmMain.00519"), GetString("U8.SA.xsglsql.frmMain.00517"))
        Case "28" 'dbd
            StrDraft = IIf(bReturnFlag, GetString("U8.SA.xsglsql.01.frmbillvouch.00379"), GetString("U8.SA.xsglsql.01.frmbillvouch.00378"))
        Case "29" 'rb
            StrDraft = IIf(bReturnFlag, GetString("U8.SA.xsglsql.frmMain.00528"), GetString("U8.SA.xsglsql.frmMain.00527"))
        Case Else
            StrDraft = Voucher.TitleCaption
    End Select
    Call m_oHelper.DraftsManagement(nMode, m_strCardNum, StrDraft)
End Sub

'删除草稿模板
Private Sub DeleteDraft(ByVal nMode As TemplateModes)
    Call m_oHelper.DeleteDraftByID(nMode, m_sCurrentDraftID)
    m_sCurrentDraftID = ""
End Sub

Private Sub SetVouchNoFormat(oDomH As DOMDocument)
    Dim ele As IXMLDOMElement
    Dim NodeData As IXMLDOMCDATASection
    Dim nd As IXMLDOMNode
    Dim ndRS As IXMLDOMNode
    Set ele = oDomH.selectSingleNode("//rs:data/单据编号")
    If ele Is Nothing Then
        Set ndRS = oDomH.selectSingleNode("//rs:data")
        Set NodeData = oDomH.createCDATASection(DomFormat.xml)
        Set nd = oDomH.createElement("单据编号")
        nd.appendChild NodeData
        ndRS.appendChild nd
    End If
End Sub

''参照单据生成当前单据数据
Private Function ReferVouch() As Boolean
    Dim tmpDomH As DOMDocument, tmpDomB As DOMDocument
    Dim strKey As String, i As Integer
    Dim ns As IXMLDOMNode
    'add by renlb20090401
    Dim C As Integer
    Dim strsql  As String
    Dim nodS As IXMLDOMNode
    Dim elelist As IXMLDOMNodeList
    Dim bReMoved As Boolean
    Dim Frm As Object
    Set Frm = New frmVouchRefers
    
    With Me.Voucher
        Set tmpDomH = New DOMDocument
        Set tmpDomB = New DOMDocument
        .getVoucherDataXML tmpDomH, tmpDomB
        Set ns = tmpDomB.selectSingleNode("//rs:data")
        Set elelist = tmpDomB.selectNodes("//z:row[@cinvcode = '']")
        If (Not ns Is Nothing) And elelist.length <> 0 Then
            For Each nodS In elelist
                ns.removeChild nodS
                bReMoved = True
            Next
        End If
    End With
    Dim iType As Integer
    
    Frm.strCardNum = strCardNum ''strVouchType
    Frm.BReferAgain = False
    
    Frm.VouchDOM = tmpDomH
    Frm.bReturnVouch = bReturnFlag
    Dim clsReferVoucher As New EFVoucherMo.clsAutoReferVouch    '' clsSaReferVoucher
    Set clsReferVoucher.m_login = m_login
    
    Select Case UCase(strCardNum)
        Case "EFFYGL040301" '费用结算单
        
            clsReferVoucher.Init "EFFYGL040301", "EFFYGL040301A"
    End Select
    Set Frm.clsReferVoucher = clsReferVoucher
    If Frm.OpenFilter Then
        Frm.Show vbModal, Me
    Else
        Frm.bcancel = True
    End If
    Set clsReferVoucher = Nothing
    If Not Frm.bcancel Then
        ReferVouch = True
'        Set Domhead = frm.Domhead
'        Set Dombody = frm.Dombody
        
        strVoucherUFTS = Frm.domHead.selectSingleNode("//z:row").Attributes.getNamedItem("ufts").nodeValue
        sCurTemplateID = ""
        FillVoucher Frm.domHead, Frm.domBody    'update by renlb
        Me.Voucher.getVoucherDataXML domHead, domBody
        Call SetItemState
        Me.Voucher.ProtectUnload2
        Voucher.row = Voucher.BodyRows
    Else
        ReferVouch = False
    End If
    
    Set Frm = Nothing
End Function

'''''参照后回填当前单据各项 add by renlb    'update by rlb20090404
Public Function FillVoucher(domSrcHead As DOMDocument, DomSrcbody As DOMDocument)
    Dim lst As IXMLDOMNodeList
    Dim nod As IXMLDOMNode
    Dim strSFieldName As String
    Dim strDFieldName As String
    Dim strSSection As String
    Dim ndListH As IXMLDOMNodeList
    Dim ndListB As IXMLDOMNodeList
    Dim elet As IXMLDOMElement
    Dim eleth As IXMLDOMElement
    Dim nodSrcB As IXMLDOMNodeList
    Dim nodSrcH As IXMLDOMNodeList
    Dim nodSrcD As IXMLDOMElement
    Dim attr As IXMLDOMAttribute
    Dim sNoReferField As String
    Dim iRow As Long
    Dim eCheck As Long
    Dim i As Long
    Set nodSrcH = domSrcHead.selectNodes("//z:row")
    Me.Voucher.getVoucherDataXML domHead, domBody
    iRow = Me.Voucher.BodyRows
    For Each elet In nodSrcH
        Select Case UCase(strCardNum)
            
            Case "EFFYGL040301"
                sNoReferField = "id,autoid,ccode,ddate,ufts,cmaker,cverifier,dverifydate,vt_id,ccloser,bbuild,coutid,cbillsign"
                sNoReferField = "," & sNoReferField & ","
                For Each attr In elet.Attributes
                    If Not IsNull(attr.value) And Not IsEmpty(Me.Voucher.headerText(attr.Name)) And InStr(sNoReferField, "," & LCase(attr.Name) & ",") = 0 Then Me.Voucher.headerText(attr.Name) = attr.value
                Next
                For Each eleth In DomSrcbody.selectNodes("//z:row")
                    Voucher.AddLine: iRow = Voucher.BodyRows
                    For Each attr In eleth.Attributes
                        If Not IsNull(attr.value) And Not IsEmpty(Me.Voucher.bodyText(iRow, attr.Name)) And InStr(sNoReferField, "," & LCase(attr.Name) & ",") = 0 Then Me.Voucher.bodyText(iRow, attr.Name) = attr.value
                    Next
                Next
        End Select
        
        For i = 1 To 16
            If Not IsNull(elet.getAttribute("cdefine" & CStr(i))) Then Me.Voucher.headerText("cdefine" & CStr(i)) = elet.getAttribute("cdefine" & CStr(i))
        Next
    Next
    
 End Function

 


'从代码中设置按钮状态
Private Sub setButtonState()
On Error Resume Next
Dim m_CardNumber As String
Dim m_Mid As String
Dim m_mcode As String
Dim m_MAuthid As String
Dim m_TablName As String
    Select Case strCardNum
            Case "EFBWGL020301"
                If GetHeadItemValue(domHead, "iswfcontrolled") = "1" Then
                If Not (GetHeadItemValue(domHead, "iverifystate") = "0") Then
                    m_MAuthid = clsVoucherCO.GetVoucherTaskID("editforverify", strVouchtype, bReturnFlag)
                    Call GetCardNumberMid(m_CardNumber, m_Mid, m_TablName, m_mcode)
                    If bVerifyCanModifyByTaskInfo(m_CardNumber, m_Mid, m_mcode, m_MAuthid) = False Then
                         Me.tbrvoucher.buttons("modify").Enabled = False
                    End If
                End If
            End If
    End Select
Me.UFToolbar1.RefreshVisible
End Sub

'刷新当前单据
Private Sub RefeshVoucher()
Dim vid As String
Dim errMsg As String
Dim UserbSuc As Boolean

    vid = GetHeadItemValue(domHead, "id")
    If val(GetHeadItemValue(domHead, "vt_id")) <> 0 And val(vid) <> 0 Then
        errMsg = clsVoucherCO.GetVoucherData(domHead, domBody, vid)
    End If
    If errMsg = "" Then
        Voucher.Visible = False
        Voucher.SkipLoadAccessories = False
        Voucher.setVoucherDataXML domHead, domBody
          
'        Userdll_UI.LoadAfter_VoucherData errMsg, UserbSuc
        
            '审批流文本
        Voucher.ExamineFlowAuditInfo = GetEAStream(Me.Voucher, strVouchtype)
        Voucher.Visible = True
'        clsTbl.ChangeButtonState Voucher, Me.tbrvoucher, Me.UFToolbar1, Voucher.VoucherStatus
        ChangeButtonState Voucher, Me.tbrvoucher, Me.UFToolbar1, Voucher.VoucherStatus
        '单据模板标题及下拉框的显示
        SetCboVtidState
        
    End If
End Sub

''校验被参照的来源单据是否存在、是否正在被修改
Public Function IsExistent() As Boolean
    Dim Rst As New ADODB.Recordset
    Dim strTable As String
    Dim strCcode As String
    
    Select Case UCase(strCardNum)
        Case "EFBWGL020301"
            strTable = "EFBWGL_V_BookSourcet"
            strCcode = Voucher.headerText("bsccode")
        Case "EFBWGL020401"
            strTable = "EFBWGL_V_SelDeclaret"
            strCcode = Voucher.headerText("selccode")
        Case "EFBWGL020501"
            strTable = "EFBWGL_V_SelRegistert"
            strCcode = Voucher.headerText("selccode")
        Case ""
            
        Case Else
            IsExistent = True
            Exit Function
    End Select
    Rst.Open "Select * from " & strTable & " where ccode = '" & strCcode & "' and ufts = '" & strVoucherUFTS & "'", DBconn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If Rst.EOF Then
        IsExistent = False
    Else
        IsExistent = True
    End If
    Rst.Close
    Set Rst = Nothing
End Function

'//------------取得发稿记录表中保存的书号(即内部编号(年+流水号))-------------
'add by Clin 2009-12-16
Public Function GetCbookno() As String
    Dim Rst As New ADODB.Recordset
    Dim strsql As String
    
    strsql = "select max(cbookno) from EFBWGL_DistRecord"
    Rst.Open strsql, DBconn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Rst.EOF Then
        GetCbookno = ""
    Else
        If IsNull(Rst(0)) Then
            GetCbookno = ""
        Else
            GetCbookno = Rst(0)
        End If
    End If
    Rst.Close
    Set Rst = Nothing
End Function

Public Sub Get_SA_VoucherListConfig()
Dim Rst As New ADODB.Recordset
Dim strsql As String
    strsql = " select * from SA_VoucherListConfig where ckey='" & strCardNum & "' "
    Rst.Open strsql, DBconn, 3, 4
    Rst.Save SA_VoucherListConfigDom, adPersistXML
    Set Rst = Nothing
End Sub

'通用单据导入模板
Public Function EXCEL_Importdate() As Boolean

    Dim excel_app As Object
    Dim excel_sheet As Object
    Dim row  As Long
    Dim col As Long
    Dim Str  As String
    Dim Str2  As String
    Dim strstr As String
    Dim Fieldname_T  As String
    Dim Fieldvale_T  As String
    Dim Fieldname_B  As String
    Dim Fieldvale_B  As String
    Dim strFile As String
    On Error GoTo Err
    Dim sql As String
    
        '打开源EXCEL表
    Dim strInputFile As String
    Dim i As Long
    Dim j As Long
    Dim jubound As Long
    Dim iUbound As Integer
    Dim strsql As String
    Dim strError As String
    Dim new_value As String
    Dim new_value1 As String
    Dim Excelcheck As Boolean
    Dim myHead As New DOMDocument    '单据表头数据
    Dim myBody As New DOMDocument    '单据表体数据
    Dim iElement As IXMLDOMElement
    Dim att As IXMLDOMAttribute
    Dim nodList As IXMLDOMNodeList
    Dim nod As IXMLDOMNode
    Dim ele As IXMLDOMElement

    Dim vtidBTMaxCol() As String               '
    Dim vtidBXCol() As String                  '
    Dim vtidBXName(10) As String
    Dim itemnum As Long
    Dim strinput As String
    Dim strErr As String
    Dim lst As ListBox
    Dim itemX As ListItem
    Dim aArray() As String                      '
    Dim titleArray() As String
    Dim m_aArray As String
    Dim intRows As Long
    Dim strXml As String
    Dim errDom As New DOMDocument
    Dim errorMsg As String
    Dim mainKey As String
    Dim AddLine As Boolean
    Dim isEnum As Boolean
    
    Dim Rst As ADODB.Recordset
    
    Dim eleMent As IXMLDOMElement
    
    EXCEL_Importdate = True
    
    strXml = "<?xml version='1.0' encoding='GB2312'?>" & Chr(13) & "<data>" & Chr(13)
    strXml = strXml & "<btrack>" & Chr(13) & "<btrackcaption rdnum='错误行号' errorreason='错误原因' />" & Chr(13)

    With Me.dlgFileOpen
        .Filter = "Excel文件(*.xls)|*.xls"
        .ShowOpen
        
        If .FileName = "" Then Exit Function
        strInputFile = .FileName
        If Len(Trim(.FileName)) <= 0 Then MsgBox "请选择您要导入的Excel文件！": Exit Function
        .FileName = ""
    End With
    '取出EXCEL数据到临时表
    Screen.MousePointer = vbHourglass
    DoEvents
    Set excel_app = CreateObject("Excel.Application")
    excel_app.workbooks.Open FileName:=strInputFile
    If val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.Activesheet
    Else
        Set excel_sheet = excel_app
    End If
    
    intRows = excel_sheet.usedrange.Rows.Count - 1
    
    iUbound = 0
    jubound = 0
    row = 1
    i = 0
    
    ButtonClick "add", "增加"
    row = 2
    Do
    '    循环读取每一行数据
        col = 4
        AddLine = False
        Str = Trim$(excel_sheet.cells(row, 3))
'        Me.Voucher.BodyMaxRows
        Do
'            读取EXECL文件中每一个字段是数据填写单据上
            '给单据赋值
            DoEvents
            '----取得第一行，第col列的字段名
            If LCase(Left(Trim$(excel_sheet.cells(1, col)), 1)) = LCase("T") Then '填充表头数据
                Fieldname_T = GetFiledNameByCardform4lxd(Right(Trim$(excel_sheet.cells(1, col)), Len(Trim$(excel_sheet.cells(1, col))) - 2), strCardNum, isEnum, "T")
                Fieldvale_T = Trim$(excel_sheet.cells(row, col))
'                Me.Voucher.headerText(Fieldname_T) = Fieldvale_T
                If Len(Fieldname_T) > 0 Then
                    If Fieldname_T = "citemcode" Or Fieldname_T = "citem_class" Then
                        Me.Voucher.headerText(Fieldname_T) = Fieldvale_T
                    Else
                        If isEnum Then
                            Me.Voucher.headerText(Fieldname_T) = Fieldvale_T
                        Else
                            Voucher.SimulateInput SectionsConstants.siHeader, 1, Fieldname_T, Fieldvale_T, 0
                        End If
                    End If
                    
                End If
            ElseIf LCase(Left(Trim$(excel_sheet.cells(1, col)), 1)) = LCase("B") Then
                If AddLine = False Then
                    Me.Voucher.AddLine
                    AddLine = True
                End If
                Fieldname_B = GetFiledNameByCardform4lxd(Right(Trim$(excel_sheet.cells(1, col)), Len(Trim$(excel_sheet.cells(1, col))) - 2), strCardNum, isEnum, "B")
                Fieldvale_B = Trim$(excel_sheet.cells(row, col))
'                Me.Voucher.bodyText(Me.Voucher.row, Fieldname_B) = Trim$(excel_sheet.cells(row, col))
                If Len(Fieldname_T) > 0 Then
                    If Fieldname_T = "citemcode" Or Fieldname_T = "citem_class" Then
                        Me.Voucher.bodyText(Me.Voucher.row, Fieldname_B) = Fieldvale_B
                    Else
                        If isEnum Then
                            Me.Voucher.bodyText(Me.Voucher.row, Fieldname_B) = Fieldvale_B
                        Else
                            Voucher.SimulateInput SectionsConstants.sibody, Voucher.row, Fieldname_B, Fieldvale_B, 0
                        End If
                    End If
                    
                End If
            End If
            col = col + 1
            If Trim$(excel_sheet.cells(1, col)) = "" Then Exit Do
        Loop
        
        
        row = row + 1
        col = 1
        Str2 = Trim$(excel_sheet.cells(row, 3))
        
        
        If (Str <> Str2) And (Str2 <> "") Then
            Me.Voucher.getVoucherDataXML domHead, domBody
            vNewID = ""
            clsSAWeb.bManualTrans = False
            clsVoucherCO.Init strCardNum, m_login, DBconn, "CS", clsSAWeb
            strErr = clsVoucherCO.Save(domHead, domBody, iVouchState, vNewID, domConfig)
            '导入状态
            If Trim(strErr) = "" Then
                excel_sheet.cells(row - 1, 1) = "数据导入成功"
                clsVoucherCO.GetVoucherData domHead, domBody, vNewID
                excel_sheet.cells(row - 1, 2) = "单据ID=" & vNewID & "/ 单据编号=" & GetHeadItemValue(domHead, "ccode")
            Else
                excel_sheet.cells(row - 1, 1) = "数据导入错误"
                excel_sheet.cells(row - 1, 2) = "错误信息： " & strErr
            End If
            ButtonClick "add", "增加"
            Str = ""
            Str2 = ""
        End If
        
        
 
        If Trim$(excel_sheet.cells(row, 3)) = "" Then
            Me.Voucher.getVoucherDataXML domHead, domBody
            vNewID = ""
            clsSAWeb.bManualTrans = False
            clsVoucherCO.Init strCardNum, m_login, DBconn, "CS", clsSAWeb
            strErr = clsVoucherCO.Save(domHead, domBody, iVouchState, vNewID, domConfig)
'            导入状态
            If Trim(strErr) = "" Then
                excel_sheet.cells(row - 1, 1) = "数据导入成功"
                clsVoucherCO.GetVoucherData domHead, domBody, vNewID
                excel_sheet.cells(row - 1, 2) = "单据ID=" & vNewID & "/ 单据编号=" & GetHeadItemValue(domHead, "ccode")
            Else
                excel_sheet.cells(row - 1, 1) = "数据导入错误"
                excel_sheet.cells(row - 1, 2) = "错误信息： " & strErr
            End If

            ButtonClick "tolast", "末张"
            Exit Do
        End If
         
    Loop
    
    excel_app.Quit
Exit Function
Err:

    Screen.MousePointer = 0
    
    If strErr <> "" Then
        MsgBox strErr
    ElseIf Err.Description <> "" Then
        Err.Description = "请检查Excel数据是否填写完整！错误原因：" & Err.Description & "  Excel第" & row & "行，第" & col & "列发生数据异常"
        MsgBox Err.Description, vbCritical, "数据导入"
    End If
    On Error Resume Next
    
    
    ButtonClick "cancel", "放弃"
    excel_app.Quit
    excel_app.activeworkbook.Close False
    
    Set excel_sheet = Nothing
    Set excel_app = Nothing
    EXCEL_Importdate = False
    Screen.MousePointer = vbDefault

End Function




Public Function GetFiledNameByCardform4lxd(ByVal Cardform As String, ByVal sType, ByRef isEnum As Boolean, Optional ByRef cardsection As String = "T") As String
    Dim rstTmp As New ADODB.Recordset
    Dim formula1 As String
    Dim i As Integer
    Dim strsql As String
    GetFiledNameByCardform4lxd = ""
    
    strsql = "select isnull(EnumType,'') as EnumType, fieldname from VoucherItems where cardformula1 ='" & Cardform & "' and cardnum='" & sType & "' and cardsection='" & cardsection & "' and ShowIt=1 and CanModify=1"
    rstTmp.Open strsql, DBconn, adOpenForwardOnly, adLockReadOnly
    
    If Not rstTmp.BOF And Not rstTmp.EOF Then
        GetFiledNameByCardform4lxd = LCase(rstTmp.Fields("fieldname"))
        isEnum = IIf(rstTmp!enumType = "", False, True)
    Else
        strsql = "select isnull(EnumType,'') as EnumType,fieldname from VoucherItems where carditemname ='" & Cardform & "' and cardnum='" & sType & "' and cardsection='" & cardsection & "' and ShowIt=1 and CanModify=1"
        If rstTmp.State = 1 Then rstTmp.Close
        rstTmp.Open strsql, DBconn, adOpenForwardOnly, adLockReadOnly
        
        If Not rstTmp.BOF And Not rstTmp.EOF Then
            GetFiledNameByCardform4lxd = LCase(rstTmp.Fields("fieldname"))
            isEnum = IIf(rstTmp!enumType = "", False, True)
        End If
    End If
    If rstTmp.State = 1 Then rstTmp.Close
    Set rstTmp = Nothing

End Function
Public Sub ChangeButtonState(Voucher As ctlVoucher, tbl As Toolbar, uftbl As UFToolbar, intVoucherState As VoucherStatusConstants)
    
    On Error Resume Next
    
        If intVoucherState = VSNormalMode Then
            '----浏览状态
            tbl.buttons("add").Enabled = True
'            tbl.buttons("voucherdesign").Enabled = True
            '----审核不能修改
            If Voucher.headerText("cverifier") <> "" Or Voucher.headerText("ccode") = "" Then
                tbl.buttons("modify").Enabled = False
                tbl.buttons("modyhd").Enabled = False
            Else
                tbl.buttons("modify").Enabled = True
                tbl.buttons("modyhd").Enabled = True
            End If
            If Voucher.headerText("cverifier") <> "" Or Voucher.headerText("ccode") = "" Then
                tbl.buttons("delete").Enabled = False
            Else
                tbl.buttons("delete").Enabled = True
            End If
            tbl.buttons("calc").Enabled = False
            tbl.buttons("refervouch").Enabled = False
            tbl.buttons("discard").Enabled = False
            tbl.buttons("copy").Enabled = True
            tbl.buttons("draft").Enabled = True
            tbl.buttons("save").Enabled = False
            If Voucher.headerText("cverifier") <> "" Then
                tbl.buttons("verify").Enabled = False
                tbl.buttons("unverify").Enabled = True
                tbl.buttons("queryconfirm").Enabled = True
                tbl.buttons("submit").Enabled = False
                tbl.buttons("resubmit").Enabled = False
                tbl.buttons("unsubmit").Enabled = False
                tbl.buttons("creatdd").Enabled = True
                If Voucher.headerText("ccloser") <> "" Then
                    tbl.buttons("closeorder").Enabled = False
                    tbl.buttons("openorder").Enabled = True
                Else
                    tbl.buttons("closeorder").Enabled = True
                    tbl.buttons("openorder").Enabled = False
                End If
                
            Else
                tbl.buttons("verify").Enabled = True
                tbl.buttons("unverify").Enabled = False
                tbl.buttons("queryconfirm").Enabled = False
                tbl.buttons("submit").Enabled = True
                tbl.buttons("resubmit").Enabled = True
                tbl.buttons("unsubmit").Enabled = True
                tbl.buttons("creatdd").Enabled = False
                tbl.buttons("closeorder").Enabled = False
                tbl.buttons("openorder").Enabled = False
            End If
            If Voucher.headerText("ccode") = "" Then
                tbl.buttons("verify").Enabled = False
                tbl.buttons("unverify").Enabled = False
                tbl.buttons("queryconfirm").Enabled = False
                tbl.buttons("submit").Enabled = False
                tbl.buttons("resubmit").Enabled = False
                tbl.buttons("unsubmit").Enabled = False
                tbl.buttons("creatdd").Enabled = False
            End If
            
            tbl.buttons("copycreating").Enabled = False
            tbl.buttons("cancel").Enabled = False
            tbl.buttons("import").Enabled = False
            tbl.buttons("formatsetup").Enabled = True
            tbl.buttons("print").Enabled = True
            tbl.buttons("preview").Enabled = True
            tbl.buttons("output").Enabled = True
            
            '----表体按钮状态
            tbl.buttons("insertrow").Enabled = False
            tbl.buttons("mnucopyitem").Enabled = False
'            tbl.buttons("mnuspiltitem").Enabled = False
            tbl.buttons("deleterecord").Enabled = False
            tbl.buttons("mnubatchmodify").Enabled = False
        Else
            '----编辑状态
'            tbl.buttons("voucherdesign").Enabled = False
            
            tbl.buttons("add").Enabled = False
            tbl.buttons("modify").Enabled = False
            tbl.buttons("modyhd").Enabled = False
            tbl.buttons("delete").Enabled = False
            tbl.buttons("copy").Enabled = False
            tbl.buttons("draft").Enabled = True
            tbl.buttons("save").Enabled = True
            
            tbl.buttons("discard").Enabled = True
            
            tbl.buttons("verify").Enabled = False
            tbl.buttons("unverify").Enabled = False
            tbl.buttons("queryconfirm").Enabled = False
            tbl.buttons("submit").Enabled = False
            tbl.buttons("resubmit").Enabled = False
            tbl.buttons("unsubmit").Enabled = False
            
            tbl.buttons("copycreating").Enabled = True
            tbl.buttons("creatdd").Enabled = False
            tbl.buttons("cancel").Enabled = True
            tbl.buttons("import").Enabled = True
            tbl.buttons("formatsetup").Enabled = False
            tbl.buttons("closeorder").Enabled = False
            tbl.buttons("openorder").Enabled = False
            tbl.buttons("print").Enabled = False
            tbl.buttons("preview").Enabled = False
            tbl.buttons("output").Enabled = False
            
            '----表体按钮状态
            tbl.buttons("insertrow").Enabled = True
            tbl.buttons("mnucopyitem").Enabled = True
'            tbl.buttons("mnuspiltitem").Enabled = True
            tbl.buttons("deleterecord").Enabled = True
            tbl.buttons("mnubatchmodify").Enabled = True
            tbl.buttons("calc").Enabled = True
            tbl.buttons("refervouch").Enabled = True
            
        End If
        tbl.buttons("ShowTemplate").Enabled = True
'        If Voucher.TitleCaption <> "" Then
'            LabelVoucherName.Caption = Voucher.TitleCaption  '//单据名称标题放入单据头的Label上。
'            Voucher.TitleCaption = ""                        '//单据的名称，清空
'        End If
'        InitComTemplate Me.tbrvoucher, ComboVTID
        SetPrintShowTemplate
'        tbrvoucher.buttons("draft").ButtonMenus("DraftManager").Visible = False
'        tbrvoucher.buttons("draft").ButtonMenus.Add , "draft1", "草稿1"
'        tbrvoucher.buttons("draft").ButtonMenus.Add , "draft2", "草稿2"
'        tbrvoucher.buttons("save").ButtonMenus("savenew").Enabled = False
        Call UFToolbar1.SetToolbar(tbrvoucher)
        uftbl.RefreshVisible
End Sub




Private Sub Voucher_SearchClick(ByVal cSearchKey As String)
    Dim tmpid As String
    Dim tmprst As New ADODB.Recordset
    Dim strsql As String
    Dim oId As String
    Dim MainTable As String
    MainTable = GetMainTable
    If MainTable = "" Then Exit Sub
    If sTmpTableName = "" Then
        sTmpTableName = "tempdb..TEMP_STSearchTableName_" & sGuid
    End If
    DropTable DBconn, sTmpTableName
          
    strsql = "select ID as cVoucherId,ccode as cVoucherCode,cast(null as nvarchar(1)) as cVoucherName,cast(null as nvarchar(1)) as cCardNum,cast(null as nvarchar(1)) as cMenu_Id,cast(null as nvarchar(1)) as cAuth_Id,cast(null as nvarchar(1)) as cSub_Id into " & sTmpTableName & " from " & MainTable & "  where  (ccode like N'%" & Trim(cSearchKey) & "%')"
    
    
    DBconn.Execute strsql
    strsql = "select cVoucherId from " & sTmpTableName
    tmprst.Open strsql, DBconn, adOpenForwardOnly, adLockReadOnly
    If Not tmprst.EOF Then
        tmpid = tmprst(0)
        Voucher.SearchTableName = sTmpTableName
        If tmprst.RecordCount = 1 Then
            sTmpTableName = ""
        End If
    Else
        tmpid = ""
        sTmpTableName = ""
    End If
    tmprst.Close
    Set tmprst = Nothing
    If tmpid <> "" Then
        Call LoadVoucher("", tmpid)
    End If
End Sub

Private Function GetMainTable() As String
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    On Error GoTo hErr
    strsql = "select  bttblname   from  Vouchers_base where cardnumber='" & m_strCardNum & "'"
    rs.Open strsql, DBconn
    If Not rs.EOF And Not rs.BOF Then
        GetMainTable = rs!bttblname & ""
    End If
    Exit Function
hErr:
    Debug.Print Err.Description
End Function

Private Sub Voucher_ReleaseSearchClick()
    sTmpTableName = ""
    Voucher.SearchTableName = ""
End Sub

'插行
Private Sub DoInsertLine()
    Dim iRow As Long
    iRow = Voucher.row
    If iRow = 0 Then
        Exit Sub
    Else
        Voucher.AddLine Voucher.row, , ALSPrevious
    End If
    ReSetBodyRowNo
End Sub

Private Sub ReSetBodyRowNo()
    Dim iRow As Long
    For iRow = 1 To Voucher.BodyRows
        Voucher.bodyText(iRow, "irowno") = iRow
    Next
End Sub
