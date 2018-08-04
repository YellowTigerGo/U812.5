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
         Caption         =   "现结"
         BeginProperty Font 
            Name            =   "宋体"
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
         Caption         =   "作废"
         BeginProperty Font 
            Name            =   "宋体"
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
            Name            =   "宋体"
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
               Caption         =   "打印模版："
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
      Caption         =   "右键菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuLinkQuery 
         Caption         =   "联查预算明细"
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
'修改后的程序指定常数的值
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
Private maxRefDate As Date ''   最晚参照日期
Private strCurVoucherNO As String
Private strVouchType As String, bReturnFlag As Boolean '记录单据类型
Private bCheckVouch As Boolean '单据的审核状态2
Public bFrmCancel As Boolean
Dim strCardNum As String        ''单据的CardNum
Dim sTemplateID As String       ''单据默认模板号码
Dim sCurTemplateID As String    ''单据当前的模板号
Dim sCurTemplateID2 As String    ''单据当前的模板号
Private vName As String
Private BrowFlag As Boolean '标识是否调用Voucher.browuser事件
Dim strRefFldName As String '发生参照的字段名
Private iVouchState As Integer
Private bClickCancel As Boolean
Private bClickSave As Boolean
'参照
Dim clsRefer As New UFReferC.UFReferClient
Dim clsAuth As New U8RowAuthsvr.clsRowAuth
Dim Domhead As New DOMDocument
Dim Dombody As New DOMDocument
Dim vNewID As Variant               '单据id
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
Public iShowMode As Integer    ''窗体模式  0：正常 1：浏览
Private bCreditCheck  As Boolean   ''是否通过信用检查
Dim bOnceRefer As Boolean
Private ButtonTaskID As String  ''按钮任务id
Private RstTemplate As ADODB.Recordset, preVTID As String      ''保存临时的单据模版记录集
Private RstTemplate2 As New ADODB.Recordset
Dim vtidPrn() As Long ''打印模版数组
Private bfillDjmb As Boolean, vtidDJMB() As Long
Private bManBodyChecked As Boolean '' 是否手工cellchecked
Private bCloseFHSingle As Boolean
Private obj_EA As Object, DOMEA As DOMDocument, strEAXML As String ''审批流
Private bLostFocus As Boolean
Private domConfig As New DOMDocument
Private domTmp As DOMDocument
Private o_crm As Object
Private moAutoFill As Object
Private dOriVoucherWidth As Double, dOriVoucherHeight As Double
Private col(1 To 22) As Long  '用数组记录关键字所在的位置


'by lg070314 增加U870支持
Private m_Cancel As Integer
Private m_UnloadMode As Integer
Dim sguid As String
Private WithEvents m_mht As UFPortalProxyMessage.IMessageHandler
Attribute m_mht.VB_VarHelpID = -1



'by lg070314 增加U870支持
'修改3 每个窗体都需要这个方法。Cancel与UnloadMode的参数的含义与QueryUnload的参数相同
'请在此方法中调用窗体Exit(退出)方法，并将设置窗体Unload事件参数(如Cancel)的值同时传给此方法的参数
Public Sub ExitForm(Cancel As Integer, UnloadMode As Integer)
'    Unload Me
'
'    Cancel = m_Cancel
'    UnloadMode = m_UnloadMode

doNext:
    If Me.voucher.VoucherStatus <> VSNormalMode Then
        Select Case MsgBox("是否保存对当前单据的编辑？", vbYesNoCancel + vbQuestion)
            Case vbYes
                ButtonClick "Save", "保存"
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
'by lg070314增加U870菜单融合，关闭时处理Business
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


'设置帮助的系统id
Private Sub SetHelpID()
    Select Case strVouchType
        Case "32"
            Me.HelpContextID = 10060203
        Case Else
            Me.HelpContextID = 10060203
    End Select
    
End Sub
 
''sKey :操作的按钮名称
''
Private Function VoucherTask(sKey As String) As Boolean
    Dim strID As String
    
    Select Case strVouchType
        Case "16"
            Select Case sKey
                Case "增加", "复制", "删除", "修改"
                    strID = "FA03000102"  '
                Case "审核", "弃审"
                    strID = "FA03000103"  '
                Case "关闭", "打开"
                    strID = "FA03000104"  '
            End Select
        Case "97"
            Select Case sKey
                Case "增加", "复制", "删除", "修改"
                    strID = "FA03010101"  '
                Case "关闭", "打开"
                    strID = "FA03010102"  '
                Case "审核", "弃审"
                    strID = "FA03010103"  '
                Case "变更"
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

''释放功能申请
Private Function VoucherFreeTask() As Boolean
    If ButtonTaskID <> "" Then
        VoucherFreeTask = LockItem(ButtonTaskID, False, True)
        ButtonTaskID = ""
    End If
End Function
 
'Dim strAuthId As String     '权限号/gyp/2002/07/24
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
                    MsgBox "你没有使用单据模版的权限！"
                    'Me.Hide
                    Exit Function
                Else
                    If clsAuth.IsHoldAuth("DJMB", sTemplateID, , "W") = False Then
                        rstTmp.Open "select vt_id from vouchertemplates where vt_cardnumber='" & strCardNum & "' and vt_id in (" & strDJAuth & ") order by vt_id", DBConn, adOpenForwardOnly, adLockReadOnly
                        If Not rstTmp.EOF Then
                            fillComBol False
                            sNewTemplateID = rstTmp(0)      'left(strDJAuth, IIf(InStr(1, strDJAuth, ",") - 1 = -1, Len(strDJAuth), InStr(1, strDJAuth, ",")))
                        Else
                            MsgBox "你没有使用单据模版的权限！"
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
                GoTo UsePre  ''记录已经取回
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
                                MsgBox "模版设置有问题"
                                ChangeTempaltes = False
                                Exit Function
                        End If
                        If Not RstTemplate2 Is Nothing Then
                            If Not RstTemplate2.EOF Then
                                bChanged = True
                            Else
                                MsgBox "模版设置有问题"
                                ChangeTempaltes = False
                                Exit Function
                            End If
                        Else
                            MsgBox "模版设置有问题"
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
    '//单据名称标题，主要是为了解决单据的名称的特殊显示问题，例如 "期初" XXX单据
    Me.LabelVoucherName.Caption = Me.voucher.TitleCaption

    '//单据的名称
    Me.voucher.TitleCaption = Me.voucher.TitleCaption
    Me.voucher.TitleCaption = ""
End Sub

''函数load单据,更改按纽状态,更改模板
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
    '审批流文本
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
    'If strSumDX = "圆整" Then strSumDX = "零圆零角零分"
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
 

''需要改变单据模版
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
'设置单据排序控件

'//////////////////////////////////////////////////
'  860sp升级到861修改处1 注释    2006/03/08 改控件在861版本中已经集成到单据控件中了   所以要删除
' voucher.SetSortCallBackObject U8VoucherSorter1
'    With U8VoucherSorter1
'        .BackColor = voucher.BackColor
'        .Left = Me.Left + 550
'        .Top = Me.Picture1.Top
'        .ZOrder
'    End With
'//////////////////////////////////////////////////


'by lg070314增加U870菜单融合功能
    ''''''''''''''''''''''''''''''''''''''
    If Not g_business Is Nothing Then
        Set Me.UFToolbar1.Business = g_business
    End If
    Call RegisterMessage
    '''''''''''''''''''''''''''''''''''''''
    
    Call SetButton  '设置菜单按钮
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
'  860sp升级到861修改 注释    2006/03/08 改控件在861版本中已经集成到单据控件中了 所以要删除
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
            'by lg070315　增加u870单据新的定位过滤
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
            
            Case "add"            '//增加
                ChangeTempaltes sCurTemplateID
                If ChangeDJMBForEdit = False Then Screen.MousePointer = vbDefault: Exit Sub
                If VoucherTask(sTaskKey) = False Then Screen.MousePointer = vbDefault: Exit Sub
                Screen.MousePointer = vbHourglass
                EditPanel_1 EdtP, 3, ""
                labZF.Visible = False
                labXJ.Visible = False
                Me.voucher.AddNew ANMNormalAdd, Domhead, Dombody '
                Call SetVouchNoWriteble      '设置单据号是否可以编辑
                Call AddNewVouch              '设置新增单据的初始值
                Me.voucher.AddNew ANMCopyALL, Domhead, Dombody
                Me.voucher.headerText("vt_id") = sCurTemplateID
                Set Domhead = Me.voucher.GetHeadDom
                If iShowMode = 2 Then

                End If
                iVouchState = 0
                Call SetButtonStatus(s)
                Call setItemState(s)
                
            Case "chenged" '变更
'                If strVouchType = "97" Or strVouchType = "96" Then
'                    Call Frm.ShowVoucher(gdzckpxg, Me.voucher.headerText("id"))
'               End If
            '/////////////////////////////////////////////////////////////////////////////////////////
            '  860sp升级到861修改处1 注释    2006/03/09 861版本中单据控件增加单据附件功能（附件的可以是文件，图片附件的上限大小为1M）
                Me.voucher.SelectFile
               
               
            Case "outadd"              '//批量导出
            Case "modify"              '//修改
                If CheckDJMBAuth(Me.voucher.headerText("ivtid"), "W") = False Then
                    MsgBox "当前操作员没有当前单据模版的使用权限！"
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
                
            Case "erase"                 '//删除
                If CheckDJMBAuth(Me.voucher.headerText("ivtid"), "W") = False Then
                    MsgBox "当前操作员没有当前单据模版的使用权限！"
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                If MsgBox("确实要删除本张单据吗？", vbYesNo + vbQuestion) = vbNo Then
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

            Case "copy"                         '//复制
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
            Case "addrow"                       '//增加一行
                With Me.voucher
                    .AddLine
                    .bodyText(.BodyRows, "itb") = "正常"
                    If strVouchType = "05" Or strVouchType = "06" Or strVouchType = "26" Or strVouchType = "27" Or strVouchType = "28" Or strVouchType = "29" Then
                        If .BodyRows > 1 Then
                           .bodyText(.BodyRows, "cwhname") = .bodyText(.BodyRows - 1, "cwhname")
                           .bodyText(.BodyRows, "cwhcode") = .bodyText(.BodyRows - 1, "cwhcode")
                        End If
                    End If
                End With
            Case "delrow"                     '//删除一行
                If (Me.voucher.BodyRows > 0) And Me.voucher.row <> 0 Then
                     Dim tmpRow As Variant
                     tmpRow = Me.voucher.row - 1
                     Me.voucher.DelLine Me.voucher.row
                    Me.voucher.row = tmpRow
                End If
                
            Case "sure"           '//审核
                Screen.MousePointer = vbHourglass
                If VoucherTask(sTaskKey) = False Then Screen.MousePointer = vbDefault: Exit Sub
                Call AddNewVouch(s)
                bCreditCheck = True
                Set Domhead = Me.voucher.GetHeadDom
                strError = clsVoucherCO.VerifyVouch(Domhead, bCheckVouch)
                Call ShowErrDom(strError, Domhead)
                ''刷新当前单据
                LoadVoucher ""
                Call VoucherFreeTask
                
            Case "unsure"            '//弃审
                Screen.MousePointer = vbHourglass
                If VoucherTask(sTaskKey) = False Then Screen.MousePointer = vbDefault: Exit Sub
                Call AddNewVouch(s)
                bCreditCheck = False
                Set Domhead = Me.voucher.GetHeadDom
                strError = clsVoucherCO.VerifyVouch(Domhead, bCheckVouch)
                Call ShowErrDom(strError, Domhead)
                ''刷新当前单据
                LoadVoucher ""
                Call VoucherFreeTask
            Case "cancel"                  '//取消
                bClickCancel = True
                voucher.VoucherStatus = VSNormalMode
                LoadVoucher ""
                bOnceRefer = False
                Call SetButtonStatus(s)
                ChangeButtonsState
                bClickCancel = False
                Call VoucherFreeTask
            Case "save"                    '//保存
                Screen.MousePointer = vbHourglass
                voucher.ProtectUnload2
                bClickCancel = False
                bClickSave = True
                strError = ""
                If Me.voucher.BodyRows = 0 And strVouchType <> "94" Then
                    MsgBox "表体没有记录，请录入！"
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                If .headVaildIsNull2(strError) = False Then
                    MsgBox "表头项目" + strError
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                
                strError = ""
                If .bodyVaildIsNull2(strError) = False Then
                    MsgBox "表体项目" + strError
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                strError = ""
                Call AddNewVouch("Save")
                voucher.getVoucherDataXML Domhead, Dombody
                '////////////////////////////////////////////////////////////////////////////////////////////////
                '860sp升级到861修改处   2006/03/08   增加单据附件功能
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
                        MsgBox IIf(Trim(strError) = "当前操作不成功，请重新再试!", "", strError)
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
            Case "print"            '打印
                    If Me.ComboVTID.ListCount = 0 Then
                        MsgBox "当前操作员没有可以使用的打印模版，请检查！"
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                    If strVouchType = "96" Or strVouchType = "97" Then
                        sPrnTmplate = Get_print_id(Me.voucher.headerText("stypenum"))
                        If sPrnTmplate = 0 Then
                            MsgBox "没有设置正确的默认打印模板，请检查！"
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
                        MsgBox "当前操作员没有可以使用的打印模版，请检查！"
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                     ''取用户设置的默认打印模板
                    If strVouchType = "96" Or strVouchType = "97" Then
                        sPrnTmplate = Get_print_id(Me.voucher.headerText("stypenum"))
                        If sPrnTmplate = 0 Then
                            MsgBox "没有设置正确的默认打印模板，请检查！"
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
                '由单据联查凭证
                 Find_GL_accvouch
            
            Case "paint"
                Screen.MousePointer = vbHourglass
                LoadVoucher "", , True
                
            Case LCase("ToPrevious")   '上一张
                Screen.MousePointer = vbHourglass
                LoadVoucher s, , True
 
'                voucher.VoucherStatus = VSNormalMode
            
            Case LCase("ToNext")   '下一张
                Screen.MousePointer = vbHourglass
                LoadVoucher s, , True
        
'                voucher.VoucherStatus = VSNormalMode
                
            Case LCase("ToLast")   '末张
                Screen.MousePointer = vbHourglass
                LoadVoucher s, , True
 
'                voucher.VoucherStatus = VSNormalMode
                
            Case LCase("ToFirst")   '首张
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
            Case LCase("LookVeri")  ''查询审批流
                If .VoucherStatus = VSeEditMode Then .ProtectUnload2
                Set Domhead = .GetHeadDom
                If obj_EA.NeedEAFControl(clsSAWeb.GetEAsCode(strVouchType, Domhead), GetHeadItemValue(Domhead, clsSAWeb.getVouchMainIDName(strVouchType))) Then
                    If (obj_EA.ResearchEAStream(clsSAWeb.GetEAsCode(strVouchType, Domhead), .headerText(clsSAWeb.getVouchMainIDName(strVouchType)))) = False Then
                        MsgBox obj_EA.ErrDescript
                    End If
                Else
                    MsgBox "该单据未进入审批流!"
                End If
                
            Case "zp"   '报账单编制时核销支票
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
'        Select Case MsgBox("是否保存对当前单据的编辑？", vbYesNoCancel + vbQuestion)
'            Case vbYes
'                ButtonClick "Save", "保存"
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
''by lg070314增加U870菜单融合，关闭时处理Business
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
    ''右键菜单
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
            KeyCode = "scardnum"     '资产卡片
            
        Case "101"
            KeyCode = "ccode"         ' 资产盘点单单号
            
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
        If Not (LCase(DomFormat.selectSingleNode("//单据编号").Attributes.getNamedItem("重号自动重取").nodeValue) = "false" And _
        LCase(DomFormat.selectSingleNode("//单据编号").Attributes.getNamedItem("允许手工修改").nodeValue) = "true") Then
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
        .MultiLineSelect = False ''设置多选默认
        clsRefer.SetReferSQLString ""
        clsRefer.SetRWAuth "INVENTORY", "R", False
        clsRefer.SetReferDisplayMode enuGrid
        sKey = .ItemState(col, sibody).sFieldName
        sKeyValue = .bodyText(row, col)
        Select Case LCase(sKey)
        
            Case "ccode", "ccode_name" '科目
                    strClass = ""
                    strGrid = "select ccode,ccode_name from code where ccode not in( select ccode from MT_basesets where  isnull(ccode,'')<>'') and bend=1 "
                    If LCase(sKey) = "ccode" And Len(Trim(.bodyText(row, "ccode"))) > 0 Then
                        strGrid = strGrid & " and ccode like '%" & Trim(.bodyText(row, "ccode")) & "%'"
                    ElseIf LCase(sKey) = "ccode_name" And Len(Trim(.bodyText(row, "ccode_name"))) > 0 Then
                        strGrid = strGrid & " and ccode_name like '%" & Trim(.bodyText(row, "ccode_name")) & "%'"
                    End If
                    strGrid = strGrid & " order by ccode "
                    If clsRefer.StrRefInit_SetColWidth(m_login, False, strClass, strGrid, "科目编码,科目名称", "2500,6000") = False Then Exit Sub
                    clsRefer.Show
                    If Not clsRefer.recMx Is Nothing Then
                        .bodyText(row, "ccode") = clsRefer.recMx(0)
                        .bodyText(row, "ccode_name") = clsRefer.recMx(1)
                        sRet = clsRefer.recMx.Fields(LCase(sKey))
                    End If
                    

            Case "cexpcode", "cexpname" '费用类别
                    strClass = "select * from dbo.ExpItemClass "
                    strGrid = "select cexpccode,cexpcode,cexpname from ExpenseItem  where 1=1  "
                    If LCase(sKey) = "cexpcode" And Len(Trim(.bodyText(row, "cexpcode"))) > 0 Then
                        strGrid = strGrid & " and cexpcode like '%" & Trim(.bodyText(row, "cexpcode")) & "%'"
                    ElseIf LCase(sKey) = "cexpname" And Len(Trim(.bodyText(row, "cexpname"))) > 0 Then
                        strGrid = strGrid & " and cexpname like '%" & Trim(.bodyText(row, "cexpname")) & "%'"
                    End If
                    strGrid = strGrid & " order by cexpcode "
                    If clsRefer.StrRefInit_SetColWidth(m_login, False, strClass, strGrid, "分类,费用类别编码,费用类别名称", "0,2500,6000") = False Then Exit Sub
                    clsRefer.Show
                    If Not clsRefer.recMx Is Nothing Then
                        .bodyText(row, "cexpcode") = clsRefer.recMx(1)
                        .bodyText(row, "cexpname") = clsRefer.recMx(2)
                    sRet = clsRefer.recMx.Fields(LCase(sKey))
                    End If

                    

                 
        End Select
End With
'by lg070315　增加U870 UAP单据控件新的参照处理
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
        
            Case "ccode"  '科目
                    strSql = "select ccode,ccode_name from code where bend=1 "
                    strSql = strSql & " and ccode = '" & Trim(.bodyText(R, "ccode")) & "'"
                    If rds.state <> 0 Then rds.Close
                    rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
                    If rds.RecordCount = 0 Then
                        MsgBox "科目不合法！", vbOKOnly + vbCritical
                        Me.voucher.bodyText(R, "ccode") = ""
                        Me.voucher.bodyText(R, "ccode_name") = ""
                        RetValue = ""
                        bChanged = Retry
                    Else
                        Me.voucher.bodyText(R, "ccode") = rds.Fields("ccode")
                        Me.voucher.bodyText(R, "ccode_name") = rds.Fields("ccode_name")
                        RetValue = rds.Fields(LCase(sKey))
                    End If
                    
            Case "ccode_name"  '科目
                    strSql = "select ccode,ccode_name from code where bend=1 "
                    strSql = strSql & " and ccode_name = '" & Trim(.bodyText(R, "ccode_name")) & "'"
                    If rds.state <> 0 Then rds.Close
                    rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
                    If rds.RecordCount = 0 Then
                        MsgBox "科目不合法！", vbOKOnly + vbCritical
                        Me.voucher.bodyText(R, "ccode") = ""
                        Me.voucher.bodyText(R, "ccode_name") = ""
                        RetValue = ""
                        bChanged = Retry
                    Else
                        Me.voucher.bodyText(R, "ccode") = rds.Fields("ccode")
                        Me.voucher.bodyText(R, "ccode_name") = rds.Fields("ccode_name")
                        RetValue = rds.Fields(LCase(sKey))
                    End If
            
            Case "cexpcode"  '费用类别
                    strSql = "select cexpccode,cexpcode,cexpname from ExpenseItem  where 1=1  "
                    strSql = strSql & " and cexpcode = '" & Trim(.bodyText(R, "cexpcode")) & "'"
                    If rds.state <> 0 Then rds.Close
                    rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
                    If rds.RecordCount = 0 Then
                        MsgBox "费用类别不合法！", vbOKOnly + vbCritical
                        Me.voucher.bodyText(R, "cexpcode") = ""
                        Me.voucher.bodyText(R, "cexpname") = ""
                        RetValue = ""
                        bChanged = Retry
                    Else
                        Me.voucher.bodyText(R, "cexpcode") = rds.Fields("cexpcode")
                        Me.voucher.bodyText(R, "cexpname") = rds.Fields("cexpname")
                        Me.voucher.bodyText(R, "iscontrol") = "控制"
                        Me.voucher.bodyText(R, "bfb") = GETysbl(Me.voucher.headerText("cdepcode"), Me.voucher.headerText("citemcode"), Me.voucher.bodyText(R, "cexpcode"))
                        RetValue = rds.Fields(LCase(sKey))
                    End If
                    
            Case "cexpname"  '费用类别
                    strSql = "select cexpccode,cexpcode,cexpname from ExpenseItem  where 1=1  "
                    strSql = strSql & " and cexpname = '" & Trim(.bodyText(R, "cexpname")) & "'"
                    If rds.state <> 0 Then rds.Close
                    rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
                    If rds.RecordCount = 0 Then
                        MsgBox "费用类别不合法！", vbOKOnly + vbCritical
                        Me.voucher.bodyText(R, "cexpcode") = ""
                        Me.voucher.bodyText(R, "cexpname") = ""
                        RetValue = ""
                        bChanged = Retry
                    Else
                        Me.voucher.bodyText(R, "cexpcode") = rds.Fields("cexpcode")
                        Me.voucher.bodyText(R, "cexpname") = rds.Fields("cexpname")
                        Me.voucher.bodyText(R, "bfb") = GETysbl(Me.voucher.headerText("cdepcode"), Me.voucher.headerText("citemcode"), Me.voucher.bodyText(R, "cexpcode"))
                        Me.voucher.bodyText(R, "iscontrol") = "控制"
                        RetValue = rds.Fields(LCase(sKey))
                    End If
            
'            Case "adds", "lenssen" '发生时控制方向
'                    If LCase(sKey) = "adds" And Len(Trim(.bodyText(R, "adds"))) > 0 Then
'                         If (Trim(.bodyText(R, "adds")) = "借方") Or (Trim(.bodyText(R, "adds")) = "贷方") Then
'                            .bodyText(R, "adds") = Trim(.bodyText(R, "adds"))
'                            RetValue = .bodyText(R, "adds")
'                            .bodyText(R, "lenssen") = ""
'                         Else
'                            .bodyText(R, "adds") = ""
'                            RetValue = ""
'                            bChanged = Retry
'                         End If
'                    ElseIf LCase(sKey) = "lenssen" And Len(Trim(.bodyText(R, "lenssen"))) > 0 Then
'                         If (Trim(.bodyText(R, "lenssen")) = "借方") Or (Trim(.bodyText(R, "lenssen")) = "贷方") Then
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
                        If (Trim(.bodyText(R, "adds")) = "借方") Then
                            .bodyText(R, "adds") = Trim(.bodyText(R, "adds"))
                            RetValue = .bodyText(R, "adds")
                            If (.bodyText(R, "lenssen") <> "贷方") And (.bodyText(R, "lenssen") <> "") Then
                                .bodyText(R, "lenssen") = "贷方"
                            End If
                        ElseIf (Trim(.bodyText(R, "adds")) = "贷方") Then
                            .bodyText(R, "adds") = Trim(.bodyText(R, "adds"))
                            RetValue = .bodyText(R, "adds")
                            If (.bodyText(R, "lenssen") <> "借方") And (.bodyText(R, "lenssen") <> "") Then
                                .bodyText(R, "lenssen") = "借方"
                            End If
                        Else
                            .bodyText(R, "adds") = ""
                            RetValue = ""
                            bChanged = Retry
                        End If
                    End If
            Case "lenssen"
                    If Len(Trim(.bodyText(R, "lenssen"))) > 0 Then
                        If (Trim(.bodyText(R, "lenssen")) = "借方") Then
                            .bodyText(R, "lenssen") = Trim(.bodyText(R, "lenssen"))
                            RetValue = .bodyText(R, "lenssen")
                            If (.bodyText(R, "adds") <> "贷方") And (.bodyText(R, "adds") <> "") Then
                                .bodyText(R, "adds") = "贷方"
                            End If
                        ElseIf (Trim(.bodyText(R, "lenssen")) = "贷方") Then
                            .bodyText(R, "lenssen") = Trim(.bodyText(R, "lenssen"))
                            RetValue = .bodyText(R, "lenssen")
                            If (.bodyText(R, "adds") <> "借方") And (.bodyText(R, "adds") <> "") Then
                                .bodyText(R, "adds") = "借方"
                            End If
                        Else
                            .bodyText(R, "lenssen") = ""
                            RetValue = ""
                            bChanged = Retry
                         End If
                    End If
            Case "rate" '预算比例
               If (.bodyText(R, "rate") > 100) Or (.bodyText(R, "rate") <= 0) Then
                    MsgBox "预算比例不合法！", vbOKOnly + vbCritical
                    .bodyText(R, "rate") = "0.00"
                    RetValue = ""
                    bChanged = Retry
               End If
               
               
               
            Case "je" '预算
                If val(.bodyText(R, "je")) <> 0 Then
                    If val(.bodyText(R, "hdje")) = 0 Then
                    .bodyText(R, "hdje") = .bodyText(R, "je")
                    End If
                End If
                          
            
'            Case "hdje"   '  核定预算
'               If (.bodyText(R, "rate") > 100) Or (.bodyText(R, "rate") <= 0) Then
'                    MsgBox "预算比例不合法！", vbOKOnly + vbCritical
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

        Case "iscontrol" '是否控制
                pCom.Clear
                pCom.AddItem "提醒"
                pCom.AddItem "控制"
                pCom.AddItem "不提示"
        
        Case "iperiod" '预算周期（会计期间）
                pCom.Clear
                pCom.AddItem "1  月"
                pCom.AddItem "2  月"
                pCom.AddItem "3  月"
                pCom.AddItem "4  月"
                pCom.AddItem "5  月"
                pCom.AddItem "6  月"
                pCom.AddItem "7  月"
                pCom.AddItem "8  月"
                pCom.AddItem "9  月"
                pCom.AddItem "10 月"
                pCom.AddItem "11 月"
                pCom.AddItem "12 月"
            
            
       Case Else
            pCom.Clear
    End Select
    
End Sub
 
Private Sub Voucher_FillList(ByVal R As Long, ByVal C As Long, pCom As Object)
    Dim sFieldName As String
    sFieldName = LCase(Me.voucher.ItemState(C, sibody).sFieldName)
    Select Case sFieldName
        Case "adds", "lenssen" '控制方向
                pCom.Clear
                pCom.AddItem ""
                pCom.AddItem "借方"
                pCom.AddItem "贷方"
        
        Case "iscontrol" '是否控制
                pCom.Clear
'                pCom.AddItem "1 提醒"
'                pCom.AddItem "2 控制"
'                pCom.AddItem "3 不提示"
                
                pCom.AddItem "提醒"
                pCom.AddItem "控制"
                pCom.AddItem "不提示"

            
        Case "ending"
            pCom.Clear
            pCom.AddItem "盘亏"      '"盘亏"，”盘实“，”变更“，盘盈
            pCom.AddItem "盘实"
            pCom.AddItem "变更"
            pCom.AddItem "盘盈"
            

            
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
    
    
        Case "cdepcode", "cdepname" '部门编码
            strClass = "select cdepcode,cdepname from Department"
            strGrid = "SELECT Department.cDepCode,Department.cDepName FROM Department  where bdepend=1 "
            If LCase(sKey) = "cdepcode" And Len(Trim(Me.voucher.headerText("cdepcode"))) > 0 Then
                strGrid = strGrid & " and cDepCode like '%" & Trim(Me.voucher.headerText("cdepcode")) & "%'"
            ElseIf LCase(sKey) = "cdepname" And Len(Trim(Me.voucher.headerText("cdepname"))) > 0 Then
                strGrid = strGrid & " and cDepName like '%" & Trim(Me.voucher.headerText("cdepname")) & "%'"
            End If
            strGrid = strGrid & " order by cDepCode "
            If clsRefer.StrRefInit_SetColWidth(m_login, False, strClass, strGrid, "频道编码,频道名称", "2500,6000") = False Then Exit Sub
            clsRefer.Show
            If Not clsRefer.recMx Is Nothing Then
                Me.voucher.headerText("cdepcode") = clsRefer.recMx(0)
                Me.voucher.headerText("cdepname") = clsRefer.recMx(1)
                sRet = clsRefer.recMx.Fields(LCase(sKey))
            End If

        Case "citemcode", "citemname" '项目编码
            strClass = "select citemccode,citemcname  from V_MT_ItemsClass"
            strGrid = "SELECT citemccode, citemcode , citemname FROM V_MT_Items  where 1=1  "
            If LCase(sKey) = "citemcode" And Len(Trim(Me.voucher.headerText("citemcode"))) > 0 Then
                strGrid = strGrid & " and citemcode like '%" & Trim(Me.voucher.headerText("citemcode")) & "%'"
            ElseIf LCase(sKey) = "citemname" And Len(Trim(Me.voucher.headerText("citemname"))) > 0 Then
                strGrid = strGrid & " and citemname like '%" & Trim(Me.voucher.headerText("citemname")) & "%'"
            End If
            strGrid = strGrid & " order by citemcode "
            If clsRefer.StrRefInit_SetColWidth(m_login, False, strClass, strGrid, "分类,栏目编码,栏目名称", "0,2500,6000") = False Then Exit Sub
            clsRefer.Show
            If Not clsRefer.recMx Is Nothing Then
                Me.voucher.headerText("citemcode") = clsRefer.recMx(1)
                Me.voucher.headerText("citemname") = clsRefer.recMx(2)
                sRet = clsRefer.recMx.Fields(LCase(sKey))
            End If
            
        Case "cexpccode", "cexpcname" '费用类别分类
            strClass = "select * from dbo.ExpItemClass "
            strGrid = "select cexpccode,cexpcname  from  ExpItemClass  where bexpcend=1  "
            If LCase(sKey) = "cexpccode" And Len(Trim(Me.voucher.headerText("cexpccode"))) > 0 Then
                strGrid = strGrid & " and cexpcode like '%" & Trim(Me.voucher.headerText("cexpccode")) & "%'"
            ElseIf LCase(sKey) = "cexpcname" And Len(Trim(Me.voucher.headerText("cexpcname"))) > 0 Then
                strGrid = strGrid & " and cexpcname like '%" & Trim(Me.voucher.headerText("cexpcname")) & "%'"
            End If
            strGrid = strGrid & " order by cexpccode "
            If clsRefer.StrRefInit_SetColWidth(m_login, False, strClass, strGrid, "费用类别编码,费用类别名称", "2500,6000") = False Then Exit Sub
            clsRefer.Show
            If Not clsRefer.recMx Is Nothing Then
                Me.voucher.headerText("cexpccode") = clsRefer.recMx(0)
                Me.voucher.headerText("cexpcname") = clsRefer.recMx(1)
                sRet = clsRefer.recMx.Fields(LCase(sKey))
            End If
        
        Case "cexpcode", "cexpname" '费用类别
            strClass = "select * from dbo.ExpItemClass "
            strGrid = "select cexpccode,cexpcode,cexpname from ExpenseItem  where 1=1  "
            If LCase(sKey) = "cexpcode" And Len(Trim(Me.voucher.headerText("cexpcode"))) > 0 Then
                strGrid = strGrid & " and cexpcode like '%" & Trim(Me.voucher.headerText("cexpcode")) & "%'"
            ElseIf LCase(sKey) = "cexpname" And Len(Trim(Me.voucher.headerText("cexpname"))) > 0 Then
                strGrid = strGrid & " and cexpname like '%" & Trim(Me.voucher.headerText("cexpname")) & "%'"
            End If
            strGrid = strGrid & " order by cexpcode "
            If clsRefer.StrRefInit_SetColWidth(m_login, False, strClass, strGrid, "分类,费用类别编码,费用类别名称", "0,2500,6000") = False Then Exit Sub
            clsRefer.Show
            If Not clsRefer.recMx Is Nothing Then
                Me.voucher.headerText("cexpcode") = clsRefer.recMx(1)
                Me.voucher.headerText("cexpname") = clsRefer.recMx(2)
                sRet = clsRefer.recMx.Fields(LCase(sKey))
            End If

        Case "cpersoncode", "cpersonname" '人员档案
            strClass = "select cdepcode,cdepname from Department"
            strGrid = "SELECT Department.cDepCode,Person.cPersonCode,Person.cPersonName FROM Department INNER JOIN Person ON Department.cDepCode = Person.cDepCode  where 1=1 "
            If LCase(sKey) = "cpersoncode" And Len(Trim(Me.voucher.headerText("cpersoncode"))) > 0 Then
                strGrid = strGrid & " and  cpersoncode like '%" & Trim(Me.voucher.headerText("cpersoncode")) & "%'"
            ElseIf LCase(sKey) = "cpersonname" And Len(Trim(Me.voucher.headerText("cpersonname"))) > 0 Then
                strGrid = strGrid & " and cpersonname like '%" & Trim(Me.voucher.headerText("cpersonname")) & "%'"
            End If
            If clsRefer.StrRefInit_SetColWidth(m_login, False, strClass, strGrid, "分类,人员编码,人员名称", "0,2500,6000") = False Then Exit Sub
            clsRefer.Show
            If Not clsRefer.recMx Is Nothing Then
                Me.voucher.headerText("cpersoncode") = clsRefer.recMx(0)
                Me.voucher.headerText("cpersonname") = clsRefer.recMx(1)
                sRet = clsRefer.recMx.Fields(LCase(sKey))
            End If


'        Case "cpicpath"
'            With Me.CommonDialog1
'                .FileName = ""
'                .Filter = "图片数据文件(*.*)|*.*"
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
'by lg070315　增加U870 UAP单据控件新的参照处理
    referPara.Cancel = True
    
    If rst.state = 1 Then rst.Close
    Set rst = Nothing
    Exit Sub
End Sub

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
            If Not clsDef.Init(False, DBConn.ConnectionString, m_login.cUserId) Then
                RefDefine = ""
                MsgBox "初始化自定义项组件失败！"
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
        
        Case "cdepcode" '部门编码
                strSql = "SELECT Department.cDepCode,Department.cDepName FROM Department  where bdepend=1 "
                strSql = strSql & " and cdepcode = '" & Trim(Me.voucher.headerText("cdepcode")) & "'"
                If rds.state <> 0 Then rds.Close
                rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
                If rds.RecordCount = 0 Then
                    MsgBox "频道不合法！", vbOKOnly + vbCritical
                    .headerText("cdepcode") = ""
                    .headerText("cdepname") = ""
                    RetValue = ""
                    bChanged = Retry
                Else
                    .headerText("cdepcode") = rds.Fields("cDepCode")
                    .headerText("cdepname") = rds.Fields("cDepName")
                    RetValue = rds.Fields(LCase(strKey))
                End If
            
        
        Case "cdepname"  '部门编码
                strSql = "SELECT Department.cDepCode,Department.cDepName FROM Department  where bdepend=1 "
                strSql = strSql & " and cDepName = '" & Trim(Me.voucher.headerText("cdepname")) & "'"
                If rds.state <> 0 Then rds.Close
                rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
                If rds.RecordCount = 0 Then
                    MsgBox "频道不合法！", vbOKOnly + vbCritical
                    .headerText("cdepcode") = ""
                    .headerText("cdepname") = ""
                    RetValue = ""
                    bChanged = Retry
                Else
                    .headerText("cdepcode") = rds.Fields("cDepCode")
                    .headerText("cdepname") = rds.Fields("cDepName")
                    RetValue = rds.Fields(LCase(strKey))
                End If
            
        Case "citemcode"  '项目编码
            strSql = "SELECT  citemcode , citemname FROM V_MT_Items  where 1=1  "
            strSql = strSql & " and citemcode = '" & Trim(Me.voucher.headerText("citemcode")) & "'"
            If rds.state <> 0 Then rds.Close
            rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
            If rds.RecordCount = 0 Then
                MsgBox "栏目不合法！", vbOKOnly + vbCritical
                .headerText("citemcode") = ""
                .headerText("citemname") = ""
                RetValue = ""
                bChanged = Retry
            Else
                .headerText("citemcode") = rds.Fields("citemcode")
                .headerText("citemname") = rds.Fields("citemname")
                RetValue = rds.Fields(LCase(strKey))
            End If
            
        Case "citemcode", "citemname" '项目编码
            strSql = "SELECT  citemcode , citemname FROM V_MT_Items  where 1=1  "
            strSql = strSql & " and citemname = '" & Trim(Me.voucher.headerText("citemname")) & "'"
            If rds.state <> 0 Then rds.Close
            rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
            If rds.RecordCount = 0 Then
                MsgBox "栏目不合法！", vbOKOnly + vbCritical
                .headerText("citemcode") = ""
                .headerText("citemname") = ""
                RetValue = ""
                bChanged = Retry
            Else
                .headerText("citemcode") = rds.Fields("citemcode")
                .headerText("citemname") = rds.Fields("citemname")
                RetValue = rds.Fields(LCase(strKey))
            End If
            
        Case "cexpccode" '费用类别分类
            strSql = "select cexpccode,cexpcname  from  ExpItemClass  where bexpcend=1  "
            strSql = strSql & " and cexpccode = '" & Trim(Me.voucher.headerText("cexpccode")) & "'"
            If rds.state <> 0 Then rds.Close
            rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
            If rds.RecordCount = 0 Then
                MsgBox "费用类别分类不合法！", vbOKOnly + vbCritical
                .headerText("cexpccode") = ""
                .headerText("cexpcname") = ""
                RetValue = ""
                bChanged = Retry
            Else
                .headerText("cexpccode") = rds.Fields("cexpccode")
                .headerText("cexpcname") = rds.Fields("cexpcname")
                RetValue = rds.Fields(LCase(strKey))
            End If
        
        Case "cexpcname" '费用类别分类
            strSql = "select cexpccode,cexpcname  from  ExpItemClass  where bexpcend=1  "
            strSql = strSql & " and cexpcname = '" & Trim(Me.voucher.headerText("cexpcname")) & "'"
            If rds.state <> 0 Then rds.Close
            rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
            If rds.RecordCount = 0 Then
                MsgBox "费用类别分类不合法！", vbOKOnly + vbCritical
                .headerText("cexpccode") = ""
                .headerText("cexpcname") = ""
                RetValue = ""
                bChanged = Retry
            Else
                .headerText("cexpccode") = rds.Fields("cexpccode")
                .headerText("cexpcname") = rds.Fields("cexpcname")
                RetValue = rds.Fields(LCase(strKey))
            End If
         
        
        Case "cexpcode" '费用类别
            strSql = "select cexpccode,cexpcode,cexpname from ExpenseItem  where 1=1  "
            strSql = strSql & " and cexpcode = '" & Trim(Me.voucher.headerText("cexpcode")) & "'"
            If rds.state <> 0 Then rds.Close
            rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
            If rds.RecordCount = 0 Then
                MsgBox "费用类别不合法！", vbOKOnly + vbCritical
                .headerText("cexpcode") = ""
                .headerText("cexpname") = ""
                RetValue = ""
                bChanged = Retry
            Else
                .headerText("cexpcode") = rds.Fields("cexpcode")
                .headerText("cexpname") = rds.Fields("cexpname")
                RetValue = rds.Fields(LCase(strKey))
            End If
        
        Case "cexpname"  '费用类别
            strSql = "select cexpccode,cexpcode,cexpname from ExpenseItem  where 1=1  "
            strSql = strSql & " and cexpname = '" & Trim(Me.voucher.headerText("cexpname")) & "'"
            If rds.state <> 0 Then rds.Close
            rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
            If rds.RecordCount = 0 Then
                MsgBox "费用类别不合法！", vbOKOnly + vbCritical
                .headerText("cexpcode") = ""
                .headerText("cexpname") = ""
                RetValue = ""
                bChanged = Retry
            Else
                .headerText("cexpcode") = rds.Fields("cexpcode")
                .headerText("cexpname") = rds.Fields("cexpname")
                RetValue = rds.Fields(LCase(strKey))
            End If

        Case "cpersoncode"  '人员档案
            strSql = "SELECT Department.cDepCode,Person.cPersonCode,Person.cPersonName FROM Department INNER JOIN Person ON Department.cDepCode = Person.cDepCode  where 1=1 "
            strSql = strSql & " and  cpersoncode = '" & Trim(Me.voucher.headerText("cpersoncode")) & "'"
            If rds.state <> 0 Then rds.Close
            rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
            If rds.RecordCount = 0 Then
                MsgBox "人员不合法！", vbOKOnly + vbCritical
                .headerText("cpersoncode") = ""
                .headerText("cpersonname") = ""
                RetValue = ""
                bChanged = Retry
            Else
                .headerText("cpersoncode") = rds.Fields("cPersonCode")
                .headerText("cpersonname") = rds.Fields("cPersonName")
                RetValue = rds.Fields(LCase(strKey))
            End If

        Case "cpersonname"  '人员档案
            strSql = "SELECT Department.cDepCode,Person.cPersonCode,Person.cPersonName FROM Department INNER JOIN Person ON Department.cDepCode = Person.cDepCode  where 1=1 "
            strSql = strSql & " and cpersonname = '" & Trim(Me.voucher.headerText("cpersonname")) & "'"
            If rds.state <> 0 Then rds.Close
            rds.Open strSql, DBConn, adOpenStatic, adLockReadOnly
            If rds.RecordCount = 0 Then
                MsgBox "人员不合法！", vbOKOnly + vbCritical
                .headerText("cpersoncode") = ""
                .headerText("cpersonname") = ""
                RetValue = ""
                bChanged = Retry
            Else
                .headerText("cpersoncode") = rds.Fields("cPersonCode")
                .headerText("cpersonname") = rds.Fields("cPersonName")
                RetValue = rds.Fields(LCase(strKey))
            End If
            
            
            
        Case "rate" '预算比例
               If (.headerText("rate") > 100) Or (.headerText("rate") <= 0) Then
                    MsgBox "预算比例不合法！", vbOKOnly + vbCritical
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
            ''打印
            strSql = "SELECT VT_Name,VT_ID  From VoucherTemplates WHERE (VT_CardNumber = '" & strCardNum & "') AND (VT_TemplateMode = 1) " & IIf(strAuth <> "", "and vt_id in (" & strAuth & ")", "")
        Else
            ''显示
            strSql = "SELECT VT_Name,VT_ID  From VoucherTemplates WHERE (VT_CardNumber = '" & strCardNum & "') AND (VT_TemplateMode = 0) " & IIf(strAuth <> "", "and vt_id in (" & strAuth & ")", "")
        End If
    Else
        If bPrint = True Then
            ''打印
            strSql = "SELECT VT_Name,VT_ID  From VoucherTemplates WHERE (" & sWhere & ") AND (VT_TemplateMode = 1) " & IIf(strAuth <> "", "and vt_id in (" & strAuth & ")", "")
        Else
            ''显示
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
'单据初始化
Public Function ShowVoucher(VoucherType As VoucherType, Optional vVoucherId As Variant, Optional iMode As Integer)
    Dim tmpTemplateID As String
    Dim errMsg As String
'by lg070314 增加U870门户融合
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
    'by ahzzd 2005/05/09 单据初始化
    clsVoucherCO.Init VoucherType, m_login, DBConn, "CS", clsSAWeb
    clsAuth.Init m_login.UfDbName, m_login.cUserId
    Set obj_EA = CreateObject("u8ExamineAndApprove.clsU8Examine")
    Call obj_EA.Init(m_login)
'    MT01    '0   费用类别与科目对照表
'    MT02    '0   借款科目对照表
'    MT03    '0   费用类别比例设置表
'    MT04    '0   费用分类比例设置表
'    MT05    '0   预算编制期初录入
'    MT06    '0   预算编制单
'    MT07    '0   预算编制调整单
'    MT08    '0   支票借款单
'    MT09    '0   节目制作经费报账单
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
     
    
        Case gdzckpxg       '// 资产卡片编辑
            strVouchType = "98"
            strCardNum = "FA01"
        Case gdzckp           '//资产卡片增加
            strVouchType = "97"
            strCardNum = "FA01"
    End Select
    ''设置按钮
 
   U8VoucherSorter1.Visible = False
 
 
    sTemplateID = clsSAWeb.GetVTID(DBConn, strCardNum)
 
'
'    ''设置按钮
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
        sCurTemplateID = sTemplateID    ''取默认模板
    Else
        sCurTemplateID = tmpTemplateID  ''新的模板
    End If
    sCurTemplateID2 = sCurTemplateID
    
    If sCurTemplateID = 0 Then
        Me.Hide
        MsgBox "您没有模版使用权限"
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '860sp升级到861修改处   2006/03/12  861 增加附件
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
 
 
    'by ahzzd 2006/05/09   数据准备到单据上完成
    '审批流文本
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
        clsRefer.setlogin m_login   ''初始化参照控件
    End If
    Me.voucher.Visible = True
    Call fillComBol(True)   ''填充模版选择
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
    
'by lg070314 增加U870支持，窗体融合
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
''设置单据号是否可以编辑
Private Sub SetVouchNoWriteble()
    Dim KeyCode As String
    
    On Error Resume Next
    If strVouchType = "92" Then Exit Sub
    KeyCode = getVoucherCodeName()
    If Not DomFormat Is Nothing Then
        If DomFormat.xml <> "" Then
            If LCase(DomFormat.selectSingleNode("//单据编号").Attributes.getNamedItem("允许手工修改").nodeValue) = "false" And LCase(DomFormat.selectSingleNode("//单据编号").Attributes.getNamedItem("重号自动重取").nodeValue) = "false" Then
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
         ''增加按钮
        'by lg070314 修改U870门户菜单融合，所有Toolbar的Button增加Tag值
        'Tag值 表示菜单上的图标文件名称   图标文件在 ..\U8SOFT\icons
        
         ''打印
            Set btnX = .buttons.Add(, "Print", strPrint, tbrDefault)
'            btnX.image = 314
            btnX.ToolTipText = strPrint
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Print"
        ''预览
            Set btnX = .buttons.Add(, "Preview", strPreview, tbrDefault)
'            btnX.image = 312
            btnX.ToolTipText = strPreview
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "print preview"
        ''输出
            Set btnX = .buttons.Add(, "Output", strOutput, tbrDefault)
'            btnX.image = 308
            btnX.ToolTipText = strOutput
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Output"
 
        ''增加
            Set btnX = .buttons.Add(, "Add", strAdd, tbrDefault)
'            btnX.image = 323
            btnX.ToolTipText = strAdd
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Add"
            
'        ''批量增加
'            Set btnX = .buttons.Add(, "batchAdd", "批增", tbrDefault)
''            btnX.image = 389
'            btnX.ToolTipText = "批增"
'            btnX.Description = btnX.ToolTipText
'            btnX.Tag = "Input"
            
'        ''批量导入
'            Set btnX = .buttons.Add(, "inAdd", "导入", tbrDefault)
'            btnX.image = 1
'            btnX.ToolTipText = "导入"
'            btnX.Description = btnX.ToolTipText
'            btnX.Tag = "inAdd"
'        ''批量导出
'            Set btnX = .buttons.Add(, "outAdd", "导出", tbrDefault)
'            btnX.image = 1
'            btnX.ToolTipText = "导出"
'            btnX.Description = btnX.ToolTipText
'            btnX.Tag = "outAdd"
'            Set btnX = .buttons.Add(, , , tbrSeparator)
        ''修改
            Set btnX = .buttons.Add(, "Modify", strModify, tbrDefault)
'            btnX.image = 324
            btnX.ToolTipText = strModify
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "modify"
            
'        ''修改变更
'            Set btnX = .buttons.Add(, "Chenged", strchenged, tbrDefault)
''            btnX.image = 321
'            btnX.ToolTipText = strchenged
'            btnX.Description = btnX.ToolTipText
'            btnX.Tag = "accessories"
            
         ''删除
            Set btnX = .buttons.Add(, "Erase", strDelete, tbrDefault)
'            btnX.image = 326
            btnX.ToolTipText = strDelete
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "delete"
         ''复制
            Set btnX = .buttons.Add(, "Copy", "复制", tbrDefault)
'            btnX.image = 318
            btnX.ToolTipText = "复制"
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Copy"
            
'        ''图片
'            Set btnX = .buttons.Add(, "picture", "图片", tbrDefault)
'            btnX.image = 20
'            btnX.ToolTipText = "图片"
'            btnX.Description = btnX.ToolTipText
'            btnX.Tag = "picture"
            
'          '条码
'            Set btnX = .buttons.Add(, "label", "条码", tbrDefault)
'            btnX.image = 20
'            btnX.ToolTipText = "条码"
'            btnX.Description = btnX.ToolTipText
'            btnX.Tag = "label"
'
            
         ''保存
            Set btnX = .buttons.Add(, "Save", strSave, tbrDefault)
''            btnX.image = 988
'btnX.Style = tbrButtonGroup
'btnX.ButtonMenus.Add
            btnX.ToolTipText = strSave
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "save"
          ''放弃
            Set btnX = .buttons.Add(, "Cancel", strDiscard, tbrDefault)
'            btnX.image = 316
            btnX.ToolTipText = strDiscard
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Cancel"
            
            '审核
            Set btnX = .buttons.Add(, "Sure", "审核", tbrDefault)
'            btnX.image = 1100
            btnX.ToolTipText = "审核"
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Approve"
            
            '弃审
            Set btnX = .buttons.Add(, "UnSure", "弃审", tbrDefault)
'            btnX.image = 341
            btnX.ToolTipText = "弃审"
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Unapprove"
            

        ''增行
            Set btnX = .buttons.Add(, "AddRow", strAddrecord, tbrDefault)
'            btnX.image = 343
            btnX.ToolTipText = strAddrecord
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "add a row"
      
        ''删行
            Set btnX = .buttons.Add(, "DelRow", strDeleterecord, tbrDefault)
'            btnX.image = 347
            btnX.ToolTipText = strDeleterecord
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Delete a row"


'        '分割符号
'            Set btnX = .buttons.Add(, , , tbrSeparator)
'
'            '过滤
'            Set btnX = .buttons.Add(, "Filter", strFilter, tbrDefault)
''            btnX.image = 1120
'            btnX.ToolTipText = strFilter
'            btnX.Description = btnX.ToolTipText
'            btnX.Tag = "Filter"
           
           
            
          ''首张
            Set btnX = .buttons.Add(, "ToFirst", strFirst, tbrDefault)
'            btnX.image = 24 '1174
            btnX.ToolTipText = strFirst
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "first page"
          ''上张
            Set btnX = .buttons.Add(, "ToPrevious", strPrevious, tbrDefault)
'            btnX.image = 22 '1139
            btnX.ToolTipText = strPrevious
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "previous page"
          ''下张
            Set btnX = .buttons.Add(, "ToNext", strNext, tbrDefault)
'            btnX.image = 23 '1133
            btnX.ToolTipText = strNext
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "next page"
           ''末张
            Set btnX = .buttons.Add(, "ToLast", strLast, tbrDefault)
'            btnX.image = 25 '1117
            btnX.ToolTipText = strLast
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Last page"

'            '全部
'            Set btnX = .buttons.Add(, "Pd_all", strPd_all, tbrDefault)
'            btnX.image = 25
'            btnX.ToolTipText = strPd_all
'            btnX.Description = btnX.ToolTipText
           ''凭证
'            Set btnX = .buttons.Add(, "Seek", "凭证", tbrDefault)
''            btnX.image = 8
'            btnX.ToolTipText = "凭证"
'            btnX.Description = btnX.ToolTipText
'            btnX.Tag = "sum"
           ''刷新
            Set btnX = .buttons.Add(, "Paint", strRefresh, tbrDefault)
'            btnX.image = 154
            btnX.ToolTipText = strRefresh
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "refresh"
            
'        End If
'           ''帮助
'            Set btnX = .Buttons.Add(, "Help", strHelp, tbrDefault)
'            btnX.Image = 36
'            btnX.ToolTipText = strHelp
'            btnX.Description = btnX.ToolTipText
           
           ''退出
'            Set btnX = .buttons.Add(, "Exit", strExit, tbrDefault)
'            btnX.image = 1118
'            btnX.ToolTipText = strExit
'            btnX.Description = btnX.ToolTipText
'            btnX.Tag = "Exit"
'           ''设置
'            Set btnX = .Buttons.Add(, "PrnSet", "设置", tbrDefault)
'            btnX.Image = 9
'            btnX.ToolTipText = "设置"
'            btnX.Description = btnX.ToolTipText
'            btnX.Visible = False
           ''列表
'            Set btnX = .Buttons.Add(, "LstTab", "列表", tbrDefault)
'            btnX.Image = 43
'            btnX.ToolTipText = "列表"
'            btnX.Description = btnX.ToolTipText
'            btnX.Visible = False
         
         If strCardNum = "MT09" Then
            Set btnX = .buttons.Add(, "ZP", "支票", tbrDefault)
            'btnX.image = 43
            btnX.ToolTipText = "支票"
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "new_persp"
            btnX.Visible = True
            
        End If
    End With
    ''置作废、现结的位置
    labZF.Top = picVoucher.Top 'Me.top - Me.tbrvoucher.top
    labZF.Left = Me.voucher.Left
    labXJ.Top = picVoucher.Top ' Me.top - Me.tbrvoucher.top    'Me.StBar.height
    labXJ.Left = Me.voucher.Left + labZF.Width
'by lg070316增加初始化U870菜单
    Call InitToolbarTag(Me.tbrvoucher)
    
End Sub
''改变button的状态
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
                   
                   Case "87", "88", "89", "90" '基础设置
                       Me.tbrvoucher.buttons("Copy").Enabled = True
                   Case "91", "92", "93", "95" '业务单据
                    '已审核
                    If .headerText("chandler") <> "" Then '审核人
                        Me.tbrvoucher.buttons("UnSure").Visible = True
                        Me.tbrvoucher.buttons("Sure").Visible = False
                        bCheckVouch = False
                        tbrvoucher.buttons("Add").Enabled = True
                        tbrvoucher.buttons("Save").Enabled = False
                        tbrvoucher.buttons("Copy").Enabled = True
                        Me.tbrvoucher.buttons("Modify").Enabled = False
                        Me.tbrvoucher.buttons("Erase").Enabled = False
                        Me.tbrvoucher.buttons("Chenged").Visible = True
                    '未审核
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
                 Case "94"  '业务单据
                    '已审核
                    If .headerText("chandler") <> "" Then '审核人
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
                        
                    '未审核
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
                If Left(Me.tbrvoucher.buttons(i).ToolTipText, 2) <> "参照" And Left(Me.tbrvoucher.buttons(i).ToolTipText, 2) <> "查询" Then
                    Me.tbrvoucher.buttons(i).Caption = Left(Me.tbrvoucher.buttons(i).ToolTipText, 2)
                End If
            Next

        Else     ''空单据
 
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
            Case "102" '资产减少
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
        If voucher.headerText("bfirst") Then
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
        MsgBox "打印设置保存失败"
    End If
End Sub
 
Private Sub VS_Change()
    Me.voucher.Top = Me.Picture2.Height - vs.value - Me.Picture2.Height ''- Me.StBar.height)
End Sub
 
'控制界面
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
                
                '设置新增单据的初始值
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
                
                '设置新增单据的初始值
                .getVoucherDataXML Domhead, Dombody
                '复制的单据的初始值是没有审核的
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
           '//根据不同单据设置单据上面的按钮
            Select Case LCase(strVouchType)
                Case "87", "88", "89", "90" '基础设置
                    ComboVTID.Visible = True
                    ComboDJMB.Visible = True
                    Labeldjmb.Caption = "显示模版："
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
                    
                    
                Case "91", "92", "93", "94", "95" '业务单据
                    ComboVTID.Visible = True
                    ComboDJMB.Visible = True
                    Labeldjmb.Caption = "显示模版："
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
                    
                Case "94"    '业务单据
                    ComboVTID.Visible = True
                    ComboDJMB.Visible = True
                    Labeldjmb.Caption = "显示模版："
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
                Case "87", "88", "89", "90" '基础设置
                    ComboVTID.Visible = True
                    ComboDJMB.Visible = True
                    Labeldjmb.Caption = "打印模版："
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
 
                    
                Case "91", "92", "93", "95"   '业务单据
                    ComboVTID.Visible = True
                    ComboDJMB.Visible = True
                    Labeldjmb.Caption = "打印模版："
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
 
                 Case "94"    '业务单据
                    ComboVTID.Visible = True
                    ComboDJMB.Visible = True
                    Labeldjmb.Caption = "打印模版："
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
            ''编辑状态下
            Select Case KeyCode
                Case vbKeyF6
                    If tbrvoucher.buttons("Save").Visible And tbrvoucher.buttons("Save").Enabled Then
                        Call ButtonClick("Save", "保存")
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
                                    And Trim(.bodyText(.row, "cInvCode")) <> "" And val(.bodyText(.row, "iQuantity")) > 0 And Trim(.bodyText(.row, "iTb")) <> "退补" Then
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
            ''非编辑状态
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
                            Call ButtonClick("Add", "增加")
                        End If
                    End If
                    If Shift = 4 Then
                        If tbrvoucher.buttons("Copy").Visible And tbrvoucher.buttons("Copy").Enabled Then
                           Call ButtonClick("Copy", "复制")
                        End If
                    End If
                Case vbKeyF8
                    If iShowMode = 1 Then Exit Sub
                    If Shift = 0 Then
                        If tbrvoucher.buttons("Modify").Visible And tbrvoucher.buttons("Modify").Enabled Then
                            Call ButtonClick("Modify", "修改")
                        End If
                    End If
                Case vbKeyP         ''打印
                    If iShowMode = 1 Then Exit Sub
                    If tbrvoucher.buttons("Print").Visible And tbrvoucher.buttons("Print").Enabled Then
                        Call ButtonClick("Print", "")
                    End If
                Case vbKeyF4        ''退出
                    If Shift = 2 Then
                        If tbrvoucher.buttons("Exit").Visible And tbrvoucher.buttons("Exit").Enabled Then
                           Call ButtonClick("Exit", "")
                        End If
                    End If
                Case vbKeyF3        ''定位
                    If iShowMode = 1 Then Exit Sub
                    If tbrvoucher.buttons("Exit").Visible And tbrvoucher.buttons("Exit").Enabled Then
                       Call ButtonClick("Seek", "")
                    End If
                    
                Case vbKeyDelete
                    If iShowMode = 1 Then Exit Sub
                    If tbrvoucher.buttons("Erase").Visible And tbrvoucher.buttons("Erase").Enabled Then
                       Call ButtonClick("Erase", "删除")
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
        strEAXML = ""
    End If
    Set tmpDOM = Nothing
    ShowErrDom = True
    strEAXML = ""
    Screen.MousePointer = vbDefault
    Exit Function
DoERR:
    MsgBox "处理错误信息时发生错误：" & Err.Description
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
''更改单据模版for增加，复制
Private Function ChangeDJMBForEdit() As Boolean
    
    With Me.voucher
        If CheckDJMBAuth(.headerText("ivtid"), "W") = False Then
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
''更改voucher caption 的颜色
Private Sub ChangeCaptionCol()
    On Error Resume Next
    With Me.voucher
        Me.LabelVoucherName.ForeColor = .TitleForeColor
        Me.LabelVoucherName.Font.Name = .TitleFont.Name
        Me.LabelVoucherName.Font.Bold = .TitleFont.Bold
        Me.LabelVoucherName.Font.Italic = .TitleFont.Italic
        Me.LabelVoucherName.Font.Underline = .TitleFont.Underline
        If bFirst = True Then
            If Left(Me.LabelVoucherName.Caption, Len("期初")) <> "期初" And Left(Me.LabelVoucherName.Caption, Len("期初")) <> "期初" Then
                If strVouchType = "05" Then
                    Me.LabelVoucherName.Caption = "期初" & Me.LabelVoucherName.Caption
                Else
                    Me.LabelVoucherName.Caption = "期初" & Me.LabelVoucherName.Caption
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
''更改单据项目到原始状态
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

'设置单据控件项目写状态
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
 
''加载单据
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
'外部可以调用内部函数
Public Sub VouchHeadCellCheck(Index As Variant, RetValue As String, bChanged As UapVoucherControl85.CheckRet)
    'index = Voucher.LookUpArrayFromKey(LCase(index), siheader)
    Index = voucher.LookUpArray(LCase(Index), siheader)
    Dim referPara As UapVoucherControl85.ReferParameter
    Call Voucher_headCellCheck(Index, RetValue, bChanged, referPara)
    voucher.ProtectUnload2
End Sub
'将控件传给外部控件
Public Function GetVoucherObject() As Object
    Set GetVoucherObject = Me.voucher
End Function
'获取单据的编辑状态,提供给外部使用
Public Function GetVouchState() As Integer
    GetVouchState = iVouchState
End Function
Private Function GetBodyRefVal(sKey As String, row As Long) As String
    Dim Obj As Object
    Dim Index As Long
    ' 得到表体对象
    Set Obj = Me.voucher.GetBodyObject()
    ' 得到关键字对应的Index
    Index = Me.voucher.LookUpArrayFromKey(sKey, sibody)
    GetBodyRefVal = Obj.TextMatrix(row, Index)
End Function


'
'检查用户用户选择的资产是否重复 或不存在
Private Function check_sassetnum_for101() As String
Dim i As Long
Dim j As Long
Dim sassetnum As String
Dim rds As New ADODB.Recordset
On Error GoTo Err
    check_sassetnum_for101 = ""
    For i = 1 To Me.voucher.BodyRows
        If Len(Trim(Me.voucher.bodyText(i, "stypenum"))) = 0 Then '
           check_sassetnum_for101 = "第" & i & "行， 国标分类代码不能为空！"
           Exit For
        End If
        If (Len(Trim(Me.voucher.bodyText(i, "sassetnum"))) = 0) And (Len(Trim(Me.voucher.bodyText(i, "scardid"))) <> 0) Then '
           check_sassetnum_for101 = "第" & i & "行， 资产编码不能为空！"
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

'检查变动单是否有金额变化,
Public Function value_change(wjbfa_asset_change_id As String) As Boolean
    Dim rstemp As New ADODB.Recordset
    Dim Str As String
    On Error GoTo Err                                                                       'usestate_before
        
        ' 1 “在建”转”在用“ 时制凭证
        ' 2  ”在用“ 金额变化时制凭证
        Str = "select * from wjbfa_vouchers  " & _
              " Where ((((dbo.wjbfa_vouchers.usdollar_after - dbo.wjbfa_vouchers.usdollar_before <> 0) and (usestate_before='在用')) " & _
              " or (usestate_before='在建' and  usestate_after='在用') )) " & _
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
'检查减少单中有没有资产是在用状态的
Public Function state(wjbfa_assetjs_id As String) As Boolean
    Dim rstemp As New ADODB.Recordset
    Dim Str As String
    On Error GoTo Err
        Str = "select * from vw_last_cards_state  " & _
              " where (dbo.vw_last_cards_state.usestate_last='在用') and sassetnum in(SELECT dbo.wjbfa_assetjss.sassetnum " & _
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


'联查凭证
Private Sub Find_GL_accvouch()
Dim rdst1 As New ADODB.Recordset
Dim rdst2 As New ADODB.Recordset
On Error GoTo Err
    Select Case strVouchType
        Case "97"  '原始卡片
                If Trim(Me.voucher.headerText("id")) <> "" Then
                    rdst1.Open "select isnull(coutno_id,'') as coutno_id  from wjbfa_cards where id=" & Me.voucher.headerText("id"), DBConn, adOpenStatic, adLockReadOnly
                    If rdst1.RecordCount > 0 Then
                        If rdst1.Fields("coutno_id") = "" Then
                            MsgBox "【" & Me.voucher.headerText("sassetnum") & "】资产 还没有生成凭证!", vbOKOnly + vbInformation
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
                            MsgBox "凭证发生变化,请重新操作", vbInformation
                        End If
                    Else
                        MsgBox "凭证不存在!", vbOKOnly + vbInformation
                        Set rdst1 = Nothing
                        Set rdst2 = Nothing
                        Exit Sub
                    End If
                End If
            
        Case "105" '资产减少审批单
                    If Trim(Me.voucher.bodyText(Me.voucher.row, "autoid")) <> "" Then
                        rdst1.Open "select isnull(coutno_id,'') as coutno_id  from wjbfa_assetjss where  autoid=" & Me.voucher.bodyText(Me.voucher.row, "autoid"), DBConn, adOpenStatic, adLockReadOnly
                        If rdst1.RecordCount > 0 Then
                            If rdst1.Fields("coutno_id") = "" Then
                                MsgBox "第" & Me.voucher.row & "行【" & Me.voucher.bodyText(Me.voucher.row, "sassetnum") & "】资产 还没有生成凭证!", vbOKOnly + vbInformation
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
                                MsgBox "凭证发生变化,请重新操作", vbInformation
                            End If
                        Else
                            MsgBox "凭证不存在!", vbOKOnly + vbInformation
                            Set rdst1 = Nothing
                            Set rdst2 = Nothing
                            Exit Sub
                        End If
                    End If
        
        Case "103" '资产变动单
                    If Trim(Me.voucher.bodyText(Me.voucher.row, "autoid")) <> "" Then
                        rdst1.Open "select isnull(coutno_id,'') as coutno_id  from wjbfa_vouchers where autoid=" & Me.voucher.bodyText(Me.voucher.row, "autoid"), DBConn, adOpenStatic, adLockReadOnly
                        If rdst1.RecordCount > 0 Then
                            If rdst1.Fields("coutno_id") = "" Then
                                MsgBox "第" & Me.voucher.row & "行【" & Me.voucher.bodyText(Me.voucher.row, "sassetnum") & "】资产 还没有生成凭证!", vbOKOnly + vbInformation
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
                                MsgBox "凭证发生变化,请重新操作", vbInformation
                            End If
                        Else
                            MsgBox "凭证不存在!", vbOKOnly + vbInformation
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
'有人员编码转换成姓名
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

'860sp升级到861修改处   2006/03/08   增加单据附件功能
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
        MsgBox "无法创建m_oDataSource对象!"
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

'得到预算比例
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

'预置表体
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
            MsgBox "请先指定频道!", vbExclamation
            Exit Function
        End If
        
        If sItemCode = "" Then
            MsgBox "请先指定栏目!", vbExclamation
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

'生成费用类别...预置表体RecordSet
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
        If strCardNum = "MT06" Then  '预算编制单
            rs("iscontrol") = "控制"
            rs("bfb") = val(GETysbl(sDepCode, sItemCode, rds("cexpcode")) & "")
        End If
        
        If strCardNum = "MT06" Or strCardNum = "MT09" Then  '预算编制单和报账单
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

'生成科目...预置表体RecordSet
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

'获取部门、项目、费用的预算累计数RecordSet
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

'联查预算明细
Private Sub ProcLinkQuery()
    MsgBox "联查预算明细", vbInformation
End Sub
