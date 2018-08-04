VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86808282-58F4-4B17-BBCA-951931BB7948}#2.82#0"; "U8VouchList.ocx"
Object = "{201FB79B-5556-47A4-AD9C-A46BA0C45A44}#6.42#0"; "UFToolBarCtrl.ocx"
Object = "{005151D4-33F6-471B-B651-222D4E411832}#4.5#0"; "UFFormPartner.ocx"
Object = "{AF8BBBB7-94C6-4772-B826-624478C37D6A}#1.5#0"; "UFKEYHOOK.ocx"
Begin VB.Form frmAutoVoucherList 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   7755
   ClientLeft      =   5700
   ClientTop       =   1965
   ClientWidth     =   10290
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   10290
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar bar 
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   4440
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   1440
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   393
      TabIndex        =   4
      Top             =   5280
      Visible         =   0   'False
      Width           =   5925
      Begin MSComctlLib.ProgressBar prgMsg 
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label3 
         Caption         =   "正在统计"
         Height          =   165
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   5235
      End
   End
   Begin UFKeyHook.UFKeyHookCtrl UFKeyHookCtrl1 
      Left            =   1800
      Top             =   6840
      _ExtentX        =   1905
      _ExtentY        =   529
   End
   Begin UFFormPartner.UFFrmCaption UFFrmCaptionMgr 
      Left            =   240
      Top             =   5400
      _ExtentX        =   1296
      _ExtentY        =   661
      Caption         =   "Form1"
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
   Begin UFToolBarCtrl.UFToolbar UFToolbar1 
      Height          =   240
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   4695
      _ExtentX        =   8281
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
   Begin U8VouchList.PagedivCtl PagedivCtl1 
      Height          =   30
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   53
   End
   Begin U8VouchList.VouchList VchLst 
      Height          =   2085
      Left            =   840
      TabIndex        =   1
      Top             =   2280
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   3678
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "frmAutoVoucherList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_coni As IPagedivConi '条件，基本上都是从U8Colset中进行初始化
Private WithEvents m_mht As UFPortalProxyMessage.IMessageHandler
Attribute m_mht.VB_VarHelpID = -1
Private WithEvents m_pagediv As Pagediv   '分页引擎
Attribute m_pagediv.VB_VarHelpID = -1
Public clsVoucherLst As New clsAutoVoucherList
Dim m_strColumnSetKey As String
Dim m_Caption As String
Dim m_strAuthID As String
Dim m_strToolBarName As String
Dim m_strHelpId As String
Dim clsTbl As New clsAutoToolBar
Dim m_strFormCaption As String
Dim strVouchtype As String
Dim m_strFormGuid As String
Private sLogName As String
Private sLogName1 As String
Dim clsVoucherCO As New EFVoucherCo.clsVoucherCO
Dim MainCode() As Long
Dim m_bshowSumType As Boolean '是否汇总显示


Public Userdll_UI As New UserDefineDll_UI

Private strUserErr As String '用户插件错误信息
Private UserbSuc As Boolean   '插件执行状态   =true 表示成功  =false 表示失败

Private fltsrv As UFGeneralFilter.FilterSrv
Private m_BehaviorObject As New clsReportCallBack

'取得单据类型对应的模版号 主表编号 870审批流
Private Sub GetCardNumberMid(CardNum As String, Mid As String, Tblname As String, row As Long)
    Select Case LCase(strVouchtype)
    Case "26", "27", "28", "29" '发票
        CardNum = "07": Mid = VchLst.TextMatrix(row, VchLst.GridColIndex("sbvid")): Tblname = "SaleBillVouch"
    Case "05", "06", "00" '发货
        CardNum = "01": Mid = VchLst.TextMatrix(row, VchLst.GridColIndex("dlid")): Tblname = "DispatchList"
    Case "97" '订单
        CardNum = "17": Mid = VchLst.TextMatrix(row, VchLst.GridColIndex("id")): Tblname = "SO_SOMain"
    Case "16" '报价单
        CardNum = "16": Mid = VchLst.TextMatrix(row, VchLst.GridColIndex("id")): Tblname = "SA_QuoMain"
    Case "98" '代垫
        CardNum = "08": Mid = VchLst.TextMatrix(row, VchLst.GridColIndex("id")): Tblname = "ExpenseVouch"
    Case "99" '费用支出
        CardNum = "09": Mid = VchLst.TextMatrix(row, VchLst.GridColIndex("id")): Tblname = "SalePayVouch"
    Case "07" '结算
        CardNum = "02": Mid = VchLst.TextMatrix(row, VchLst.GridColIndex("id")): Tblname = "SA_SettleVouch"
    Case "sa18"
        CardNum = "SA18": Mid = VchLst.TextMatrix(row, VchLst.GridColIndex("id")): Tblname = "SA_InvPriceJustMain"
    Case "sa19"
        CardNum = "SA19": Mid = VchLst.TextMatrix(row, VchLst.GridColIndex("id")): Tblname = "SA_CusPriceJustMain"
'    Case "95", "92" '包装物
'        cardNum = "10": Mid = "autoid"
    End Select
End Sub
Private Function GetReturnFlagBool(RfName As String) As Boolean
    Dim RfName_ As String
    Dim sql As String
    Dim Rst As New ADODB.Recordset
    sql = "select top 1 enumcode from aa_enum where enumtype=N'sa.boolean' and enumname=N'" + RfName + "'"
    Rst.Open sql, DBConn, adOpenStatic, adLockReadOnly
    
    If Rst.BOF And Rst.EOF Then
        RfName_ = LCase(RfName)
    Else
        RfName_ = Rst.Fields(0).value
    End If
    Select Case RfName_
    Case "是", "1", "yes", "是"
        GetReturnFlagBool = True
    Case "否", "0", "否", "no"
        GetReturnFlagBool = False
    End Select
    Rst.Close
    Set Rst = Nothing
End Function

Private Function setVouchtype() As String
    Dim rstFilter As String
    On Error Resume Next
    Toolbar1.buttons("closebatch").Visible = False
    Toolbar1.buttons("openbatch").Visible = False
    Toolbar1.buttons("lockbatch").Visible = False
    Toolbar1.buttons("unlockbatch").Visible = False
    If getIsWfControl(DBConn, clsVoucherLst.strKey, rstFilter) Then
        Toolbar1.buttons("submit").Visible = True
        Toolbar1.buttons("unsubmit").Visible = True
        Toolbar1.buttons("viewverify").Visible = True
    Else
        Toolbar1.buttons("submit").Visible = False
        Toolbar1.buttons("unsubmit").Visible = False
        Toolbar1.buttons("viewverify").Visible = False
    End If
    Select Case LCase(clsVoucherLst.strKey)
        Case "17"
'            strVouchType = "97"
'            Toolbar1.buttons("closebatch").Visible = True
'            Toolbar1.buttons("openbatch").Visible = True
'            If clsSAWeb.bMQStart Or clsSAWeb.bMPStart Then
'                Toolbar1.buttons("lockbatch").Visible = True
'                Toolbar1.buttons("unlockbatch").Visible = True
'            End If
        Case "01", "03"
            strVouchtype = "05"
            Toolbar1.buttons("closebatch").Visible = True
            Toolbar1.buttons("openbatch").Visible = True
        Case "02", "04"
            strVouchtype = "07"

        Case "05", "06"
            strVouchtype = "06"
        Case "07"
            strVouchtype = "26"
        Case "08"
            strVouchtype = "98"
        Case "09"
            strVouchtype = "99"
        Case "13"
            strVouchtype = "27"
       Case "15"
            strVouchtype = "28"
        Case "14"
            strVouchtype = "29"
        Case "16"
            strVouchtype = "16"
            Toolbar1.buttons("closebatch").Visible = True
            Toolbar1.buttons("openbatch").Visible = True
        Case "sa18"
            strVouchtype = "sa18"
        Case "sa19"
            strVouchtype = "sa19"
        Case "sa26"
            strVouchtype = "sa26"
            Toolbar1.buttons("closebatch").Visible = True
            Toolbar1.buttons("openbatch").Visible = True
        'LDX    2009-07-03  Add Beg
        Case "mt011"
            strVouchtype = "98"
        Case "mt006"
            strVouchtype = "92"
        Case "mt007"
            strVouchtype = clsVoucherLst.strKey
        Case Else
            strVouchtype = clsVoucherLst.strKey
        'LDX    2009-07-03  Add End
    End Select
    UFToolbar1.RefreshVisible
End Function

Private Function setrstFilter(ii As Long) As String
Dim rstFilter As String
    Select Case LCase(clsVoucherLst.strKey)
        Case "17"
            rstFilter = "id=" & VchLst.TextMatrix(ii, VchLst.GridColIndex("id"))
        Case "01", "03"
            rstFilter = "DLID=" & VchLst.TextMatrix(ii, VchLst.GridColIndex("DLID"))
        Case "02", "04"
            rstFilter = "id=" & VchLst.TextMatrix(ii, VchLst.GridColIndex("ID"))
        Case "05", "06"
            rstFilter = "DLID=" & VchLst.TextMatrix(ii, VchLst.GridColIndex("DLID"))
        Case "07"
            rstFilter = "SBVID=" & VchLst.TextMatrix(ii, VchLst.GridColIndex("SBVID"))
        Case "08"
            rstFilter = "id=" & VchLst.TextMatrix(ii, VchLst.GridColIndex("id"))
        Case "09"
             rstFilter = "id=" & VchLst.TextMatrix(ii, VchLst.GridColIndex("id"))
        Case "13"
            rstFilter = "SBVID=" & VchLst.TextMatrix(ii, VchLst.GridColIndex("SBVID"))
        Case "15"
            rstFilter = "SBVID=" & VchLst.TextMatrix(ii, VchLst.GridColIndex("SBVID"))
        Case "14"
            rstFilter = "SBVID=" & VchLst.TextMatrix(ii, VchLst.GridColIndex("SBVID"))
        Case "16"
            rstFilter = "id=" & VchLst.TextMatrix(ii, VchLst.GridColIndex("id"))
        Case "sa18"
            rstFilter = "id=" & VchLst.TextMatrix(ii, VchLst.GridColIndex("id"))
        Case "sa19"
            rstFilter = "id=" & VchLst.TextMatrix(ii, VchLst.GridColIndex("id"))
        'LDX    2009-07-03  Add Beg
        Case "mt011", "mt006", "mt007"
            rstFilter = "id=" & VchLst.TextMatrix(ii, VchLst.GridColIndex("id"))
        'LDX    2009-07-03  Add End
    End Select
    setrstFilter = rstFilter
End Function
Private Function setMainkey(ii As Long) As String
Dim rstFilter As String
    Select Case LCase(clsVoucherLst.strKey)
        Case "17"
            rstFilter = VchLst.TextMatrix(ii, VchLst.GridColIndex("id"))

        Case "sa18"
            rstFilter = VchLst.TextMatrix(ii, VchLst.GridColIndex("id"))
        Case "sa19"
            rstFilter = VchLst.TextMatrix(ii, VchLst.GridColIndex("id"))
        Case "sa26"
            rstFilter = VchLst.TextMatrix(ii, VchLst.GridColIndex("id"))
        'LDX    2009-07-03  Add Beg
        Case "mt011", "mt006", "mt007"
            rstFilter = VchLst.TextMatrix(ii, VchLst.GridColIndex("id"))
        Case Else
            rstFilter = VchLst.TextMatrix(ii, VchLst.GridColIndex("id"))
        'LDX    2009-07-03  Add End
    End Select
    setMainkey = rstFilter
End Function

Private Function getVouchcode(ii As Long) As String
Dim rstFilter As String
    Select Case LCase(clsVoucherLst.strKey)
        Case "17"
            rstFilter = VchLst.TextMatrix(ii, VchLst.GridColIndex("csocode"))
        Case "01", "03"
            rstFilter = VchLst.TextMatrix(ii, VchLst.GridColIndex("cdlcode"))
        Case "02", "04"
            rstFilter = VchLst.TextMatrix(ii, VchLst.GridColIndex("ccode"))
        Case "05", "06"
            rstFilter = VchLst.TextMatrix(ii, VchLst.GridColIndex("cdlcode"))
        Case "07"
            rstFilter = VchLst.TextMatrix(ii, VchLst.GridColIndex("csbvcode"))
        Case "08"
            rstFilter = VchLst.TextMatrix(ii, VchLst.GridColIndex("cevcode"))
        Case "09"
            rstFilter = VchLst.TextMatrix(ii, VchLst.GridColIndex("cspvcode"))
        Case "13"
            rstFilter = VchLst.TextMatrix(ii, VchLst.GridColIndex("csbvcode"))
        Case "15"
            rstFilter = VchLst.TextMatrix(ii, VchLst.GridColIndex("csbvcode"))
        Case "14"
            rstFilter = VchLst.TextMatrix(ii, VchLst.GridColIndex("csbvcode"))
        Case "16"
            rstFilter = VchLst.TextMatrix(ii, VchLst.GridColIndex("ccode"))
        Case "sa18"
            rstFilter = VchLst.TextMatrix(ii, VchLst.GridColIndex("ccode"))
        Case "sa19"
            rstFilter = VchLst.TextMatrix(ii, VchLst.GridColIndex("ccode"))
        Case "sa26"
            rstFilter = VchLst.TextMatrix(ii, VchLst.GridColIndex("ccode"))
        'LDX    2009-07-03  Add Beg
        Case "mt011", "mt006", "mt007"
            rstFilter = VchLst.TextMatrix(ii, VchLst.GridColIndex("ccode"))
        'LDX    2009-07-03  Add End
    End Select
    getVouchcode = rstFilter
End Function
Public Property Get strColumnSetKey() As String
    strColumnSetKey = m_strColumnSetKey
End Property

Public Property Let strColumnSetKey(ByVal vNewValue As String)
    m_strColumnSetKey = vNewValue
End Property

Public Property Get formCaption() As String
    formCaption = m_Caption
End Property

Public Property Let formCaption(ByVal vNewValue As String)
    m_Caption = vNewValue
End Property
Public Property Get strAuthID() As String
    strAuthID = m_strAuthID
End Property

Public Property Let strAuthID(ByVal vNewValue As String)
    m_strAuthID = vNewValue
End Property

 
Private Sub ctlVoucher1_SaveSettingEvent(ByVal varDevice As Variant)
    Dim TmpUFTemplate As Object ' UFVoucherServer85.clsVoucherTemplate
'    Set TmpUFTemplate = New UFVoucherServer85.clsVoucherTemplate
    Set TmpUFTemplate = CreateObject("UFVoucherServer85.clsVoucherTemplate")
    If TmpUFTemplate.SaveDeviceCapabilities(DBConn.ConnectionString, 31180, varDevice) <> 0 Then
        MsgBox GetString("U8.SA.xsglsql.01.frmbillvouch.00361"), vbInformation 'zh-CN：打印设置保存失败
    End If
End Sub



Private Sub Form_Activate()
    VchLst_RowColChange
End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim strButtonKey As String
'    strButtonKey = clsTbl.GetKeyCodeByHotKey(KeyCode, Shift)
'    ButtonClick strButtonKey
'End Sub

Private Sub Form_Load()
    Dim strErrorRes As String
    SkinSE_Start hwnd
    If Not g_business Is Nothing Then
        Set Me.UFToolbar1.Business = g_business
    End If
    RegisterMessage

    sLogName = "VouchVerify" + Trim(getCurrentSessionID()) + ".log"
    sLogName1 = "Submit" + Trim(getCurrentSessionID()) + ".log"
    Me.VchLst.SumStyle = vlGridSum
    Me.HelpContextID = val(strHelpId)
    Set Me.Icon = frmMain.Icon
    
    Dim oDicTmp As Object
    Set oDicTmp = CreateObject("Scripting.Dictionary")
    Call Me.UFToolbar1.Settoolbarfromdata(Me.Toolbar1, DBConn, m_login, strColumnSetKey & "_list", strColumnSetKey & "_list", oDicTmp)
    Me.UFToolbar1.SetToolbar Me.Toolbar1
    Me.UFToolbar1.Height = 0
    Me.Toolbar1.Visible = False
    UFToolbar1.SetFormInfo VchLst, Me


'    clsVoucherLst.Init strColumnSetKey, strErrorRes
'    clsTbl.initToolBar Me.Toolbar1, strToolBarName, strErrorRes
'    ChangeOneFormTbr Me, Me.Toolbar1, Me.UFToolbar1, strColumnSetKey
'    Call Me.UFToolbar1.SetFormInfo(Me.VchLst, Me)
'    clsTbl.SetListButtonState clsVoucherLst.strKey, Me.Toolbar1, Me.UFToolbar1
    
    Set m_pagediv = New Pagediv
    m_bshowSumType = clsVoucherLst.bShowSumType
     '----单据列列表 打开时不显示过滤
    Dim ErrInfo As String
    If fltsrv Is Nothing Then
        Set fltsrv = New UFGeneralFilter.FilterSrv
        Dim tempflag As Boolean
        m_BehaviorObject.StrReportName = clsVoucherLst.strKey
        Set fltsrv.BehaviorObject = m_BehaviorObject
        tempflag = fltsrv.InitBaseVarValue(m_login, clsVoucherLst.strKey, , "SA", ErrInfo)
    End If
    Call VchLst.InitFlt(m_login, fltsrv, "", "", "", "")
    
    Dim rstSum As New ADODB.Recordset
    Set rstSum = clsVoucherLst.GetSumRecord("1=2")
    VchLst.SumStyle = vlRecordAndGridsum
    VchLst.SetSumRst rstSum
    InitConi "1=2"
    VchLst.HiddenRefreshView = False
    
'    GetDatas
    Call VchLst.BindPagediv(m_pagediv)
    Call m_pagediv.Initialize(DBConn, m_coni)
    m_pagediv.LoadData
    Me.VchLst.AdJustGridWidth
    Set rstSum = Nothing
    
    UFFrmCaptionMgr.Caption = formCaption ' Me.VchLst.Title
'    Dim strDWName As String
    Dim sXML As String
'    m_Login.GetAccInfo 105, strDWName
    sXML = "<表尾>"
    sXML = sXML + "<字段>" + GetString("U8.SA.xsglsql.frmvouchlist.00041") + clsSAWeb.scName + "</字段> " '"<字段>" + "单位：" + strDWName + "</字段> "
    sXML = sXML + "<字段>" + GetString("U8.SA.xsglsql.frmvouchlist.00044") + m_login.cUserName + "</字段> " '"<字段>" + "制表：" + m_lstLogin.cUserName + "</字段> "
    sXML = sXML + "<字段>" + GetString("U8.SA.xsglsql.frmvouchlist.00047") + Format(m_login.CurDate, "YYYY-MM-DD") + "</字段> " '"<字段>" + "打印日期：" + Format(m_lstLogin.CurDate, "YYYY-MM-DD") + "</字段> "
    sXML = sXML + "</表尾>"
    Me.VchLst.SetPrintOtherInfo sXML
    setVouchtype
    
    
    
    
    '    UserbSuc = True
    Userdll_UI.Userdll_Init g_business, m_login, DBConn, Me, strColumnSetKey, strUserErr, UserbSuc
    If UserbSuc = False Then
        MsgBox strUserErr
    End If
    
End Sub

Private Sub Form_Resize()
    Me.UFToolbar1.Align = 0
    Me.UFToolbar1.SetDisplayStyle PictureText
    Me.UFToolbar1.Left = 0
    Me.UFToolbar1.Top = 0
    Me.UFToolbar1.Width = Me.ScaleWidth
    Me.UFToolbar1.Height = Me.Toolbar1.Height
    
    If Me.Toolbar1.Visible Then
        Me.VchLst.Top = Me.Toolbar1.Height
    Else
        Me.VchLst.Top = Me.UFToolbar1.Height
    End If
    
    
    Me.VchLst.Left = 0
    Me.VchLst.Width = Me.ScaleWidth
    PagedivCtl1.Left = 0
    PagedivCtl1.Width = Me.ScaleWidth
    PagedivCtl1.Top = Me.ScaleHeight - PagedivCtl1.Height
    
    
    If (Me.UFToolbar1.Top + Me.UFToolbar1.Height) > 0 Then VchLst.Top = Me.UFToolbar1.Top + Me.UFToolbar1.Height
    If (Me.Height - Me.UFToolbar1.Height - PagedivCtl1.Height - 50) > 0 Then VchLst.Height = Me.Height - Me.UFToolbar1.Height - PagedivCtl1.Height - 50
    
'    If (Me.UFToolbar1.Top + Me.UFToolbar1.Height + VchLst.Height) > 0 Then PagedivCtl1.Top = Me.UFToolbar1.Top + Me.UFToolbar1.Height + VchLst.Height
    
'
'
'
'    PagedivCtl1.Top = Me.ScaleHeight - PagedivCtl1.Height
'
    

    Me.bar.Left = 0
    Me.bar.Width = Me.ScaleWidth
    Me.bar.Top = Me.ScaleHeight - bar.Height
    Me.bar.Visible = False
'    If Me.ScaleHeight >= VchLst.Top + PagedivCtl1.Height Then Me.VchLst.Height = Me.ScaleHeight - VchLst.Top - PagedivCtl1.Height - 1000
'    PagedivCtl1.Top = Me.VchLst.Top + Me.VchLst.Height
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Set m_oColSet = Nothing
    LockItem Me.strAuthID, False
    Set clsVoucherLst = Nothing
    Set clsVoucherCO = Nothing
    Set clsTbl = Nothing
'    SendMessgeToPortal "DocEditorClosed"
    UnRegisterMessage
End Sub

'初始化分页条件
Private Sub InitConi(strwhere As String)
    Dim i As Long
    
    If m_coni Is Nothing Then
        Set m_coni = New DefaultPagedivConi
    End If
    m_coni.From = clsVoucherLst.strFrom  '相当于From部分
    clsVoucherLst.m_oColSet.ShowSumType = m_bshowSumType
    clsVoucherLst.m_oColSet.setColMode clsVoucherLst.strColumnSetKey, 0
    If m_bshowSumType Then
        m_coni.SelectConi = Replace(clsVoucherLst.m_oColSet.GetSqlSumString, "''", "' '")
        If strwhere = "" Then strwhere = "1=1"
        m_coni.Where = strwhere
        m_coni.GroupBy = clsVoucherLst.m_oColSet.GetSqlGroupString
        m_coni.OrderID = clsVoucherLst.m_oColSet.GetOrderStringEx
        m_coni.RealTableName = clsVoucherLst.GetVoucherListSet("maintbl")
        m_coni.RealPrimaryKey = clsVoucherLst.GetVoucherListSet("mainkey")
    Else
        m_coni.SelectConi = clsVoucherLst.m_oColSet.GetSqlString
        m_coni.OrderID = clsVoucherLst.m_oColSet.GetOrderStringEx
        m_coni.Where = strwhere '相当于where部分
        m_coni.GroupBy = ""
        m_coni.RealTableName = clsVoucherLst.GetVoucherListSet("detailtbl")
        m_coni.RealPrimaryKey = clsVoucherLst.GetVoucherListSet("detailkey")
    End If
End Sub

Private Sub m_pagediv_AfterGetData(Rst As ADODB.Recordset, cnt As Long)
    Me.VchLst.InitHead clsVoucherLst.strColFormatXml
    clsVoucherLst.FormatVouchList Me.VchLst
    VchLst.DoFormat
End Sub

Private Sub m_pagediv_BeforeGetCount()
    VchLst.FillMode = FillOverwrite
End Sub

Private Sub m_pagediv_GetData(ByVal vltable As U8VouchList.VouchListTable)
    VchLst.ClearDataSource
    VchLst.SetVchLstRst vltable.DataRecordset
    VchLst.RecordCount = vltable.DataCount
End Sub

Private Sub PagedivCtl1_BeforeSendCommand(cmdType As U8VouchList.UFCommandType, pageSize As Long, pageCurrent As Long)
    Me.VchLst.SetVchLstRst Nothing
    Me.VchLst.FillMode = FillOverwrite
End Sub

Private Sub GetDatas(Optional fltsrv As Object)
    Dim strwhere As String
    
    If fltsrv Is Nothing Then Exit Sub
    strwhere = clsVoucherLst.strDefaultFilter
    If strwhere = "" Then
        strwhere = fltsrv.GetSQLWhere
    Else
        If fltsrv.GetSQLWhere <> "" Then
            strwhere = "(" + strwhere + ") and (" + fltsrv.GetSQLWhere + ")"
        End If
    End If
     
    Dim rstSum As New ADODB.Recordset
    Set rstSum = clsVoucherLst.GetSumRecord(strwhere)
    VchLst.SumStyle = vlRecordAndGridsum
    VchLst.SetSumRst rstSum
    InitConi strwhere
    
'    Call PagedivCtl1.BindPagediv(m_pagediv)
    Call VchLst.BindPagediv(m_pagediv)
    Call m_pagediv.Initialize(DBConn, m_coni)
    m_pagediv.LoadData
    Me.VchLst.AdJustGridWidth
    Set rstSum = Nothing
End Sub


Public Property Get strToolBarName() As String
    strToolBarName = m_strToolBarName
End Property

Public Property Let strToolBarName(ByVal vNewValue As String)
    m_strToolBarName = vNewValue
End Property

Public Property Get strHelpId() As String
    strHelpId = m_strHelpId
End Property

Public Property Let strHelpId(ByVal vNewValue As String)
    m_strHelpId = vNewValue
End Property

Private Sub ShowVerifyHistory(Optional blnShowWindows As Boolean = True)
    Dim m_CardNumber As String
    Dim m_Mid As String
    Dim m_TablName As String
    Dim introw As Long
    Call GetCardNumberMid(m_CardNumber, m_Mid, m_TablName, VchLst.row)
    If blnShowWindows Then
        ShowWorkFlowView strFormGuid, m_CardNumber, "UFIDA.U8.Audit.AuditHistoryView"
    End If
    SendPortalMessage strFormGuid, m_CardNumber, m_Mid, "DocQueryAuditHistory", VchLst.TextMatrix(VchLst.row, VchLst.GridColIndex("cmaker")), VchLst.TextMatrix(VchLst.row, VchLst.GridColIndex("ufts")), getVouchcode(VchLst.row), strVouchtype
End Sub

Public Sub ButtonClick(strButtonKey As String)
    Dim c_row As Long
    Dim c_col As Long
 
    '功能申请
    On Error GoTo ErrHandle
    
    

    
    
    
    If clsTbl.ButtonKeyDown(m_login, strButtonKey) Then
    
    
        UserbSuc = False
        Userdll_UI.Before_ButtonClick Me.VchLst, strButtonKey, strUserErr, UserbSuc
        If UserbSuc Then
            GetDatas fltsrv
            Exit Sub
        End If
    
        '业务处理
        Select Case LCase(strButtonKey)
            Case "creating"
                Import_arapvouchers "ap"
                GetDatas fltsrv
 
            Case "viewverify"
                ShowVerifyHistory
            Case "filter"
                If clsVoucherLst.ShowFilter() Then
                    m_bshowSumType = clsVoucherLst.bShowSumType
                    GetDatas fltsrv
                End If
            Case "filterset"
                clsVoucherLst.SetFilter
                GetDatas fltsrv
            Case "columnset"
                If clsVoucherLst.ColumnSet Then
                    GetDatas fltsrv
                End If
            Case "print", "preview", "output"
                clsVoucherLst.PrintVoucherList Me.VchLst, strButtonKey, Me.strColumnSetKey
            Case "exit"
                Unload Me
                Exit Sub
            Case "locate"
                If Not VchLst.LocateState Then
                    VchLst.locate True
                Else
                    VchLst.locate False
                End If
            Case "search"
              If LCase(clsVoucherLst.strKey) = "mt020" Or LCase(clsVoucherLst.strKey) = "mt021" Then
                clsVoucherLst.ShowVoucher Me.VchLst
              Else
                VchLst_DblClick
              End If
            Case "selectall"
                Me.VchLst.AllSelect
            Case "unselectall"
                Me.VchLst.AllNone
            Case "refresh"
                GetDatas fltsrv
            Case "help"
                'LDX    2009-05-31  Add Beg
'                SendKeys "{F1}"
                On Error Resume Next
                ShowContextHelp Me.hwnd, App.HelpFile, Me.HelpContextID
                'LDX    2009-05-31  Add Beg
            Case "新增", "修改", "删除", "作废"
                clsVoucherLst.CreateNewVoucher Me.VchLst, strButtonKey
            Case "verifybatch"
'                strOperStatus = "VouchVerifyBatch"
                DoAllVerCloseLock "VouchVerifyBatch"
            Case "unverifybatch"
'                strOperStatus = "VouchUnVerifyBatch"
                DoAllVerCloseLock "VouchUnVerifyBatch"
            Case "deletebatch"
'                strOperStatus = "VouchDeleteBatch"
                DoAllVerCloseLock "VouchDeleteBatch"
            Case "closebatch"
'                strOperStatus = "VouchCloseBatch"
                DoAllVerCloseLock "VouchCloseBatch"
            Case "openbatch"
'                strOperStatus = "VouchOpen"
                DoAllVerCloseLock "VouchOpen"
            Case "lockbatch"
'                strOperStatus = "VouchLockBatch"
                DoAllVerCloseLock "VouchLockBatch"
            Case "unlockbatch"
'                strOperStatus = "VouchUnLockBatch"
                DoAllVerCloseLock "VouchUnLockBatch"
            Case "prnbatch"
                Call clsVoucherLst.PrnBatch(Me.VchLst)
            Case "submit"
                Call submitBatch
            Case "unsubmit"
                Call backBatch
         
        End Select
        clsTbl.ButtonKeyUp m_login, strButtonKey
        
    End If
    Exit Sub
ErrHandle:
    MsgBox Err.Description
    clsTbl.ButtonKeyUp m_login, strButtonKey
    Screen.MousePointer = vbDefault
End Sub
  
 
'做审核弃审 复核弃复 关闭打开 锁定解锁
Private Sub DoAllVerCloseLock(strOperStatus As String)
    Dim introw As Integer
    Dim i As Long
    For i = 1 To Me.VchLst.Rows - 1
         If VchLst.TextMatrix(i, VchLst.GridColIndex("selcol")) = "Y" Then
            introw = 1
            ReDim MainCode(1 To 1)
            MainCode(introw) = -1
            
            Exit For
         End If
    Next
    If introw = 0 Then
       MsgBox GetString("U8.SA.xsglsql.frmrefpur.02491"), vbInformation 'zh-CN： 请选择单据!
       Exit Sub
    End If
    introw = 0
    Select Case strOperStatus
        Case "VouchVerifyBatch", "VouchUnVerifyBatch"
              If VouchCheckBatch(strOperStatus) Then 'Unload Me
                   If strVouchtype = "26" Or strVouchtype = "27" Or strVouchtype = "28" Or strVouchtype = "29" Then
                       MsgBox GetStringPara("U8.SA.xsglsql.frmrefpur.02492.01", IIf(strOperStatus = "VouchVerifyBatch", GetString("U8.SA.xsglsql.clsreportcallback.00097"), GetString("U8.SA.xsglsql.ModSale.00430"))), vbInformation 'Para zh-CN：批量{0}完毕。 'zh-CN：复核 'zh-CN：弃复
                   Else
                       MsgBox GetStringPara("U8.SA.xsglsql.frmrefpur.02492.01", IIf(strOperStatus = "VouchVerifyBatch", GetString("U8.SA.xsglsql.clsreportcallback.00100"), GetString("U8.SA.xsglsql.ModSale.00385"))), vbInformation  'Para zh-CN：批量{0}完毕。 'zh-CN：审核 'zh-CN：弃审
                   End If
              End If
        Case "VouchCloseBatch", "VouchOpen"
              If VouchCloseBatch(strOperStatus) Then 'Unload Me
                 MsgBox GetStringPara("U8.SA.xsglsql.frmrefpur.02492.01", IIf(strOperStatus = "VouchCloseBatch", GetString("U8.SA.xsglsql.clsreportcallback.00110"), GetString("U8.SA.xsglsql.clsreportcallback.00111"))), vbInformation  'Para zh-CN：批量{0}完毕。 'zh-CN：关闭 'zh-CN：打开
              End If
        Case "VouchLockBatch", "VouchUnLockBatch"
           If VouchCloseBatch(strOperStatus) Then 'Unload Me
                 MsgBox GetStringPara("U8.SA.xsglsql.frmrefpur.02492.01", IIf(strOperStatus = "VouchLockBatch", GetString("U8.SA.xsglsql.frmrefpur.02505"), GetString("U8.SA.xsglsql.frmrefpur.02506"))), vbInformation  'Para zh-CN：批量{0}完毕。 'zh-CN：锁定 'zh-CN：解锁
           End If
       Case "VouchDeleteBatch"
           If VouchDeleteBatch(strOperStatus) Then 'Unload Me
               
           End If
        Case Else
    End Select
    Call ButtonClick("refresh")
End Sub

''批量审核

Private Function VouchCheckBatch(strOperStatus As String) As Boolean
    Dim strsql As String
    Dim strAutoID As String
    Dim bVerifyVouch As Boolean
    Dim i As Long
    Dim domHead As New DOMDocument
    Dim domBody As New DOMDocument
    Dim domResult As New DOMDocument
    Dim strError As String
    Dim iCount As Long
    Dim iElement As IXMLDOMElement
    Dim bEAContinue As Boolean, strEAXML As String
    Dim oNodeEA As IXMLDOMNode, oEleEA As IXMLDOMElement
    Dim bContinue As Boolean, oDomEA As New DOMDocument, bEAContinueCheck As Boolean
    Dim strEAOCode As String, bEACancel As Boolean, vMsgRet As Variant
    Dim bWorkflow As Boolean
    Dim rstFilter As String
    
    On Error GoTo Err_Handle
    
    bEACancel = False
    bEAContinue = False
    bEAContinueCheck = False
    If strVouchtype = "26" Or strVouchtype = "27" Or strVouchtype = "28" Or strVouchtype = "29" Then
        strsql = IIf(strOperStatus = "VouchUnVerifyBatch", GetString("U8.SA.xsglsql.ModSale.00430"), GetString("U8.SA.xsglsql.clsreportcallback.00097")) 'zh-CN：弃复 'zh-CN：复核
    Else
        strsql = IIf(strOperStatus = "VouchUnVerifyBatch", GetString("U8.SA.xsglsql.ModSale.00385"), GetString("U8.SA.xsglsql.clsreportcallback.00100")) 'zh-CN：弃审 'zh-CN：审核
    End If
    
    If strOperStatus = "VouchUnVerifyBatch" Then
       bVerifyVouch = False
     Else
       bVerifyVouch = True
      End If
    On Error Resume Next
    If Dir(App.Path & "\" + sLogName) <> "" Then Kill App.Path & "\" + sLogName
    On Error GoTo 0
'    If strVouchType <> "16" And strVouchType <> "99" And UCase(strVouchType) <> "SA18" And UCase(strVouchType) <> "SA19" Then
'        If clsSAWeb.bCrCtrWShow = False And clsSAWeb.bCrCheckWhen = False And (clsSAWeb.bCredit Or clsSAWeb.bCrCheckDe Or clsSAWeb.bCrCheckPe) Then  '不需要信用审批,信用检查点   false在单据审核时审核
'            If ((strVouchType = "26" Or strVouchType = "27") And clsSAWeb.bCreditBillF) Or (strVouchType = "28" And clsSAWeb.bCreditBillD) Or (strVouchType = "29" And clsSAWeb.bCreditBillL) Or (strVouchType = "97" And clsSAWeb.bCreditOrder) Or (strVouchType = "05" And clsSAWeb.bCreditDisp) Or ((strVouchType = "06" Or strVouchType = "00") And clsSAWeb.bCreditDispWT) Or (strVouchType = "07" And clsSAWeb.bCreditDispJS) Or (strVouchType = "98" And clsSAWeb.bCreditExpense) Then
'                If MsgBox(GetString("U8.SA.xsglsql.frmmakebill.01975"), vbYesNo + vbQuestion, GetString("U8.SA.xsglsql.frmmakebill.01977")) = vbYes Then 'zh-CN：如果信用控制选项设置为不需要信用审批,当超过信用期限或者信用额度时，是否继续？ 'zh-CN：注意
'                      bContinue = True
'                End If
'            End If
'         End If
'    End If
    
    Dim m_CardNumber As String
    Dim m_Mid As String
    Dim errMsg As String
    Dim m_TablName As String
    Dim primBizData As String
    Dim Authid As String
    Dim AbbAuthid As String
    Dim j As Long
    Dim strBillType As String
    Dim auditResult As String
    Dim Action As String
    Dim State As Integer
    Dim strAuditOpinion As String
    Dim strSuccess As String
    Dim iSuccess As Integer
    Dim strErrors As String
    Dim iFailed As Integer
    
    Dim eleResult As IXMLDOMElement
    
    Dim AuditServiceProxy As Object
    
    Set AuditServiceProxy = CreateObject("UFIDA.U8.Audit.ServiceProxy.AuditServiceProxy")
    '调用批审之前，调用该对象的填写审批意见服务
    'LDX    2009-07-02  注释    Beg
    Dim calledCtx As New UFSoft_U8_Framework_LoginContext.CalledContext
    'LDX    2009-07-02  注释    End
    
    bar.Max = VchLst.Rows - 1
    bar.Visible = True
    strBillType = ""
    For i = 1 To VchLst.Rows - 1
            If bEACancel = True Then
                Exit For
            End If
    
            If VchLst.TextMatrix(i, VchLst.GridColIndex("selcol")) = "Y" Then
                If CheckCode(val(setMainkey(i))) Then
                    j = j + 1
                    ReDim Preserve MainCode(1 To j)
                    MainCode(j) = val(setMainkey(i))
                    If VchLst.TextMatrix(i, VchLst.GridColIndex("iswfcontrolled")) = "1" Then '启用审批流的时候 否则走 终审代码
                        If Not strOperStatus = "VouchUnVerifyBatch" Then
                            If VchLst.TextMatrix(i, VchLst.GridColIndex("iswfcontrolled")) = "1" And VchLst.TextMatrix(i, VchLst.GridColIndex("iVerifyState")) = "1" Then
                                primBizData = "         <KeySet>" & Chr(13)
                                primBizData = primBizData & "                   <Key name=""VoucherId"" value=""" & setMainkey(i) & """/>" & Chr(13)
                                primBizData = primBizData & "                   <Key name=""VoucherType"" value=""" & UCase(getCardNumber(strVouchtype)) & """/>" & Chr(13)
                                primBizData = primBizData & "                   <Key name=""VoucherCode"" value=""" & getVouchcode(i) & """/>" & Chr(13)
                                primBizData = primBizData & "                   <Key name=""Ufts"" value=""" & VchLst.TextMatrix(i, VchLst.GridColIndex("ufts")) & """/>" & Chr(13)
                                primBizData = primBizData & "                   <Key name=""AuditAuthId"" value=""""/>" & Chr(13)
                                primBizData = primBizData & "                   <Key name=""AbandonAuthId"" value=""""/>" & Chr(13)
                                primBizData = primBizData & "         </KeySet>" & Chr(13)
    
                                If Not bWorkflow Then
                                    calledCtx.subid = m_login.cSub_Id
                                    calledCtx.TaskID = m_login.TaskID
                                    calledCtx.token = m_login.userToken
                                    If strOperStatus = "VouchVerifyBatch" Then
                                        If AuditServiceProxy.ShowAuditSimpleUI(calledCtx, Action, State, strAuditOpinion) = False Then
                                            VouchCheckBatch = False
                                            Exit Function
                                        End If
                                    End If
                                    bWorkflow = True
                                End If
                                Call AuditServiceProxy.Audit(primBizData, Action, State, strAuditOpinion, calledCtx, auditResult)
                                domResult.loadXML auditResult
                                For Each eleResult In domResult.documentElement.selectNodes("//Result")
                                    If CBool(eleResult.getAttribute("AuditResult")) = True Then
                                        VchLst.TextMatrix(i, VchLst.GridColIndex("selcol")) = ""
                                    Else
                                        Wrtlog GetString("U8.SA.xsglsql.frmrefpur.02365") & getVouchcode(i) & eleResult.getAttribute("errMsg"), sLogName   '"单据" 'zh-CN：单据号：
    '                                    Wrtlog GetString("U8.SA.xsglsql.frmrefpur.02365") & getVouchcode(i) & eleResult.getAttribute("errMsg"), sLogName  '"单据" 'zh-CN：单据号：
                                    End If
                                Next
                            ''                                                    处理工作流 (没有提交)
                            ElseIf VchLst.TextMatrix(i, VchLst.GridColIndex("iswfcontrolled")) = "1" And VchLst.TextMatrix(i, VchLst.GridColIndex("iVerifyState")) = "0" Then
                                Wrtlog GetString("U8.SA.xsglsql.frmrefpur.02365") & getVouchcode(i) & GetStringPara("U8.pu.prjpu860.04715", getVouchcode(i)), sLogName   '单据{0}没有提交，请首先提交！
                            End If
                        Else
                            If VchLst.TextMatrix(i, VchLst.GridColIndex("iswfcontrolled")) = "1" And VchLst.TextMatrix(i, VchLst.GridColIndex("iVerifyState")) <> "0" Then
                                primBizData = "         <KeySet>" & Chr(13)
                                primBizData = primBizData & "                   <Key name=""VoucherId"" value=""" & setMainkey(i) & """/>" & Chr(13)
                                primBizData = primBizData & "                   <Key name=""VoucherType"" value=""" & UCase(getCardNumber(strVouchtype)) & """/>" & Chr(13)
                                primBizData = primBizData & "                   <Key name=""VoucherCode"" value=""" & getVouchcode(i) & """/>" & Chr(13)
                                primBizData = primBizData & "                   <Key name=""AuditAuthId"" value=""""/>" & Chr(13)
                                primBizData = primBizData & "                   <Key name=""Ufts"" value=""" & VchLst.TextMatrix(i, VchLst.GridColIndex("ufts")) & """/>" & Chr(13)
                                primBizData = primBizData & "                   <Key name=""AbandonAuthId"" value=""""/>" & Chr(13)
                                primBizData = primBizData & "         </KeySet>" & Chr(13)
                                If Not bWorkflow Then
                                    calledCtx.subid = m_login.cSub_Id
                                    calledCtx.TaskID = m_login.TaskID
                                    calledCtx.token = m_login.userToken
                                    If strOperStatus = "VouchUnVerifyBatch" Then
                                        If Not AuditServiceProxy.ShowAuditAbandonUI(calledCtx, State, strAuditOpinion) Then
                                            VouchCheckBatch = False
                                            Exit Function
                                        End If
                        '                primBizData = ""
                                    End If
                                    bWorkflow = True
                                End If
                                Call AuditServiceProxy.Abandon(primBizData, strAuditOpinion, State, calledCtx, auditResult)
                                domResult.loadXML auditResult
                                For Each eleResult In domResult.documentElement.selectNodes("//Result")
                                    If CBool(eleResult.getAttribute("AuditResult")) = True Then
                                        VchLst.TextMatrix(i, VchLst.GridColIndex("selcol")) = ""
                                    Else
                                        Wrtlog GetString("U8.SA.xsglsql.frmrefpur.02365") & getVouchcode(i) & eleResult.getAttribute("errMsg"), sLogName   '"单据" 'zh-CN：单据号：
    '                                    Wrtlog GetString("U8.SA.xsglsql.frmrefpur.02365") & getVouchcode(i) & eleResult.getAttribute("errMsg"), sLogName  '"单据" 'zh-CN：单据号：
                                    End If
                                Next
                            ElseIf VchLst.TextMatrix(i, VchLst.GridColIndex("iswfcontrolled")) = "1" And VchLst.TextMatrix(i, VchLst.GridColIndex("iVerifyState")) = "0" Then
                                Wrtlog GetString("U8.SA.xsglsql.frmrefpur.02365") & getVouchcode(i) & GetStringPara("U8.pu.prjpu860.04715", getVouchcode(i)), sLogName   '单据{0}没有提交，请首先提交！
                            End If
                        End If
                    Else
                        If IIf(strOperStatus = "VouchUnVerifyBatch", 1 = 1, VchLst.TextMatrix(i, VchLst.GridColIndex("cVerifier")) = "") Then
''                                StartProcessLog "initclsVoucherCO"
                                Select Case LCase(strVouchtype)
'                                Case "05"
'                                    clsVoucherCO.Init DispatchBlue, m_Login, DBConn, "CS", clsSAWeb
'                                Case "06"
'                                    clsVoucherCO.Init WTDispatchBlue, m_Login, DBConn, "CS", clsSAWeb
'                                Case "00"
'                                    clsVoucherCO.Init WTTZ, m_Login, DBConn, "CS", clsSAWeb
'                                Case "26"
'                                    clsVoucherCO.Init SaleBillSpecBlue, m_Login, DBConn, "CS", clsSAWeb
'                                Case "27"
'                                    clsVoucherCO.Init SaleBillCommBlue, m_Login, DBConn, "CS", clsSAWeb
'                                Case "28"
'                                    clsVoucherCO.Init SAMoveBlue, m_Login, DBConn, "CS", clsSAWeb
'                                Case "29"
'                                    clsVoucherCO.Init SARetailBlue, m_Login, DBConn, "CS", clsSAWeb
'                                Case "07"
'                                    If strOperStatus = "VouchVerifyBatch" And strBillType = "" Then
'                                        frmSelectFP.Show vbModal
'                                        If frmSelectFP.bCancel = False Then
'                                            strBillType = frmSelectFP.strVouchType
'                                        Else
'                                            Exit Function
'                                        End If
'                                    End If
'                                    clsVoucherCO.Init EntrustBlue, m_Login, DBConn, "CS", clsSAWeb
'                                Case "97"
'                                    clsVoucherCO.Init SODetails, m_Login, DBConn, "CS", clsSAWeb
'                                Case "98"
'                                    clsVoucherCO.Init ExpenseVouch, m_Login, DBConn, "CS", clsSAWeb
'                                Case "99"
'                                    clsVoucherCO.Init SalePayVouch, m_Login, DBConn, "CS", clsSAWeb
'                                Case "16"
'                                    clsVoucherCO.Init SAQuo, m_Login, DBConn, "CS", clsSAWeb
'                                Case "sa18", "sa19", "sa26"
'                                    Dim clsVoucher As New clsSaVoucher
'                                    clsVoucher.Init LCase(strVouchType), strError
'                                    clsVoucherCO.InitSys m_Login, , clsSAWeb
                                Case "98"
                                    clsVoucherCO.Init "MT66", m_login, DBConn, "CS", clsSAWeb
                                Case "92"
                                    clsVoucherCO.Init "MT06", m_login, DBConn, "CS", clsSAWeb
                                Case "93"
                                    clsVoucherCO.Init "MT07", m_login, DBConn, "CS", clsSAWeb
                                Case Else
                                    clsVoucherCO.Init strVouchtype, m_login, DBConn, "CS", clsSAWeb
                                End Select
'                            EndProcessLog "initclsVoucherCO"
                            strError = ""
'                            StartProcessLog "GetVoucherData"
                            If LCase(strVouchtype) = "sa18" Or LCase(strVouchtype) = "sa19" Or LCase(strVouchtype) = "sa26" Then
'                                strError = clsVoucher.LoadVoucherHead(CStr(MainCode(j)), Domhead)
'                                Set clsVoucher = Nothing
                            Else
                                strError = clsVoucherCO.GetVoucherData(domHead, domBody, MainCode(j))
                            End If
'                            EndProcessLog "GetVoucherData"
    
                            If strError = "" Then
                                Set iElement = domHead.selectSingleNode("//z:row") '
                                If bContinue Then
                                    iElement.setAttribute "ccrechppass", "ufsoft"
    
                                End If
                                If strBillType <> "" Then
                                    iElement.setAttribute "billvouchtype", strBillType
                                End If
                                iElement.setAttribute "cvouchtype", strVouchtype
                                strError = ""
                                If strError = "" Then
'                                    StartProcessLog "VerifyVouch"
                                    strError = clsVoucherCO.VerifyVouch(domHead, bVerifyVouch)
                                    If strError <> "" Then
                                        Dim lngPos As Long
                                        Dim strXml As String
                                        lngPos = InStr(1, strError, "<", vbTextCompare)
                                        If lngPos > 0 Then
                                            Dim ele As IXMLDOMElement
                                            Dim domTmp As New DOMDocument
                                            strXml = Mid(strError, lngPos)
                                            domTmp.loadXML strXml
                                            Set ele = domTmp.selectSingleNode("//rs:data/zeroout")
                                        ''填充
                                            If ele.Attributes.getNamedItem("okonly") Is Nothing Then
                                                '如果不控制可用量继续
                                                iElement.setAttribute "saveafterok", "1"
                                                strError = clsVoucherCO.VerifyVouch(domHead, bVerifyVouch)
                                            Else
                                                strError = Mid(strError, 1, InStr(1, strError, "<", vbTextCompare) - 1)
                                                Wrtlog GetString("U8.SA.xsglsql.frmrefpur.02365") & getVouchcode(i) & strError, sLogName  '"单据" 'zh-CN：单据号
                                            End If
                                            Set domTmp = Nothing
                                        End If
                                    End If
'                                    EndProcessLog "VerifyVouch"
                                End If
                                If strError = "" Then
                                    VchLst.TextMatrix(i, VchLst.GridColIndex("selcol")) = ""
                                    iCount = iCount + 1
                                Else
                                    If InStr(1, strError, "<", vbTextCompare) <> 0 Then
                                            strError = Mid(strError, 1, InStr(1, strError, "<", vbTextCompare) - 1)
                                            If InStr(1, strError, "详细错误如下：", vbTextCompare) <> 0 Then
                                                strError = Mid(strError, 1, InStr(1, strError, "错误出现在", vbTextCompare) - 1) & "可用量不够"  '，也可能信用检查不过
                                            End If
                                    End If
                                    If strVouchtype = "16" Then
                                        Wrtlog GetString("U8.SA.xsglsql.frmrefpur.02365") & strError, sLogName  '"单据" 'zh-CN：单据号：
                                    Else
                                        Wrtlog GetString("U8.SA.xsglsql.frmrefpur.02365") & getVouchcode(i) & strError, sLogName  '"单据" 'zh-CN：单据号：
                                    End If
                                End If
                            Else
                                If strVouchtype = "16" Then
                                    Wrtlog GetString("U8.SA.xsglsql.frmrefpur.02365") & strError, sLogName
                                Else
                                    Wrtlog GetString("U8.SA.xsglsql.frmrefpur.02365") & getVouchcode(i) & strError, sLogName
                                End If
                            End If
                        Else
                            Wrtlog GetStringPara("U8.SA.xsglsql.frmrefpur.02365.02", getVouchcode(i), IIf(strOperStatus = "VouchUnVerifyBatch", GetString("U8.SA.xsglsql.ModSale.00385"), GetString("U8.SA.xsglsql.clsreportcallback.00100"))), sLogName 'Para zh-CN：单据号{0} 已作{1}操作。 'zh-CN：弃审 'zh-CN：审核
                        End If
                    End If
                End If
            End If
        bar.value = i
    Next
    bar.Visible = False
    Set domHead = Nothing
    
    VouchCheckBatch = True
'    If Dir(App.Path & "\" + sLogName) <> "" Then
'        frmMsg.ShowMsg (App.Path & "\" + sLogName)
'        VouchCheckBatch = False
'    End If
    Exit Function
Err_Handle:
    MsgBox Err.Description, vbExclamation
    VouchCheckBatch = False
End Function


''批量删除

Private Function VouchDeleteBatch(strOperStatus As String) As Boolean

    Dim strsql As String
    Dim strAutoID As String
    Dim bVerifyVouch As Boolean
    Dim i As Long
    Dim domHead As New DOMDocument
    Dim domBody As New DOMDocument
    Dim domResult As New DOMDocument
    Dim strError As String
    Dim iCount As Long
    Dim iElement As IXMLDOMElement
    Dim bEAContinue As Boolean, strEAXML As String
    Dim oNodeEA As IXMLDOMNode, oEleEA As IXMLDOMElement
    Dim bContinue As Boolean, oDomEA As New DOMDocument, bEAContinueCheck As Boolean
    Dim strEAOCode As String, bEACancel As Boolean, vMsgRet As Variant
    Dim bWorkflow As Boolean
    Dim rstFilter As String
    
    On Error GoTo Err_Handle
    
    bEACancel = False
    bEAContinue = False
    bEAContinueCheck = False
    
    On Error Resume Next
    On Error GoTo 0
    
    Dim m_CardNumber As String
    Dim m_Mid As String
    Dim errMsg As String
    Dim m_TablName As String
    Dim primBizData As String
    Dim Authid As String
    Dim AbbAuthid As String
    Dim j As Long
    Dim strBillType As String
    Dim auditResult As String
    Dim Action As String
    Dim State As Integer
    Dim strAuditOpinion As String
    Dim strSuccess As String
    Dim iSuccess As Integer
    Dim strErrors As String
    Dim iFailed As Integer
    
    Dim eleResult As IXMLDOMElement
    
    Dim AuditServiceProxy As Object
    
    Set AuditServiceProxy = CreateObject("UFIDA.U8.Audit.ServiceProxy.AuditServiceProxy")
    '调用批审之前，调用该对象的填写审批意见服务
    Dim calledCtx As New UFSoft_U8_Framework_LoginContext.CalledContext
    
    bar.Max = VchLst.Rows - 1
    bar.Visible = True
    strBillType = ""
    For i = 1 To VchLst.Rows - 1
            If bEACancel = True Then
                Exit For
            End If
            If VchLst.TextMatrix(i, VchLst.GridColIndex("selcol")) = "Y" Then
                If CheckCode(val(setMainkey(i))) Then
                    j = j + 1
                    ReDim Preserve MainCode(1 To j)
                    MainCode(j) = val(setMainkey(i))
                    
                    Select Case LCase(strVouchtype)
                    Case "98"
                        clsVoucherCO.Init "MT66", m_login, DBConn, "CS", clsSAWeb
                    Case "92"
                        clsVoucherCO.Init "MT06", m_login, DBConn, "CS", clsSAWeb
                    Case "93"
                        clsVoucherCO.Init "MT07", m_login, DBConn, "CS", clsSAWeb
                    End Select
                    
                    strError = ""
                    If LCase(strVouchtype) = "98" Or LCase(strVouchtype) = "92" Or LCase(strVouchtype) = "93" Then
                        strError = clsVoucherCO.GetVoucherData(domHead, domBody, MainCode(j))
                    End If
                    If strError = "" Then
                        'LDX    2009-07-17  Add Beg
'                        If LCase(strVouchType) = "92" Then
'                           If MsgBox("是否删除月度预算编制单?", vbYesNo) = vbYes Then
'
'                           Else
'                                Exit Function
'                           End If
'                        End If
'                        If LCase(strVouchType) = "93" Then
'                           If MsgBox("是否删除预算调整单?", vbYesNo) = vbYes Then
'
'                           Else
'                                Exit Function
'                           End If
'                        End If
                        If LCase(strVouchtype) = "98" Then
                           If MsgBox("是否删除年度计划?", vbYesNo) = vbYes Then
                                Dim Rst As New ADODB.Recordset
                                strsql = "select * from mt_budget where cvouchtype='02' and left(iperiod,2)<>13 and year(ddate)=year(getdate())"
                                If Rst.State <> 0 Then Rst.Close
                                Rst.Open strsql, DBConn, adOpenStatic, adLockReadOnly
                                If Rst.RecordCount > 0 Then
                                    MsgBox "月度预算编制已存在，不允许删除!"
                                    Exit Function
                                End If
                           Else
                                Exit Function
                           End If
                        End If
                        strError = clsVoucherCO.Delete(domHead)
                        'LDX    2009-07-17  Add End
                        If strError = "" Then
                            VchLst.TextMatrix(i, VchLst.GridColIndex("selcol")) = ""
                            iCount = iCount + 1
                        Else
                            MsgBox "删除错误!"
                            Exit Function
                        End If
                    Else
                        MsgBox "获取数据错误!"
                        Exit Function
                    End If
            End If
        End If
        bar.value = i
    Next
    bar.Visible = False
    Set domHead = Nothing
    
    VouchDeleteBatch = True
    Exit Function
Err_Handle:
    MsgBox Err.Description, vbExclamation
    VouchDeleteBatch = False
End Function

'取得单据类型对应的模版号
Private Function getCardNumber(sVouchType As String) As String
    Select Case sVouchType
    Case "26", "27", "28", "29" '发票
        getCardNumber = "07"
    Case "05", "06", "00" '发货
        getCardNumber = "01"
    Case "97" '订单
        getCardNumber = "17"
    Case "16" '报价单
        getCardNumber = "16"
    Case "98" '代垫
        getCardNumber = "08"
    Case "99" '费用支出
        getCardNumber = "09"
    Case "07" '结算
        getCardNumber = "02"
    Case "00"
        getCardNumber = "28"
    Case Else
        getCardNumber = sVouchType
'    Case "95", "92" '包装物
'        cardNum = "10": Mid = "autoid"
    End Select
End Function



Private Function GetKeyFieldName(strOperStatus As String) As String
    Select Case LCase(clsVoucherLst.strKey)
        Case "sa26"
            If clsVoucherLst.bShowSumType Then
                GetKeyFieldName = "id"
            Else
                If strOperStatus = "VouchClose" Or strOperStatus = "VouchCloseBatch" Or strOperStatus = "VouchOpen" Then
                    GetKeyFieldName = "autoid"
                Else
                    GetKeyFieldName = "id"
                End If
            End If
        Case "17"
            If clsVoucherLst.bShowSumType Then
                GetKeyFieldName = "id"
            Else
                If strOperStatus = "VouchClose" Or strOperStatus = "VouchCloseBatch" Or strOperStatus = "VouchOpen" Then
                    GetKeyFieldName = "isosid"
                Else
                    GetKeyFieldName = "id"
                End If
            End If
        Case "01", "03"
            If clsVoucherLst.bShowSumType Then
                GetKeyFieldName = "DLID"
            Else
                If strOperStatus = "VouchClose" Or strOperStatus = "VouchCloseBatch" Or strOperStatus = "VouchOpen" Then
                    GetKeyFieldName = "idlsid"
                Else
                    GetKeyFieldName = "DLID"
                End If
            End If
        Case "02", "04"
            GetKeyFieldName = "id"
        Case "05", "06"
            If clsVoucherLst.bShowSumType Then
                GetKeyFieldName = "DLID"
            Else
                If strOperStatus = "VouchClose" Or strOperStatus = "VouchCloseBatch" Or strOperStatus = "VouchOpen" Then
                    GetKeyFieldName = "idlsid"
                Else
                    GetKeyFieldName = "DLID"
                End If
            End If
        Case "07"
            GetKeyFieldName = "sbvid"
        Case "08"
            GetKeyFieldName = "id"
        Case "09"
             GetKeyFieldName = "id"
        Case "13"
            GetKeyFieldName = "sbvid"
        Case "15"
            GetKeyFieldName = "sbvid"
        Case "14"
            GetKeyFieldName = "sbvid"
        Case "16"
            GetKeyFieldName = "id"
        Case "sa18"
            GetKeyFieldName = "id"
        Case "sa19"
            GetKeyFieldName = "id"
    End Select
End Function
Private Function VouchCloseBatch(strOperStatus As String, Optional bLock As Boolean = False) As Boolean
    'Dim strsql As String
    'Dim strKey As String
    'Dim strKeyValue As String
    'Dim i As Long
    'Dim iCount As Long
    'Dim strTblName As String
    'Dim rst As New ADODB.Recordset
    'Dim strError As String
    'Dim domHead As New DOMDocument
    ''        Dim Dombody As New DOMDocument
    'Dim errDom As New DOMDocument
    'If strVouchType = "97" Then
    '    strTblName = "SO_SOMain"
    'Else
    '    If strVouchType = "16" Then
    '        strTblName = "SA_QuoMain"
    '    Else
    '        strTblName = "DispatchList"
    '    End If
    'End If
    'On Error Resume Next
    'If Dir(App.Path & "\" + sLogName) <> "" Then Kill App.Path & "\" + sLogName
    'On Error GoTo Err_Handle
    '
    'VouchCloseBatch = True
    'Select Case LCase(strVouchType)
    '    Case "05"
    '        clsVoucherCO.Init DispatchBlue, m_Login, DBConn, "CS", clsSAWeb
    ''                                  clsVoucherCO.InitSys m_Login, DispatchBlue, clsSAWeb
    '    Case "97"
    '        clsVoucherCO.Init SODetails, m_Login, DBConn, "CS", clsSAWeb
    ''                                  clsVoucherCO.InitSys m_Login, SODetails, clsSAWeb
    '    Case "16"
    '        clsVoucherCO.Init SAQuo, m_Login, DBConn, "CS", clsSAWeb
    '    Case "sa26"
    '        clsVoucherCO.InitSys m_Login, , clsSAWeb
    'End Select
    'strKey = GetKeyFieldName(strOperStatus)
    'Dim j As Long
    'iCount = 0
    '
    'bar.Max = VchLst.Rows - 1
    'bar.Visible = True
    'For i = 1 To VchLst.Rows - 1
    '        If VchLst.TextMatrix(i, VchLst.GridColIndex("selcol")) = "Y" Then
    '        strKeyValue = VchLst.TextMatrix(i, VchLst.GridColIndex(strKey))
    '        If CheckCode(val(strKeyValue)) Then
    '            j = j + 1
    '            ReDim Preserve MainCode(1 To j)
    '            MainCode(j) = val(strKeyValue)
    '            strError = ""
    '            If (strOperStatus = "VouchClose" Or strOperStatus = "VouchCloseBatch" Or strOperStatus = "VouchOpen") And (strVouchType = "97" Or strVouchType = "05" Or strVouchType = "sa26") Then
    '                If strVouchType = "05" Or strVouchType = "97" Or strVouchType = "16" Or strVouchType = "sa26" Then
    '                    If strVouchType = "05" Then
    '                        strsql = " SELECT  DispatchList.bfirst,DispatchList.dlid as dlid,cDLCode, cvouchtype, cSTCode,  dDate, cRdCode, cDepCode, cPersonCode, SBVID,  cSOCode, DispatchList.cCusCode AS cCusCode  , cPayCode, cSCCode, cShipAddress, " & _
    '                                 " cexch_name, iExchRate, iTaxRate, 0 AS bFirst, bReturnFlag, bSettleAll, cMemo, cSaleOut, cDefine1, cDefine2, cDefine3, cDefine4, cDefine5, cDefine6, cDefine7, cDefine8," & _
    '                                 " cDefine9, cDefine10, ISNULL(cVerifier,'') AS cVerifier , cMaker, iSale,  cCusName , " & _
    '                                 " iVTid, convert(char,convert(money,dispatchlist.ufts),2) as ufts, cBusType, cAccounter, cCreChpName, cDefine11, cDefine12, cDefine13, cDefine14, cDefine15, " & _
    '                                 " cDefine16 " & _
    '                                 " From DispatchList WHERE dlid=" & VchLst.TextMatrix(i, VchLst.GridColIndex("dlid"))      'setrstFilter(i)
    '                    ElseIf strVouchType = "97" Then
    '                        strsql = " SELECT  SO_SOMain.ID as ID,cSOCode,'97' AS cvouchtype, cSTCode,  dDate,  cDepCode, cPersonCode,  SO_SOMain.cCusCode AS cCusCode  , cPayCode, cSCCode,  " & _
    '                                 " cexch_name, iExchRate, iTaxRate,  cMemo,  cDefine1, cDefine2, cDefine3, cDefine4, cDefine5, cDefine6, cDefine7, cDefine8," & _
    '                                 " cDefine9, cDefine10, ISNULL(cVerifier,'') AS cVerifier , cMaker,   cCusName , " & _
    '                                 " iVTid, convert(char,convert(money,SO_SOMain.ufts),2) as ufts, cBusType,  cCreChpName, cDefine11, cDefine12, cDefine13, cDefine14, cDefine15, " & _
    '                                 " cDefine16 " & _
    '                                 " From SO_SOMain WHERE  id=" & VchLst.TextMatrix(i, VchLst.GridColIndex("id"))                '& setrstFilter(i)
    '                    ElseIf strVouchType = "16" Then
    '                        strsql = " SELECT  ID,cCode,'16' AS cvouchtype, cSTCode,  dDate,  cDepCode, cPersonCode,  cCusCode AS cCusCode  , cPayCode, cSCCode,  " & _
    '                                 " cexch_name, iExchRate, iTaxRate,  cMemo,  cDefine1, cDefine2, cDefine3, cDefine4, cDefine5, cDefine6, cDefine7, cDefine8," & _
    '                                 " cDefine9, cDefine10,  cMaker,    " & _
    '                                 " iVTid, convert(char,convert(money,ufts),2) as ufts, cBusType,   cDefine11, cDefine12, cDefine13, cDefine14, cDefine15, " & _
    '                                 " cDefine16 " & _
    '                                 " From sa_quomain WHERE  id=" & VchLst.TextMatrix(i, VchLst.GridColIndex("id"))                '& setrstFilter(i)
    '                    Else
    '                        strsql = "select id,cCode,'SA26' AS cvouchtype, cSTCode,  dDate,  cDepCode, cPersonCode,  cCusCode AS cCusCode,cMemo,  cDefine1, cDefine2, cDefine3, cDefine4, cDefine5, cDefine6, cDefine7, cDefine8,  cDefine9, cDefine10, ISNULL(cVerifier,'') AS cVerifier , cMaker,   cCusName ,iVTid, convert(char,convert(money,ufts),2) as ufts, cDefine11, cDefine12, cDefine13, cDefine14, cDefine15,cDefine16 "
    '                        strsql = strsql & " from sa_preordermain where id=" & VchLst.TextMatrix(i, VchLst.GridColIndex("id"))
    '                    End If
    '                    If rst.State <> adStateClosed Then
    '                        rst.Close
    '                    End If
    '                    rst.Open ConvertSQLString(LCase(strsql)), DBConn, adOpenForwardOnly, adLockReadOnly
    '                    If rst.EOF Then GoTo DoNext  '单据已经不存在了
    '                    Set domHead = New DOMDocument
    '                    rst.Save domHead, adPersistXML
    '
    '                End If
    '            Else
    '                strError = clsVoucherCO.GetVoucherDataHead(domHead, MainCode(j))
    '            End If
    '            If strError <> "" Then
    '                Wrtlog GetString("U8.SA.xsglsql.frmrefpur.02365") & getVouchcode(i) & strError, sLogName  '"单据" 'zh-CN：单据号
    '                GoTo DoNext
    '            End If
    '            If strOperStatus = "VouchOpen" Then
    '               If strVouchType = "16" Then    'strVouchType = "97" Or
    '                    strError = clsVoucherCO.OrderClose(domHead, False)
    '                    If strError <> "" Then
    '                        VouchCloseBatch = False
    '                        Wrtlog GetString("U8.SA.xsglsql.frmrefpur.02365") & strError, sLogName   '"单据" 'zh-CN：单据号
    '                    Else
    '                        iCount = iCount + 1
    '                    End If
    ''                            strsql = "UPDATE " & strTblName & " SET  cCloser=NULL WHERE " & strKey & "=" & strKey
    ''                            DBConn.Execute strsql
    '                    iCount = iCount + 1
    '               Else
    '                    If clsVoucherLst.bShowSumType Then
    '                        strError = clsVoucherCO.OrderClose(domHead, False)
    '                    Else
    '                        strError = clsVoucherCO.OrderClose(domHead, False, val(strKeyValue))
    '                    End If
    '                    If strError <> "" Then
    '                        Dim lngPos As Long
    '                        Dim strXml As String
    '                        lngPos = InStr(1, strError, "<", vbTextCompare)
    '                        If lngPos > 0 Then
    '                            Dim ele As IXMLDOMElement
    '                            Dim domTmp As New DOMDocument
    '                            Dim iElement As IXMLDOMElement
    '                            strXml = Mid(strError, lngPos)
    '                            domTmp.loadXML strXml
    '                            Set ele = domTmp.selectSingleNode("//rs:data/zeroout")
    '                        ''填充
    '                            If ele.Attributes.getNamedItem("okonly") Is Nothing Then
    '                                '如果不控制可用量继续
    '                                Set iElement = domHead.selectSingleNode("//z:row") '
    '                                iElement.setAttribute "saveafterok", "1"
    '                                If clsVoucherLst.bShowSumType Then
    '                                    strError = clsVoucherCO.OrderClose(domHead, False)
    '                                Else
    '                                    strError = clsVoucherCO.OrderClose(domHead, False, val(strKeyValue))
    '                                End If
    '                                If strError <> "" Then
    '                                    If InStr(1, strError, "<", vbTextCompare) <> 0 Then
    '                                            strError = Mid(strError, 1, InStr(1, strError, "<", vbTextCompare) - 1)
    '                                            If InStr(1, strError, "详细错误如下：", vbTextCompare) <> 0 Then
    '                                                strError = Mid(strError, 1, InStr(1, strError, "错误出现在", vbTextCompare) - 1) & "可用量不够"  '，也可能信用检查不过
    '                                            End If
    '                                    End If
    '                                Else
    '                                    iCount = iCount + 1
    '                                End If
    '                            Else
    '                                strError = Mid(strError, 1, InStr(1, strError, "<", vbTextCompare) - 1)
    '                                Wrtlog GetString("U8.SA.xsglsql.frmrefpur.02365") & getVouchcode(i) & strError, sLogName  '"单据" 'zh-CN：单据号
    '                            End If
    '                            Set domTmp = Nothing
    '                        Else
    '                            Wrtlog GetString("U8.SA.xsglsql.frmrefpur.02365") & getVouchcode(i) & strError, sLogName  '"单据" 'zh-CN：单据号
    '                            VouchCloseBatch = False
    '                        End If
    '                    Else
    '                        iCount = iCount + 1
    '                    End If
    ''                            strSQL = "UPDATE DispatchList SET bSettleAll=0,cCloser=NULL WHERE DLID=" & VchLst.TextMatrix(i, 2)
    ''                            DBConn.Execute strSQL
    ''                            strSQL = "UPDATE DispatchLists SET bSettleAll=0 WHERE DLID=" & VchLst.TextMatrix(i, 2)
    ''                            DBConn.Execute strSQL
    '
    '               End If
    '            ElseIf strOperStatus = "VouchClose" Or strOperStatus = "VouchCloseBatch" Then
    ''                       If strVouchType = "16" Then   'strVouchType = "97" Or
    ''                            strError = clsVoucherCO.OrderClose(domHead, True)
    '''                            strsql = "UPDATE " & strTblName & " SET cCloser=N'" & m_Login.cUserName & "' WHERE " & strKey & "=" & strKey
    '''                            DBConn.Execute strsql
    ''                            iCount = iCount + 1
    ''                       Else
    '                    If clsVoucherLst.bShowSumType Or strVouchType = "16" Then
    '                        strError = clsVoucherCO.OrderClose(domHead, True)
    '                    Else
    '                        strError = clsVoucherCO.OrderClose(domHead, True, val(strKeyValue))
    '                    End If
    '                    If strError <> "" Then
    '                        VouchCloseBatch = False
    '                        If strVouchType = "16" Then
    '                            Wrtlog GetString("U8.SA.xsglsql.frmrefpur.02365") & strError, sLogName   '"单据" 'zh-CN：单据号：
    '                        Else
    ''                                    Dim lngPos As Long
    ''                                    Dim strXml As String
    '                            lngPos = InStr(1, strError, "<", vbTextCompare)
    '                            If lngPos > 0 Then
    ''                                        Dim ele As IXMLDOMElement
    ''                                        Dim domTmp As New DOMDocument
    ''                                        Dim iElement As IXMLDOMElement
    '                                strXml = Mid(strError, lngPos)
    '                                domTmp.loadXML strXml
    '                                Set ele = domTmp.selectSingleNode("//rs:data/zeroout")
    '                            ''填充
    '                                If ele.Attributes.getNamedItem("okonly") Is Nothing Then
    '                                    '如果不控制可用量继续
    '                                    Set iElement = domHead.selectSingleNode("//z:row") '
    '                                    iElement.setAttribute "saveafterok", "1"
    '                                    If clsVoucherLst.bShowSumType Or strVouchType = "16" Then
    '                                        strError = clsVoucherCO.OrderClose(domHead, True)
    '                                    Else
    '                                        strError = clsVoucherCO.OrderClose(domHead, True, val(strKeyValue))
    '                                    End If
    '                                    If strError <> "" Then
    '                                        If InStr(1, strError, "<", vbTextCompare) <> 0 Then
    '                                                strError = Mid(strError, 1, InStr(1, strError, "<", vbTextCompare) - 1)
    '                                                If InStr(1, strError, "详细错误如下：", vbTextCompare) <> 0 Then
    '                                                    strError = Mid(strError, 1, InStr(1, strError, "错误出现在", vbTextCompare) - 1) & "可用量不够"  '，也可能信用检查不过
    '                                                End If
    '                                        End If
    '                                    Else
    '                                        iCount = iCount + 1
    '                                    End If
    '                                Else
    '                                    strError = Mid(strError, 1, InStr(1, strError, "<", vbTextCompare) - 1)
    '                                    Wrtlog GetString("U8.SA.xsglsql.frmrefpur.02365") & getVouchcode(i) & strError, sLogName  '"单据" 'zh-CN：单据号
    '                                End If
    '                                Set domTmp = Nothing
    '                            Else
    '                                Wrtlog GetString("U8.SA.xsglsql.frmrefpur.02365") & getVouchcode(i) & MidErrorStr(strError), sLogName  '"单据" 'zh-CN：单据号：
    '                            End If
    '                        End If
    '                    Else
    '                        iCount = iCount + 1
    '                    End If
    ''                            strSQL = "UPDATE DispatchList SET bSettleAll=1,cCloser='" & m_login.cUserName & "' WHERE DLID=" & VchLst.TextMatrix(i, 2)
    ''                            DBConn.Execute strSQL
    ''                            strSQL = "UPDATE DispatchLists SET bSettleAll=1 WHERE DLID=" & VchLst.TextMatrix(i, 2)
    ''                            DBConn.Execute strSQL
    '
    ''                       End If
    '            ElseIf strOperStatus = "VouchLockBatch" Then
    '                strError = clsVoucherCO.LockVouch(domHead, True)
    '                If strError <> "" Then
    '                    Wrtlog GetString("U8.SA.xsglsql.frmrefpur.02365") & getVouchcode(i) & strError, sLogName 'zh-CN：单据号：
    '                End If
    '            ElseIf strOperStatus = "VouchUnLockBatch" Then
    '                strError = clsVoucherCO.LockVouch(domHead, False)
    '                If strError <> "" Then
    '                    Wrtlog GetString("U8.SA.xsglsql.frmrefpur.02365") & getVouchcode(i) & strError, sLogName 'zh-CN：单据号：
    '                End If
    '            End If
    '            VchLst.TextMatrix(i, VchLst.GridColIndex("selcol")) = ""
    '        End If
    '
    '    End If
    '    bar.value = i
    '
    '
    'DoNext:
    '
    'Next
    'bar.Visible = False
    '
    'If Dir(App.Path & "\" + sLogName) <> "" Then
    '    VouchCloseBatch = False
    '    frmMsg.ShowMsg (App.Path & "\" + sLogName)
    'End If
    'Set domHead = Nothing
    ''        Set Dombody = Nothing
    '
    'Exit Function
'Err_Handle:
'    MsgBox Err.Description, vbExclamation
'    VouchCloseBatch = False

End Function
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    ButtonClick Button.Key
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    ButtonClick ButtonMenu.Key
End Sub

Private Sub UFKeyHookCtrl1_ContainerKeyDown(KeyCode As Integer, ByVal Shift As Integer)
    Dim strButtonKey As String
    Dim strButtonStyle As String
    
    Dim ele As IXMLDOMElement
    Set ele = clsTbl.GetKeyCodeByHotKey(KeyCode, Shift)
    If ele Is Nothing Then Exit Sub
    strButtonKey = ele.Attributes.getNamedItem("buttonkey").nodeValue
    If Not ele.Attributes.getNamedItem("buttonstyle") Is Nothing Then
        strButtonStyle = ele.Attributes.getNamedItem("buttonstyle").nodeValue
    End If
    If strButtonKey <> "" Then
        If strButtonStyle <> "0" And strButtonStyle <> "5" Then
            If Me.Toolbar1.buttons(strButtonStyle).ButtonMenus(strButtonKey).Enabled Then
                ButtonClick strButtonKey
            End If
        Else
            If strButtonStyle = "5" Then
                If Me.Toolbar1.buttons(strButtonKey).Enabled Then
                    Call g_business.SelectToolbarButton(Me.Toolbar1.buttons(strButtonKey))
                End If
            Else
                If Me.Toolbar1.buttons(strButtonKey).Enabled Then
                    ButtonClick strButtonKey
                End If
            End If
        End If
    End If
End Sub

Private Sub UFToolbar1_OnCommand(ByVal enumType As UFToolBarCtrl.ENUM_MENU_OR_BUTTON, ByVal cButtonId As String, ByVal cmenuid As String)
    ButtonClick IIf(enumType = enumButton, cButtonId, cmenuid)
End Sub

'LDX    2009-07-03  Add Beg
Private Sub VchLst_CellValueChanged(ByVal row As Integer, ByVal col As Integer, NewValue As Variant, OldValue As Variant, KeepFocus As Boolean)
    Dim Skey As String
    Dim strsql As String
    Dim oRs As ADODB.Recordset
    Dim i As Integer
    
    If LCase(clsVoucherLst.strKey) = "mt020" Or LCase(clsVoucherLst.strKey) = "mt021" Then
        '合计行
        If row = VchLst.Rows Then
            NewValue = OldValue
            Exit Sub
        End If
        If VchLst.TextMatrix(row, VchLst.GetColIndex("selcol")) = "" Then
            NewValue = OldValue
            MsgBox "请先选择该条数据！"
            Exit Sub
        End If
        Skey = LCase(VchLst.GetColName(col))
        Select Case Skey
            Case "hdje"  '核定金额
                If NewValue <> "" Then
                    If Not IsNumeric(NewValue) Then
                        NewValue = OldValue
                        MsgBox "录入的数据必须是数值型！"
                        Exit Sub
                    ElseIf CDbl(NewValue) < 0 Then
                        NewValue = OldValue
                        MsgBox "核定金额必须大于等于零！"
                        Exit Sub
                    End If
                End If
        End Select
    End If
End Sub
'LDX    2009-07-03  Add End

Private Sub VchLst_DblClick()
    Userdll_UI.VchLst_DblClick Me.VchLst, strUserErr, UserbSuc
    If UserbSuc Then
        Exit Sub
    End If
    Select Case LCase(clsVoucherLst.strKey)
        Case "mt020", "mt021"
        Case Else
            clsVoucherLst.ShowVoucher Me.VchLst
    End Select
End Sub

Private Sub VchLst_FilterClick(fldsrv As Object)
    GetDatas fldsrv
End Sub

Private Sub VchLst_PrintSettingChanged(ByVal varLocalSettings As Variant, ByVal varModuleSettings As Variant)
    clsVoucherLst.SavePrnSet varLocalSettings + varModuleSettings, Me.strColumnSetKey
End Sub
Public Property Get strFormCaption() As String
    strFormCaption = m_strFormCaption
End Property

Public Property Let strFormCaption(ByVal vNewValue As String)
    m_strFormCaption = vNewValue
End Property
Private Function CheckID(id As Long) As Boolean
    Dim i As Long
    CheckID = True
        For i = 1 To UBound(DLID)
            If id = DLID(i) Then
                CheckID = False
                Exit For
            End If
        Next i
End Function
Private Function CheckCode(id As Long) As Boolean
    Dim i As Long
    CheckCode = True
        For i = 1 To UBound(MainCode)
            If id = MainCode(i) Then
                CheckCode = False
                Exit For
            End If
        Next i
End Function

Public Sub ExitForm(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub
'870 added for 工作流批量提交
Private Sub submitBatch()
    Dim strErrorResId As String '审批流的错误信息870 added
    Dim m_CardNumber As String
    Dim m_Mid As String
    Dim m_TablName As String
    Dim introw As Integer
    Dim i As Long
    Dim j As Long
    Dim strBillType As String
'    Dim buttonId As String
On Error Resume Next
'申请权限
If Dir(App.Path & "\" + sLogName1) <> "" Then Kill App.Path & "\" + sLogName1
On Error GoTo Err_Handle:
    For i = 1 To Me.VchLst.Rows - 1
         If VchLst.TextMatrix(i, VchLst.GridColIndex("selcol")) = "Y" Then
            introw = 1
            ReDim MainCode(1 To 1)
            MainCode(introw) = -1
            Exit For
         End If
    Next
    If introw = 0 Then
       MsgBox GetString("U8.SA.xsglsql.frmrefpur.02491"), vbInformation 'zh-CN： 请选择单据!
       Exit Sub
    End If
    'If clsVoucherLst.strKey = "02" Then
    '   frmSelectFP.Show vbModal
    '   If frmSelectFP.bcancel = False Then
    '       strBillType = frmSelectFP.strVouchType
    '   Else
    '       Exit Sub
    '   End If
    'End If
'    Select Case LCase(clsVoucherLst.strKey)
'        Case "17", "16", "08", "09", "sa18", "sa19"
'        If Not VoucherTask(GetString("U8.SA.xsglsql_2.button.submit"), buttonId, strVouchType, True) Then
'            Wrtlog GetString("U8.SA.USSASERVER.clssystem.00656"), sLogName1
'            GoTo showErr
'        End If
'    End Select
    
    For i = 1 To VchLst.Rows - 1
        If VchLst.TextMatrix(i, VchLst.GridColIndex("selcol")) = "Y" Then
            If CheckCode(val(setMainkey(i))) Then
                j = j + 1
                ReDim Preserve MainCode(1 To j)
                MainCode(j) = val(setMainkey(i))
                If VchLst.TextMatrix(i, VchLst.GridColIndex("iswfcontrolled")) = "1" And (VchLst.TextMatrix(i, VchLst.GridColIndex("iverifystate")) = "0" _
                Or (VchLst.TextMatrix(i, VchLst.GridColIndex("iverifystate")) = "1" And val(VchLst.TextMatrix(i, VchLst.GridColIndex("ireturncount"))) > 0)) Then
                    If VchLst.TextMatrix(i, VchLst.GridColIndex("ccloser")) = "" Then
                        Call GetCardNumberMid(m_CardNumber, m_Mid, m_TablName, i)
                        DoUndoSubmit True, m_CardNumber, m_Mid, m_TablName, VchLst.TextMatrix(i, VchLst.GridColIndex("ufts")), CBool(VchLst.TextMatrix(i, VchLst.GridColIndex("iswfcontrolled"))), strErrorResId
                        If strErrorResId <> "" Then
                            Wrtlog GetString("U8.SA.xsglsql.frmrefpur.02365") & getVouchcode(i) & strErrorResId, sLogName1  '"单据" 'zh-CN：单据号：
                            strErrorResId = ""
                        Else
                            Wrtlog GetString("U8.SA.xsglsql.frmrefpur.02365") & getVouchcode(i) & GetString("U8.SA.xsglsql_2.saworkflowsrv.011"), sLogName1 ' & "提交成功！"
                        End If
                    Else
                        Wrtlog GetStringPara("U8.SA.USSASERVER.clsotherdj.01403", getVouchcode(i)), sLogName1
                    End If
                Else
                    Wrtlog GetStringPara("U8.SA.xsglsql_2.saworkflowsrv.008", getVouchcode(i)), sLogName1 ' "单据" & getVouchcode(I) & "已经提交或未启用审批流！"
                End If
            End If
        End If
    Next
'    VoucherFreeTask buttonId '释放任务
showErr:

'If Dir(App.Path & "\" + sLogName1) <> "" Then
'    frmMsg.ShowMsg (App.Path & "\" + sLogName1)
'End If
Call ButtonClick("REFRESH")
Exit Sub
Err_Handle:
    MsgBox Err.Description, vbExclamation
End Sub


'870 added for 工作流批量撤销
Private Sub backBatch()
    Dim strErrorResId As String '审批流的错误信息870 added
    Dim m_CardNumber As String
    Dim m_Mid As String
    Dim m_TablName As String
    Dim introw As Integer
    Dim i As Long
    Dim j As Long
'    Dim buttonId As String
On Error Resume Next
If Dir(App.Path & "\" + sLogName1) <> "" Then Kill App.Path & "\" + sLogName1
On Error GoTo Err_Handle:
    For i = 1 To Me.VchLst.Rows - 1
         If VchLst.TextMatrix(i, VchLst.GridColIndex("selcol")) = "Y" Then
            introw = 1
            ReDim MainCode(1 To 1)
            MainCode(introw) = -1
            Exit For
         End If
    Next
    If introw = 0 Then
       MsgBox GetString("U8.SA.xsglsql.frmrefpur.02491"), vbInformation 'zh-CN： 请选择单据!
       Exit Sub
    End If
    
'    Select Case LCase(clsVoucherLst.strKey)
'    Case "17", "16", "08", "09", "sa18", "sa19"
'        If Not VoucherTask(GetString("U8.SA.xsglsql_2.button.unsubmit"), buttonId, strVouchType, True) Then
'            Wrtlog GetString("U8.SA.USSASERVER.clssystem.00656"), sLogName1
'            GoTo showErr
'        End If
'    End Select
    
    For i = 1 To VchLst.Rows - 1
        If VchLst.TextMatrix(i, VchLst.GridColIndex("selcol")) = "Y" Then
            If CheckCode(val(setMainkey(i))) Then
                j = j + 1
                ReDim Preserve MainCode(1 To j)
                MainCode(j) = val(setMainkey(i))
                If VchLst.TextMatrix(i, VchLst.GridColIndex("iswfcontrolled")) = "1" And VchLst.TextMatrix(i, VchLst.GridColIndex("iverifystate")) <> "0" Then
'                    Select Case LCase(clsVoucherLst.strKey)
'                    Case "01", "03"
'                        If VoucherTask(GetString("U8.SA.xsglsql_2.button.unsubmit"), buttonId, "05", GetReturnFlagBool(VchLst.TextMatrix(I, VchLst.GridColIndex("breturnflag")))) = False Then
'                            Wrtlog GetString("U8.SA.USSASERVER.clssystem.00656"), sLogName1
'                            GoTo showErr
'                        End If
'                    Case "05", "06"
'                        If VoucherTask(GetString("U8.SA.xsglsql_2.button.unsubmit"), buttonId, "06", GetReturnFlagBool(VchLst.TextMatrix(I, VchLst.GridColIndex("breturnflag")))) = False Then
'                            Wrtlog GetString("U8.SA.USSASERVER.clssystem.00656"), sLogName1
'                            GoTo showErr
'                        End If
'                    Case "02", "04"
'                        If VoucherTask(GetString("U8.SA.xsglsql_2.button.unsubmit"), buttonId, "07", GetReturnFlagBool(VchLst.TextMatrix(I, VchLst.GridColIndex("breturnflag")))) = False Then
'                            Wrtlog GetString("U8.SA.USSASERVER.clssystem.00656"), sLogName1
'                            GoTo showErr
'                        End If
'                    Case "07", "13", "14", "15"
'                        If VoucherTask(GetString("U8.SA.xsglsql_2.button.unsubmit"), buttonId, VchLst.TextMatrix(I, VchLst.GridColIndex("cvouchtype")), GetReturnFlagBool(VchLst.TextMatrix(I, VchLst.GridColIndex("breturnflag")))) = False Then
'                            Wrtlog GetString("U8.SA.USSASERVER.clssystem.00656"), sLogName1
'                            GoTo showErr
'                        End If
'                    Case Else
'                        If VoucherTask(GetString("U8.SA.xsglsql_2.button.unsubmit"), buttonId, strVouchType, False) = False Then
'                            Wrtlog GetString("U8.SA.USSASERVER.clssystem.00656"), sLogName1
'                            GoTo showErr
'                        End If
'                    End Select
                    
                    Call GetCardNumberMid(m_CardNumber, m_Mid, m_TablName, i)
                    DoUndoSubmit False, m_CardNumber, m_Mid, m_TablName, VchLst.TextMatrix(i, VchLst.GridColIndex("ufts")), CBool(VchLst.TextMatrix(i, VchLst.GridColIndex("iswfcontrolled"))), strErrorResId, getVouchcode(i)
                    If strErrorResId <> "" Then
                        Wrtlog GetString("U8.SA.xsglsql.frmrefpur.02365") & getVouchcode(i) & strErrorResId, sLogName1  '"单据" 'zh-CN：单据号：
                        strErrorResId = ""
                    Else
                        Wrtlog GetString("U8.SA.xsglsql.frmrefpur.02365") & getVouchcode(i) & GetString("U8.SA.xsglsql_2.saworkflowsrv.012"), sLogName1 '"撤销成功！"
                    End If
                Else
                    Wrtlog GetStringPara("U8.SA.xsglsql_2.saworkflowsrv.009", getVouchcode(i)), sLogName1  '"单据" & getVouchcode(I) & "已经撤销或未启用审批流！"
                End If
            End If
        End If
    Next
    
'    VoucherFreeTask buttonId '释放任务
    
showErr:

'If Dir(App.Path & "\" + sLogName1) <> "" Then
'    frmMsg.ShowMsg (App.Path & "\" + sLogName1)
'End If
Call ButtonClick("REFRESH")
Exit Sub
Err_Handle:
    MsgBox Err.Description, vbExclamation
End Sub
Private Function MidErrorStr(strError As String) As String
    Dim lngPos As Long
    lngPos = InStr(1, strError, "<rs:data")
    If lngPos > 0 Then
        lngPos = InStr(1, strError, "，")
        If lngPos > 0 Then
            MidErrorStr = Left(strError, lngPos - 1)
        End If
        lngPos = InStr(1, strError, ",")
        If lngPos > 0 Then
            MidErrorStr = Left(strError, lngPos - 1)
        End If
    Else
        MidErrorStr = strError
    End If
End Function

Public Property Get strFormGuid() As String
    strFormGuid = m_strFormGuid
End Property

Public Property Let strFormGuid(ByVal vNewValue As String)
    m_strFormGuid = vNewValue
End Property

Private Sub VchLst_RowColChange()
    On Error Resume Next
    If Toolbar1.buttons("verifybatch").Visible Then
        If Err.Number > 0 Then
            Err.Clear
            Exit Sub
        End If
        ShowVerifyHistory False
    End If
End Sub
'注销消息函数
Private Sub UnRegisterMessage()
    If m_mht Is Nothing Then Exit Sub
    If Not g_business Is Nothing Then
        Call g_business.UnregisterMessageHandler(m_mht)
    End If
    Set m_mht = Nothing
End Sub
 
Private Sub RegisterMessage()
    Set m_mht = New UFPortalProxyMessage.IMessageHandler
    m_mht.MessageType = "DocAuditComplete"
    If Not g_business Is Nothing Then
        Call g_business.RegisterMessageHandler(m_mht)
    End If
End Sub

'导入应付单
Public Sub Import_arapvouchers(cFlag As String)
'    Dim clsAPCO  As UFAPBO.clsAPVouch
''    CreateObject("ScmPublicSrv.clsAutoFill")
'    Dim oAcc As UFAPBO.clsAccount_AP
    Dim clsAPCO  As Object
    Dim oAcc As Object

    Dim strsql As String:   Dim rs As Object
    Dim domHeadAp As DOMDocument: Dim domBodyAp As DOMDocument
    Dim domHeadSrc As DOMDocument: Dim domBodySrc As DOMDocument
    Dim eleSrcH As IXMLDOMElement: Dim eleSrcB As IXMLDOMElement
    Dim eleAp As IXMLDOMElement
    Dim bTran As Boolean: bTran = False
    Dim iprg As Long
    Dim indexTrue As Long: Dim indexFalse As Long
    Dim busr As Boolean
    Dim errStr  As String
        
    On Error GoTo errExit
    
    Set domHeadSrc = VchLst.GetListDom(True)
    If domHeadSrc.selectNodes("//z:row").length = 0 Then
        MsgBox "请选择需要生单的数据源": Exit Sub
    End If
    
    iprg = 0
    Picture3.Visible = True
    prgMsg.value = iprg:    Label3.Caption = "正在准备生单,请等待.........": DoEvents
    
    Set oAcc = CreateObject("UFAPBO.clsAccount_AP") ' New UFAPBO.clsAccount_AP
    oAcc.Init m_login, cFlag
    Set clsAPCO = CreateObject("UFAPBO.clsAPVouch") ' New UFAPBO.clsAPVouch
    clsAPCO.Init m_login, DBConn, oAcc
        
    strsql = "select * from AP_ApVouch where 1=2"
    Set rs = DBConn.Execute(strsql): Set domHeadAp = New DOMDocument: rs.Save domHeadAp, 1: rs.Close: Set rs = Nothing
    strsql = "select * from AP_ApVouchs where 1=2"
    Set rs = DBConn.Execute(strsql): Set domBodyAp = New DOMDocument: rs.Save domBodyAp, 1: rs.Close: Set rs = Nothing
    
    prgMsg.Max = domHeadSrc.selectNodes("//z:row").length:     Label3.Caption = "正在生单,请等待.........": DoEvents
    
    For Each eleSrcH In domHeadSrc.selectNodes("//z:row")
        iprg = iprg + 1
        prgMsg.value = iprg:    Label3.Caption = "正在处理[" & eleSrcH.getAttribute("ccode") & "] ": DoEvents
        
        strsql = "select * from EFFYGL_V_SettleVouchs where id=" & val(eleSrcH.getAttribute("id"))
        Set rs = DBConn.Execute(strsql): Set domBodySrc = New DOMDocument: rs.Save domBodySrc, 1: rs.Close: Set rs = Nothing
        '构造dom
        Set eleAp = domHeadAp.selectSingleNode("//rs:data").parentNode
        eleAp.removeChild domHeadAp.selectSingleNode("//rs:data"): eleAp.appendChild domHeadAp.createElement("rs:data")
        Set eleAp = domBodyAp.selectSingleNode("//rs:data").parentNode
        eleAp.removeChild domBodyAp.selectSingleNode("//rs:data"): eleAp.appendChild domBodyAp.createElement("rs:data")
        
        Set eleAp = domHeadAp.selectSingleNode("//rs:data")
        Set eleAp = eleAp.appendChild(domHeadAp.createElement("z:row"))
        eleAp.setAttribute "cVouchType", "P0"
        eleAp.setAttribute "dVouchDate", m_login.CurDate
        
        eleAp.setAttribute "bd_c", "0"                              '借贷方向
        eleAp.setAttribute "cFlag", "AP"                              '应收应付标志
        
        eleAp.setAttribute "cDwCode", eleSrcH.getAttribute("cvencode") & ""    '供应商编码
        eleAp.setAttribute "cDeptCode", eleSrcH.getAttribute("cdepcode") & ""   '  部门编码
        eleAp.setAttribute "cPerson", eleSrcH.getAttribute("cpersoncode") & ""   '  人员编码
        eleAp.setAttribute "cDigest", eleSrcH.getAttribute("cdigest") & ""      '  摘要
        eleAp.setAttribute "cexch_name", eleSrcH.getAttribute("cexch_name") & ""      '  币种
        eleAp.setAttribute "iExchRate", eleSrcH.getAttribute("iexchrate") & ""      '  汇率
        
        eleAp.setAttribute "iAmount_f", eleSrcH.getAttribute("je") & ""      '  金额
        eleAp.setAttribute "iRAmount_f", eleSrcH.getAttribute("je") & ""      '  余额
        eleAp.setAttribute "iAmount", val(eleSrcH.getAttribute("iexchrate") & "") * val(eleSrcH.getAttribute("je") & "")      '  金额
        eleAp.setAttribute "iRAmount", val(eleSrcH.getAttribute("iexchrate") & "") * val(eleSrcH.getAttribute("je") & "")    '  本币余额

        eleAp.setAttribute "cOperator", eleSrcH.getAttribute("cmaker") & ""      '  录入人
        eleAp.setAttribute "dcreatesystime", m_login.CurDate                 '  制单时间
        
        For Each eleSrcB In domBodySrc.selectNodes("//z:row")
            Set eleAp = domBodyAp.selectSingleNode("//rs:data")
            Set eleAp = eleAp.appendChild(domBodyAp.createElement("z:row"))
                        
            eleAp.setAttribute "bd_c", "1"                      '借贷方向
            eleAp.setAttribute "cexch_name", eleSrcH.getAttribute("cexch_name") & ""      '  币种
            eleAp.setAttribute "iExchRate", eleSrcH.getAttribute("iexchrate") & ""      '  汇率
            eleAp.setAttribute "cDwCode", eleSrcB.getAttribute("cvencode") & ""    '供应商编码
            eleAp.setAttribute "iAmount_f", eleSrcB.getAttribute("imoney") & ""      '  金额
            eleAp.setAttribute "iAmount", val(eleSrcH.getAttribute("iexchrate") & "") * val(eleSrcB.getAttribute("imoney") & "")    '  本币金额
        Next
        
        Set eleAp = domHeadAp.selectSingleNode("//z:row")
        If bPeriod(DBConn, eleAp.getAttribute("dVouchDate"), "AP") = False Then   '如果年度结转以后导上年的票，此处改为true;导本年的票改为false
            eleAp.setAttribute "dVouchDate", clsSAWeb.getBeginDate(CurrentAccMonth(DBConn, "AP"))
        End If
        busr = False
        DBConn.BeginTrans: bTran = True
        
        '生成XML文件，临时使用
'        domHeadAp.Save App.Path & "\domHeadAp.xml"
'        domBodyAp.Save App.Path & "\domBodyAp.xml"
        busr = clsAPCO.VouchCheck(domHeadAp, domBodyAp, errStr, True)
        If busr = False Then GoTo oneErr
        
        busr = clsAPCO.SaveVouch(domHeadAp, domBodyAp, errStr, True) '保存应收单据
        If busr Then
            strsql = "update EFFYGL_SettleVouch set bbuild=1,coutid='" & domHeadAp.selectSingleNode("//z:row").Attributes.getNamedItem("cVouchID").Text & "' where id=" & val(eleSrcH.getAttribute("id"))
            DBConn.Execute strsql
            indexTrue = indexTrue + 1
            If bTran Then DBConn.CommitTrans: bTran = False
        Else
oneErr:
            indexFalse = indexFalse + 1
            If bTran Then DBConn.RollbackTrans: bTran = False
            If MsgBox("处理[" & eleSrcH.getAttribute("ccode") & "]时发生如下错误：" & vbCrLf & errStr & vbCrLf & "是否继续？", vbYesNo + vbQuestion) = vbNo Then
                GoTo trueExit
            End If
        End If
    Next
trueExit:
    If Picture3.Visible = True Then Picture3.Visible = False
    MsgBox "共处理数据[" & iprg & "]条，成功[" & indexTrue & "]条，[" & indexFalse & "]条", vbInformation
    Exit Sub
errExit:
    If bTran Then DBConn.RollbackTrans: bTran = False
    If Picture3.Visible = True Then Picture3.Visible = False
    If Err.Number <> 0 Then errStr = Err.Description
    MsgBox "处理过程发生如下错误：" & errStr
    
    Exit Sub
    
    'Dim strsql As String
    Dim temp_rs As ADODB.Recordset
    Dim temp_domHead As ADODB.Recordset
    Dim temp_DomBody As ADODB.Recordset
    Dim objsys As Object
    Dim R As Long
    Dim i As Long
    Dim Addtrue As Boolean
    
    Dim k As Integer
On Error GoTo errStr
    R = 2
    k = 0
    Set temp_rs = New ADODB.Recordset
    Set temp_domHead = New DOMDocument
    Set temp_DomBody = New DOMDocument
'    Set oAcc = New UFAPBO.clsAccount_AP
    Set oAcc = CreateObject("UFAPBO.clsAccount_AP")
    oAcc.Init m_login, cFlag
'    Set clsVoucherCO = New UFAPBO.clsAPVouch
    Set clsVoucherCO = CreateObject("UFAPBO.clsAPVouch")
    
    clsVoucherCO.Init m_login, DBConn, oAcc
    objsys.Init m_login
    
    Picture3.Visible = True
    prgMsg.value = i
    Label3.Caption = "正在生单,请等待........."
    DoEvents
    Dim UfGrid1 As Object
    Dim BFlag As Boolean
    Set UfGrid1 = Me.VchLst
    For i = 1 To Me.VchLst.Rows - 1
        If UCase(Trim(UfGrid1.TextMatrix(i, 0))) = "√" Then
        If BFlag Then
            strsql = " update tmparapvouch set sel=1 where    [类型]='" & cFlag & "'  and BFlag = 1 and [票号]='" & Trim(UfGrid1.TextMatrix(i, 2)) & "' and [部门编码]='" & Trim(UfGrid1.TextMatrix(i, 10)) & "' and [制单人]='" & m_login.cUserName & "'"
        Else
            strsql = " update tmparapvouch set sel=1 where    [类型]='" & cFlag & "'  and BFlag = 0 and [票号]='" & Trim(UfGrid1.TextMatrix(i, 2)) & "' and [部门编码]='" & Trim(UfGrid1.TextMatrix(i, 10)) & "' and [制单人]='" & m_login.cUserName & "'"
        End If
        DBConn.Execute strsql, R
        End If
    Next
    
    If BFlag Then
        strsql = "select    * from tmparapvouch where sel=1 and  [类型]='" & cFlag & "'  and BFlag = 1 and [制单人]='" & m_login.cUserName & "'"
    Else
        strsql = "select    * from tmparapvouch where sel=1 and  [类型]='" & cFlag & "'  and BFlag = 0 and [制单人]='" & m_login.cUserName & "'"
    End If
    If temp_rs.State <> 0 Then temp_rs.Close
    temp_rs.CursorLocation = adUseClient
    temp_rs.Open strsql, DBConn, adOpenStatic, adLockReadOnly
    If Not temp_rs.EOF Then
        prgMsg.Max = temp_rs.RecordCount
        i = 1
    End If
    
    Debug.Print temp_rs.Fields("票号") & Time
    
    While Not temp_rs.EOF
        Debug.Print temp_rs.Fields("票号") & Time
        Label3.Caption = "正在处理[" & temp_rs.Fields("票号") & "] 机票,请等待........."
        DoEvents
        Debug.Print temp_rs.Fields("票号") & "GetDom" & Time
        Call GetDom(cFlag, temp_rs.Fields("票号"), temp_rs.Fields("部门编码"), temp_domHead, temp_DomBody)
        temp_domHead.Save "d:\u8soft\domh_tmp.xml"
        temp_DomBody.Save "d:\u8soft\domb_tmp.xml"
        prgMsg.value = i
'        temp_domHead.Load "c:\domh1.xml"
'        temp_DomBody.Load "c:\domb1.xml"
        errStr = ""
        Addtrue = True
        R = 0
'        Debug.Print temp_rs.Fields("票号") & "BeginTrans" & Time
'        DBconn.BeginTrans
        '将当前操作员信息的写入domh中
        SetHeadItemValue temp_domHead, "cOperator", m_login.cUserName
        
        '20071217 单楠提出的将导入日期写入表头自定义2上
        '6、在形成应收和应付单的时候（包括红字）在表头自定义项上加一个导入日期。
        SetHeadItemValue temp_domHead, "cDefine2", m_login.CurDate
        '判断当前业务单据日期所在月份是否结帐False表示已结帐，若已结帐就将单据日期改成下月1号
        If bPeriod(DBConn, Replace(GetHeadItemValue(temp_domHead, "dVouchDate"), "T", " "), cFlag) = False Then    '如果年度结转以后导上年的票，此处改为true;导本年的票改为false
            SetHeadItemValue temp_domHead, "dVouchDate", clsSAWeb.getBeginDate(CurrentAccMonth(DBConn, cFlag))
        End If
        Debug.Print temp_rs.Fields("票号") & "VouchCheck" & Time
        'busr = clsVoucherCO.VouchCheck(temp_domHead, temp_DomBody, Errstr, Addtrue) '保存前的系统校验
        busr = True  'sl 修改 2008/02/28 将单据校验取消
        Debug.Print temp_rs.Fields("票号") & "VouchCheck" & Time
        
        
        
        
        If busr Then
            Debug.Print temp_rs.Fields("票号") & "BeginTrans" & Time
            DBConn.BeginTrans
            k = 1
            Debug.Print temp_rs.Fields("票号") & "SaveVouch" & Time
            busr = clsAPCO.SaveVouch(temp_domHead, temp_DomBody, errStr, Addtrue) '保存应收单据
            Debug.Print temp_rs.Fields("票号") & "SaveVouch" & Time
            If busr Then '成功时清除原始应收EXECL 数据
                If BFlag Then
                    strsql = " delete tmparapvouch where  sel=1 and [类型]='" & cFlag & "' and BFlag = 1 and  [票号]='" & temp_rs.Fields("票号") & "' and [部门编码]='" & temp_rs.Fields("部门编码") & "' and [制单人]='" & m_login.cUserName & "'"
                Else
                    strsql = " delete tmparapvouch where  sel=1 and [类型]='" & cFlag & "' and BFlag = 0 and  [票号]='" & temp_rs.Fields("票号") & "' and [部门编码]='" & temp_rs.Fields("部门编码") & "' and [制单人]='" & m_login.cUserName & "'"
                End If
                DBConn.Execute strsql, R
            Else
                R = 2
            End If
        End If
        If R = 1 Then
            DBConn.CommitTrans
            k = 0
            Debug.Print temp_rs.Fields("票号") & "CommitTrans" & Time
        Else
            DBConn.RollbackTrans
            k = 0
            Debug.Print temp_rs.Fields("票号") & "RollbackTrans" & Time
            If MsgBox("['" & temp_rs.Fields("票号") & "']" & errStr & "是否继续？", vbYesNo, Me.Caption) = vbNo Then
'                R = 0
                GoTo errStr
            End If
        End If
        temp_rs.MoveNext
        i = i + 1
    Wend
    Picture3.Visible = False
    Exit Sub
errStr:
    If k = 1 Then DBConn.RollbackTrans
    Picture3.Visible = False
    
End Sub





 
'by ahzzd 20071018 将当前的记录转换成dom
Private Sub GetDom(cFlags As String, Str As String, Str2 As String, domHead As DOMDocument, domBody As DOMDocument)

End Sub
