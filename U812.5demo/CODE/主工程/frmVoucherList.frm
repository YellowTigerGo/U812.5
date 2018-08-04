VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{A0C292A3-118E-11D2-AFDF-000021730160}#1.0#0"; "UFEDIT.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86808282-58F4-4B17-BBCA-951931BB7948}#2.30#0"; "U8VouchList.ocx"
Object = "{201FB79B-5556-47A4-AD9C-A46BA0C45A44}#6.25#0"; "UFToolBarCtrl.ocx"
Begin VB.Form frmVoucherList 
   AutoRedraw      =   -1  'True
   Caption         =   "单据列表"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   8670
   WindowState     =   2  'Maximized
   Begin EDITLib.Edit txtMoCode 
      Height          =   330
      Left            =   9990
      TabIndex        =   13
      Top             =   8145
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   582
      _StockProps     =   253
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      Property        =   1
   End
   Begin EDITLib.Edit txtPageSize 
      Height          =   240
      Left            =   1710
      TabIndex        =   10
      Top             =   7650
      Width           =   810
      _Version        =   65536
      _ExtentX        =   1429
      _ExtentY        =   423
      _StockProps     =   253
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Property        =   3
      MaxLength       =   10
      Appearance      =   0
   End
   Begin U8VouchList.VouchList VouchList 
      Height          =   2295
      Left            =   2040
      TabIndex        =   1
      Top             =   1440
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   4048
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   6960
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   953
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      _Version        =   393216
   End
   Begin EDITLib.Edit txtGoto 
      Height          =   240
      Left            =   3390
      TabIndex        =   11
      Top             =   7650
      Width           =   810
      _Version        =   65536
      _ExtentX        =   1429
      _ExtentY        =   423
      _StockProps     =   253
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Property        =   3
      MaxLength       =   10
      Appearance      =   0
   End
   Begin UFToolBarCtrl.UFToolbar UFToolbar1 
      Height          =   240
      Left            =   240
      TabIndex        =   14
      Top             =   720
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
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   7800
      Top             =   2640
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Label LabMoCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "生产订单号"
      Height          =   195
      Left            =   8820
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   8190
      Width           =   900
   End
   Begin VB.Label labLast 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "末页"
      Height          =   180
      Left            =   6930
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   7650
      Width           =   360
   End
   Begin VB.Label labNext 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "下一页"
      Height          =   180
      Left            =   6210
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   7650
      Width           =   540
   End
   Begin VB.Label labPrevious 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上一页"
      Height          =   180
      Left            =   5580
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   7650
      Width           =   540
   End
   Begin VB.Label labFirst 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "首页"
      Height          =   180
      Left            =   5085
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   7650
      Width           =   360
   End
   Begin VB.Label labOk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "确定"
      Height          =   180
      Left            =   4365
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   7650
      Width           =   360
   End
   Begin VB.Label labGoto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "转到页"
      Height          =   180
      Left            =   2790
      TabIndex        =   4
      Top             =   7650
      Width           =   540
   End
   Begin VB.Label labPageSize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "页大小"
      Height          =   180
      Left            =   1140
      TabIndex        =   3
      Top             =   7650
      Width           =   540
   End
   Begin VB.Label labPage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "第1/1页"
      Height          =   180
      Left            =   180
      TabIndex        =   2
      Top             =   7650
      Width           =   630
   End
End
Attribute VB_Name = "frmVoucherList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim rssum As ADODB.Recordset
Dim m_nCurrentPage As Long
Dim m_nPageCount As Long
Dim m_nRecordCount As Long
Dim m_sTypeWhere As String
Dim m_sWhere As String
Dim M_Vouchlist As New clsVouchlistGDZC
Dim m_strTaskId As String
Dim cKey As String   '单据类型的关键字值(单据列表用)
Dim strVouchType As String '单据类型
Dim objCols As U8colset.clsCols
Dim m_bSum  As Boolean
Dim sTableName As String '临时表名
Dim bRenew As Boolean '是否需要 重新建立临时表
Dim m_bPushMo As Boolean

Private m_Cancel As Integer
Private m_UnloadMode As Integer
Dim vfd As Object
Dim sguid As String
Private WithEvents m_mht As UFPortalProxyMessage.IMessageHandler
Attribute m_mht.VB_VarHelpID = -1

'每个窗体都需要这个方法。Cancel与UnloadMode的参数的含义与QueryUnload的参数相同
'请在此方法中调用窗体Exit(退出)方法，并将设置窗体Unload事件参数(如Cancel)的值同时传给此方法的参数
Public Sub ExitForm(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    Cancel = m_Cancel
    UnloadMode = m_UnloadMode
End Sub

Public Property Let VouchKey(ByVal Str As String)
    cKey = Str
End Property

Public Property Get VouchKey() As String
    VouchKey = cKey
End Property

Public Property Let VouchType(ByVal Str As String)
    strVouchType = Str
End Property

Public Property Get VouchType() As String
    strVouchType = cKey
End Property


Public Property Let strsguid(ByVal Str As String)
    sguid = Str
End Property

Public Property Get strsguid() As String
   strsguid = sguid
End Property

Public Property Let Object_vfd(ByVal Obj As Object)
    Set vfd = Obj
End Property

Public Property Get Object_vfd() As Object
   Set Object_vfd = vfd
End Property

 




Public Property Get cols() As U8colset.clsCols
    Set cols = objCols
End Property

Public Property Get bPushMo() As Boolean
    bPushMo = m_bPushMo
End Property

Public Property Let bPushMo(ByVal bValue As Boolean)
    m_bPushMo = bValue
End Property

'***************************
'系统号
'***************************
Public Property Let Sysid(ByVal Str As String)
    m_sSysID = Str
End Property

Public Property Get Sysid() As String
    Sysid = m_sSysID
End Property

Private Sub setVchFormat()
    On Error GoTo err_log
    Dim sXML As String
    Dim a As Variant
    sXML = VchlstSaPu.GetHeadXml(cKey, m_lstLogin, oColSet)
    Set objCols = oColSet.GetColProperties
'    If cKey = "17" And m_bPushMo Then
'        VouchList.ShowSelCol = True
'    Else
        VouchList.ShowSelCol = True
'    End If
    VouchList.InitHead sXML
'by lg070324 此句不能加上，否则窗口融合功能失效
'    Me.Caption = VouchList.Title
    
    Exit Sub
err_log:
    MsgBox Err.Description, vbOKOnly + vbInformation, Me.Caption
End Sub


Private Function setVchData(Optional sMode As String) As Boolean
On Error GoTo err_log
    M_Vouchlist.bCusAuth = m_bCusAuth
    M_Vouchlist.bDepAuth = m_bDepAuth
    M_Vouchlist.bInvAuth = m_bInvAuth
    M_Vouchlist.bPerAuth = m_bPerAuth
    M_Vouchlist.bUseAuth = m_bUseAuth
    M_Vouchlist.bVenAuth = m_bVenAuth
    If m_sWhere = "" Then
        m_sWhere = m_sWhere & m_sTypeWhere
    Else
        If m_sTypeWhere <> "" Then
            m_sWhere = m_sWhere & " And " & m_sTypeWhere
        End If
    End If
    If m_bSum Then
        Set rssum = Nothing
    End If
    m_nPageSize = val(GetSetting(App.EXEName, "pagesize", "pagesize"))
    If m_nPageSize = 0 Then
        m_nPageSize = 500
    End If
    Dim sErrRet As String
    If m_nCurrentPage <= 0 Then m_nCurrentPage = 1
    Randomize
    If sTableName = "" Then sTableName = "Lst" & m_lstLogin.cUserId & Fix(Rnd * 1000)
    bRenew = Not m_bSum
    If Not m_bSum Then
        If Not M_Vouchlist.GetVchListData(m_sSysID, cKey, m_lstLogin, False, m_nPageSize, m_nCurrentPage, m_nPageCount, rs, rssum, oColSet, m_sWhere, sErrRet, sTableName, bRenew, m_nRecordCount, strVouchType) Then
            If sErrRet <> "" Then
                MsgBox sErrRet, vbOKOnly + vbExclamation, Me.Caption
                Exit Function
            End If
        End If
    Else
        If Not VchlstSaPu.GetVchListData(m_sSysID, cKey, m_lstLogin, False, m_nPageSize, m_nCurrentPage, m_nPageCount, rs, , oColSet, m_sWhere, sErrRet, sTableName, bRenew, m_nRecordCount, strVouchType) Then
            If sErrRet <> "" Then
                MsgBox sErrRet, vbOKOnly + vbExclamation, Me.Caption
                Exit Function
            End If
        End If
    End If
    m_bSum = True
    Call setDataFormat
    If UCase(sMode) = "APPEND" Then
        VouchList.FillMode = FillAppend
    Else
        VouchList.FillMode = FillOverwrite
    End If
    'by ahzzd 200607021 单据列表只要合计行 不要小计行
    VouchList.SumStyle = vlGridSum  '
    VouchList.SetSumRst rssum
    VouchList.SetVchLstRst rs
    Set rsVouchlist = rs
    setVchFormat
    setVchData = True
    'txtGoto.Text = m_nCurrentPage
    If m_nPageCount = 0 Then m_nPageCount = 1
    If m_nCurrentPage > m_nPageCount Then m_nCurrentPage = 1
    labPage.Caption = "第" & m_nCurrentPage & "/" & m_nPageCount & "页"
    txtPageSize.Text = m_nPageSize
    txtGoto.Text = m_nCurrentPage
    If m_nCurrentPage = 1 Then
        labFirst.Enabled = False
        labPrevious.Enabled = False
    Else
        labFirst.Enabled = True
        labPrevious.Enabled = True
    End If
    If m_nCurrentPage = m_nPageCount Then
        labNext.Enabled = False
        labLast.Enabled = False
    Else
        labNext.Enabled = True
        labLast.Enabled = True
    End If
    VouchList.RecordCount = m_nRecordCount
    Exit Function
err_log:
    MsgBox Err.Description, vbOKOnly + vbInformation, Me.Caption
End Function

Private Function getPrnSet() As String
    Dim rst As New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open "select * from prn_format where moduleid='" & cKey & "_" & "Print_" & m_lstLogin.cUserId & "'", m_connLst, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rst.EOF Then
        getPrnSet = ""
    Else
        getPrnSet = rst.Fields("formatxml")
    End If
    rst.Close
    Set rst = Nothing
End Function
Private Sub SavePrnSet(strPrnXml As String)
    Dim rst As New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open "select * from prn_format where moduleid='" & cKey & "_" & "Print_" & m_lstLogin.cUserId & "'", m_connLst, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rst.EOF Then
        m_connLst.Execute "insert into prn_format (moduleid,formatxml) values ('" & cKey & "_" & "Print_" & m_lstLogin.cUserId & "','" & strPrnXml & "')"
    Else
        m_connLst.Execute "update prn_format set formatxml='" & strPrnXml & "' where moduleid='" & cKey & "_" & "Print_" & m_lstLogin.cUserId & "'"
    End If
    rst.Close
    Set rst = Nothing
End Sub

Private Sub OnButtonClick(strButtonKey As String)
    Dim strPrnXml As String
    Dim cAuthId As String
    Dim strVouchID As String
    Dim rsMytemp As New ADODB.Recordset
    Dim domMyHeadDom As New DOMDocument
    Dim domMybodyDom As New DOMDocument
    Dim strTempMyID As String
    Dim i As Long
    Dim sqlstr As String
    Dim falg As Boolean
    On Error GoTo Err
    Select Case UCase(strButtonKey)
    
        Case "SELECTALL"  '全选
            For i = 1 To VouchList.Rows - 1
                 VouchList.TextMatrix(VouchList.Rows - i, objCols("selCol").iColPos) = "Y"
            Next
        
        Case "SELECTED_MODE" '反选
            For i = 1 To VouchList.Rows - 1
                If VouchList.TextMatrix(VouchList.Rows - i, objCols("selCol").iColPos) = "Y" Then
                    VouchList.TextMatrix(VouchList.Rows - i, objCols("selCol").iColPos) = ""
                Else
                    VouchList.TextMatrix(VouchList.Rows - i, objCols("selCol").iColPos) = "Y"
                End If
            Next
        
        Case "UNSELECTALL" '全不选
            For i = 1 To VouchList.Rows - 1
                 VouchList.TextMatrix(VouchList.Rows - i, objCols("selCol").iColPos) = ""
            Next
        Case "MAKE_SURE"  '确定
            MsgBox "例子还没有做任何处理,您可以按照您的业务需要进行对应的处理!"
       
        Case "HELP"
            SendKeys "{F1}"
        Case "FILTERSETTING"
            Call FilterSetting
        Case "COLUMN"  '栏目column
            If sTableName <> "" Then Call M_Vouchlist.DropTmpTable("TempDB.." & "TMPUF_" & m_lstLogin.TaskId & "_" & sTableName, m_lstLogin)
            m_bSum = False '过滤重算求和，重新创建临时表
            sTableName = ""
            Call ColumnSet
        Case "FILTER"
            Call Filter
        Case "QUERY" 'query
            Call Locate
        Case "NEXT"
            VouchList.Find
        Case "PRING"
            If cKey = "FA180" Then
           
            Else
                strPrnXml = getPrnSet()
                If strPrnXml <> "" Then VouchList.InitPrintSetup strPrnXml
                VouchList.VchLstPrint
            End If
        Case "PREVIEW"
            If cKey = "FA180" Then
               
            Else
                strPrnXml = getPrnSet()
                If strPrnXml <> "" Then VouchList.InitPrintSetup strPrnXml
                VouchList.VchLstPreview
            End If
        Case "EXPORT"
            If VouchList.Rows > 1 Then
                VouchList.VchLstPrintToFile
            Else
                MsgBox "没有可以输出的数据", vbCritical
            End If
        Case "REFRESH"
            Screen.MousePointer = vbHourglass
            Toolbar.buttons("refresh").Enabled = False
'            UFToolbar1.RefreshEnable
            If sTableName <> "" Then Call M_Vouchlist.DropTmpTable("TempDB.." & "TMPUF_" & m_lstLogin.TaskId & "_" & sTableName, m_lstLogin)
            m_bSum = False '过滤重算求和，重新创建临时表
            sTableName = ""
            Call RefreshVchLst
            Toolbar.buttons("refresh").Enabled = True
'            UFToolbar1.RefreshEnable
            Screen.MousePointer = vbDefault
                    
        Case "EXIT"
            Unload Me
        Case "BATCHPRINT"
            Call BatchOperation(VouchList, "print", cKey, objCols)
        Case "BATCHOPEN"
            Call BatchOperation(VouchList, "open", cKey, objCols)
        Case "BATCHCLOSE"
            Call BatchOperation(VouchList, "close", cKey, objCols)
        Case "BATCHCONFIRM"
            Call BatchOperation(VouchList, "confirm", cKey, objCols)
        Case "BATCHUNCONFIRM"
            Call BatchOperation(VouchList, "unconfirm", cKey, objCols)
        Case UCase(strKSelectAll)
            Me.VouchList.AllSelect
        Case UCase(strKUnSelectAll)
            Me.VouchList.AllNone
    End Select
    Exit Sub
Err:
    MsgBox Err.Description
End Sub

Private Sub UFToolbar1_OnCommand(ByVal enumType As UFToolBarCtrl.ENUM_MENU_OR_BUTTON, ByVal cButtonId As String, ByVal cMenuId As String)
    OnButtonClick IIf(enumType = enumButton, cButtonId, cMenuId) ', ""
End Sub

'Private Sub UFToolbar1_OnCommand(ByVal enumType As prjTBCtrl.ENUM_MENU_OR_BUTTON, ByVal cButtonId As String, ByVal cMenuId As String)
'    OnButtonClick cButtonId
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            If Me.Toolbar.buttons("filter").Enabled = True And Shift <> 2 Then
                Call Toolbar_ButtonClick(Me.Toolbar.buttons("filter"))
            ElseIf Me.Toolbar.buttons("Locate").Enabled = True And Shift = 2 Then
                Call Toolbar_ButtonClick(Me.Toolbar.buttons("Locate"))
            End If
        Case vbKeyP
            If Shift = 2 Then
                If Me.Toolbar.buttons("Pring").Enabled = True Then
                    Call Toolbar_ButtonClick(Me.Toolbar.buttons("Pring"))
                End If
            End If
    End Select
End Sub

Private Sub Form_Load()
    Dim sXML As String
    Dim strDWName As String
    m_nPageSize = val(GetSetting(App.EXEName, "pagesize", "pagesize"))
    If m_nPageSize = 0 Then
        m_nPageSize = 500
    End If
    'by lg070314增加U870菜单融合功能
    ''''''''''''''''''''''''''''''''''''''
    If Not g_business Is Nothing Then
        Set Me.UFToolbar1.Business = g_business
    End If
    Call RegisterMessage
    '''''''''''''''''''''''''''''''''''''''
    Call initToolBar
    Set Me.labFirst.MouseIcon = MyRes.LoadResPic(998, 2)
    Set Me.labGoto.MouseIcon = MyRes.LoadResPic(998, 2)
    Set Me.labLast.MouseIcon = MyRes.LoadResPic(998, 2)
    Set Me.labNext.MouseIcon = MyRes.LoadResPic(998, 2)
    Set Me.labOk.MouseIcon = MyRes.LoadResPic(998, 2)
    Set Me.labPrevious.MouseIcon = MyRes.LoadResPic(998, 2)
    m_lstLogin.GetAccInfo 105, strDWName
    sXML = "<表尾>"
    sXML = sXML + "<字段>" + "单位：" + strDWName + "</字段> "
    sXML = sXML + "<字段>" + "制表：" + m_lstLogin.cUserName + "</字段> "
    sXML = sXML + "<字段>" + "打印日期：" + Format(m_lstLogin.CurDate, "YYYY-MM-DD") + "</字段> "
    sXML = sXML + "</表尾>"
    Me.VouchList.SetPrintOtherInfo sXML
    Set Me.Icon = frmMain.Icon
End Sub

Private Sub initToolBar()
    iniImageList
    AddButtons
'    Toolbar.ImageList = ImageList
    With Toolbar
        Set .ImageList = frmMain.imgBmp
        Dim i As Long
        For i = 1 To Toolbar.buttons.Count
        Next
        If cKey = "17" And m_bPushMo Then
            Me.Toolbar.buttons("pushmo").Visible = True
            Me.Toolbar.buttons(strKSelectAll).Visible = True
            Me.Toolbar.buttons(strKUnSelectAll).Visible = True
            LabMoCode.Visible = False
            txtMoCode.Visible = False
        Else
'            Me.Toolbar.buttons("pushmo").Visible = False
            Me.Toolbar.buttons(strKSelectAll).Visible = False
            Me.Toolbar.buttons(strKUnSelectAll).Visible = False
            LabMoCode.Visible = False
            txtMoCode.Visible = False
        End If
            Me.Toolbar.buttons("SelectAll").Visible = True
            Me.Toolbar.buttons("UnSelectAll").Visible = True
            Me.Toolbar.buttons("selected_mode").Visible = True
            Me.Toolbar.buttons("make_sure").Visible = True
     End With
'by增加 870功菜单功能  zzd 0324
    ChangeOneFormTbr Me, Me.Toolbar, Me.UFToolbar1
End Sub
Private Sub AddButtons()
    Dim btnX As MSComctlLib.Button
    With Toolbar.buttons
        .Clear
         ''打印
        Set btnX = .Add(, "Print", strPrint, tbrDefault)
'            btnX.image = 314
        btnX.ToolTipText = strPrint
        btnX.Description = btnX.ToolTipText
        btnX.Tag = "Print"
                 
         '预览
         Set btnX = .Add(, "Preview", strPreview, tbrDefault)
'          btnX.image = 314
         btnX.ToolTipText = strPreview
         btnX.Description = btnX.ToolTipText
         btnX.Tag = "Preview"
         

        '全选
        Set btnX = .Add(, "SelectAll", strSelectAll, tbrDefault)
'            btnX.image = 314
        btnX.ToolTipText = strSelectAll
        btnX.Description = btnX.ToolTipText
        btnX.Tag = "Select all"
        
        '反选
        Set btnX = .Add(, "selected_mode", strconSelectAll, tbrDefault)
'            btnX.image = 314
        btnX.ToolTipText = strconSelectAll
        btnX.Description = btnX.ToolTipText
        btnX.Tag = "selected_mode"
 
        '全不选
         Set btnX = .Add(, "UnSelectAll", strUnSelectAll, tbrDefault)
'          btnX.image = 314
         btnX.ToolTipText = strUnSelectAll
         btnX.Description = btnX.ToolTipText
         btnX.Tag = "Select none"
         
        '确定
         Set btnX = .Add(, "make_sure", strmake_sure, tbrDefault)
'          btnX.image = 314
         btnX.ToolTipText = strmake_sure
         btnX.Description = btnX.ToolTipText
         btnX.Tag = "add_obj"

         '输出
         Set btnX = .Add(, "export", strOutput, tbrDefault)
'          btnX.image = 314
         btnX.ToolTipText = strOutput
         btnX.Description = btnX.ToolTipText
         btnX.Tag = "export"
         
         '过滤
         Set btnX = .Add(, "filter", strFilter, tbrDefault)
'          btnX.image = 314
         btnX.ToolTipText = strFilter
         btnX.Description = btnX.ToolTipText
         btnX.Tag = "filter"
 
        '定位
         Set btnX = .Add(, "query", strLocate, tbrDefault)
'          btnX.image = 314
         btnX.ToolTipText = strLocate
         btnX.Description = btnX.ToolTipText
         btnX.Tag = "query"
         
         '滤设
         Set btnX = .Add(, "filtersetting", strfiltersetting, tbrDefault)
'          btnX.image = 314
         btnX.ToolTipText = strfiltersetting
         btnX.Description = btnX.ToolTipText
         btnX.Tag = "design"
         
         '栏目
         Set btnX = .Add(, "column", strColumn, tbrDefault)
'          btnX.image = 314
         btnX.ToolTipText = strColumn
         btnX.Description = btnX.ToolTipText
         btnX.Tag = "column"
         
'         '处理
'         Set btnX = .Add(, "addvouth", addvouth, tbrDefault)
''          btnX.image = 314
'         btnX.ToolTipText = addvouth
'         btnX.Description = btnX.ToolTipText
'         btnX.Tag = "control"
            
         '刷新
          Set btnX = .Add(, "refresh", strRefresh, tbrDefault)
'          btnX.image = 314
         btnX.ToolTipText = strRefresh
         btnX.Description = btnX.ToolTipText
         btnX.Tag = "refresh"
         
        '帮助
         Set btnX = .Add(, "help", strHelp, tbrDefault)
'          btnX.image = 314
         btnX.ToolTipText = strHelp
         btnX.Description = btnX.ToolTipText
         btnX.Tag = "help"
        
         '退出
'                 Set btnX = .Add(, "strExit", strExit, tbrDefault)
''          btnX.image = 314
'         btnX.ToolTipText = strExit
'         btnX.Description = btnX.ToolTipText
'         btnX.Tag = "strExit"
 
    End With
    
'by lg070324 增加菜单融合

    Call InitToolbarTag(Me.Toolbar)
End Sub
Private Sub iniImageList()
    With ImageList.ListImages
        .Add , , MyRes.LoadResPic(IDB_Print, 0)
        .Add , , MyRes.LoadResPic(IDB_Preview, 0)
        .Add , , MyRes.LoadResPic(IDB_Output, 0)
        
        .Add , , MyRes.LoadResPic(IDB_Filter, 0)
        .Add , , MyRes.LoadResPic(IDB_Locate, 0)
'        .Add , , MyRes.LoadResPic(IDB_Next, 0)
        .Add , , MyRes.LoadResPic(IDB_FilterSet, 0)
        .Add , , MyRes.LoadResPic(IDB_Set, 0)
        
        .Add , , MyRes.LoadResPic(IDB_SelectAll, 0)
        .Add , , MyRes.LoadResPic(IDB_UnSelectAll, 0)
'        .Add , , MyRes.LoadResPic(IDB_BatchPrint, 0)
'        .Add , , MyRes.LoadResPic(IDB_BatchOpen, 0)
'        .Add , , MyRes.LoadResPic(IDB_BatchClose, 0)
'        .Add , , MyRes.LoadResPic(IDB_BatchVeri, 0)
        .Add , , MyRes.LoadResPic(IDB_BatchUnVeri, 0)
        .Add , , MyRes.LoadResPic(IDB_Payment, 0) '下达生产
        .Add , , MyRes.LoadResPic(IDB_Refresh, 0)
        .Add , , MyRes.LoadResPic(IDB_Help, 0)
        .Add , , MyRes.LoadResPic(IDB_Exit, 0)
    End With
End Sub

Private Sub ResizeForm()
    On Error Resume Next
'    Me.UFToolbar1.Move 0, 0, Me.ScaleWidth
    
    Me.UFToolbar1.Top = 0
    Me.UFToolbar1.Width = Me.ScaleWidth
'    Me.VouchList.Move 0, Me.UFToolbar1.Height, Me.ScaleWidth, IIf(Me.Height - 1300 < 0, 0, Me.Height - 1800)
    
    Me.VouchList.Move 0, Me.UFToolbar1.Height, Me.ScaleWidth, IIf(Me.Height - 1300 < 0, 0, Me.Height - 600)
'    Me.VouchList.Top = Me.UFToolbar1.Height
'    Me.VouchList.Height = Me.ScaleHeight
'    Me.VouchList.Width = Me.ScaleWidth

    labPage.Top = VouchList.Top + VouchList.Height + 300
    labPage.Left = VouchList.Left
    labPageSize.Top = labPage.Top
    labPageSize.Top = labPage.Top
    txtPageSize.Top = labPage.Top
    labGoto.Top = labPage.Top
    txtGoto.Top = labPage.Top
    labOk.Top = labPage.Top
    labFirst.Top = labPage.Top
    labPrevious.Top = labPage.Top
    labNext.Top = labPage.Top
    labLast.Top = labPage.Top
    LabMoCode.Top = labPage.Top
    LabMoCode.Left = labLast.Left + 500
    txtMoCode.Top = labPage.Top
    txtMoCode.Left = LabMoCode.Left + LabMoCode.Width + 100
End Sub

'///////////////////////////////////////////////////////////////////////////////
'功能设置单据列表的对应字段的显示格式设置格式化串
'by 客户化开发中心 2006/03/01
'//////////////////////////////////////////////////////////////////////////////
Private Sub setDataFormat()
    Select Case cKey
        Case "FA110"
            Call SetDataFormatFA110(VouchList)
        Case Else
            Call SetDataFormatFA110(VouchList)
    End Select
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetLabBlack
End Sub

Private Sub Form_Resize()

    Me.ForeColor = &H80000003
    Me.Line (0, Me.Toolbar.Height)-(Me.ScaleWidth, Me.Toolbar.Height)
    Me.ForeColor = vbWhite
    Me.Line (0, Me.Toolbar.Height + 15)-(Me.ScaleWidth, Me.Toolbar.Height + 15)
    ResizeForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set UFToolbar1.Business = Nothing
    m_lstLogin.TaskExec strTaskId, 0, val(m_lstLogin.cIYear)
End Sub

Private Sub labFirst_Click()
    If m_nCurrentPage = 1 Then Exit Sub
    m_nCurrentPage = 1
    Call setVchData
End Sub

Private Sub labFirst_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetLabBlue(labFirst)
End Sub

Private Sub labLast_Click()
    If m_nCurrentPage = m_nPageCount Then Exit Sub
    m_nCurrentPage = m_nPageCount
    Call setVchData
End Sub

Private Sub labLast_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetLabBlue(labLast)
End Sub

Private Sub labNext_Click()
    If m_nCurrentPage = m_nPageCount Then Exit Sub
    m_nCurrentPage = m_nCurrentPage + 1
    Call setVchData
End Sub

Private Sub labNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call SetLabBlue(labNext)
End Sub

Private Sub labOk_Click()
    If Not IsNumeric(txtGoto.Text) Or val(txtGoto.Text) > m_nPageCount Then txtGoto.Text = 1
    m_nCurrentPage = val(txtGoto.Text)
    Call SaveSetting(App.EXEName, "pagesize", "pagesize", txtPageSize.Text)
    Call setVchData
End Sub

Private Sub labOk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetLabBlue(labOk)
End Sub

Private Sub labPrevious_Click()
    If m_nCurrentPage = 1 Then Exit Sub
    m_nCurrentPage = m_nCurrentPage - 1
    Call setVchData
End Sub

Private Sub labPrevious_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetLabBlue(labPrevious)
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    OnButtonClick Button.key
End Sub

Private Sub RefreshVchLst()
    Call setVchData
    Call setVchFormat
End Sub


'*********过滤设置*************
Private Sub FilterSetting()
    If cKey <> "24" Then
        Call objfltint.SetFilter(cKey, m_sSysID, m_connLst, , False)
    Else
        Call objfltint.SetFilter("单据列表", "FA", m_connLst, , False)
    End If
    objfltint.DeleteFilter
    Set objfltint = Nothing
End Sub

'**********栏目设置*****************
Private Sub ColumnSet()
    oColSet.setColMode cKey, 1 'CvouType
    oColSet.isShowTitle = True
    If oColSet.setCol Then
        setVchData
        'setVchFormat
    End If
End Sub

'***********定位***********************
Private Sub Locate()
    If Not VouchList.LocateState Then
        VouchList.Locate True
    Else
        VouchList.Locate False
    End If
End Sub

'调用过滤
Public Function Filter() As Boolean
    Dim objfltint As New clsFilterInterface
    Dim objFilter As clsReportFilter
    Dim sWhere As String
    Filter = False
    
    Filter = objfltint.OpenFilter(cKey, m_sSysID, m_connLst, , True, True, m_lstLogin.cUserId, m_lstLogin)
    Call Unload_frms(Me.Name)
    If Filter Then
        If sTableName <> "" Then Call M_Vouchlist.DropTmpTable("TempDB.." & "TMPUF_" & m_lstLogin.TaskId & "_" & sTableName, m_lstLogin)
        m_bSum = False '过滤重算求和，重新创建临时表
        sTableName = ""
        Set objFilter = objfltint.GetReportFilterObject(cKey)
        m_sWhere = getWhereStrFromHeron(cKey, objFilter.SQLString)
        objfltint.DeleteFilter
        Set objfltint = Nothing
        Call setVchData(sWhere)
        Filter = True
    End If
End Function

Private Function getAuthString(strBusObId As String) As String
    Dim objRowAuthsrv As U8RowAuthsvr.clsRowAuth
    Dim strTmp As String
    Set objRowAuthsrv = New U8RowAuthsvr.clsRowAuth
    objRowAuthsrv.Init m_lstLogin.UfDbName, "UFSOFT"
    strTmp = objRowAuthsrv.getAuthString(strBusObId)
    'If Len(strTmp) = 0 Then Exit Function
    
    Select Case UCase(strBusObId)
    Case "CUSTOMER"
        getAuthString = " WHERE " & strTmp
    Case "VENDOR"

    Case "INVENTORY"
        
    Case "DEPARTMENT"
        getAuthString = " cDepCode in (" & strTmp & ")"
        
    Case "PERSON"
    
    Case "USER"
    End Select
    
End Function

Private Sub Command1_Click()
    Dim i As Long
    Dim Str As String
    For i = 1 To 50
        Str = Str & "AAAA" & Chr(9)
    Next
    VouchList.AddItem Str, VouchList.Rows
End Sub

Private Sub txtGoto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call labOk_Click
    End If
End Sub


Private Sub txtGoto_Validate(Cancel As Boolean)
    If (Not IsNumeric(txtGoto.Text) Or val(txtGoto.Text) > m_nPageCount) And Trim(txtGoto.Text) <> "" Or val(txtGoto.Text) <= 0 Then
        Cancel = True
        txtGoto.ForeColor = vbRed
    Else
        txtGoto.ForeColor = vbBlack
    End If
End Sub

Private Sub txtPageSize_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call labOk_Click
    End If
End Sub
Private Sub txtPageSize_Validate(Cancel As Boolean)
    If txtPageSize.Text = "" Then Cancel = True
    If Not IsNumeric(txtPageSize.Text) Then Cancel = True
    If val(txtPageSize.Text) >= 2147483647 Or val(txtPageSize.Text) <= 0 Then Cancel = True
    If val(txtPageSize.Text) <> val(GetSetting(App.EXEName, "pagesize", "pagesize")) Then
        m_nCurrentPage = 0
        Call SaveSetting(App.EXEName, "pagesize", "pagesize", txtPageSize.Text)
    End If
End Sub

Private Sub VouchList_DblClick()
    Call ShowVouch(VouchList, cKey, objCols)
End Sub

Private Sub VouchList_PrintSettingChanged(ByVal varLocalSettings As Variant, ByVal varModuleSettings As Variant)
    SavePrnSet varLocalSettings + varModuleSettings
End Sub

Private Sub VouchList_Scroll()
'    Call Scroll
End Sub

Private Sub Scroll()
    If VouchList.TopRow >= m_nCurrentPage * m_nPageSize - 100 Then
        If m_nCurrentPage <= m_nPageCount Then
            m_nCurrentPage = m_nCurrentPage + 1
            Call setVchData("append")
        End If

    End If
    
End Sub

Public Property Let TypeWhere(ByVal sValue As String)
    m_sTypeWhere = sValue
End Property

Public Property Get TypeWhere() As String
    TypeWhere = m_sTypeWhere
End Property

Private Sub SetLabBlue(ByVal lab As Label)
    lab.ForeColor = vbBlue
End Sub

Private Sub SetLabBlack()
    labOk.ForeColor = vbBlack
    labFirst.ForeColor = vbBlack
    labNext.ForeColor = vbBlack
    labPrevious.ForeColor = vbBlack
    labLast.ForeColor = vbBlack
    labGoto.ForeColor = vbBlack
End Sub

Public Property Get strTaskId() As String
    strTaskId = m_strTaskId
End Property

Public Property Let strTaskId(ByVal vNewValue As String)
    m_strTaskId = vNewValue
End Property

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

 
 
 
 Private Sub RegisterMessage()
    Set m_mht = New UFPortalProxyMessage.IMessageHandler
    m_mht.MessageType = "DocAuditComplete"
    If Not g_business Is Nothing Then
        Call g_business.RegisterMessageHandler(m_mht)
    End If
End Sub

