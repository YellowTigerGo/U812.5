VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{201FB79B-5556-47A4-AD9C-A46BA0C45A44}#6.42#0"; "UFToolBarCtrl.ocx"
Object = "{005151D4-33F6-471B-B651-222D4E411832}#4.5#0"; "UFFormPartner.ocx"
Object = "{AF8BBBB7-94C6-4772-B826-624478C37D6A}#1.5#0"; "UFKEYHOOK.ocx"
Object = "{86808282-58F4-4B17-BBCA-951931BB7948}#2.79#0"; "U8VouchList.ocx"
Begin VB.Form FrmList 
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   9975
   StartUpPosition =   3  '窗口缺省
   Begin UFFormPartner.UFFrmCaption UFFrmCaptionMgr 
      Left            =   5880
      Top             =   2640
      _ExtentX        =   1931
      _ExtentY        =   1085
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
   Begin U8VouchList.VouchList VouchList1 
      Height          =   3135
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5530
   End
   Begin UFToolBarCtrl.UFToolbar UFToolbar1 
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   975
      _ExtentX        =   1720
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   1085
      ButtonWidth     =   609
      ButtonHeight    =   926
      Appearance      =   1
      _Version        =   393216
   End
   Begin UFKeyHook.UFKeyHookCtrl UFKeyHookCtrl1 
      Left            =   5640
      Top             =   5280
      _ExtentX        =   1905
      _ExtentY        =   529
   End
End
Attribute VB_Name = "FrmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************
'模块功能说明：
'               1、实现列表的基本按钮功能:打印、输出、预览、
'                                       批审、批弃、批开、批关、过滤、定位、栏目、全选、全消、联查、刷新
'               2、支持列表分页功能
'               3、支持功能权限

'创建时间：2008-11-21
'创建人：xuyan
'****************************************

'列表变量
Public objColset As U8ColumnSet.clsColSet

Private WithEvents m_pagediv As Pagediv                    '分页
Attribute m_pagediv.VB_VarHelpID = -1
Private m_coni As IPagedivConi                             '条件，基本上都是从U8Colset中进行初始化

Private m_Cancel, m_UnloadMode As Integer
Attribute m_UnloadMode.VB_VarUserMemId = 1073938434
Private ListTitle As String
Attribute ListTitle.VB_VarUserMemId = 1073938436

Dim cMenuId, cMenuName, cAuthId As String    ' 单据节点
Attribute cMenuId.VB_VarUserMemId = 1073938437
Attribute cMenuName.VB_VarUserMemId = 1073938437
Attribute cAuthId.VB_VarUserMemId = 1073938437

'功能权限
Private Const AuthBrowselist = "FYSL02050603"    '浏览
Private Const AuthBrowseLink = "ST02JC020406"    '联查
Private Const AuthPrint = "FSFY02010201"    '打印
Private Const AuthOut = "FSFY02010202"    '输出
Private Const AuthVerify = "ST02JC020105"   '审核
Private Const AuthUnVerify = "ST02JC020106"    '弃审
Private Const AuthReturn = "ST02JC020302"    '批量归还


Private Const AuthOpen = "SAM0302004"    '打开
Private Const AuthClose = "SAM0302005"    '关闭

' Private Const AuthVerify = "ST02JC020105" '审核
' Private Const AuthUnVerify = "ST02JC020106" '弃审

Private strMsg As String
Attribute strMsg.VB_VarUserMemId = 1073938440
Private oldccode As String
Private strStatus3 As String
Attribute strStatus3.VB_VarUserMemId = 1073938442

'by zhangwchb 20110719 列表纬度扩展
Dim sExtendField As String
Dim sExtendJoinSQL As String
Dim oExtend As Object

Private o_FilterObject As Object

Private Sub Form_Activate()
    Form_Resize
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF1
        Call LoadHelpId(Me, "15031400")
    End Select
End Sub

Private Sub Form_Load()

'初始化窗体

    
    Call InitMulText
    Call InitForm
    '1、初始化列表：加载列表模板 InitList
    '2、加载过滤条件
    '3、初始化分页控件 InitConi
    '4、填充分页控件数据，加载列表数据 FillList(m_pagediv.LoadData)、m_pagediv_GetData、m_pagediv_AfterGetData

    '初始化列表
    Call InitList

    '初始化分页控件
    Call InitConi(strWhere)

    '填充分页控件数据，加载列表数据
    Call FillList
    
    '11.1合并显示
    Call SetToolbarForColumn

End Sub

Private Sub Form_Resize()

    UFToolbar1.Move 0, 0, Me.ScaleWidth, Me.Toolbar1.Height
    'VouchList1.Move 0, 0, Me.ScaleWidth, IIf(Me.ScaleHeight - Toolbar1.Height < 0, 0, Me.ScaleHeight - Toolbar1.Height)
    VouchList1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
    'PagedivCtl1.Move 0, Me.VouchList1.Height, Me.ScaleWidth, 600

End Sub

Private Sub InitForm()
    On Error Resume Next


    ListTitle = "费用预算单列表"  '列表标题
    cMenuId = "FSFY020102"
    cMenuName = "费用预算单列表"
   ' cAuthId = "ST02JC020101"
    '*************
    
    '*******************
    
     Dim ErrInfo As String
     Dim bSuccess As Boolean
     Dim sListName As String
     sListName = "费用预算单列表参照"
     
     Dim errStr As String
    Set clsbill = CreateObject("USERPCO.VoucherCO")        'New USERPCO.VoucherCO
    clsbill.IniLogin g_oLogin, errStr
    Set mologin = clsbill.login
    Set UFToolbar1.Business = goBusiness
    
    '****************wangfb 11.0Toobar迁移2012-03-21 start ************************
    Call InitToolBar(mologin, "HY_FSFY_Costbudget002", Toolbar1, UFToolbar1)
    
    
    Call UFToolbar1.InitExternalButton("Collection002_List", mologin.OldLogin)
    Call UFToolbar1.SetFormInfo(Me.VouchList1, Me)
    If sListName <> "" Then
        If o_FilterObject Is Nothing Then
           Set o_FilterObject = CreateObject("UFGeneralFilter.FilterSrv")
        End If
       
        bSuccess = o_FilterObject.InitBaseVarValue(g_oLogin, "", sListName, "ST", ErrInfo)
        VouchList1.InitFlt g_oLogin, o_FilterObject, "", "", "", ErrInfo
     End If
     
    UFFrmCaptionMgr.Caption = "费用预算单列表"
    Me.Caption = UFFrmCaptionMgr.Caption

    VouchList1.HiddenRefreshView = False
    
    VouchList1.formCode = "HY_FSFY_Costbudget002"
    '工具栏初始化
    '11.0toolbar迁移，借入借出业务单据标准化之后，原来客开的按钮初始化
    'Call Init_Toolbarlist(Me.Toolbar1)
    
    Call ChangeOneFormTbrlist(Me, Me.Toolbar1, Me.UFToolbar1)
    Call SetWFControlBrnsList(g_oLogin, g_Conn, Me.Toolbar1, Me.UFToolbar1, gstrCardNumber)
    
    SetToolbarVisible
    UFToolbar1.RefreshVisible
    '****************wangfb 11.0Toobar迁移2012-03-21 start ************************
    
    '获取U8版本 -chenliangc
    gU8Version = GetU8Version(gConn)

End Sub


Private Sub InitList()
    On Error GoTo Err_Handler

    Set objColset = New U8ColumnSet.clsColSet
    objColset.Init gConn, goLogin.cUserId

    '11.1合并显示
    Set VouchList1.ColSetObj = objColset
    
    '加载列表模板
    objColset.setColMode gstrCardNumberlist
    Me.VouchList1.InitHead objColset.getColInfo()

    '合计方式
    If Replace(strWhere, " ", "") = "(1=2)" Then
        VouchList1.SumStyle = vlSumNone '   vlRecordAndGridsum                     ' vlGridSum
    Else
        VouchList1.SumStyle = vlRecordAndGridsum
    End If

    'FillAppend 附加
    'FillOverwrite 此种填充方式可以使用选择功能，[选择]列显示为固定列
    Me.VouchList1.FillMode = FillOverwrite
    VouchList1.ShowSelCol = True
    '    VouchList1.SumStyle = vlSumNone

    '列表标题名称
    Me.VouchList1.Title = ListTitle

    'Public m_sQuantityFmt As String '数量格式
    'Public m_sNumFmt As String      ' 数值格式
    'Public m_iExchRateFmt As String   ' 换算率
    'Public m_iRateFmt As String   ' 税率
    'Public m_sPriceFmt As String  ' 金额格式
    'Public m_sPriceFmtSA As String  ' 金额格式（销售用）
    With VouchList1
        .SetFormatString "cfreightCost", m_sPriceFmt
        .SetFormatString "iinvexchrate", m_iExchRateFmt

        .SetFormatString "iquantity", m_sQuantityFmt
        .SetFormatString "inum", m_sNumFmt
        .SetFormatString "iQtyBackSum", m_sQuantityFmt
        .SetFormatString "iQtyBack2Sum", m_sNumFmt
        .SetFormatString "iQtyCFreeSum", m_sQuantityFmt
        .SetFormatString "iQtyCFree2Sum", m_sNumFmt
        .SetFormatString "iQtyCOutSum", m_sQuantityFmt
        .SetFormatString "iQtyCOut2Sum", m_sNumFmt
        .SetFormatString "iQtyCOverSum", m_sQuantityFmt
        .SetFormatString "iQtyCOver2Sum", m_sNumFmt
        .SetFormatString "iQtyCSaleSum", m_sQuantityFmt
        .SetFormatString "iQtyCSale2Sum", m_sNumFmt
        .SetFormatString "iQtyOutSum", m_sQuantityFmt
        .SetFormatString "iQtyOut2Sum", m_sNumFmt

        .SetFormatString "cdefine7", m_sQuantityFmt
        .SetFormatString "cdefine16", m_sQuantityFmt
        .SetFormatString "cdefine26", m_sQuantityFmt
        .SetFormatString "cdefine27", m_sQuantityFmt
    End With

    Exit Sub

Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Sub
Private Sub FillList()

    Me.VouchList1.SetVchLstRst Nothing
    '加载数据
    m_pagediv.LoadData

End Sub
Private Sub InitConi(mwhere As String)
    On Error GoTo Err_Handler


    Set m_pagediv = New Pagediv
    
    If m_coni Is Nothing Then
        Set m_coni = New DefaultPagedivConi
    End If
    
    'by zhangwchb 20110719 列表纬度扩展
    Set oExtend = CreateObject("VoucherExtendService.ClsExtendServer")
    Call oExtend.GetExtendInfo(gConn, gstrCardNumberlist, "L", sExtendField, sExtendJoinSQL)

    m_coni.From = VoucherList & sExtendJoinSQL   'MainView '相当与from
    m_coni.SelectConi = objColset.GetSqlString    '相当与查询字段
    m_coni.SumConi = objColset.GetSumString
    m_coni.where = "(1=1)"

    If mwhere <> "" Then m_coni.where = m_coni.where & " and " & mwhere    '查询条件
    '权限处理
    m_coni.where = m_coni.where

    m_coni.OrderID = objColset.GetOrderString    '排序字段

    'Call PagedivCtl1.BindPagediv(m_pagediv)
    Call m_pagediv.Initialize(gConn, m_coni)
    Call VouchList1.BindPagediv(m_pagediv)
    DropTable "tempdb..TEMP_STSearchTableNameList_" & sGUID
    g_Conn.Execute "select id as cVoucherId,ccode as cVoucherCode,cast(null as nvarchar(1)) as cVoucherName,cast(null as nvarchar(1)) as cCardNum,cast(null as nvarchar(1)) as cMenu_Id,cast(null as nvarchar(1)) as cAuth_Id,cast(null as nvarchar(1)) as cSub_Id into tempdb..TEMP_STSearchTableNameList_" & sGUID & " from " & m_coni.From & " where 1=1 " & IIf(m_coni.where = "", "", " and " & m_coni.where)
    Exit Sub

Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set o_FilterObject = Nothing
    VouchList1.Dispose
    Set m_pagediv = Nothing
End Sub

Private Sub m_pagediv_BeforeGetCount()
    VouchList1.FillMode = FillOverwrite
End Sub

Private Sub m_pagediv_GetData(ByVal vltable As U8VouchList.VouchListTable)
    On Error GoTo Err_Handler

    'Dim recclass As New ADODB.Recordset

    VouchList1.SetVchLstRst vltable.DataRecordset  '

    'Set recclass = vltable.DataRecordset
    VouchList1.SetSumRst vltable.SumRecordset
    VouchList1.RecordCount = vltable.DataCount
    Me.VouchList1.Title = ListTitle
    Exit Sub

Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Sub

Private Sub m_pagediv_AfterGetData(rst As ADODB.Recordset, cnt As Long)
On Error GoTo Err_Handler

    Me.VouchList1.InitHead objColset.getColInfo()
    If Replace(strWhere, " ", "") = "(1=2)" Then
        VouchList1.SumStyle = vlSumNone '   vlRecordAndGridsum                     ' vlGridSum
    Else
        VouchList1.SumStyle = vlRecordAndGridsum
    End If
       Me.VouchList1.Title = ListTitle
    Exit Sub
   
Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
    
End Sub
'每个窗体都需要这个方法。Cancel与UnloadMode的参数的含义与QueryUnload的参数相同。
Public Sub ExitForm(Cancel As Integer, UnloadMode As Integer)
    
    Dim sXML As String

    sXML = VouchList1.GetColumnWidthXML()

    If sXML <> "" Then

        If MsgBox(GetResString("U8.AA.U8VouchList.00103"), vbYesNo, GetString("U8.DZ.JA.Res030")) = vbYes Then
            If Not objColset Is Nothing Then objColset.SaveColWidth2 sXML
        End If
    End If
    
    Unload Me
    Cancel = m_Cancel
    UnloadMode = m_UnloadMode

    strWhere = ""

End Sub

Private Sub PagedivCtl1_BeforeSendCommand(cmdType As U8VouchList.UFCommandType, pageSize As Long, PageCurrent As Long)
    Me.VouchList1.SetVchLstRst Nothing
    Me.VouchList1.FillMode = FillOverwrite
End Sub
'打印，预览，输出
Private Sub ListPrint(flag As Integer)
    On Error Resume Next

    If flag = 1 Then
        VouchList1.VchLstPrint
    ElseIf flag = 2 Then
        VouchList1.VchLstPreview
    Else
        VouchList1.VchLstPrintToFile
    End If

End Sub

Private Sub ExecSelectAll(flag As Boolean)

'全选
    If flag = True Then
        VouchList1.AllSelect
        '全消
    Else
        VouchList1.AllNone
    End If

End Sub

'反选
Private Sub ReverseSelection()
    Dim i As Long
    For i = 1 To VouchList1.rows - 1
        If VouchList1.TextMatrix(i, VouchList1.GetColIndex("selcol")) = "Y" Then
            VouchList1.TextMatrix(i, VouchList1.GetColIndex("selcol")) = ""
        Else
            VouchList1.TextMatrix(i, VouchList1.GetColIndex("selcol")) = "Y"
        End If
    Next

End Sub


Private Sub ExecLink()
    If VouchList1.rows < 2 Then Exit Sub
    If VouchList1.row < 1 Then VouchList1.row = 1

    Dim ClsIn As New ClsInterFace
    
    sID = VouchList1.TextMatrix(VouchList1.row, VouchList1.GetColIndex(HeadPKFld))
    sID = sID & Chr(9) & "tempdb..TEMP_STSearchTableNameList_" & sGUID

    If g_oBusiness Is Nothing Then
        Set g_oBusiness = goBusiness
    End If

    ClsIn.ILoginable_Login
    Call ClsIn.ILoginable_CallFunction(cMenuId, cMenuName, cAuthId, sID)
End Sub

Private Sub ExecRefresh()
'初始化列表
    Call InitList


    '初始化分页控件
    Call InitConi(strWhere)

    '填充分页控件数据，加载列表数据
    Call FillList

End Sub

'读取时间戳，并与旧时间戳比较
Private Function ExecFunCompareUfts(Optional strcode As String) As Boolean

'读取时间戳
    TimeStamp = GetTimeStamp(g_Conn, MainTable, lngVoucherID)

    If TimeStamp = RecordDeleted Then
        '        MsgBox "单据(" & strcode & ")已被其他用户删除,不可修改", vbInformation, getstring("U8.DZ.JA.Res030")
        ExecFunCompareUfts = False
        Exit Function
    ElseIf TimeStamp = RecordError Then
        '         MsgBox "单据(" & strcode & ")数据出现错误,请刷新", vbInformation, getstring("U8.DZ.JA.Res030")
        ExecFunCompareUfts = False
        Exit Function
    ElseIf OldTimeStamp <> TimeStamp Then
        '        MsgBox "单据(" & strcode & ")已被其他用户修改,请刷新", vbInformation, getstring("U8.DZ.JA.Res030")
        ExecFunCompareUfts = False
        Exit Function
    Else
        OldTimeStamp = TimeStamp
        ExecFunCompareUfts = True
    End If
End Function

'审核
Private Sub ExecConfirm(flag As Boolean, bWorkflow As Boolean)

    On Error GoTo Err_Handler

    Dim sql As String
    Dim i As Integer
    Dim cHandlervalue As String
    Dim dVeriDatevalue As String
    Dim iStatusvalue As String

    '工作流
    Dim Action As String
    Dim State As Integer
    Dim strAuditOpinion As String
    Dim primBizData As String
    Dim eleResult As IXMLDOMElement
    Dim domResult As New DOMDocument
    Dim AuditServiceProxy As Object
    Dim auditResult As String
    Dim oldcode As String
    Dim oldCode4PushOtherOut As String
    Dim bPushOut As Boolean
    Dim bSuccess As Boolean
    bPushOut = getAccinformation("ST", "bautolendout", g_Conn)
    
    '调用批审之前，调用该对象的填写审批意见服务
    Dim calledCtx As Object
    Set calledCtx = CreateObject("UFSoft.U8.Framework.LoginContext.CalledContext")
    '

    FrmProgress.Show
  
    FrmProgress.Label1.Caption = GetString("U8.DZ.JA.Res340") & IIf(flag = True, GetString("U8.DZ.JA.Res350"), GetString("U8.DZ.JA.Res360")) & GetString("U8.DZ.JA.Res370")
    FrmProgress.ProgressBar1.Max = VouchList1.rows - 1



    '审核
    If flag = True Then
        cHandlervalue = goLogin.cUserId
        dVeriDatevalue = Date
        iStatusvalue = 2
        '弃审
    Else
        cHandlervalue = ""
        dVeriDatevalue = "null"
        iStatusvalue = 1
    End If

    If bWorkflow Then        '需要走工作流
        Set AuditServiceProxy = CreateObject("UFIDA.U8.Audit.ServiceProxy.AuditServiceProxy")

        calledCtx.SubId = goLogin.cSub_Id
        calledCtx.TaskId = goLogin.TaskId
        calledCtx.token = goLogin.userToken
        If flag Then        '审核
            If AuditServiceProxy.ShowAuditSimpleUI(calledCtx, Action, State, strAuditOpinion) = False Then
                Exit Sub
            End If
        End If
        If Not flag Then    '弃审
            If Not AuditServiceProxy.ShowAuditAbandonUI(calledCtx, State, strAuditOpinion) Then
                Exit Sub
            End If
            primBizData = ""
        End If
    End If

   oldcode = ""
   oldCode4PushOtherOut = ""
    For i = 1 To VouchList1.rows - 1                 '多行审核
    
        bSuccess = True
        If VouchList1.TextMatrix(i, VouchList1.GetColIndex("selcol")) = "Y" Then
            If oldcode = VouchList1.TextMatrix(i, VouchList1.GetColIndex("cCODE")) Then
               GoTo DoNext
            End If
             oldcode = VouchList1.TextMatrix(i, VouchList1.GetColIndex("cCODE"))
        
            '关闭的跳出循环
            If VouchList1.TextMatrix(i, VouchList1.GetColIndex("CloseUser")) <> "" Then
                GoTo DoNext
            End If
            'enum by modify
            '-----------------------------------------------------------------------------------------
            If flag = False Then
                If VouchList1.TextMatrix(i, VouchList1.GetColIndex("cCreateType")) = "转换单据" Then
                    '               MsgBox getstring("U8.DZ.JA.Res240"), vbInformation, getstring("U8.DZ.JA.Res030")
                    '               Exit Sub
                    GoTo DoNext
                End If

                If VouchList1.TextMatrix(i, VouchList1.GetColIndex("cCreateType")) = "期初单据" Then
                    If VoucherIsCreate2(VouchList1.TextMatrix(i, VouchList1.GetColIndex("ID"))) Then
                        '                   MsgBox getstring("U8.DZ.JA.Res250"), vbInformation, getstring("U8.DZ.JA.Res030")
                        '                   Exit Sub
                        GoTo DoNext
                    End If
                Else
                    If VoucherIsCreate(VouchList1.TextMatrix(i, VouchList1.GetColIndex("ID"))) Then
                        '                   MsgBox getstring("U8.DZ.JA.Res250"), vbInformation, getstring("U8.DZ.JA.Res030")
                        '                   Exit Sub
                        GoTo DoNext
                    End If
                End If
            End If

            lngVoucherID = VouchList1.TextMatrix(i, VouchList1.GetColIndex("ID"))
            OldTimeStamp = VouchList1.TextMatrix(i, VouchList1.GetColIndex("ufts"))
'            If ExecFunCompareUfts(VouchList1.TextMatrix(i, VouchList1.GetColIndex("cCODE"))) = False Then GoTo DoNext  'Exit Sub  'liwqa 注释
            '-----------------------------------------------------------------------------------------

            If VouchList1.TextMatrix(i, VouchList1.GetColIndex("iswfcontrolled")) = "1" Then
                If flag Then    '审核
                    If VouchList1.TextMatrix(i, VouchList1.GetColIndex("iVerifyState")) = "1" Then           '已提交
                        primBizData = "         <KeySet>" & Chr(13)
                        primBizData = primBizData & "                   <Key name=""VoucherId"" value=""" & VouchList1.TextMatrix(i, VouchList1.GetColIndex(HeadPKFld)) & """/>" & Chr(13)
                        primBizData = primBizData & "                   <Key name=""VoucherType"" value=""" & gstrCardNumberlist & """/>" & Chr(13)
                        primBizData = primBizData & "                   <Key name=""VoucherCode"" value=""" & VouchList1.TextMatrix(i, VouchList1.GetColIndex("ccode")) & """/>" & Chr(13)
                        primBizData = primBizData & "                   <Key name=""Ufts"" value=""" & VouchList1.TextMatrix(i, VouchList1.GetColIndex("ufts")) & """/>" & Chr(13)
                        primBizData = primBizData & "                   <Key name=""AuditAuthId"" value=""""/>" & Chr(13)
                        primBizData = primBizData & "                   <Key name=""AbandonAuthId"" value=""""/>" & Chr(13)
                        primBizData = primBizData & "         </KeySet>" & Chr(13)
                        Call AuditServiceProxy.Audit(primBizData, Action, State, strAuditOpinion, calledCtx, auditResult)
                        domResult.loadXML auditResult
                        For Each eleResult In domResult.documentElement.selectNodes("//Result")
                            If CBool(eleResult.getAttribute("AuditResult")) = True Then
                                VouchList1.TextMatrix(i, VouchList1.GetColIndex("selcol")) = ""
                                 strMsg = strMsg & GetStringPara(IIf(flag = True, "U8.DZ.JA.Res430", "U8.DZ.JA.Res480"), VouchList1.TextMatrix(i, VouchList1.GetColIndex("cCODE"))) & vbCrLf
                            Else
                                bSuccess = False
                                strMsg = strMsg & eleResult.getAttribute("errMsg") ', vbInformation, GetString("U8.DZ.JA.Res030")
                            End If
                        Next
                        ''                                                    处理工作流 (没有提交)
                    ElseIf VouchList1.TextMatrix(i, VouchList1.GetColIndex("iVerifyState")) = "0" Or VouchList1.TextMatrix(i, VouchList1.GetColIndex("iVerifyState")) = "" Then
                        bSuccess = False
                        strMsg = strMsg & GetStringPara("U8.pu.prjpu860.04715", VouchList1.TextMatrix(i, VouchList1.GetColIndex("ccode"))) ', vbInformation, GetString("U8.DZ.JA.Res030")    '单据{0}没有提交，请首先提交！
                    End If
                Else
                    If VouchList1.TextMatrix(i, VouchList1.GetColIndex("iVerifyState")) <> "0" Then
                        primBizData = "         <KeySet>" & Chr(13)
                        primBizData = primBizData & "                   <Key name=""VoucherId"" value=""" & VouchList1.TextMatrix(i, VouchList1.GetColIndex(HeadPKFld)) & """/>" & Chr(13)
                        primBizData = primBizData & "                   <Key name=""VoucherType"" value=""" & gstrCardNumberlist & """/>" & Chr(13)
                        primBizData = primBizData & "                   <Key name=""VoucherCode"" value=""" & VouchList1.TextMatrix(i, VouchList1.GetColIndex("ccode")) & """/>" & Chr(13)
                        primBizData = primBizData & "                   <Key name=""AuditAuthId"" value=""""/>" & Chr(13)
                        primBizData = primBizData & "                   <Key name=""Ufts"" value=""" & VouchList1.TextMatrix(i, VouchList1.GetColIndex("ufts")) & """/>" & Chr(13)
                        primBizData = primBizData & "                   <Key name=""AbandonAuthId"" value=""""/>" & Chr(13)
                        primBizData = primBizData & "         </KeySet>" & Chr(13)
                        Call AuditServiceProxy.Abandon(primBizData, strAuditOpinion, State, calledCtx, auditResult)
                        domResult.loadXML auditResult
                        For Each eleResult In domResult.documentElement.selectNodes("//Result")
                            If CBool(eleResult.getAttribute("AuditResult")) = True Then
                                VouchList1.TextMatrix(i, VouchList1.GetColIndex("selcol")) = ""
                                 strMsg = strMsg & GetStringPara(IIf(flag = True, "U8.DZ.JA.Res430", "U8.DZ.JA.Res480"), VouchList1.TextMatrix(i, VouchList1.GetColIndex("cCODE"))) & vbCrLf
                            Else
                                bSuccess = False
                                strMsg = strMsg & eleResult.getAttribute("errMsg") ', vbInformation, GetString("U8.DZ.JA.Res030")
                            End If
                        Next
                    ElseIf VchLst.TextMatrix(i, VchLst.GridColIndex("iswfcontrolled")) = "1" And VchLst.TextMatrix(i, VchLst.GridColIndex("iVerifyState")) = "0" Then
                        bSuccess = False
                        strMsg = strMsg & GetStringPara("U8.pu.prjpu860.04715", VouchList1.TextMatrix(i, VouchList1.GetColIndex("ccode"))) ', vbInformation, GetString("U8.DZ.JA.Res030")    '单据{0}没有提交，请首先提交！
                    End If

                End If    '匹配审核

            Else
                If (flag = True And VouchList1.TextMatrix(i, VouchList1.GetColIndex(StrcHandler)) = "") Or (flag = False And VouchList1.TextMatrix(i, VouchList1.GetColIndex(StrcHandler)) <> "") Then

                    FrmProgress.ProgressBar1.Value = i
                    DoEvents

                    sql = "Update " & MainTable & " set " & StrcHandler & "='" & cHandlervalue & "' ," & StrdVeriDate & "=" & IIf(flag = True, "'" & dVeriDatevalue & "'", "null") & " , " & StriStatus & "=" & iStatusvalue & vbCrLf & _
                            " where convert(nchar,convert(money,ufts),2)  = " & OldTimeStamp & " and " & HeadPKFld & "=" & VouchList1.TextMatrix(i, VouchList1.GetColIndex(HeadPKFld))
                    Dim lngAct As Long
                    lngAct = 0
                    gConn.Execute sql, lngAct
                    '添加并发控制 by liwqa
                    If lngAct = 0 Then
                        bSuccess = False
                        If InStr(strMsg, VouchList1.TextMatrix(i, VouchList1.GetColIndex("cCODE"))) <= 0 Then      '已经做过提示的不在提示
                            strMsg = strMsg & VouchList1.TextMatrix(i, VouchList1.GetColIndex("cCODE")) & GetString("U8.DZ.JA.Res290") & vbCrLf
                        End If
                    Else
                         strMsg = strMsg & GetStringPara(IIf(flag = True, "U8.DZ.JA.Res430", "U8.DZ.JA.Res480"), VouchList1.TextMatrix(i, VouchList1.GetColIndex("cCODE"))) & vbCrLf
                         '业务通知
                        NotifySrvSend "HYJCGH001", "HYJCGH001" & IIf(flag, ".Audit", ".UnAudit"), VouchList1.TextMatrix(i, VouchList1.GetColIndex(HeadPKFld)), goLogin
                    End If
                Else
                    bSuccess = False
                End If
            End If    '匹配iswfcontrolled=“1”
            
            '审核成功的才推单(只有非工作流单据审核才推单)
            If Not bWorkflow And flag And bSuccess _
                And VouchList1.TextMatrix(i, VouchList1.GetColIndex(StrcHandler)) = "" _
                And VouchList1.TextMatrix(i, VouchList1.GetColIndex("cCreateType")) <> "期初单据" Then
                '借出借用单审核自动生成其他出库单
                If LCase(bPushOut) = "true" Then
                    If VouchList1.TextMatrix(i, VouchList1.GetColIndex(HeadPKFld)) <> "" And _
                        oldCode4PushOtherOut <> VouchList1.TextMatrix(i, VouchList1.GetColIndex("cCODE")) Then
                        oldCode4PushOtherOut = VouchList1.TextMatrix(i, VouchList1.GetColIndex("cCODE"))
                        strMsg = strMsg & ExecPushOtherOut(VouchList1.TextMatrix(i, VouchList1.GetColIndex(HeadPKFld))) & vbCrLf
                    End If
                End If
            End If
        End If     '匹配="Y"
DoNext:
    Next i

    Unload FrmProgress

    '    If flag = True Then
    '        MsgBox "批审完成", vbInformation, getstring("U8.DZ.JA.Res030")
    '    Else
    '        MsgBox "批弃完成", vbInformation, getstring("U8.DZ.JA.Res030")
    '    End If

    Load FrmMsgBox
    FrmMsgBox.Text1 = strMsg
    FrmMsgBox.Show 1

    Exit Sub

Err_Handler:
    Unload FrmProgress
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")

End Sub

Private Sub ExecOpenClose(flag As Boolean)

    On Error GoTo Err_Handler

    Dim sql As String
    Dim i As Integer
    Dim CloseUservalue As String
    Dim dCloseDatevalue As String
    Dim iStatusvalue As String

    FrmProgress.Show
    FrmProgress.Label1.Caption = GetString("U8.DZ.JA.Res340") & IIf(flag = True, GetString("U8.DZ.JA.Res390"), GetString("U8.DZ.JA.Res380")) & GetString("U8.DZ.JA.Res370")
    FrmProgress.ProgressBar1.Max = VouchList1.rows - 1



    '关闭
    If flag = True Then
        CloseUservalue = goLogin.cUserId
        dCloseDatevalue = Date
        iStatusvalue = 4    '关闭
        '打开
    ElseIf flag = False Then
        CloseUservalue = ""
        dCloseDatevalue = "null"
        If VouchList1.TextMatrix(i, VouchList1.GetColIndex(StrIntoUser)) <> "" Then
            iStatusvalue = 3    '推单
        ElseIf VouchList1.TextMatrix(i, VouchList1.GetColIndex(StrcHandler)) <> "" Then
            iStatusvalue = 2    '审核
        Else
            iStatusvalue = 1    '新建
        End If
    End If


    For i = 1 To VouchList1.rows - 1
        If VouchList1.TextMatrix(i, VouchList1.GetColIndex("selcol")) = "Y" _
           And ((flag = True And VouchList1.TextMatrix(i, VouchList1.GetColIndex(StrCloseUser)) = "") _
                Or (flag = False And VouchList1.TextMatrix(i, VouchList1.GetColIndex(StrCloseUser)) <> "")) Then


            FrmProgress.ProgressBar1.Value = i
            DoEvents


            sql = "Update " & MainTable & " set " & StrCloseUser & "='" & CloseUservalue & "' ," & _
                  StrdCloseDate & "=" & IIf(flag = True, "'" & dCloseDatevalue & "'", "null") & " , " & _
                  StriStatus & "=" & IIf(flag = False And VouchList1.TextMatrix(i, VouchList1.GetColIndex(StrcHandler)) <> "", 2, iStatusvalue) & _
                " where " & HeadPKFld & "=" & VouchList1.TextMatrix(i, VouchList1.GetColIndex(HeadPKFld))

            gConn.Execute sql

        End If
    Next i

    Unload FrmProgress

    If flag = True Then
        MsgBox GetString("U8.DZ.JA.Res400"), vbInformation, GetString("U8.DZ.JA.Res030")
    Else
        MsgBox GetString("U8.DZ.JA.Res410"), vbInformation, GetString("U8.DZ.JA.Res030")
    End If

    Exit Sub

Err_Handler:
    Unload FrmProgress
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")


End Sub

Private Sub ExecFilter()

    If GetFilterList(goLogin, o_FilterObject) = False Then Exit Sub

    '初始化列表
    Call InitList


    '初始化分页控件
    Call InitConi(strWhere)

    '填充分页控件数据，加载列表数据
    Call FillList

End Sub


Private Sub SetCatalog()
    On Error GoTo 0

    Dim sXML As String

    If objColset.setCol <> enmCancel Then
        sXML = objColset.getColInfo
        VouchList1.InitHead sXML
        
        '11.1合并显示
        Call SetToolbarForColumn '栏目设置后调用设置"合并显式"按钮状态
    End If
    
End Sub


Private Sub UFKeyHookCtrl1_ContainerKeyDown(KeyCode As Integer, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyF12
            Dim strFilterID As String
            Dim strSrcMenuID As String
'            Dim oMenuPub As UFIDA_U8_UI_PubFilterSolutionMenu.PubFilterSolutionMenu
'            Set oMenuPub = New UFIDA_U8_UI_PubFilterSolutionMenu.PubFilterSolutionMenu
'            Set oMenuPub.U8Login = g_oLogin
            strFilterID = "ST[__]项目发布单列表参照"
            strSrcMenuID = "FYSL020104"
'            oMenuPub.FilterID = strFilterID
'            oMenuPub.SourceMenuID = strSrcMenuID
'            Call oMenuPub.ShowForm
    End Select
End Sub

Private Sub UFToolbar1_OnCommand(ByVal enumType As UFToolBarCtrl.ENUM_MENU_OR_BUTTON, ByVal cButtonId As String, ByVal cMenuId As String)
    If enumType = enumButton Then
        Call ButtonClick(Toolbar1.Buttons(cButtonId).Key)
    Else
        Call ButtonClick(Toolbar1.Buttons(cButtonId).ButtonMenus(cMenuId).Key)
    End If
End Sub

Private Sub ButtonClick(strbuttonkey As String)
    Dim bFlowControl As Boolean                        '是否工作流控制
    Dim HasWfcontrolled As Boolean                              '选中行是否有不受工作流控制的记录，有则置1，没则置0；
    Dim i As Long, Selected As Boolean
    Dim errMsg As String
    Dim oDomHead As New DOMDocument
    Dim oDomBody As New DOMDocument

    Dim bReturn As Boolean
    Select Case strbuttonkey
        '打印
    Case sKey_Print
        If ZwTaskExec(goLogin, AuthPrint, 1) = False Then Exit Sub
        ListPrint (1)
        Call ZwTaskExec(goLogin, AuthPrint, 0)
        '预览
    Case sKey_Preview
        If ZwTaskExec(goLogin, AuthPrint, 1) = False Then Exit Sub
        ListPrint (2)
        Call ZwTaskExec(goLogin, AuthPrint, 0)
        '输出
    Case sKey_Output
        If ZwTaskExec(goLogin, AuthOut, 1) = False Then Exit Sub
        ListPrint (3)
        Call ZwTaskExec(goLogin, AuthOut, 0)


        '定位
    Case sKey_Locate
        VouchList1.locate True

        '过滤
    Case strKfilter
        ExecFilter

        '栏目
    Case sKey_Column
        SetCatalog
        '刷新
        ExecRefresh

        '?批打
    Case sKey_Batchprint
        If VouchList1.rows <= 1 Then Exit Sub
        '全选
    Case strKSelectAll
        ExecSelectAll (True)

        '反选
    Case sKey_ReverseSelection
        ReverseSelection

        '全消
    Case strKUnSelectAll
        ExecSelectAll (False)

        '单据
    Case sKey_Link
        If ZwTaskExec(goLogin, AuthBrowseLink, 1) = False Then Exit Sub
        ExecLink
        Call ZwTaskExec(goLogin, AuthBrowseLink, 0)
    
        '审核
    Case sKey_Confirm
        '            If ZwTaskExec(goLogin, "ST02JC020105", 1) = False Then Exit Sub


        If getIsWfControl(goLogin, gConn, errMsg, gstrCardNumberlist) Then          '工作流控制
            bFlowControl = True
        Else
            bFlowControl = False
        End If

        'AllOutWfcontrolled 标示全部不受工作流控制
        HasWfcontrolled = False
        strMsg = ""
        oldccode = ""
        'enum by modify
        For i = 1 To Me.VouchList1.rows - 1
            If Me.VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("selcol")) = "Y" Then
                If oldccode <> VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE")) Then
                    oldccode = VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE"))
                    strStatus3 = CheckVoucherStatus(VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("ID")), VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCreateType")))

                     If CheckUserAuth(gConn.ConnectionString, goLogin.cUserId, VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cMaker")), "V") = False Then
                        ReDim varArgs(0)
                        varArgs(0) = VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE"))
                        strMsg = strMsg & GetStringPara("U8.pu.VoucherCommon.00115", varArgs(0)) & vbCrLf
                        
                    Else
                        If strStatus3 = "生单" Or strStatus3 = "审核" Then
                        
                            ReDim varArgs(0)
                            varArgs(0) = VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE"))
                            strMsg = strMsg & GetStringPara("U8.DZ.JA.Res420", varArgs(0)) & vbCrLf
                           
                            '                        strMsg = strMsg & "单据 " & VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE")) & " 已经审核！" & vbCrLf
                        ElseIf strStatus3 = "新建" Then
                            ReDim varArgs(0)
                            varArgs(0) = VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE"))
'                            strMsg = strMsg & GetStringPara("U8.DZ.JA.Res430", varArgs(0)) & vbCrLf    'by liwq 注释，审核成功不能再这里报出
                            'strMsg = strMsg & "单据 " & VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE")) & " 审核成功！" & vbCrLf
                        ElseIf strStatus3 = "关闭" Then
                            ReDim varArgs(0)
                            varArgs(0) = VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE"))
                            strMsg = strMsg & GetStringPara("U8.DZ.JA.Res445", varArgs(0)) & vbCrLf
                            'strMsg = strMsg & "单据 " & VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE")) & " 已关闭！" & vbCrLf
                        Else
                            ReDim varArgs(0)
                            varArgs(0) = VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE"))
                            strMsg = strMsg & GetStringPara("U8.DZ.JA.Res440", varArgs(0)) & vbCrLf
                            'strMsg = strMsg & "单据 " & VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE")) & " 不存在！" & vbCrLf
                        End If
                    End If
                End If

                Selected = True
                If Me.VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("iswfcontrolled")) = "1" And Me.VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("iVerifyState")) = "1" Then
                    HasWfcontrolled = True
                    Exit For
                End If
            End If
        Next


        If Selected Then
            If ZwTaskExec(goLogin, AuthVerify, 1) = False Then Exit Sub
            ExecConfirm True, HasWfcontrolled
            ExecRefresh
            Call ZwTaskExec(goLogin, AuthVerify, 0)
        Else
            MsgBox GetString("U8.DZ.JA.Res450"), vbInformation, GetString("U8.DZ.JA.Res030")
        End If

        '            Call ZwTaskExec(goLogin, "ST02JC020105", 0) 'ST02JC020105

        '弃审
    Case sKey_Cancelconfirm
        '            If ZwTaskExec(goLogin, "ST02JC020106", 1) = False Then Exit Sub

        If getIsWfControl(goLogin, gConn, errMsg, gstrCardNumberlist) Then          '工作流控制
            bFlowControl = True
        Else
            bFlowControl = False
        End If

        strMsg = ""
        oldccode = ""
        AllOutWfcontrolled = True
        For i = 1 To Me.VouchList1.rows - 1
            If Me.VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("selcol")) = "Y" Then

                If oldccode <> VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE")) Then
                    oldccode = VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE"))
                    strStatus3 = CheckVoucherStatus(VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("ID")), VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCreateType")))
                    
                     If CheckUserAuth(gConn.ConnectionString, goLogin.cUserId, VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cMaker")), "U") = False Then
                        ReDim varArgs(0)
                        varArgs(0) = VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE"))
                        strMsg = strMsg & GetStringPara("U8.pu.VoucherCommon.00129", varArgs(0)) & vbCrLf
                        
                    Else
                        'enum by modify
                        If strStatus3 = "新建" Then
                            ReDim varArgs(0)
                            varArgs(0) = VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE"))
                            strMsg = strMsg & GetStringPara("U8.DZ.JA.Res460", varArgs(0)) & vbCrLf
                            'strMsg = strMsg & "单据 " & VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE")) & " 没有审核！" & vbCrLf
                        ElseIf strStatus3 = "生单" Then
                        If VouchList1.TextMatrix(i, VouchList1.GetColIndex("cCreateType")) = "转换单据" Then
                         ReDim varArgs(0)
                            varArgs(0) = VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE"))
                            strMsg = strMsg & GetStringPara("U8.DZ.JA.Res2030", varArgs(0)) & vbCrLf
                          Else
                            ReDim varArgs(0)
                            varArgs(0) = VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE"))
                            strMsg = strMsg & GetStringPara("U8.DZ.JA.Res470", varArgs(0)) & vbCrLf
                           End If
                            'strMsg = strMsg & "单据 " & VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE")) & " 已生单，不能弃审！" & vbCrLf
                        ElseIf strStatus3 = "审核" Then
                            ReDim varArgs(0)
                            varArgs(0) = VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE"))
'                            strMsg = strMsg & GetStringPara("U8.DZ.JA.Res480", varArgs(0)) & vbCrLf    'by liwqa 注释弃审成功放到执行语句后面提示
                            'strMsg = strMsg & "单据 " & VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE")) & " 弃审成功！" & vbCrLf
                        ElseIf strStatus3 = "关闭" Then
                            ReDim varArgs(0)
                            varArgs(0) = VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE"))
                            strMsg = strMsg & GetStringPara("U8.DZ.JA.Res445", varArgs(0)) & vbCrLf
                            'strMsg = strMsg & "单据 " & VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE")) & " 已关闭！" & vbCrLf
                        Else
                            ReDim varArgs(0)
                            varArgs(0) = VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE"))
                            strMsg = strMsg & GetStringPara("U8.DZ.JA.Res440", varArgs(0)) & vbCrLf
                            'strMsg = strMsg & "单据 " & VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE")) & " 不存在！" & vbCrLf
                        End If
                    End If
                End If

                Selected = True
                If Me.VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("iswfcontrolled")) = "1" Then
                    AllOutWfcontrolled = False
                    Exit For
                End If
            End If
        Next

        If Selected Then
            If ZwTaskExec(goLogin, AuthUnVerify, 1) = False Then Exit Sub
            ExecConfirm False, Not AllOutWfcontrolled
            ExecRefresh
            Call ZwTaskExec(goLogin, AuthUnVerify, 0)
        Else
            MsgBox GetString("U8.DZ.JA.Res450"), vbInformation, GetString("U8.DZ.JA.Res030")
        End If

        '            Call ZwTaskExec(goLogin, "ST02JC020106", 0)
        '刷新
    Case sKey_Refresh
        ExecRefresh

    Case gstrHelpCode
        SendKeys "{F1}"
    
    '批量归还
    Case sKey_BatchReturn
        strMsg = ""
        oldccode = ""
        For i = 1 To Me.VouchList1.rows - 1
            If Me.VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("selcol")) = "Y" Then
                If oldccode <> VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE")) Then
                    oldccode = VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE"))
                    Selected = True
                End If
            End If
        Next

        If Selected Then
            If ZwTaskExec(goLogin, AuthReturn, 1) = False Then Exit Sub
            ExecBatchReturn
            ExecRefresh
            Call ZwTaskExec(goLogin, AuthReturn, 0)
        Else
            MsgBox GetString("U8.DZ.JA.Res450"), vbInformation, GetString("U8.DZ.JA.Res030")
        End If
    End Select
End Sub

Private Sub VouchList1_BeforeSendCommand(cmdType As U8VouchList.UFCommandType, pageSize As Long, PageCurrent As Long)
   
    Me.VouchList1.SetVchLstRst Nothing
    Me.VouchList1.FillMode = FillOverwrite
    
    
End Sub



Private Sub VouchList1_CopySelect(bAuther As Boolean)
    bAuther = ZwTaskExec(goLogin, AuthOut, 1, True)
End Sub

Private Sub VouchList1_DblClick()
    ExecLink
End Sub

Private Sub VouchList1_FilterClick(fldsrv As Object)
    Set o_FilterObject = fldsrv
    ExecFilter
End Sub
'wangfb 11.0Toolbar迁移 2012-03-31
Private Function SetToolbarVisible()
    On Error Resume Next
    With Toolbar1
        '11.0本版暂不支持借出借入单列表工作流操作
'        If isWfcontrolled() Then
'            .Buttons("Submit").Visible = True
'            .Buttons("Resubmit").Visible = True
'            .Buttons("Unsubmit").Visible = True
'            .Buttons("ViewVerify").Visible = True
'        Else
            .Buttons("Submit").Visible = False
            .Buttons("Resubmit").Visible = False
            .Buttons("Unsubmit").Visible = False
            .Buttons("ViewVerify").Visible = False
'        End If
        
        .Buttons("Refresh").Visible = False
        .Buttons("SelectAll").Visible = False
        .Buttons("ReverseSelection").Visible = False
        .Buttons("UnSelectAll").Visible = False
        .Buttons("CreateVoucher").Visible = False
        '.Buttons("Cancelconfirm").Visible = False  '批弃 改为 弃审
        .Buttons("Open").Visible = False
        .Buttons("Close").Visible = False
        '.Buttons("Confirm").Visible = False        '批审 改为 审核
        .Buttons("Preview").Visible = False
        .Buttons("PrintBatch").Visible = False
        .Buttons("Help").Visible = False
        
         '暂不支持打印模板
        .Buttons("DesignPT").Visible = False
        .Buttons("PrintTemplate").Visible = False
        
        'toolbar 迁移 从Init_Toolbar里抽取
        .Buttons("Print").ButtonMenus(sKey_Batchprint).Visible = False   '暂不支持批打
    End With
    VBA.Err.Clear
End Function
Private Sub SetToolbarForColumn()
On Error GoTo ErrHandle
    If objColset Is Nothing Then Exit Sub
    If objColset.IsSupportTotalTableMerge = True Then '支持整表合并
        If objColset.TotalTableMerge = True Then '处于整表合并状态
             Me.Toolbar1.Buttons("MergeFullList").Value = 1
             Me.UFToolbar1.RefreshChecked
        Else
             Me.Toolbar1.Buttons("MergeFullList").Value = 0
             Me.UFToolbar1.RefreshChecked
        End If
    End If
    Exit Sub
ErrHandle:
    Err.Clear
End Sub

'批量归还
Private Sub ExecBatchReturn()

    On Error GoTo Err_Handler

    Dim sql As String
    Dim i As Integer
    Dim CloseUservalue As String
    Dim dCloseDatevalue As String
    Dim iStatusvalue As String
    Dim sMsg As String
    Dim lngID As Long
    Dim cCode As String
    Dim cOldCode As String
    Dim cCreateType As String

    FrmProgress.Show
    FrmProgress.Label1.Caption = GetString("U8.ST.Default.00122")
    FrmProgress.ProgressBar1.Max = VouchList1.rows - 1
    
    Dim IsBackWfcontrolled As Boolean '借出归还单是否工作流控制
    If getIsWfControl(goLogin, gConn, sMsg, "HYJCGH005") Then          '工作流控制
        IsBackWfcontrolled = True
    Else
        IsBackWfcontrolled = False
    End If
    
    For i = 1 To VouchList1.rows - 1
        If VouchList1.TextMatrix(i, VouchList1.GetColIndex("selcol")) = "Y" Then
        
            FrmProgress.ProgressBar1.Value = i
            DoEvents
            
            lngID = VouchList1.TextMatrix(i, VouchList1.GetColIndex(HeadPKFld))
            cCode = VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE"))
            cCreateType = VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCreateType"))
            sMsg = ""
            If cCode <> cOldCode Then
                cOldCode = cCode
                If CheckCanBack(lngID, cCode, cCreateType, sMsg) Then
                    If ExecReturn(VouchList1.TextMatrix(i, VouchList1.GetColIndex(HeadPKFld)), sMsg, IsBackWfcontrolled, VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("ufts"))) Then
                        strMsg = strMsg & GetStringPara("U8.ST.V870.00759", cCode) & vbCrLf '
                        'strMsg = strMsg & "单据 " & VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE")) & " 归还成功！" & vbCrLf
                        If sMsg <> "" Then strMsg = strMsg & sMsg & vbCrLf
                    Else
                        strMsg = strMsg & GetStringPara("U8.ST.V870.00760", cCode) & vbCrLf '
                        'strMsg = strMsg & "单据 " & VouchList1.TextMatrix(i, Me.VouchList1.GetColIndex("cCODE")) & " 归还失败！" & vbCrLf
                        If sMsg <> "" Then strMsg = strMsg & sMsg & vbCrLf
                    End If
                Else
                    If sMsg <> "" Then strMsg = strMsg & sMsg & vbCrLf
                End If
            End If
        End If
    Next i

    Unload FrmProgress
    
    Load FrmMsgBox
    FrmMsgBox.Text1 = strMsg
    FrmMsgBox.Show 1
    
    Exit Sub

Err_Handler:
    Unload FrmProgress
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Sub

