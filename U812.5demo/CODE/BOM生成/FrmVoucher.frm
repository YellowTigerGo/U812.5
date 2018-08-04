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
         Name            =   "宋体"
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
         Caption         =   "单据名称"
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
         Caption         =   "打印模版："
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
         Name            =   "宋体"
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
'模块功能说明：
'1、实现单据的基本按钮功能:打印、输出、预览、新增、修改、删除、增行、删行、复制行、复制、保存、放弃、提交、重新提交、撤销、审核、弃审、打开、关闭、刷新、翻页、定位、附件、合并显示（右键菜单）、批改（右键菜单）、记录定位（右键菜单）、参照生单、推单、取价（单行，整单）。
'2、预制单据模板：包含单据表头表体的基本栏目，如1-16表头自定义项，1-10存货自由项，1-16表体自定义项、项目编码、批号、保值期等
'3  单据号设置:     流水号、部门、仓库、单据日期、制单人
'4、表头各个栏目的参照、有效性校验
'5、表体各个栏目的参照、有效性校验
'6、保存时的有效性校验 , 如: 自由项结构性是否合法 , 保质期存货的保质期校验, 批次存货和跟踪型存货的批号与入库单号校验、
'7、并发处理
'8、数量、件数换算
'9、支持表体排序、多个显示、打印模板
'10、支持打印设置保存
'11、支持功能权限设置
'12、支持数据权限(记录级权限，暂不支持存货权限)
'13、支持取价功能，价格参照，金额计算
'14、支持审批流
'15、支持站点加密
'16、支持单据模板与列表数据精度设置
'创建时间：2008-11-21
'创建人：xuyan
'****************************************

Option Explicit

'单据服务组件
Private VchSrv As New clsVouchServer

Private m_Cancel, m_UnloadMode As Integer

' * 本地属性变量副本
Private m_strVT_ID As String
Private m_strVT_PRN_ID As String

'by liwqa Template
Private dicTemplate As New Dictionary                      '记录单据显示模板与combox的对应关系
Private dicTemplatePrint As New Dictionary
Private bInitForm As Boolean

Private objVoucherTemplate As New UFVoucherServer85.clsVoucherTemplate    'UFVoucherServer85.clsVoucherTemplate
Private objVoucher85 As UFVoucherD85.clsVoucher85
Private objBill As UFBillComponent.clsBillComponent
'功能权限
Private Const AuthBrowse = "FYSL02050301"                  '浏览
Private Const AuthAdd = "FYSL02050302"                     '新增
Private Const AuthModify = "FYSL02050303"                  '修改
Private Const AuthDelete = "FYSL02050304"                  '删除
Private Const AuthVerify = "FYSL02050305"                  '审核
Private Const AuthUnVerify = "FYSL02050306"                '弃审
Private Const AuthOpen = "FYSL02050310"                    '打开
Private Const AuthClose = "FYSL02050311"                   '关闭
Private Const AuthPrint = "FYSL02050307"                   '打印
Private Const AuthOut = "FYSL02050308"                     '输出
Private Const AuthProrefer = "FYSL02050312"                '参照
Private Const AuthProrefer1 = "FYSL02050313"                '参照

Dim wfcBack As Integer                                     '是否因工作流刷新界面标示
Dim inited As Boolean                                      '窗体初始化完成

'by zhangwchb 20110718 扩展字段
Dim sExtendField As String
Dim sExtendJoinSQL As String
Dim sExtendBodyField As String
Dim sExtendBodyJoinSQL As String
Dim oExtend As Object

Dim objRelation As Object

Private mobjSubServ As New ScmPublicSrv.clsAutoFill        '20110822 by zhangwchb

'向平台发送信息接口 20110812
Private WithEvents m_mht As UFPortalProxyMessage.IMessageHandler
Attribute m_mht.VB_VarHelpID = -1

Private isSavedOK As Boolean                               '单据是否保存成功 'by zhangwchb 20110829 提交保存

'在新增状态下，如果用户单击显示模板。此时判断表体是否有数据
Dim sPreVTID As Integer
Public bexitload As Boolean '单据模板数据权限控制
Private mdOldSelRow As Long
Private mdOldSelCol As Long
Public bAlter As Boolean

'删除
Private Sub ExecSubDelete()
    On Error GoTo Err_Handler
    Dim sMessage, sSource As String
    
    '12.0支持扩展自定义项
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
'    If Voucher.headerText("cCreateType") = "转换单据" Then
'        MsgBox GetString("U8.DZ.JA.Res040"), vbInformation, GetString("U8.DZ.JA.Res030")
'        Exit Sub
'    End If

    If MsgBox(GetString("U8.DZ.JA.Res050"), vbYesNo, GetString("U8.DZ.JA.Res030")) = vbNo Then
        Exit Sub
    End If

    'by liwqa 并发
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
'     '更新项目发布单累计数量和金额
'
'     strSql = " update HY_FYSL_Contract set totalappmoney= " & Null2Something(Voucher.headerText("totalappmoney"), 0) - Null2Something(Voucher.headerText("appprice"), 0) & " where  ccode= '" & Voucher.headerText("concode") & "'"
'       g_Conn.Execute strSql
'
'     '***********************************
'
    
    
      '附件删除
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
    '上一张单据的 ID
    lngVoucherID = GetTheLastID(login:=g_oLogin, _
            oConnection:=g_Conn, _
            sTable:=MainTable, _
            sField:=HeadPKFld & " asc", _
            sDataNumFormat:="0", _
            sWhereStatement:="" & IIf(PageCurrent > 1, HeadPKFld & " < " & lngVoucherID & "", ""))



    '**********************************************************
    '新增(增加/复制),删除更新全局变量
    '增加保存,重取lngvoucherid,修改保存不用重取
    '**********************************************************

    pageCount = pageCount - 1
    If PageCurrent > 1 Then PageCurrent = PageCurrent - 1


    '**********************************************************
    '新增,修改,删除 更新单据状态
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

    ' * 显示友好的错误信息
    Call ShowErrorInfo( _
            sHeaderMessage:=sMessage, _
            lMessageType:=vbExclamation, _
            lErrorLevel:=ufsELHeaderAndDescription _
            )

    ' * 定义错误数据源
    If Left(Err.Source, 3) = "***" Then
        sSource = Err.Source
    Else
        sSource = "***Sub DeleteVoucher of Form frmVoucher"
    End If

    ' * 调试模式时，显示调试窗口，用于跟踪错误
    If g_blnDEBUG_MODE Then
        Call ShowDebugForm( _
                bErrorMode:=True, _
                sProcedure:=sSource)
    End If
    GoTo Exit_Label
End Sub

'修改
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

    '设置单据编号是否可编辑
    Dim manual As Boolean                                  ' 是否完全手工编号
    Call SetVouchCodeEnable(manual)


    Call SetCtlStyle(Me, Voucher, Me.Toolbar, Me.UFToolbar, mOpStatus)

    'by zhangwchb 20110829 提交保存
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

    ' * 显示友好的错误信息
    Call ShowErrorInfo( _
            sHeaderMessage:=sMessage, _
            lMessageType:=vbExclamation, _
            lErrorLevel:=ufsELHeaderAndDescription _
            )

    ' * 定义错误数据源
    If Left(Err.Source, 3) = "***" Then
        sSource = Err.Source
    Else
        sSource = "***Sub DoEdit of Form frmVoucher"
    End If

    ' * 调试模式时，显示调试窗口，用于跟踪错误
    If g_blnDEBUG_MODE Then
        Call ShowDebugForm( _
                bErrorMode:=True, _
                sProcedure:=sSource)
    End If
    GoTo Exit_Label

End Sub

'保存
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
    
    'by liwqa 并发
    Dim m_bTrans     As Boolean

    Dim lngAct       As Long
    
    '12.0支持扩展自定义项
    Dim skeyfld      As String

    Dim skeySubfld   As String

    Dim objExtend    As Object

    Dim fldextends   As ADODB.Fields

    Set objExtend = CreateObject("VoucherExtendService.ClsExtendServer")

    Dim oDomBody       As New DOMDocument

    Dim oHeadElement   As IXMLDOMElement

    Dim oBodyElement() As IXMLDOMElement

    '    Me.Voucher.RemoveEmptyRow   '清除空行

    isSavedOK = False                                      'by zhangwchb 20110829 提交保存

CHECK:
    '    For i = 1 To Voucher.BodyRows
    '        '        Voucher.BodyRowIsEmpty i
    '        If Trim(Voucher.bodyText(i, "cinvcode")) = "" Then
    '            Voucher.row = i
    '            Voucher.DelLine
    '            GoTo CHECK
    '        End If
    '    Next

    '有效性校验
    If ExecFunSaveCheck(Voucher) = False Then Exit Sub

    '修改保存，需要比较时间戳，避免出现并发
    If Voucher.VoucherStatus = VSeEditMode Then
        If ExecFunCompareUfts = False Then Exit Sub
    End If

    '读取单据字段和数据
    ' 读取字段
    Set HeadData = GetHeadVouchData(g_Conn, Voucher, MainTable)
    '    Set BodyData = GetBodyVouchData(g_Conn, Voucher, DetailsTable)

    g_Conn.BeginTrans
    m_bTrans = True

    '新增保存,获得最大id,autoid
    '修改保存时,不需要更新id
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

    '更新单据号流水号
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
    
    '12.0表头扩展自定义项
    '    Set oHeadElement = oDomHead.selectSingleNode("//z:row")
    '    skeyfld = objExtend.getKeyid(g_Conn, MainTable & "_ExtraDefine")
    '    skeySubfld = objExtend.getKeyid(g_Conn, DetailsTable & "_ExtraDefine")

    '保存表头
    HeadData.Item(1).Item("dmDate").Value = g_oLogin.CurDate    'Now() '制单时间 Format(Now(), "YYYY-MM-DD HH:MM:SS")
    HeadData.Item(1).Item("iStatus").Value = 1             '状态

    If bAlter Then
        HeadData.Item(1).Item("iStatus").Value = 2             '状态
    End If
    
    Set DMO = New CDMO

    '新增
    If Voucher.VoucherStatus = VSeAddMode Then
        
        Set Rult = DMO.Insert(g_Conn, HeadData)
        
              
       
'       strSql = " update HY_FYSL_Contract set totalappmoney=isnull(totalappmoney,0)+" & Voucher.headerText("appprice") & " where  ccode= '" & Voucher.headerText("concode") & "'"
'       g_Conn.Execute strSql
        
        
        '        '12.0表头扩展自定义项-新增
        '        oHeadElement.setAttribute skeyfld, sID
        '        Set fldextends = objExtend.getVoucherExtendSaveInfo(g_Conn, MainTable & "_ExtraDefine")
        '        objExtend.SavebyInsert oHeadElement, MainTable & "_ExtraDefine", g_Conn, fldextends, , skeyfld
        '
        '修改
    Else
        '修改前做并发
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
        
        '12.0表头扩展自定义项-update
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
    '回写历史单据
      

     
     
    '*************************



   
    

    '表头附件保存
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
    '    If gcCreateType = "期初单据" Then
    '        strSql = "update " & DetailsTable & " set iqtyout=iquantity,iqtyout2=inum where id = " & sID    '" and cCreateType = '期初单据'"
    '        g_Conn.Execute strSql
    '    End If

    '检验单据编号
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
    '新增(增加/复制),删除更新全局变量
    '保存,重取lngvoucherid
    '**********************************************************

    If mOpStatus = ADD_MAIN Then
        lngVoucherID = sID
        pageCount = pageCount + 1
        PageCurrent = PageCurrent + 1
    End If

    '**********************************************************
    '新增,修改,删除 更新单据状态
    '**********************************************************
    mOpStatus = SHOW_ALL
    Voucher.VoucherStatus = VSNormalMode
    Call ExecSubRefresh
    Call SetCtlStyle(Me, Voucher, Me.Toolbar, Me.UFToolbar, mOpStatus)

    isSavedOK = True                                       'by zhangwchb 20110829 提交保存

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

    '单据模板未变化直接退出
    'If m_strVT_ID = Mid(ComTemplateShow.Text, 1, InStr(1, ComTemplateShow.Text, " ") - 1) Or bInitForm = True Then Exit Sub
    If CStr(m_strVT_ID) = CStr(ComTemplateShow.ItemData(ComTemplateShow.ListIndex)) Or bInitForm = True Then Exit Sub
    
    If Voucher.BodyRows > 0 And Voucher.VoucherStatus <> VSNormalMode Then
        MsgBox GetResString("U8.ST.USKCGLSQL.frmbill.03350"), vbOKOnly + vbExclamation, GetResString("U8.ST.USKCGLSQL.modmain.03048") '表体已经存在数据,不允许进行模版切换!
        ComTemplateShow.ListIndex = sPreVTID
        UFToolbar.RefreshCombobox
        Exit Sub
    End If
    
    Dim domHead As New DOMDocument, domBody As New DOMDocument
    Screen.MousePointer = vbHourglass
    m_strVT_ID = ComTemplateShow.ItemData(ComTemplateShow.ListIndex) 'Mid(ComTemplateShow.Text, 1, InStr(1, ComTemplateShow.Text, " ") - 1)
    Voucher.headerText("ivtid") = m_strVT_ID
    Me.Voucher.getVoucherDataXML domHead, domBody
    ' * 创建单据后台服务对象
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

'调整界面显示
Private Sub SetLayOut()

    Me.UFToolbar.Move 0, 0, Me.ScaleWidth

    '重要，单据自动调整大小
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

'快捷键操作 -chenliangc
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim strSql As String
'    Dim rs As New ADODB.Recordset
'
'
'    If Shift = vbAltMask Then
'        '快捷键ALT + E输出
'        If KeyCode = vbKeyE Then
'            If Me.Toolbar.Buttons(sKey_Output).Enabled = True Then
'                ButtonClick sKey_Output
'            End If
'            '弃审
'        ElseIf KeyCode = vbKeyU Then
'            If Me.Toolbar.Buttons(sKey_Cancelconfirm).Enabled = True Then
'                ButtonClick sKey_Cancelconfirm
'            End If
'            '撤销
'        ElseIf KeyCode = vbKeyJ Then
'            If Me.Toolbar.Buttons(sKey_Unsubmit).Enabled = True Then
'                ButtonClick sKey_Unsubmit
'            End If
'        ElseIf KeyCode = vbKeyPageUp Then                  '首页
'            If Me.Toolbar.Buttons(sKey_First).Enabled = True Then
'                ButtonClick sKey_First
'            End If
'        ElseIf KeyCode = vbKeyPageDown Then                '末页
'            If Me.Toolbar.Buttons(sKey_Last).Enabled = True Then
'                ButtonClick sKey_Last
'            End If
'        ElseIf KeyCode = vbKeyC Then                       '关闭
'            If Me.Toolbar.Buttons(sKey_Close).Enabled = True Then
'                ButtonClick sKey_Close
'            End If
'        ElseIf KeyCode = vbKeyO Then                       '打开
'            If Me.Toolbar.Buttons(sKey_Open).Enabled = True Then
'                ButtonClick sKey_Open
'            End If
'        End If
'    ElseIf Shift = vbCtrlMask Then
'        '快捷键Ctrl + P打印
'        If KeyCode = vbKeyP Then
'            If Me.Toolbar.Buttons(sKey_Print).Enabled = True Then
'                ButtonClick sKey_Print
'            End If
'            '快捷键Ctrl + W打印预览
'        ElseIf KeyCode = vbKeyW Then
'            If Me.Toolbar.Buttons(sKey_Preview).Enabled = True Then
'                ButtonClick sKey_Preview
'            End If
'            '快捷键Ctrl + F3过滤
'        ElseIf KeyCode = vbKeyF3 Then
'            If Me.Toolbar.Buttons(sKey_Locate).Enabled = True Then
'                ButtonClick sKey_Locate
'            End If
'            '快捷键Ctrl + G生单
'        ElseIf KeyCode = vbKeyG Then
'            If Me.Toolbar.Buttons(sKey_ReferVoucher).Enabled = True Then
'                ButtonClick sKey_ReferVoucher
'            End If
'            '保存新增
'        ElseIf KeyCode = vbKeyS Then
'            ' ButtonClick sKey_ReferVoucher
'            If Me.Toolbar.Buttons(sKey_Save).Enabled = True Then
'                ButtonClick sKey_Save
'                Call g_oBusiness.SelectToolbarButton(Toolbar.Buttons(sKey_Add))
'            End If
'
'            '放弃
'        ElseIf KeyCode = vbKeyZ Then
'            If Me.Toolbar.Buttons(sKey_Discard).Enabled = True Then
'                ButtonClick sKey_Discard
'            End If
'            '提交
'        ElseIf KeyCode = vbKeyJ Then
'            If Me.Toolbar.Buttons(sKey_Submit).Enabled = True Then
'                ButtonClick sKey_Submit
'            End If
'            '审核
'        ElseIf KeyCode = vbKeyU Then
'            If Me.Toolbar.Buttons(sKey_Confirm).Enabled = True Then
'                ButtonClick sKey_Confirm
'            End If
'
'            '复制
'        ElseIf KeyCode = vbKeyF5 Then
'            If Me.Toolbar.Buttons(sKey_Copy).Enabled = True Then
'                ButtonClick sKey_Copy
'            End If
'            '生单
'        ElseIf KeyCode = vbKeyG Then
'            'ButtonClick sKey_ReferVoucher
'        ElseIf KeyCode = vbKeyR Then                       '刷新
'            If Me.Toolbar.Buttons(sKey_Refresh).Enabled = True Then
'                ButtonClick sKey_Refresh
'            End If
'
'        ElseIf KeyCode = vbKeyD Then                       '删行
'            If Me.Toolbar.Buttons(sKey_Deleterecord).Enabled = True Then
'                ButtonClick sKey_Deleterecord
'
'            End If
'        ElseIf KeyCode = vbKeyA Then                       '删行
'            If Me.Toolbar.Buttons(sKey_Addrecord).Enabled = True Then
'                ButtonClick sKey_Addrecord
'            End If
'        ElseIf KeyCode = vbKeyF4 Then
'            Call ExitForm(0, 0)
'            '快捷键Ctrl+E或者Ctrl+B，自动指定批号,入库单号
'        ElseIf KeyCode = vbKeyE Or KeyCode = vbKeyB Or KeyCode = vbKeyQ Or KeyCode = vbKeyO Then
'            Call GetBatchInfoFun(Voucher, KeyCode, Shift)
'            KeyCode = 0
'        End If
'        ' End If
'    End If
'
'    Select Case KeyCode
'        Case vbKeyF1                                       '帮助
'            Call LoadHelpId(Me, "15030910")
'        Case vbKeyF5                                       '新增
'
'            If Me.Toolbar.Buttons(sKey_Add).Enabled = True Then
'                ButtonClick sKey_Add
'                'Call g_oBusiness.SelectToolbarButton(Toolbar.Buttons(sKey_Add))
'                '         ElseIf Me.Toolbar.Buttons(sKey_Add1).Enabled = True Then
'                '              Call g_oBusiness.SelectToolbarButton(Toolbar.Buttons(sKey_Add1))
'            End If
'            '  ButtonClick sKey_Add
'        Case vbKeyF6                                       '保存
'            If Me.Toolbar.Buttons(sKey_Save).Enabled = True Then
'                ButtonClick sKey_Save
'            End If
'        Case vbKeyF8                                       '修改
'            If Me.Toolbar.Buttons(sKey_Modify).Enabled = True Then
'                ButtonClick sKey_Modify
'            End If
'        Case vbKeyDelete                                   '删除
'            If Me.Toolbar.Buttons(sKey_Delete).Enabled = True Then
'                ButtonClick sKey_Delete
'            End If
'        Case vbKeyPageUp
'            If Me.Toolbar.Buttons(sKey_Previous).Enabled = True Then
'                ButtonClick sKey_Previous                  '上一页
'            End If
'        Case vbKeyPageDown
'            If Me.Toolbar.Buttons(sKey_Next).Enabled = True Then
'                ButtonClick sKey_Next                      '下一页
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

''加载帮助文件及ID
'Private Sub LoadHelpId(HelpID As String)
'    App.HelpFile = GetWinSysPath & "ufcomsql\行业插件\电子行业\应用说明\电子行业插件帮助.chm"
'    Me.HelpContextID = HelpID
'End Sub

Private Sub Form_Load()
    bInitForm = True
    '窗体初始化函数
    Call InitForm
    ' * 加载单据数据

    '隐藏frm菜单 modify by chenliangc 把Me.PopMenu1.Visible("retmenu") = False放到load事件中
'    Me.PopMenu1.Visible("retmenu") = False
'    '

    '    MsgBox "LoadData"
    If Not LoadData() Then
        '单据模板数据权限控制
        
''        MsgBox GetString("U8.DZ.JA.Res130"), vbExclamation, GetString("U8.DZ.JA.Res030")
        If bexitload = False Then
            ExitForm 0, 0
        End If
        bexitload = False
        Exit Sub
    End If
    '窗体名称
    Me.Caption = GetString("U8.DZ.JA.Res140")
    Call RegisterMessage                                   '20110812

    'wangfb 11.0Toobar迁移2012-03-20 设置可见性
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
'初始化窗体的常用属性
'如:各种全局变量的初始化,单据控件的初始化,工具栏的初始化,主表、字表
Private Sub InitForm()
    On Error GoTo Err_Handler
    Dim sSource As String
    '单据模版id,单据模板号必须在4位以内，否则不能加载附件

    '允许相应键盘事件
    Me.KeyPreview = True
    '加载时浏览状态
    mOpStatus = SHOW_ALL

    '过滤条件
    strwhereVou = ""

    Set VchSrv = New clsVouchServer




    '设置单据格式
    '    VchSrv.SetVouchStyle g_Conn, Voucher, gstrCardNumber

    pageCount = VchSrv.GetPageCount(g_Conn, gstrCardNumber, HeadPKFld, Replace(sAuth_ALL, "HY_FYSL_Payment.id", "V_HY_FYSL_Payment.id"))
    '由列表进入
    If CLng(sID) <> 0 Then
        UpdatePageCurrent (sID)
    Else
        PageCurrent = pageCount
    End If

    '不可编辑颜色
    Voucher.DisabledColor = &HE8E8E8

    '表体排序
    Voucher.ShowSorter = True

    '窗体名称
    Me.Caption = Me.Voucher.TitleCaption

    '初始化单据pco对象,单据参照时使用 USERPCO.VoucherCO
    Dim errStr As String
    Set clsbill = CreateObject("USERPCO.VoucherCO")        'New USERPCO.VoucherCO
    clsbill.IniLogin g_oLogin, errStr
    Set mologin = clsbill.login
    '*************wangfb 11.0Toobar迁移2012-03-20 start ****************
    Set UFToolbar.Business = g_oBusiness
    Call InitToolBar(mologin, "HY_FYSL_Payment001", Toolbar, UFToolbar, Me.Voucher)
    Call UFToolbar.InitExternalButton("Payment001", mologin.OldLogin)
    Call UFToolbar.SetFormInfo(Me.Voucher, Me)
    '在调用InitExternalButton方法后需要重新调用SetToolbar方法，否则自定义按钮加载不上
    UFToolbar.SetToolbar Toolbar
    
    '11.0Toolbar迁移初始化显示或打印模板按钮
    Call InitComTemplate
    Call InitComTemplatePRN
'    If IsObject(Toolbar.Buttons("PrintTemplate").Tag) Then
'        Set Toolbar.Buttons("PrintTemplate").Tag.Tag = Me.ComTemplatePRN
'    End If
'    If IsObject(Toolbar.Buttons("ShowTemplate").Tag) Then
'        Set Toolbar.Buttons("ShowTemplate").Tag.Tag = Me.ComTemplateShow
'    End If
    
    '工具栏初始化
    '11.0toolbar迁移，借入借出业务单据标准化之后，原来客开的按钮初始化
    'Call Init_Toolbar(Me.Toolbar)
   
    Call ChangeOneFormTbr(Me, Me.Toolbar, Me.UFToolbar)
'    Call SetWFControlBrns(g_oLogin, g_Conn, Me.Toolbar, Me.UFToolbar, gstrCardNumber)
 
    '*************wangfb 11.0Toobar迁移2012-03-20 end ****************
'
'    '初始化菜单,此处必须使用call方法
'    Call Me.PopMenu1.SubClassMenu(Me)

    '获取U8版本 -chenliangc
    gU8Version = GetU8Version(g_Conn)


Exit_Label:
    On Error GoTo 0
    Exit Sub



    '容错处理
Err_Handler:

    ' * 定义错误数据源
    If Left(Err.Source, 3) = "***" Then
        sSource = Err.Source
    Else
        sSource = "***Sub InitComTemplate of Form frmVoucher"
    End If

    ' * 抛出异常
    Err.Raise _
            Number:=Err.Number, _
            Source:=sSource, _
            Description:=Err.Description
End Sub
'初始化单据显示模板
Private Sub InitComTemplate()

    Dim sql As String
    Dim oRecordset As New Recordset
    
    Dim sMessage As String
    Dim i As Integer
    ' * 错误源
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

    '容错处理
Err_Handler:
    If Not oRecordset Is Nothing Then
        If oRecordset.State = adStateOpen Then _
                Call oRecordset.Close
    End If
    Set oRecordset = Nothing

    ' * 定义错误数据源
    If Left(Err.Source, 3) = "***" Then
        sSource = Err.Source
    Else
        sSource = "***Sub InitComTemplate of Form frmVoucher"
    End If

    ' * 抛出异常
    Err.Raise _
            Number:=Err.Number, _
            Source:=sSource, _
            Description:=Err.Description

End Sub
'初始化单据打印模板 'by liwqa Template
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

    '重置标题


    ' *******************************************************
    ' * 读取当前表单的模板ID (VT_ID) 值
    '
    '    Call LoadVTID

    ' *******************************************************
    ' * 创建单据后台服务对象
    '
    If objVoucherTemplate Is Nothing Then _
            Set objVoucherTemplate = _
            New UFVoucherServer85.clsVoucherTemplate



    '    ' 创建单据数据源对象
    Set oDataSource = CreateObject("IDataSource.DefaultDataSource")

    If oDataSource Is Nothing Then
        MsgBox GetString("U8.DZ.JA.Res160"), vbExclamation, GetString("U8.DZ.JA.Res030")
    End If

    Set oDataSource.SetLogin = g_oLogin

    Set Voucher.SetDataSource = oDataSource

    '请注意:SetTemplateData  必须放在 Set oDataSource.SetLogin = g_oLogin 之后, 即必须先给单据数据源初始化
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

'加载单据模版和数据
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


    '由列表进入
    If CLng(sID) <> 0 Then
        lngVoucherID = sID


        '直接进入单据
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
                '                If tmpLinkTbl = "" Then '单据联查 时 按钮状态控制 by zhangwchb 20110809
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
'    '如果单据数据中没有保存就取默认模板 by liwq
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

    Dim oRecordset As ADODB.Recordset                      '模版数据记录集
    Dim sAuth As String                                    '字段权限字符串
    Dim sNumber As String                                  '单据编号规则字符串
    Dim lColor1 As Long
    Dim lColor2 As Long

    'On Error GoTo Err_Handler

    ' *******************************************************
    ' * 得到单据模版数据,根据指定的单据类型和模版ID取得单据数据
    '
    Set oRecordset = objVoucherTemplate.GetTemplateData2( _
            conn:=g_Conn, _
            sBillName:=gstrCardNumber, _
            vTemplateID:=m_strVT_ID)


    ' *******************************************************
    ' * 取得指定操作员对单前单据的权限,以便进行权限控制
    ' *
    ' * 注:
    ' *     1)  每次换模版的时候需要应用一次
    sAuth = objVoucherTemplate.getAuthString( _
            ologin:=g_oLogin, _
            nID:=gstrCardNumber)


    ' *******************************************************
    ' * 取得 Rule 颜色
    '
    Call objVoucherTemplate.GetRuleColor( _
            strConn:=g_Conn, _
            clrDisable:=lColor1, _
            clrNeed:=lColor2)


    ' *******************************************************
    ' * 设置单据控件不可见
    '
    Voucher.Visible = False


    ' *******************************************************
    ' * 注:
    ' *     1)  SetVoucherAuth 方法必需在 SetTemplateData 方
    ' *         法前使用
    '
    Call Voucher.SetVoucherAuth(sAuth)
    Call Voucher.SetRuleColor(lColor1, lColor2)

    '    FormatVouchList oRecordset    ' 处理单据模板精度问题


    Call Voucher.SetTemplateData(oRecordset)

    Voucher.Visible = True

    'by liwq
    If gcCreateType = "期初单据" Then
        Me.LabTitle = GetString("U8.DZ.JA.Res170")
    Else
        Me.LabTitle = Voucher.TitleCaption                 ' GetString("U8.DZ.JA.Res140")
    End If

    Me.UFFrmCaptionMgr.Caption = Voucher.TitleCaption
    'Toolbar迁移 wangfb 2012-03-30 title移到Voucher上
    'Voucher.TitleCaption = ""
    Me.LabTitle = ""


    ' *******************************************************
    ' * 设置单据编号 Rule
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

'by zhangwchb 20110718 扩展字段
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

    ' * 错误源
    Dim sSource As String
    On Error GoTo Err_Handler
    
    numappprice = 0

    Screen.MousePointer = vbHourglass

    oRecordset.CursorLocation = adUseClient

    'by zhangwchb 20110718 扩展字段
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

    ' * 转换成 XML 数据格式
    oRecordset.Save oDomHead, adPersistXML

    'by zhangwchb 20110718 扩展字段
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
'    ' * 转换成 XML 数据格式
'    oRecordset.Save oDomBody, adPersistXML
    Voucher.setVoucherDataXML oDomHead, oDomBody
    SetStamp Voucher

    mOpStatus = SHOW_ALL
    Call SetCtlStyle(Me, Voucher, Me.Toolbar, Me.UFToolbar, mOpStatus)

    '读取时间戳
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

    ' * 定义错误数据源
    If Left(Err.Source, 3) = "***" Then
        sSource = Err.Source
    Else
        sSource = "***Sub LoadVoucherData of Form frmVoucher"
    End If

    ' * 抛出异常
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

'每个窗体都需要这个方法。Cancel与UnloadMode的参数的含义与QueryUnload的参数相同。
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
    '    tmpLinkTbl = "" ''单据联查 时 按钮状态控制 by zhangwchb 20110809
    Call UnRegisterMessage                                 '20110812

End Sub

Private Sub Form_Resize()
    Call SetLayOut
End Sub

'右键菜单事件调用
Private Sub PopMenu1_MenuClick(sMenuKey As String)

    On Error Resume Next

    Select Case LCase(sMenuKey)

        Case "addr"                                        '增行
            Voucher.AddLine Voucher.BodyRows + 1

        Case "delr"                                        '删行
            Voucher.DelLine Voucher.row

        Case "rslocate"                                    '定位记录
            Voucher.ShowFindDlg

        Case "incor"                                       '合并显示
            Call Execincor

        Case "batchmodify"                                 '批改
            Call ExecBathModify

        Case "copyr"                                       '复制行
            Voucher.DuplicatedLine Voucher.row
            Voucher.bodyText(Voucher.row, "AutoID") = ""
            Voucher.bodyText(Voucher.row, "cbsysbarcode") = "" '行复制需要清除掉表体的行条码
    End Select

End Sub

Public Sub ExecBathModify()    '批量修改
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

Public Sub Execincor()    '合并显示

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

    If Voucher.ShowVoucherDesign = True Then               '‘true代表更新，false代表不更新

        '        Dim domHead As New DOMDocument
        '
        '        Dim domBody As New DOMDocument

        '        Voucher.getVoucherDataXML domHead, domBody

        '重新设置单据格式

        Call SetTemplateData
        
        Call LoadVoucherData

        '        Voucher.setVoucherDataXML domHead, domBody

        SetLayOut
    End If

    '   Voucher.ShowVoucherDesign
End Sub

'处理所有的按钮事件

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

        '默认新增单据
        If LCase(strbuttonkey) = "add" Then strbuttonkey = sKey_Add2

        Select Case strbuttonkey

                '打印
            Case sKey_Print

                If ZwTaskExec(g_oLogin, AuthPrint, 1) = False Then Exit Sub
                Call ExecSubVoucherPrint(oConnection:=g_Conn, oVoucher:=Voucher, sBillNumber:=gstrCardNumber, sTemplateID:=m_strVT_PRN_ID, bPreview:=False)
                ZwTaskExec g_oLogin, AuthPrint, 0

                '预览
            Case sKey_Preview

                If ZwTaskExec(g_oLogin, AuthPrint, 1) = False Then Exit Sub
                Call ExecSubVoucherPrint(oConnection:=g_Conn, oVoucher:=Voucher, sBillNumber:=gstrCardNumber, sTemplateID:=m_strVT_PRN_ID, bPreview:=True)
                ZwTaskExec g_oLogin, AuthPrint, 0

                '输出
            Case sKey_Output

                If ZwTaskExec(g_oLogin, AuthOut, 1) = False Then Exit Sub
                Call ExportVoucherDataToFile(oConnection:=g_Conn, oVoucher:=Voucher, sBillNumber:=gstrCardNumber, sTemplateID:=m_strVT_PRN_ID)
                ZwTaskExec g_oLogin, AuthOut, 0

            Case sKey_VoucherDesign
                ShowVoucherDesign

            Case sKey_SaveVoucherDesign
                Call SaveVoucherTemplateInfo

                '第一页
            Case sKey_First
                Call ExecSubPageFirst

                '上一页
            Case sKey_Previous
                Call ExecSubPageUp

                '下一页
            Case sKey_Next
                Call ExecSubPageDown

                '最后一页
            Case sKey_Last
                Call ExecSubPageLast

            Case sKey_RefVoucher                         'zhangwchb 20110714 增加关联单据
                ShowRefVouchers False

            Case "tlbLinkAllVouch"
                ShowRefVouchers True

                '增加
            Case sKey_Add2
                '新增时申请权限
                '                If Me.ComTemplateShow.ListCount = 0 Then
                '                    'MsgBox "对不起，你没有模版权限，无法增加！", vbInformation, pustrMsgTitle
                '                    MsgBox GetResString("U8.ST.Default.00156"), vbInformation, GetResString("U8.SCM.ST.KCGLSQL.FrmMainST.rop.Caption") 'pustrMsgTitle
                '                    Exit Sub
                '                End If
                
                If ZwTaskExec(g_oLogin, AuthAdd, 1) = False Then Exit Sub
                Call initCustomRelation                    '20110822 by zhangwchb
                Call ExecSubAdd
                Call ZwTaskExec(g_oLogin, AuthAdd, 0)

            Case sKey_Add1                                 'strAdd1
                '新增时申请权限
                '                If Me.ComTemplateShow.ListCount = 0 Then
                '                    'MsgBox "对不起，你没有模版权限，无法增加！", vbInformation, pustrMsgTitle
                '                    MsgBox GetResString("U8.ST.Default.00156"), vbInformation, GetResString("U8.SCM.ST.KCGLSQL.FrmMainST.rop.Caption") 'pustrMsgTitle
                '                    Exit Sub
                '                End If
                
                If ZwTaskExec(g_oLogin, AuthAdd, 1) = False Then Exit Sub
                Call initCustomRelation                    '20110822 by zhangwchb
                Call ExecSubAdd(0)
                Call ZwTaskExec(g_oLogin, AuthAdd, 0)

                '修改
            Case sKey_Modify

                If Check_Auth = False Then Exit Sub

                If ZwTaskExec(g_oLogin, AuthModify, 1) = False Then Exit Sub

                '审批流中修改单据 20110817 by zhangwchb
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

                '删除
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

                '参照生单(合同)
            Case "Prorefer1"
            
                frmExcelDR.Show vbModal
                frmExcelDR.ZOrder 0

                Call ZwTaskExec(g_oLogin, AuthProrefer, 0)
            '参照生单(工程)
            Case "Prorefer2"
            
                If ZwTaskExec(g_oLogin, AuthProrefer1, 1) = False Then Exit Sub
                Call initCustomRelation                    '20110822 by zhangwchb
                Call ExecSubAdd

                If ReferVoucheng Then
                    '显示参照返回的数据
                    '        Debug.Print gDomReferHead.xml
                    '        Debug.Print gDomReferBody.xml
                    '生单处理
                    ProcessDataeng Voucher
                    '置灰生单按钮
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
                
                '参照生单(项目)
            Case "Prorefer3"
            
                If ZwTaskExec(g_oLogin, AuthProrefer, 1) = False Then Exit Sub
                Call initCustomRelation                    '20110822 by zhangwchb
                Call ExecSubAdd

                If ReferVouchpro Then
                    '显示参照返回的数据
                    '        Debug.Print gDomReferHead.xml
                    '        Debug.Print gDomReferBody.xml
                    '生单处理
                    ProcessDatapro Voucher
                    '置灰生单按钮
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
                  
 
                '推单-销售订单
            Case sKey_CreateSAVoucher

                '推单前比较时间戳,相等则跟新时间戳
                If SdFlg(lngVoucherID, OldTimeStamp) <> "" Then
                    ReDim varArgs(0)
                    varArgs(0) = Voucher.headerText("cCODE")
                    MsgBox GetStringPara("U8.DZ.JA.Res230", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    '                MsgBox "单据号为" & Voucher.headerText("cCODE") & "的单据已生成下游单据！", vbInformation, GetString("U8.DZ.JA.Res030")
                    Exit Sub

                End If

                Voucher.getVoucherDataXML oDomHead, oDomBody    '获得数据
                ExecmakeDom oDomHead, oDomBody, g_Conn     '组织数据

                g_Conn.BeginTrans

                If ExecCreateVoucher(oDomHead, oDomBody, g_Conn, g_oLogin, 销售, 销售订单) Then    '推单并回写生单状态
                    g_Conn.CommitTrans
                Else
                    g_Conn.RollbackTrans                   '推单失败回滚
                End If

                Call ExecSubRefresh

                '推单-采购订单
            Case sKey_CreatePUVoucher

                '推单前比较时间戳,相等则跟新时间戳
                If SdFlg(lngVoucherID, OldTimeStamp) <> "" Then
                    ReDim varArgs(0)
                    varArgs(0) = Voucher.headerText("cCODE")
                    MsgBox GetStringPara("U8.DZ.JA.Res230", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    '                MsgBox "单据号为" & Voucher.headerText("cCODE") & "的单据已生成下游单据！", vbInformation, GetString("U8.DZ.JA.Res030")
                    Exit Sub

                End If

                Voucher.getVoucherDataXML oDomHead, oDomBody    '获得数据
                ExecmakeDom oDomHead, oDomBody, g_Conn     '组织数据

                g_Conn.BeginTrans

                If ExecCreateVoucher(oDomHead, oDomBody, g_Conn, g_oLogin, 采购, 采购订单) Then
                    g_Conn.CommitTrans
                Else
                    g_Conn.RollbackTrans                   '推单失败回滚
                End If

                Call ExecSubRefresh

                '推单-其他入库单
            Case sKey_CreateSCVoucher

                '推单前比较时间戳,相等则跟新时间戳
                If SdFlg(lngVoucherID, OldTimeStamp) <> "" Then
                    ReDim varArgs(0)
                    varArgs(0) = Voucher.headerText("cCODE")
                    MsgBox GetStringPara("U8.DZ.JA.Res230", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    '               MsgBox "单据号为" & Voucher.headerText("cCODE") & "的单据已生成下游单据！", vbInformation, GetString("U8.DZ.JA.Res030")
                    Exit Sub

                End If

                Voucher.getVoucherDataXML oDomHead, oDomBody    '获得数据
                ExecmakeDom oDomHead, oDomBody, g_Conn     '组织数据

                g_Conn.BeginTrans

                If ExecCreateVoucher(oDomHead, oDomBody, g_Conn, g_oLogin, 库存, 其他入库单) Then
                    g_Conn.CommitTrans
                Else
                    g_Conn.RollbackTrans                   '推单失败回滚
                End If

                Call ExecSubRefresh

                '推单-应付单
            Case sKey_CreateAPVoucher

                '推单前比较时间戳,相等则跟新时间戳
                If SdFlg(lngVoucherID, OldTimeStamp) <> "" Then
                    ReDim varArgs(0)
                    varArgs(0) = Voucher.headerText("cCODE")
                    MsgBox GetStringPara("U8.DZ.JA.Res230", varArgs(0)), vbInformation, GetString("U8.DZ.JA.Res030")

                    '               MsgBox "单据号为" & Voucher.headerText("cCODE") & "的单据已生成下游单据！", vbInformation, GetString("U8.DZ.JA.Res030")
                    Exit Sub

                End If

                Voucher.getVoucherDataXML oDomHead, oDomBody    '获得数据
                ExecmakeDom oDomHead, oDomBody, g_Conn     '组织数据

                g_Conn.BeginTrans

                If ExecCreateVoucher(oDomHead, oDomBody, g_Conn, g_oLogin, 应付) Then
                    g_Conn.CommitTrans
                Else
                    g_Conn.RollbackTrans                   '推单失败回滚
                End If

                Call ExecSubRefresh

                '复制
            Case sKey_Copy
                '新增时申请权限
                '                If Me.ComTemplateShow.ListCount = 0 Then
                '                    'MsgBox "对不起，你没有模版权限，无法增加！", vbInformation, pustrMsgTitle
                '                    MsgBox GetResString("U8.ST.Default.00156"), vbInformation, GetResString("U8.SCM.ST.KCGLSQL.FrmMainST.rop.Caption") 'pustrMsgTitle
                '                    Exit Sub
                '                End If
                
                If Check_Auth = False Then Exit Sub

                If ZwTaskExec(g_oLogin, AuthAdd, 1) = False Then Exit Sub
                Call initCustomRelation                    '20110822 by zhangwchb
                Call ExecSubCopy
                Call ZwTaskExec(g_oLogin, AuthAdd, 0)

                '保存
            Case sKey_Save

                If Voucher.ProtectUnload2 <> 2 Then
                    Voucher.SetFocus

                    Exit Sub

                End If

                Call ExecSubSave

                '放弃
            Case sKey_Discard

                Call ExecSubDiscard

                '审核
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
                
                '弃审
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
                '                If Voucher.headerText("cCreateType") = "转换单据" Then
                '                    MsgBox GetString("U8.DZ.JA.Res240"), vbInformation, GetString("U8.DZ.JA.Res030")
                '
                '                    Exit Sub
                '
                '                End If

                '                If gcCreateType = "期初单据" Then
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

                '弃审
                '                If CheckSubmit(MainTable, "ID", CStr(lngVoucherID)) Then
                '                    wfcBack = 2
                '                    Call ExecCancelAudit(gsGUIDForVouch)
                '                Else

                If ZwTaskExec(g_oLogin, AuthUnVerify, 1) = False Then Exit Sub
                If ExecFunCompareUfts = True Then ExecSubCancelconfirm
                Call ZwTaskExec(g_oLogin, AuthUnVerify, 0)
                '                End If

                '打开
            Case sKey_Open

                If ZwTaskExec(g_oLogin, AuthOpen, 1) = False Then Exit Sub
                If ExecFunCompareUfts = True Then ExecSubOpen
                Call ZwTaskExec(g_oLogin, AuthOpen, 0)

                '关闭
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

                '定位
            Case sKey_Locate
                Call ExecLocate

                '刷新
            Case sKey_Refresh
                Call ExecSubRefresh

                '帮助
            Case gstrHelpCode
                SendKeys "{F1}"

                '增行
            Case sKey_Addrecord
                Call ExecSubAddRecord

                '插行
            Case sKey_InsertRecord
                Call ExecSubInsertRecord

                '删行
            Case sKey_Deleterecord
                Call ExecSubDeleterecord

                '?附件
            Case sKey_Acc
                Voucher.SelectFile

                '审批流 chenliangc
            Case sKey_Submit

                'by zhangwchb 20110829 提交保存
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

                '取价
            Case "rowprice", "allprice"
                Call GetPrice(LCase(strbuttonkey), "97", Voucher)    '97销售订单
                
                '批改
            Case "BatchModify"
                Call PopMenu1_MenuClick(strbuttonkey)
                
                '行复制
            Case "mnuCopyLine"
                Call PopMenu1_MenuClick("copyr")
                
                '拆分行
            Case "mnuSplitLine"
                Call PopMenu1_MenuClick("bsplit")
            
            Case "QueryStockAll"
                Call QueryStockAll(Voucher)
                
            Case "QueryStock"
                Call QueryStock(Voucher)
                
                '刷新表体现存量
            Case "RefreshStock"
                Call ShowBodyStockAll(Voucher)
                
                '归还
            Case "Return"

                Dim sTmp               As String

                Dim strMsg             As String

                Dim cCode              As String

                Dim IsBackWfcontrolled As Boolean '借出归还单是否工作流控制

                If getIsWfControl(g_oLogin, g_Conn, sTmp, "HYJCGH005") Then          '工作流控制
                    IsBackWfcontrolled = True
                Else
                    IsBackWfcontrolled = False
                End If

                cCode = Voucher.headerText("cCODE")
                
                If CheckCanBack(lngVoucherID, cCode, gcCreateType, sTmp) Then
                    If ExecReturn(lngVoucherID, sTmp, IsBackWfcontrolled, Voucher.headerText("ufts")) Then
                        strMsg = strMsg & GetStringPara("U8.ST.V870.00759", cCode) & vbCrLf '

                        'strMsg = strMsg & "单据 " & cCode & " 归还成功！" & vbCrLf
                        If sTmp <> "" Then strMsg = strMsg & sTmp & vbCrLf
                    Else
                        strMsg = strMsg & GetStringPara("U8.ST.V870.00760", cCode) & vbCrLf '

                        ' strMsg = strMsg & "单据 " & cCode & " 归还失败！" & vbCrLf
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

        ' * 显示友好的错误信息
        Call ShowErrorInfo(sHeaderMessage:=sMessage, lMessageType:=vbExclamation, lErrorLevel:=ufsELHeaderAndDescription)
        GoTo Exit_Label

    End If

End Sub

'审批流中修改单据 20110817 by zhangwchb
Private Function CheckVerModify() As Boolean
    On Error GoTo lerr
    Dim AuditServiceProxy As Object
    Dim objCalledContext As Object
    Dim strErr As String
    Dim IsChangeableVoucher As Boolean

    If Voucher.headerText("iswfcontrolled") = 1 Then
        If Voucher.headerText("iverifystate") = "0" Or Voucher.headerText("iverifystate") = "" Then
            GoTo ExitOK
            '       Else ' 审批流控制,已经提交,进入下面的判断
        End If
    Else
        GoTo ExitOK
    End If

    '判断终审前是否允许修改**********************************************************************************
    Set objCalledContext = CreateObject("UFSoft.U8.Framework.LoginContext.CalledContext")
    objCalledContext.SubId = g_oLogin.cSub_Id
    objCalledContext.TaskId = g_oLogin.TaskId
    objCalledContext.token = g_oLogin.userToken
    Set AuditServiceProxy = CreateObject("UFIDA.U8.Audit.ServiceProxy.AuditServiceProxy")

    IsChangeableVoucher = AuditServiceProxy.IsChangeableVoucher(Voucher.headerText("ID"), LCase$(gstrCardNumber), Voucher.headerText("cCode"), objCalledContext, strErr)

    '判断终审前是否允许修改**********************************************************************************

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

    'by liwqa MainView 改为MainTable


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

    '更新当前页变量PageCurrent
    Call UpdatePageCurrent(lngVoucherID)

    mOpStatus = SHOW_ALL
    Call SetCtlStyle(Me, Voucher, Me.Toolbar, Me.UFToolbar, mOpStatus)

    '读取时间戳
    OldTimeStamp = GetTimeStamp(g_Conn, MainTable, lngVoucherID)

    Exit Sub


Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Sub

'读取时间戳，并与旧时间戳比较
Private Function ExecFunCompareUfts() As Boolean

    '读取时间戳
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

'关闭
Private Sub ExecSubClose()
    On Error GoTo Err_Handler
    'by liwqa 并发
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
'打开
Private Sub ExecSubOpen()

    '状态:如果生单人不为空,置为3,即已生单;如果生单人为空:
    '                                                   如果审核人不为空,置为2,即审核;如果审核人为空,则置为1,即新建

    On Error GoTo Err_Handler
    'by liwqa 并发
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

'生单
Private Sub ExecSubKCreateBill()

    g_Conn.Execute "update " & MainTable & " set " & StrIntoUser & "='" & g_oLogin.cUserName & "' , " & StrdIntoDate & "='" & g_oLogin.CurDate & "' , " & StriStatus & "=3 where " & HeadPKFld & "=" & lngVoucherID
    Call ExecSubRefresh

End Sub

'审核
Private Sub ExecSubConfirm()
    On Error GoTo Err_Handler
    'by liwqa 并发
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
        sMsg = GetString("U8.DZ.JA.Res1940") & vbCrLf '"单据审核成功!"
          
    End If
    
'    '审核自动生成其他出库单
'    If gcCreateType <> "期初单据" Then
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


'弃审
Private Sub ExecSubCancelconfirm()
    On Error GoTo Err_Handler
    'by liwqa 并发
    Dim m_bTrans As Boolean
    Dim lngAct As Long
    g_Conn.BeginTrans
    m_bTrans = True
    Dim sql As String
    
    
        
      '************************************
     '更新项目发布单累计数量和金额

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
'        '业务通知
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
 


'增行
Private Sub ExecSubAddRecord()
    Voucher.AddLine
End Sub
'插行
Private Sub ExecSubInsertRecord()
    Dim iRow As Long
    iRow = Voucher.row
    If iRow = 0 Then
        Exit Sub
    Else
        Voucher.AddLine Voucher.row, , ALSPrevious
    End If
End Sub
'删行
Private Sub ExecSubDeleterecord()
    Voucher.DelLine Voucher.row
End Sub

'放弃'
Private Sub ExecSubDiscard()

    If MsgBox(GetString("U8.DZ.JA.Res300"), vbYesNo, GetString("U8.DZ.JA.Res030")) = vbYes Then
        mOpStatus = SHOW_ALL
        Call ExecSubRefresh
    End If

End Sub


'首页
Public Sub ExecSubPageFirst()

    '用页数定位会导致取得id是错误的
    Dim sql As String
    Dim rsid As New ADODB.Recordset

    PageCurrent = 1

    '获取数据权限
    '    Dim sRet As String
    '    sRet = GetRowAuthAlls(g_Conn, g_oLogin.cUserId, "R")
    ''    If sRet <> "" Then
    ''    sRet = " where cCreateType<>'期初单据' and 1=1 " & sRet
    '    sRet = " where  1=1 " & sRet


    '    If tmpLinkTbl <> "" Then  '单据联查 时 按钮状态控制 by zhangwchb 20110809
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

'末页
Public Sub ExecSubPageLast()
    '用页数定位会导致取得id是错误的
    Dim sql As String
    Dim rsid As New ADODB.Recordset

    PageCurrent = pageCount

    '获取数据权限
    '    Dim sRet As String
    '    sRet = GetRowAuthAlls(g_Conn, g_oLogin.cUserId, "R")
    ''    If sRet <> "" Then
    ''    sRet = " where cCreateType<>'期初单据' and  1=1 " & sRet
    '    sRet = " where  1=1 " & sRet

    '    If tmpLinkTbl <> "" Then  '单据联查 时 按钮状态控制 by zhangwchb 20110809
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

'上一页
Public Sub ExecSubPageUp()
    '用页数定位会导致取得id是错误的
    Dim sql As String
    Dim rsid As New ADODB.Recordset

    If PageCurrent > 1 Then
        PageCurrent = PageCurrent - 1

        '获取数据权限
        '        Dim sRet As String
        '        sRet = GetRowAuthAlls(g_Conn, g_oLogin.cUserId, "R")


        '        SQL = "select top 1 " & HeadPKFld & " as id from " & MainTable & " where cCreateType<>'期初单据' and " & HeadPKFld & "<" & lngVoucherID & sRet & " order by " & HeadPKFld & " desc"

        '        If tmpLinkTbl <> "" Then  '单据联查 时 按钮状态控制 by zhangwchb 20110809
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

'下一页
Public Sub ExecSubPageDown()
    '用页数定位会导致取得id是错误的
    Dim sql As String
    Dim rsid As New ADODB.Recordset

    If PageCurrent < pageCount Then

        '获取数据权限
        Dim sRet As String
        '        sRet = GetRowAuthAlls(g_Conn, g_oLogin.cUserId, "R")

        '        SQL = "select top 1 " & HeadPKFld & " as id from " & MainTable & " where  cCreateType<>'期初单据' and " & HeadPKFld & ">" & lngVoucherID & sRet & " order by " & HeadPKFld & " asc "

        '        If tmpLinkTbl <> "" Then  ''单据联查 时 按钮状态控制 by zhangwchb 20110809
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

        'mod不能调用frm 中函数
        Call ExecSubRefresh

    End If


End Sub

'刷新
'根据当前lngvoucherID更新
'加载,翻页,删除时必须更新lngvoucherID
Public Sub ExecSubRefresh()
    bAlter = False
    
    Call LoadVoucherData

    If Voucher.headerText("cCODE") = "" Then
        Call ExecSubPageLast
    End If

    '刷新时重新得到后台数据 by liwq
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
        '快捷键ALT + E输出
        If KeyCode = vbKeyE Then
            If Me.Toolbar.Buttons(sKey_Output).Enabled = True Then
                ButtonClick sKey_Output
            End If
            '弃审
        ElseIf KeyCode = vbKeyU Then
            If Me.Toolbar.Buttons(sKey_Cancelconfirm).Enabled = True Then
                ButtonClick sKey_Cancelconfirm
            End If
            '撤销
        ElseIf KeyCode = vbKeyJ Then
            If Me.Toolbar.Buttons(sKey_Unsubmit).Enabled = True Then
                ButtonClick sKey_Unsubmit
            End If
        ElseIf KeyCode = vbKeyPageUp Then                  '首页
            If Me.Toolbar.Buttons(sKey_First).Enabled = True Then
                ButtonClick sKey_First
            End If
        ElseIf KeyCode = vbKeyPageDown Then                '末页
            If Me.Toolbar.Buttons(sKey_Last).Enabled = True Then
                ButtonClick sKey_Last
            End If
        ElseIf KeyCode = vbKeyC Then                       '关闭
            If Me.Toolbar.Buttons(sKey_Close).Enabled = True Then
                ButtonClick sKey_Close
            End If
        ElseIf KeyCode = vbKeyO Then                       '打开
            If Me.Toolbar.Buttons(sKey_Open).Enabled = True Then
                ButtonClick sKey_Open
            End If
        End If
    ElseIf Shift = vbCtrlMask Then
        '快捷键Ctrl + P打印
        If KeyCode = vbKeyP Then
            If Me.Toolbar.Buttons(sKey_Print).Enabled = True Then
                ButtonClick sKey_Print
            End If
            '快捷键Ctrl + W打印预览
        ElseIf KeyCode = vbKeyW Then
            If Me.Toolbar.Buttons(sKey_Preview).Enabled = True Then
                ButtonClick sKey_Preview
            End If
            '快捷键Ctrl + F3过滤
        ElseIf KeyCode = vbKeyF3 Then
            If Me.Toolbar.Buttons(sKey_Locate).Enabled = True Then
                ButtonClick sKey_Locate
            End If
            '快捷键Ctrl + G生单
        ElseIf KeyCode = vbKeyG Then
            If Me.Toolbar.Buttons(sKey_ReferVoucher).Enabled = True Then
                ButtonClick sKey_ReferVoucher
            End If
            '保存新增
        ElseIf KeyCode = vbKeyS Then
            ' ButtonClick sKey_ReferVoucher
            If Me.Toolbar.Buttons(sKey_Save).Enabled = True Then
                ButtonClick sKey_Save
                Call g_oBusiness.SelectToolbarButton(Toolbar.Buttons(sKey_Add))
            End If

            '放弃
        ElseIf KeyCode = vbKeyZ Then
            If Me.Toolbar.Buttons(sKey_Discard).Enabled = True Then
                ButtonClick sKey_Discard
            End If
            '提交
        ElseIf KeyCode = vbKeyJ Then
            If Me.Toolbar.Buttons(sKey_Submit).Enabled = True Then
                ButtonClick sKey_Submit
            End If
            '审核
        ElseIf KeyCode = vbKeyU Then
            If Me.Toolbar.Buttons(sKey_Confirm).Enabled = True Then
                ButtonClick sKey_Confirm
            End If

            '复制
        ElseIf KeyCode = vbKeyF5 Then
            If Me.Toolbar.Buttons(sKey_Copy).Enabled = True Then
                ButtonClick sKey_Copy
            End If
            '生单
        ElseIf KeyCode = vbKeyG Then
            'ButtonClick sKey_ReferVoucher
        ElseIf KeyCode = vbKeyR Then                       '刷新
            If Me.Toolbar.Buttons(sKey_Refresh).Enabled = True Then
                ButtonClick sKey_Refresh
            End If

        ElseIf KeyCode = vbKeyD Then                       '删行
            If Me.Toolbar.Buttons(sKey_Deleterecord).Enabled = True Then
                ButtonClick sKey_Deleterecord

            End If
        ElseIf KeyCode = vbKeyA Then                       '删行
            If Me.Toolbar.Buttons(sKey_Addrecord).Enabled = True Then
                ButtonClick sKey_Addrecord
            End If
        ElseIf KeyCode = vbKeyF4 Then
            Call ExitForm(0, 0)
            '快捷键Ctrl+E或者Ctrl+B，自动指定批号,入库单号
        ElseIf KeyCode = vbKeyE Or KeyCode = vbKeyB Or KeyCode = vbKeyQ Or KeyCode = vbKeyO Then
            'Call GetBatchInfoFun(Voucher, KeyCode, Shift)
            KeyCode = 0
        End If
        ' End If
    End If

    Select Case KeyCode
        Case vbKeyF1                                       '帮助
            Call LoadHelpId(Me, "15030910")
        Case vbKeyF5                                       '新增

            If Me.Toolbar.Buttons(sKey_Add).Enabled = True Then
                ButtonClick sKey_Add
                'Call g_oBusiness.SelectToolbarButton(Toolbar.Buttons(sKey_Add))
                '         ElseIf Me.Toolbar.Buttons(sKey_Add1).Enabled = True Then
                '              Call g_oBusiness.SelectToolbarButton(Toolbar.Buttons(sKey_Add1))
            End If
            '  ButtonClick sKey_Add
        Case vbKeyF6                                       '保存
            If Me.Toolbar.Buttons(sKey_Save).Enabled = True Then
                ButtonClick sKey_Save
            End If
        Case vbKeyF8                                       '修改
            If Me.Toolbar.Buttons(sKey_Modify).Enabled = True Then
                ButtonClick sKey_Modify
            End If
        Case vbKeyDelete                                   '删除
            If Me.Toolbar.Buttons(sKey_Delete).Enabled = True Then
                ButtonClick sKey_Delete
            End If
        Case vbKeyPageUp
            If Me.Toolbar.Buttons(sKey_Previous).Enabled = True Then
                ButtonClick sKey_Previous                  '上一页
            End If
        Case vbKeyPageDown
            If Me.Toolbar.Buttons(sKey_Next).Enabled = True Then
                ButtonClick sKey_Next                      '下一页
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

'增加
Public Sub ExecSubAdd(Optional byt1 As Byte = 1)
    On Error GoTo Err_Handler:
    Dim rs As New ADODB.Recordset
    Dim sql, sCache, sMessage, sSource As String
    Dim oDomHead, oDomBody As DOMDocument
    Set oDomHead = New DOMDocument
    Set oDomBody = New DOMDocument
    numappprice = 0
    bAlter = False
    
    '单据初始化
    'enum by modify
    If byt1 = 1 Then
        gcCreateType = "新增单据"
    Else
        gcCreateType = "期初单据"
    End If

    Call setTemplate("")                                   'by liwqa Template

    Call InitVoucher
    Call SetLayOut

    '表头
    sql = "select * from " & MainView & " where 1=2"
    Call rs.Open(sql, g_Conn, 1, 1)
    rs.Save oDomHead, adPersistXML
    rs.Close
    Set rs = Nothing


    '表体
    sql = "select * from " & DetailsView & " where 1=2"
    Call rs.Open(sql, g_Conn, 1, 1)
    rs.Save oDomBody, adPersistXML
    rs.Close
    Set rs = Nothing

    '新增单据
    'enum by modify
    Voucher.AddNew ANMNormalAdd, oDomHead, oDomBody
    Voucher.SetBillNumberRule sCache
    '制单人,单据日期
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

'    Voucher.headerText("cexch_name") = "人民币"
'    Voucher.headerText("iexchrate") = 1#
    symbol = "*"                                           '汇率折算方式
    Voucher.headerText("iStatus") = "1"
    '    Voucher.headerText("cCreateType") = "新增单据"
  
  If isfromcon = True Then
   Voucher.headerText("sourcetype") = "FYSL0004"
 
 End If
     

    '    Voucher.headerText("cAboutVoucher") = "销售出库单"
'    Voucher.headerText("isengdec") = "否"
    Voucher.EnableHead "supengcode", False
    Voucher.EnableHead "supcname", False
    

    Voucher.headerText("ivtid") = m_strVT_ID               'by liwqa Template

    Voucher.headerText("VoucherType") = gstrCardNumber     '"HY99"
    Dim errMsg As String
    If getIsWfControl(g_oLogin, g_Conn, errMsg, gstrCardNumber) Then
        Voucher.headerText("iswfcontrolled") = 1
    End If

    '标志状态
    Voucher.VoucherStatus = VSeAddMode
    '设置单据编号是否可编辑
    Dim manual As Boolean                                  ' 是否完全手工编号
    Call SetVouchCodeEnable(manual)

    mOpStatus = ADD_MAIN

    Call SetCtlStyle(Me, Voucher, Me.Toolbar, Me.UFToolbar, mOpStatus)

    'by zhangwchb 20110829 提交保存  If bInWorkFlow = True Then
'        Me.Toolbar.Buttons("Submit").Enabled = True
'        UFToolbar.RefreshEnable
'    Else
'        Me.Toolbar.Buttons("Submit").Enabled = False
'        UFToolbar.RefreshEnable
'    End If
'

    '给单据表体单元格加载图片
    '    Dim body As Object
    '    Set body = Voucher.GetBodyObject
    '
    '    body.Cell(flexcpPicture, 1, 1, 1, 1) = Me.ImageList1.ListImages.Item(2).Picture



Exit_Label:
    On Error GoTo 0

    Exit Sub
Err_Handler:
    sMessage = GetString("U8.DZ.JA.Res310")

    ' * 显示友好的错误信息
    Call ShowErrorInfo( _
            sHeaderMessage:=sMessage, _
            lMessageType:=vbExclamation, _
            lErrorLevel:=ufsELHeaderAndDescription _
            )

    ' * 定义错误数据源

    If Left(Err.Source, 3) = "***" Then
        sSource = Err.Source
    Else
        sSource = "***Sub AddNewVoucher of Form frmPMRecord"
    End If

    ' * 调试模式时，显示调试窗口，用于跟踪错误
    If g_blnDEBUG_MODE Then
        Call ShowDebugForm( _
                bErrorMode:=True, _
                sProcedure:=sSource)
    End If
    GoTo Exit_Label
End Sub

'复制
Private Sub ExecSubCopy()

    Dim oDomHead, oDomBody As New DOMDocument
    Dim sCache As String

    bAlter = False
    
    Voucher.getVoucherDataXML oDomHead, oDomBody

    Voucher.AddNew ANMCopyALL, oDomHead, oDomBody
    Voucher.SetBillNumberRule sCache

    '制单人,单据日期
    Voucher.headerText(StrcMaker) = g_oLogin.cUserName
   
    Voucher.headerText(HeadPKFld) = ""                     '主表标志id
    Voucher.headerText(StrcHandler) = ""                   '审核人
    Voucher.headerText(StrdVeriDate) = ""                  '审核日期
    Voucher.headerText(StrCloseUser) = ""                  '关闭人
    Voucher.headerText(StrdCloseDate) = ""                 '关闭日期
    Voucher.headerText(StrIntoUser) = ""                   '生单人
    Voucher.headerText(StrdIntoDate) = ""                  '生单日期
    Voucher.headerText("iStatus") = "新建"
    '    Voucher.headerText("cCreateType") = "新增单据"
    '只有转换单复制时才改为新增单据
  
    '    Voucher.headerText("cType") = "客户"
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

 

    '设置单据编号是否可编辑
    Dim manual As Boolean                                  ' 是否完全手工编号
    Call SetVouchCodeEnable(manual)

    If manual Then                                         '完全手工编号则编号置空
        Voucher.headerText(strcCode) = ""
    End If


    '标志状态
    Voucher.VoucherStatus = VSeAddMode
    mOpStatus = ADD_MAIN

'    Call setAllDisable
    Call SetCtlStyle(Me, Voucher, Me.Toolbar, Me.UFToolbar, mOpStatus)

    '参照生单置灰
    Me.Toolbar.Buttons("Prorefer").Enabled = False
    Me.UFToolbar.RefreshEnable

    'by zhangwchb 20110829 提交保存
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
                    Case "ccusinvcode", "ccusinvname"      ''客户存货代码
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

'自动获取单据号
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

        Set oelement = oDomFormat.selectSingleNode("//单据编号")

        '支持完全手工编号
        '允许手工修改得含义为 完全手工编号， 重号自动重取的含义为 手工修改，重号自动重取
        bManualCode = oelement.getAttribute("允许手工修改")
        bCanModyCode = oelement.getAttribute("允许手工修改") Or oelement.getAttribute("重号自动重取")
    Else
        MsgBox sError, vbInformation, GetString("U8.DZ.JA.Res030")
    End If

    '支持完全手工编号，此时不取单据号 2003-07-16 黄朝阳
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
    '    '使用新的参照服务
    '    Dim objRefer As New U8RefService.IService
    '    Set vis = Voucher.ItemState(Col, 1)                    '记住此处，是关键
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
    '    '这句重要，将参照事件设置为不启动状态
    '    referpara.Cancel = True
    '
    ''    On Error GoTo errhandle
    '    '设置参照是否多选
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
'        pCom.AddItem "客户"
'        pCom.AddItem "供应商"
'        pCom.AddItem "部门"
'        '        pCom.AddItem "人员"
'    End If
'
'    If LCase(Voucher.ItemState(Index, siheader).sFieldName) = LCase("cAboutVoucher") Then
'        pCom.AddItem "销售出库单"
'        pCom.AddItem "服务单"
'        pCom.AddItem "委外出库单"
'    End If
'
'    If LCase(Voucher.ItemState(Index, siheader).sFieldName) = LCase("cfreight") Then
'        pCom.AddItem "否"
'        pCom.AddItem "是"
'    End If
'End Sub


Private Sub Voucher_headBrowUser(ByVal Index As Variant, sRet As Variant, referpara As UAPVoucherControl85.ReferParameter)
    Call VoucherheadBrowUser(Voucher, Index, sRet, referpara)
End Sub


Private Sub Voucher_headCellCheck(Index As Variant, retvalue As String, bChanged As UAPVoucherControl85.CheckRet, referpara As UAPVoucherControl85.ReferParameter)
    Call VoucherheadCellCheckFun(Voucher, Index, retvalue, bChanged, referpara)
End Sub

Private Sub Voucher_MouseUp(ByVal section As UAPVoucherControl85.SectionsConstants, ByVal Index As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    '调出右键菜单
'    If Button = 2 Then
'
'        If Me.Voucher.VoucherStatus = VSNormalMode Then
'
'            Me.PopMenu1.Visible("AddR") = False            '增行
'            Me.PopMenu1.Visible("DelR") = False            '删行
'            Me.PopMenu1.Visible("batchModify") = False     '批改
'            Me.PopMenu1.Visible("copyR") = False           '复制行
'            '        Me.PopMenu1.Visible("bsplit") = False
'            Me.PopMenu1.Visible("A") = False               '批改
'            Me.PopMenu1.Visible("B") = False               '批改
'
'        Else
'
'            Me.PopMenu1.Visible("AddR") = False             '增行
'            Me.PopMenu1.Visible("DelR") = True             '删行
'            Me.PopMenu1.Visible("copyR") = True            '复制行
'            Me.PopMenu1.Visible("batchModify") = True      '批改
'            '        Me.PopMenu1.Visible("bsplit") = True
'
'            Me.PopMenu1.Visible("A") = True                '批改
'            Me.PopMenu1.Visible("B") = True                '批改
'
'        End If
'
'        'retmenu:菜单根节点名称
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

    '保存打印格式设置
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
    If Voucher.headerText("isengdec") = "否" Then
 
    Voucher.EnableHead "supengcode", False
    Voucher.EnableHead "supcname", False
    End If

  

    Set rs = Nothing
    Exit Function

Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Function



'根据表头的单据模板
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
    '与当前模板一致直接退出
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

'初始化自定义项回填类 20110822
Private Sub initCustomRelation()

    Voucher.SetCustomRelation mobjSubServ.GetCustomRelationRecord(g_Conn, gstrCardNumber)

End Sub

'单据控件的自定义回填事件 20110822
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
    '权限检查
    Dim rs As New ADODB.Recordset
    Dim sql As String
    Dim ctype  As String, bObjectCode As String
    Dim i As Long
    Check_Auth = True
'    ctype = Voucher.headerText("cType")
'    bObjectCode = Voucher.headerText("bObjectCode") & ""
'    If Voucher.headerText("bObjectCode") & "" <> "" Then
'        Select Case ctype
'            Case "供应商"
'                sql = "SELECT distinct bObjectCode  from HY_DZ_BorrowOut a left join vendor b on a.bObjectCode=b.cvencode where a.bObjectCode='" & Voucher.headerText("bObjectCode") & "' "
'                sql = sql & IIf(sAuth_vendorW = "", "", " and b.iid in (" & sAuth_vendorW & ")")
'            Case "部门"
'                sql = "SELECT distinct bObjectCode  from HY_DZ_BorrowOut a left join department b on a.bObjectCode=b.cdepcode where a.bObjectCode='" & Voucher.headerText("bObjectCode") & "' "
'                sql = sql & IIf(sAuth_depW = "", "", " and b.cdepcode in (" & sAuth_depW & ")")
'            Case "客户"
'                sql = "SELECT distinct bObjectCode  from HY_DZ_BorrowOut a left join customer b on a.bObjectCode=b.ccuscode where a.bObjectCode='" & Voucher.headerText("bObjectCode") & "' "
'                sql = sql & IIf(sAuth_CusW = "", "", " and b.iid in (" & sAuth_CusW & ")")
'        End Select
'        If sql <> "" Then
'        rs.Open sql, g_Conn
'        If rs.EOF Or rs.BOF Then
'
'            Select Case ctype
'                Case "供应商"
'                    MsgBox GetStringPara(("U8.DZ.JA.Res2040"), bObjectCode), vbInformation, GetString("U8.DZ.JA.Res030")
'                    Check_Auth = False
'                    Exit Function
'                Case "部门"
'                    MsgBox GetStringPara(("U8.DZ.JA.Res2050"), bObjectCode), vbInformation, GetString("U8.DZ.JA.Res030")
'                    Check_Auth = False
'                    Exit Function
'                Case "客户"
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
    '部门
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
'        '存货
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
'        '仓库
'        sql = "select a.cinvcode from HY_DZ_BorrowOuts a  left join Warehouse b on a.cwhcode=b.cwhcode where cinvcode='" & Voucher.bodyText(i, "cinvcode") & "'" & IIf(sAuth_WareHouseW = "", "", " and( ISNULL(b.cwhcode,N'')=N'' OR b.cwhcode in (" & sAuth_WareHouseW & "))")
'
'        Set rs = New ADODB.Recordset
'        rs.Open sql, g_Conn
'        If rs.EOF Or rs.BOF Then
'            MsgBox GetStringPara(("U8.DZ.JA.Res2090"), Voucher.bodyText(i, "cwhcode")), vbInformation, GetString("U8.DZ.JA.Res030")
'            Check_Auth = False
'            Exit Function
'        End If
'        '货位
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

'判断单据类型是否启用审批流  'by zhangwchb 20110829 提交保存
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

'wangfb 11.0Toolbar迁移 2012-03-31
Private Function SetToolbarVisible()
    On Error Resume Next
    With Toolbar
        '11.0新规范启用工作流才显示
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
        '附件、预览、帮助、生单、推单、取价按钮隐藏
 
        .Buttons("Preview").Visible = False
        .Buttons("Help").Visible = False
        .Buttons("Locate").Visible = False
        
        If Voucher.VoucherStatus <> VSNormalMode Then
            '11.0去掉增行按钮
            '显示格式下的合并显示，取消合并和设置合并规则隐藏
     
            Toolbar.Buttons(sKey_RefVoucher).Visible = False
            
            If Voucher.VoucherStatus = VSeAddMode Then
                '新增时讨论可用，但批注和通知不可用
                Toolbar.Buttons("Notes").Enabled = False
                Toolbar.Buttons("Discuss").Enabled = True
                Toolbar.Buttons("Notify").Enabled = False
            Else
                '修改时批注、讨论可用，但通知不可用
                Toolbar.Buttons("Notes").Enabled = True
                Toolbar.Buttons("Discuss").Enabled = True
                Toolbar.Buttons("Notify").Enabled = False
            End If
            Toolbar.Buttons("tlbLinkAllVouch").Enabled = False
        Else
             
            
            '空白单据时批注、讨论和通知都不可用，否则都可用
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

    If Voucher.headerText("ccreatetype") = "转换单据" Then
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
