VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{201FB79B-5556-47A4-AD9C-A46BA0C45A44}#6.42#0"; "UFToolBarCtrl.ocx"
Object = "{005151D4-33F6-471B-B651-222D4E411832}#4.5#0"; "UFFormPartner.ocx"
Object = "{AF8BBBB7-94C6-4772-B826-624478C37D6A}#1.5#0"; "UFKEYHOOK.ocx"
Object = "{86808282-58F4-4B17-BBCA-951931BB7948}#2.82#0"; "U8VouchList.ocx"
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
   StartUpPosition =   3  'Windows Default
   Begin U8VouchList.VouchList VouchList2 
      Height          =   1575
      Left            =   480
      TabIndex        =   3
      Top             =   3960
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   2778
   End
   Begin UFFormPartner.UFFrmCaption UFFrmCaptionMgr 
      Left            =   7320
      Top             =   1680
      _ExtentX        =   1931
      _ExtentY        =   1085
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
   Begin U8VouchList.VouchList VouchList1 
      Height          =   3135
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   9015
      _ExtentX        =   15901
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
         Name            =   "����"
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
'ģ�鹦��˵����
'               1��ʵ���б�Ļ�����ť����:��ӡ�������Ԥ����
'                                       �������������������ء����ˡ���λ����Ŀ��ȫѡ��ȫ�������顢ˢ��
'               2��֧���б��ҳ����
'               3��֧�ֹ���Ȩ��

'����ʱ�䣺2008-11-21
'�����ˣ�xuyan
'****************************************

'�б����
Public objColset As U8ColumnSet.clsColSet
Public objColset2 As U8ColumnSet.clsColSet

Private WithEvents m_pagediv As Pagediv                    '��ҳ
Attribute m_pagediv.VB_VarHelpID = -1
Private WithEvents m_pagediv2 As Pagediv                    '��ҳ
Attribute m_pagediv2.VB_VarHelpID = -1
Private m_coni As IPagedivConi                             '�����������϶��Ǵ�U8Colset�н��г�ʼ��
Private m_coni2 As IPagedivConi                             '�����������϶��Ǵ�U8Colset�н��г�ʼ��

Private m_Cancel, m_UnloadMode As Integer
Attribute m_UnloadMode.VB_VarUserMemId = 1073938434
Private ListTitle As String
Attribute ListTitle.VB_VarUserMemId = 1073938436

Dim cMenuId, cMenuName, cAuthId As String    ' ���ݽڵ�
Attribute cMenuId.VB_VarUserMemId = 1073938437
Attribute cMenuName.VB_VarUserMemId = 1073938437
Attribute cAuthId.VB_VarUserMemId = 1073938437

'����Ȩ��
Private Const AuthBrowselist = "PD01030101"    '���
Private Const AuthBrowseLink = "ST02JC020406"    '����
Private Const AuthPrint = "LSDG000101"    '��ӡ
Private Const AuthOut = "LSDG000102"    '���
Private Const AuthVerify = "ST02JC020105"   '���
Private Const AuthUnVerify = "ST02JC020106"    '����
Private Const AuthReturn = "ST02JC020302"    '�����黹

Private AuthEdit As String '�༭

Private Const AuthOpen = "SAM0302004"    '��
Private Const AuthClose = "SAM0302005"    '�ر�

' Private Const AuthVerify = "ST02JC020105" '���
' Private Const AuthUnVerify = "ST02JC020106" '����

Private strMsg As String
Attribute strMsg.VB_VarUserMemId = 1073938440
Private oldccode As String
Private strStatus3 As String
Attribute strStatus3.VB_VarUserMemId = 1073938442

'by zhangwchb 20110719 �б�γ����չ
Dim sExtendField As String
Dim sExtendJoinSQL As String
Dim oExtend As Object

Private o_FilterObject As Object
Public gaiz As Boolean

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

'��ʼ������
    strWhere2 = "1=2"
    strWhere = "1=2"
'    Call CreateTableView
    Call InitMulText
    Call InitForm
    '1����ʼ���б������б�ģ�� InitList
    '2�����ع�������
    '3����ʼ����ҳ�ؼ� InitConi
    '4������ҳ�ؼ����ݣ������б����� FillList(m_pagediv.LoadData)��m_pagediv_GetData��m_pagediv_AfterGetData

    '��ʼ���б�
    Call InitList
        InitList2
    '��ʼ����ҳ�ؼ�
    Call InitConi(strWhere)
'    Call InitConi2(strWhere2)
    '����ҳ�ؼ����ݣ������б�����
    Call FillList
'    Call FillList2
    '11.1�ϲ���ʾ
    Call SetToolbarForColumn

End Sub

Private Sub Form_Resize()

    UFToolbar1.Move 0, 0, Me.ScaleWidth, Me.Toolbar1.Height
    'VouchList1.Move 0, 0, Me.ScaleWidth, IIf(Me.ScaleHeight - Toolbar1.Height < 0, 0, Me.ScaleHeight - Toolbar1.Height)
    VouchList1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight / 2
    VouchList2.Move 0, VouchList1.Height, Me.ScaleWidth, Me.ScaleHeight / 2
    'PagedivCtl1.Move 0, Me.VouchList1.Height, Me.ScaleWidth, 600

End Sub

'������ʱ�����Ӧ����ͼ
Private Sub CreateTableView()
    If Me.gaiz Then
        TblName = CreateTableName("EF_InScanDetail11")
    Else
        TblName = CreateTableName("EF_InScanDetail10")
    End If
    
    ViewDetailName = "v_" & TblName & "Detail"
    ViewMainName = "v_" & TblName & "Main"
    CreateTable (TblName)
    Call CreateView(ViewDetailName, ViewMainName, TblName)
End Sub

Private Sub CreateTable(talName As String)
    sql = "if exists (select * from sysobjects where id = object_id('" & talName & "') and sysstat & 0xf = 3)" & _
        " drop table " & talName
    g_Conn.Execute sql
    sql = "select * into " & talName & " from EF_InScanDetail where 1=2"
    g_Conn.Execute sql
End Sub

Private Sub CreateView(ViewDetailName As String, ViewMainName As String, TblName As String)
    Dim adocomm As New ADODB.Command
    Set adocomm.ActiveConnection = g_Conn
    adocomm.CommandText = "EF_proc_v_InScantmp"
    adocomm.CommandType = adCmdStoredProc
    adocomm.Parameters(1) = ViewDetailName
    adocomm.Parameters(2) = ViewMainName
    adocomm.Parameters(3) = TblName
    adocomm.Execute
End Sub

Private Sub InitForm()
    On Error Resume Next

    
     Dim ErrInfo As String
     Dim bSuccess As Boolean
     Dim sListName As String
     
    ListTitle = "����嵥�б�"  '�б����
'   cMenuId = "QR0215"
    cMenuName = "����嵥�б�"
    sListName = "����嵥�б����"
    UFFrmCaptionMgr.Caption = "����嵥�б�"
    AuthEdit = "PD01030102"
    
    
   ' cAuthId = "ST02JC020101"
    '*************
    
    '*******************
    
     
     Dim errStr As String
    Set clsbill = CreateObject("USERPCO.VoucherCO")        'New USERPCO.VoucherCO
    clsbill.IniLogin g_oLogin, errStr
    Set mologin = clsbill.login
    Set UFToolbar1.Business = goBusiness
    
    'TODO:
    '****************wangfb 11.0ToobarǨ��2012-03-21 start ************************
    Call InitToolBar(mologin, "EF_V_HZDesignList", Toolbar1, UFToolbar1)
    
    
'    Call UFToolbar1.InitExternalButton("InputpoApp001_List", mologin.OldLogin)
    Call UFToolbar1.SetFormInfo(Me.VouchList1, Me)
'    If sListName <> "" Then
'        If o_FilterObject Is Nothing Then
'           Set o_FilterObject = CreateObject("UFGeneralFilter.FilterSrv")
'        End If
'
'        bSuccess = o_FilterObject.InitBaseVarValue(g_oLogin, "", sListName, "GL", ErrInfo)
'        VouchList1.InitFlt g_oLogin, o_FilterObject, "", "", "", ErrInfo
'     End If
     
    Me.Caption = UFFrmCaptionMgr.Caption

    
    
    VouchList1.formCode = "EF_V_HZDesignList"
'    VouchList1.HiddenRefreshView = True
'    VouchList1.HiddenPageDiv = False
'
'    VouchList2.HiddenRefreshView = True
'    VouchList2.HiddenTotalView = True
'    VouchList2.HiddenPageDiv = True
'    VouchList2.HiddenFoldView = True
    VouchList2.HideTitleCaption = True

    '��������ʼ��
    '11.0toolbarǨ�ƣ�������ҵ�񵥾ݱ�׼��֮��ԭ���Ϳ��İ�ť��ʼ��
    'Call Init_Toolbarlist(Me.Toolbar1)
    
    Call ChangeOneFormTbrlist(Me, Me.Toolbar1, Me.UFToolbar1)
'    Call SetWFControlBrnsList(g_oLogin, g_Conn, Me.Toolbar1, Me.UFToolbar1, gstrCardNumber)
    
    SetToolbarVisible
    UFToolbar1.RefreshVisible
    '****************wangfb 11.0ToobarǨ��2012-03-21 start ************************
    
    '��ȡU8�汾 -chenliangc
    gU8Version = GetU8Version(gConn)

End Sub


Private Sub InitList()
    On Error GoTo Err_Handler

    Set objColset = New U8ColumnSet.clsColSet
    objColset.Init gConn, goLogin.cUserId

    '11.1�ϲ���ʾ
    Set VouchList1.ColSetObj = objColset
    
    '�����б�ģ��
    objColset.setColMode gstrCardNumberlist
    Me.VouchList1.InitHead objColset.getColInfo()

    '�ϼƷ�ʽ
    If Replace(strWhere, " ", "") = "(1=2)" Then
        VouchList1.SumStyle = vlSumNone '   vlRecordAndGridsum                     ' vlGridSum
    Else
        VouchList1.SumStyle = vlRecordAndGridsum
    End If

    'FillAppend ����
    'FillOverwrite ������䷽ʽ����ʹ��ѡ���ܣ�[ѡ��]����ʾΪ�̶���
    Me.VouchList1.FillMode = FillOverwrite
    VouchList1.ShowSelCol = True
    '    VouchList1.SumStyle = vlSumNone

    '�б��������
    Me.VouchList1.Title = ListTitle
 
    With VouchList1
 

        .SetFormatString "iqty", m_sQuantityFmt
       
    End With

    Exit Sub

Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Sub

Private Sub InitList2()
    On Error GoTo Err_Handler

    Set objColset2 = New U8ColumnSet.clsColSet
    objColset2.Init gConn, goLogin.cUserId

    '11.1�ϲ���ʾ
    Set VouchList2.ColSetObj = objColset2
    
    '�����б�ģ��
    objColset2.setColMode "PD010301T"
    Me.VouchList2.InitHead objColset2.getColInfo()

    '�ϼƷ�ʽ
    If Replace(strWhere, " ", "") = "(1=2)" Then
        VouchList2.SumStyle = vlSumNone '   vlRecordAndGridsum                     ' vlGridSum
    Else
        VouchList2.SumStyle = vlRecordAndGridsum
    End If

    'FillAppend ����
    'FillOverwrite ������䷽ʽ����ʹ��ѡ���ܣ�[ѡ��]����ʾΪ�̶���
    Me.VouchList2.FillMode = FillOverwrite
    VouchList2.ShowSelCol = True
    '    VouchList1.SumStyle = vlSumNone

    With VouchList2
 

        .SetFormatString "num", m_sQuantityFmt
        .SetFormatString "imqty", m_sQuantityFmt
        .SetFormatString "iwmqty", m_sQuantityFmt
        .SetFormatString "isendqty", m_sQuantityFmt
        .SetFormatString "iwsendqty", m_sQuantityFmt
       
    End With

    Exit Sub

Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Sub

Private Sub FillList()

    Me.VouchList1.SetVchLstRst Nothing
    '��������
    m_pagediv.LoadData

End Sub


Private Sub FillList2()

    Me.VouchList2.SetVchLstRst Nothing
    '��������
    m_pagediv2.LoadData

End Sub

Private Sub InitConi(mwhere As String)
    On Error GoTo Err_Handler


    Set m_pagediv = New Pagediv
    
    If m_coni Is Nothing Then
        Set m_coni = New DefaultPagedivConi
    End If
    
    'by zhangwchb 20110719 �б�γ����չ
    Set oExtend = CreateObject("VoucherExtendService.ClsExtendServer")
    Call oExtend.GetExtendInfo(gConn, gstrCardNumberlist, "L", sExtendField, sExtendJoinSQL)

    m_coni.From = "EF_V_HZDesignList" & sExtendJoinSQL   'MainView '�൱��from
    m_coni.SelectConi = objColset.GetSqlString    '�൱���ѯ�ֶ�
    m_coni.SumConi = objColset.GetSumString
    m_coni.where = " 1=1  and ismaterial='��' and (billtype='�����嵥' OR billtype='�����嵥')"
    If mwhere <> "" Then m_coni.where = m_coni.where & " and " & mwhere    '��ѯ����
    'Ȩ�޴���
'    m_coni.where = m_coni.where

'    m_coni.OrderID = "cdeptcode,cinvcode "   '�����ֶ�

    'Call PagedivCtl1.BindPagediv(m_pagediv)
    Call m_pagediv.Initialize(gConn, m_coni)
    Call VouchList1.BindPagediv(m_pagediv)
'    DropTable "tempdb..TEMP_STSearchTableNameList_" & sGUID
'    g_Conn.Execute "select id as cVoucherId,ccode as cVoucherCode,cast(null as nvarchar(1)) as cVoucherName,cast(null as nvarchar(1)) as cCardNum,cast(null as nvarchar(1)) as cMenu_Id,cast(null as nvarchar(1)) as cAuth_Id,cast(null as nvarchar(1)) as cSub_Id into tempdb..TEMP_STSearchTableNameList_" & sGUID & " from " & m_coni.From & " where 1=1 " & IIf(m_coni.where = "", "", " and " & m_coni.where)
    Exit Sub

Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Sub

Private Sub InitConi2(mwhere As String)
    On Error GoTo Err_Handler


    Set m_pagediv2 = New Pagediv
    
    If m_coni2 Is Nothing Then
        Set m_coni2 = New DefaultPagedivConi
    End If
    
    'by zhangwchb 20110719 �б�γ����չ
    Set oExtend = CreateObject("VoucherExtendService.ClsExtendServer")
    Call oExtend.GetExtendInfo(gConn, "INSCANDETAIL001", "L", sExtendField, sExtendJoinSQL)

    m_coni2.From = "EF_V_HZDesignList" & sExtendJoinSQL   'MainView '�൱��from
    m_coni2.SelectConi = objColset2.GetSqlString    '�൱���ѯ�ֶ�
    m_coni2.SumConi = objColset2.GetSumString
    m_coni2.where = " 1=1 and ismaterial='��'"
    If mwhere <> "" Then m_coni2.where = m_coni2.where & " and " & mwhere    '��ѯ����
    'Ȩ�޴���
'    m_coni2.where = m_coni2.where

'    m_coni2.OrderID = "cdeptcode,cinvcode "   '�����ֶ�

    'Call PagedivCtl1.BindPagediv(m_pagediv2)
    Call m_pagediv2.Initialize(gConn, m_coni2)
    Call VouchList2.BindPagediv(m_pagediv2)
'    DropTable "tempdb..TEMP_STSearchTableNameList_" & sGUID
'    g_Conn.Execute "select id as cVoucherId,ccode as cVoucherCode,cast(null as nvarchar(1)) as cVoucherName,cast(null as nvarchar(1)) as cCardNum,cast(null as nvarchar(1)) as cMenu_Id,cast(null as nvarchar(1)) as cAuth_Id,cast(null as nvarchar(1)) as cSub_Id into tempdb..TEMP_STSearchTableNameList_" & sGUID & " from " & m_coni2.From & " where 1=1 " & IIf(m_coni2.where = "", "", " and " & m_coni2.where)
    Exit Sub

Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set o_FilterObject = Nothing
    VouchList1.Dispose
    Set m_pagediv = Nothing
    Set m_pagediv2 = Nothing
End Sub

Private Sub m_pagediv_BeforeGetCount()
    VouchList1.FillMode = FillOverwrite
End Sub

Private Sub m_pagediv2_BeforeGetCount()
    VouchList2.FillMode = FillOverwrite
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

Private Sub m_pagediv2_GetData(ByVal vltable As U8VouchList.VouchListTable)
    On Error GoTo Err_Handler

    'Dim recclass As New ADODB.Recordset

    VouchList2.SetVchLstRst vltable.DataRecordset  '

    'Set recclass = vltable.DataRecordset
    VouchList2.SetSumRst vltable.SumRecordset
    VouchList2.RecordCount = vltable.DataCount

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

Private Sub m_pagediv2_AfterGetData(rst As ADODB.Recordset, cnt As Long)
On Error GoTo Err_Handler

    Me.VouchList2.InitHead objColset2.getColInfo()
    If Replace(strWhere, " ", "") = "(1=2)" Then
        VouchList2.SumStyle = vlSumNone '   vlRecordAndGridsum                     ' vlGridSum
    Else
        VouchList2.SumStyle = vlRecordAndGridsum
    End If

    Exit Sub
   
Err_Handler:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
    
End Sub

'ÿ�����嶼��Ҫ���������Cancel��UnloadMode�Ĳ����ĺ�����QueryUnload�Ĳ�����ͬ��
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
    ClearInScanDetailtmp
    
End Sub

Private Function ClearInScanDetailtmp() As Boolean
    On Error GoTo hErr
    gConn.Execute "drop table " & TblName
    Exit Function
hErr:
    
End Function

Private Sub PagedivCtl1_BeforeSendCommand(cmdType As U8VouchList.UFCommandType, pageSize As Long, PageCurrent As Long)
    Me.VouchList1.SetVchLstRst Nothing
    Me.VouchList1.FillMode = FillOverwrite
End Sub

Private Sub PagedivCtl2_BeforeSendCommand(cmdType As U8VouchList.UFCommandType, pageSize As Long, PageCurrent As Long)
    Me.VouchList2.SetVchLstRst Nothing
    Me.VouchList2.FillMode = FillOverwrite
End Sub

'��ӡ��Ԥ�������
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

'ȫѡ
    If flag = True Then
        VouchList1.AllSelect
        'ȫ��
    Else
        VouchList1.AllNone
    End If

End Sub

'��ѡ
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

Public Sub ExecRefresh()
'��ʼ���б�
    Call InitList
    InitList2

    '��ʼ����ҳ�ؼ�
    Call InitConi(strWhere)
    Call InitConi2(strWhere2)
    '����ҳ�ؼ����ݣ������б�����
    Call FillList
    Call FillList2
    If VouchList2.rows <= 1 Then
        gMoCode = ""
    End If
End Sub

'��ȡʱ����������ʱ����Ƚ�
Private Function ExecFunCompareUfts(Optional strcode As String) As Boolean

'��ȡʱ���
    TimeStamp = GetTimeStamp(g_Conn, MainTable, lngVoucherID)

    If TimeStamp = RecordDeleted Then
        '        MsgBox "����(" & strcode & ")�ѱ������û�ɾ��,�����޸�", vbInformation, getstring("U8.DZ.JA.Res030")
        ExecFunCompareUfts = False
        Exit Function
    ElseIf TimeStamp = RecordError Then
        '         MsgBox "����(" & strcode & ")���ݳ��ִ���,��ˢ��", vbInformation, getstring("U8.DZ.JA.Res030")
        ExecFunCompareUfts = False
        Exit Function
    ElseIf OldTimeStamp <> TimeStamp Then
        '        MsgBox "����(" & strcode & ")�ѱ������û��޸�,��ˢ��", vbInformation, getstring("U8.DZ.JA.Res030")
        ExecFunCompareUfts = False
        Exit Function
    Else
        OldTimeStamp = TimeStamp
        ExecFunCompareUfts = True
    End If
End Function

'���
Private Sub ExecConfirm(flag As Boolean, bWorkflow As Boolean)

    On Error GoTo Err_Handler

    Dim sql As String
    Dim i As Integer
    Dim cHandlervalue As String
    Dim dVeriDatevalue As String
    Dim iStatusvalue As String

    '������
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
    
    '��������֮ǰ�����øö������д�����������
    Dim calledCtx As Object
    Set calledCtx = CreateObject("UFSoft.U8.Framework.LoginContext.CalledContext")
    '

    FrmProgress.Show
  
    FrmProgress.Label1.Caption = GetString("U8.DZ.JA.Res340") & IIf(flag = True, GetString("U8.DZ.JA.Res350"), GetString("U8.DZ.JA.Res360")) & GetString("U8.DZ.JA.Res370")
    FrmProgress.ProgressBar1.Max = VouchList1.rows - 1



    '���
    If flag = True Then
        cHandlervalue = goLogin.cUserId
        dVeriDatevalue = Date
        iStatusvalue = 2
        '����
    Else
        cHandlervalue = ""
        dVeriDatevalue = "null"
        iStatusvalue = 1
    End If

    If bWorkflow Then        '��Ҫ�߹�����
        Set AuditServiceProxy = CreateObject("UFIDA.U8.Audit.ServiceProxy.AuditServiceProxy")

        calledCtx.SubId = goLogin.cSub_Id
        calledCtx.TaskId = goLogin.TaskId
        calledCtx.token = goLogin.userToken
        If flag Then        '���
            If AuditServiceProxy.ShowAuditSimpleUI(calledCtx, Action, State, strAuditOpinion) = False Then
                Exit Sub
            End If
        End If
        If Not flag Then    '����
            If Not AuditServiceProxy.ShowAuditAbandonUI(calledCtx, State, strAuditOpinion) Then
                Exit Sub
            End If
            primBizData = ""
        End If
    End If

   oldcode = ""
   oldCode4PushOtherOut = ""
    For i = 1 To VouchList1.rows - 1                 '�������
    
        bSuccess = True
        If VouchList1.TextMatrix(i, VouchList1.GetColIndex("selcol")) = "Y" Then
            If oldcode = VouchList1.TextMatrix(i, VouchList1.GetColIndex("cCODE")) Then
               GoTo DoNext
            End If
             oldcode = VouchList1.TextMatrix(i, VouchList1.GetColIndex("cCODE"))
        
            '�رյ�����ѭ��
            If VouchList1.TextMatrix(i, VouchList1.GetColIndex("CloseUser")) <> "" Then
                GoTo DoNext
            End If
            'enum by modify
            '-----------------------------------------------------------------------------------------
            If flag = False Then
                If VouchList1.TextMatrix(i, VouchList1.GetColIndex("cCreateType")) = "ת������" Then
                    '               MsgBox getstring("U8.DZ.JA.Res240"), vbInformation, getstring("U8.DZ.JA.Res030")
                    '               Exit Sub
                    GoTo DoNext
                End If

                If VouchList1.TextMatrix(i, VouchList1.GetColIndex("cCreateType")) = "�ڳ�����" Then
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
'            If ExecFunCompareUfts(VouchList1.TextMatrix(i, VouchList1.GetColIndex("cCODE"))) = False Then GoTo DoNext  'Exit Sub  'liwqa ע��
            '-----------------------------------------------------------------------------------------

            If VouchList1.TextMatrix(i, VouchList1.GetColIndex("iswfcontrolled")) = "1" Then
                If flag Then    '���
                    If VouchList1.TextMatrix(i, VouchList1.GetColIndex("iVerifyState")) = "1" Then           '���ύ
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
                        ''                                                    �������� (û���ύ)
                    ElseIf VouchList1.TextMatrix(i, VouchList1.GetColIndex("iVerifyState")) = "0" Or VouchList1.TextMatrix(i, VouchList1.GetColIndex("iVerifyState")) = "" Then
                        bSuccess = False
                        strMsg = strMsg & GetStringPara("U8.pu.prjpu860.04715", VouchList1.TextMatrix(i, VouchList1.GetColIndex("ccode"))) ', vbInformation, GetString("U8.DZ.JA.Res030")    '����{0}û���ύ���������ύ��
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
                        strMsg = strMsg & GetStringPara("U8.pu.prjpu860.04715", VouchList1.TextMatrix(i, VouchList1.GetColIndex("ccode"))) ', vbInformation, GetString("U8.DZ.JA.Res030")    '����{0}û���ύ���������ύ��
                    End If

                End If    'ƥ�����

            Else
                If (flag = True And VouchList1.TextMatrix(i, VouchList1.GetColIndex(StrcHandler)) = "") Or (flag = False And VouchList1.TextMatrix(i, VouchList1.GetColIndex(StrcHandler)) <> "") Then

                    FrmProgress.ProgressBar1.Value = i
                    DoEvents

                    sql = "Update " & MainTable & " set " & StrcHandler & "='" & cHandlervalue & "' ," & StrdVeriDate & "=" & IIf(flag = True, "'" & dVeriDatevalue & "'", "null") & " , " & StriStatus & "=" & iStatusvalue & vbCrLf & _
                            " where convert(nchar,convert(money,ufts),2)  = " & OldTimeStamp & " and " & HeadPKFld & "=" & VouchList1.TextMatrix(i, VouchList1.GetColIndex(HeadPKFld))
                    Dim lngAct As Long
                    lngAct = 0
                    gConn.Execute sql, lngAct
                    '��Ӳ������� by liwqa
                    If lngAct = 0 Then
                        bSuccess = False
                        If InStr(strMsg, VouchList1.TextMatrix(i, VouchList1.GetColIndex("cCODE"))) <= 0 Then      '�Ѿ�������ʾ�Ĳ�����ʾ
                            strMsg = strMsg & VouchList1.TextMatrix(i, VouchList1.GetColIndex("cCODE")) & GetString("U8.DZ.JA.Res290") & vbCrLf
                        End If
                    Else
                         strMsg = strMsg & GetStringPara(IIf(flag = True, "U8.DZ.JA.Res430", "U8.DZ.JA.Res480"), VouchList1.TextMatrix(i, VouchList1.GetColIndex("cCODE"))) & vbCrLf
                         'ҵ��֪ͨ
                        NotifySrvSend "HYJCGH001", "HYJCGH001" & IIf(flag, ".Audit", ".UnAudit"), VouchList1.TextMatrix(i, VouchList1.GetColIndex(HeadPKFld)), goLogin
                    End If
                Else
                    bSuccess = False
                End If
            End If    'ƥ��iswfcontrolled=��1��
            
            '��˳ɹ��Ĳ��Ƶ�(ֻ�зǹ�����������˲��Ƶ�)
            If Not bWorkflow And flag And bSuccess _
                And VouchList1.TextMatrix(i, VouchList1.GetColIndex(StrcHandler)) = "" _
                And VouchList1.TextMatrix(i, VouchList1.GetColIndex("cCreateType")) <> "�ڳ�����" Then
                '������õ�����Զ������������ⵥ
                If LCase(bPushOut) = "true" Then
                    If VouchList1.TextMatrix(i, VouchList1.GetColIndex(HeadPKFld)) <> "" And _
                        oldCode4PushOtherOut <> VouchList1.TextMatrix(i, VouchList1.GetColIndex("cCODE")) Then
                        oldCode4PushOtherOut = VouchList1.TextMatrix(i, VouchList1.GetColIndex("cCODE"))
                        strMsg = strMsg & ExecPushOtherOut(VouchList1.TextMatrix(i, VouchList1.GetColIndex(HeadPKFld))) & vbCrLf
                    End If
                End If
            End If
        End If     'ƥ��="Y"
DoNext:
    Next i

    Unload FrmProgress

    '    If flag = True Then
    '        MsgBox "�������", vbInformation, getstring("U8.DZ.JA.Res030")
    '    Else
    '        MsgBox "�������", vbInformation, getstring("U8.DZ.JA.Res030")
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



    '�ر�
    If flag = True Then
        CloseUservalue = goLogin.cUserId
        dCloseDatevalue = Date
        iStatusvalue = 4    '�ر�
        '��
    ElseIf flag = False Then
        CloseUservalue = ""
        dCloseDatevalue = "null"
        If VouchList1.TextMatrix(i, VouchList1.GetColIndex(StrIntoUser)) <> "" Then
            iStatusvalue = 3    '�Ƶ�
        ElseIf VouchList1.TextMatrix(i, VouchList1.GetColIndex(StrcHandler)) <> "" Then
            iStatusvalue = 2    '���
        Else
            iStatusvalue = 1    '�½�
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

    If GetFilterList(goLogin, o_FilterObject, "sa") = False Then Exit Sub

    '��ʼ���б�
    Call InitList
    InitList2
    strWhere2 = "1=2"
    '��ʼ����ҳ�ؼ�
    Call InitConi(strWhere)
    Call InitConi2(strWhere2)
    '����ҳ�ؼ����ݣ������б�����
    Call FillList
    Call FillList2
End Sub


Private Sub SetCatalog()
    On Error GoTo 0

    Dim sXML As String

    If objColset.setCol <> enmCancel Then
        sXML = objColset.getColInfo
        VouchList1.InitHead sXML
        
        '11.1�ϲ���ʾ
        Call SetToolbarForColumn '��Ŀ���ú��������"�ϲ���ʽ"��ť״̬
    End If
    
End Sub

Private Sub SetCatalog2()
    On Error GoTo 0

    Dim sXML As String

    If objColset2.setCol <> enmCancel Then
        sXML = objColset2.getColInfo
        VouchList2.InitHead sXML
        
        '11.1�ϲ���ʾ
'        Call SetToolbarForColumn '��Ŀ���ú��������"�ϲ���ʽ"��ť״̬
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
            strFilterID = "ST[__]��Ŀ�������б����"
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

    Dim bFlowControl    As Boolean                        '�Ƿ���������

    Dim HasWfcontrolled As Boolean                              'ѡ�����Ƿ��в��ܹ��������Ƶļ�¼��������1��û����0��

    Dim i               As Long, Selected As Boolean

    Dim errMsg          As String

    Dim oDomHead        As New DOMDocument

    Dim oDomBody        As New DOMDocument

    Dim Index           As Long

    Dim bReturn         As Boolean

    Select Case strbuttonkey

            '��ӡ
        Case sKey_Print

'            If ZwTaskExec(goLogin, AuthPrint, 1) = False Then Exit Sub
            ListPrint (1)
'            Call ZwTaskExec(goLogin, AuthPrint, 0)

            'Ԥ��
        Case sKey_Preview

'            If ZwTaskExec(goLogin, AuthPrint, 1) = False Then Exit Sub
            ListPrint (2)
'            Call ZwTaskExec(goLogin, AuthPrint, 0)

            '���
        Case sKey_Output

'            If ZwTaskExec(goLogin, AuthOut, 1) = False Then Exit Sub
            ListPrint (3)
'            Call ZwTaskExec(goLogin, AuthOut, 0)

        Case "Input"
   
            frmExcelDR.Show vbModal
            frmExcelDR.ZOrder 0
                
            Index = InStr(1, strWhere, "1=2")
           
            If Index > 0 Then
                strWhere = Replace$(strWhere, "1=2", "1=1")
            End If

            If strWhere = "" Then
                strWhere = " (1=1) "
            End If
                
            Call ExecRefresh

        Case "poappadd"
   
'            Call poAppvouchAdd
            Call ExecRefresh

            '��λ
        Case sKey_Locate
            VouchList1.locate True

            '����
        Case strKfilter, "Filter"
            ExecFilter
         
            '��Ŀ
        Case sKey_Column
            SetCatalog
            'ˢ��
            ExecRefresh

            '?����
        Case "Column2"
            SetCatalog2
            'ˢ��
            ExecRefresh
        Case sKey_Batchprint

            If VouchList1.rows <= 1 Then Exit Sub

            'ȫѡ
        Case strKSelectAll
            ExecSelectAll (True)

            '��ѡ
        Case sKey_ReverseSelection
            ReverseSelection

            'ȫ��
        Case strKUnSelectAll
            ExecSelectAll (False)

            'ˢ��
        Case sKey_Refresh
            ExecRefresh

        Case gstrHelpCode
            SendKeys "{F1}"
            
        Case "scan"
            If ZwTaskExec(goLogin, AuthEdit, 1) = False Then Exit Sub
            
            Dim Frm As New frmScan
            Frm.gaiz = Me.gaiz
            Set Frm.Frm = Me
            Frm.Show vbModal
            Frm.ZOrder 0
            ExecRefresh
            Set Frm = Nothing
            Call ZwTaskExec(goLogin, AuthEdit, 0)
        Case sKey_Deleterecord
            If ZwTaskExec(goLogin, AuthEdit, 1) = False Then Exit Sub
            DoDel
            ExecRefresh
            Call ZwTaskExec(goLogin, AuthEdit, 0)
        Case sKey_Save
            If ZwTaskExec(goLogin, AuthEdit, 1) = False Then
                MsgBox "û�в���Ȩ�ޡ�", vbInformation, "��ʾ"
                Exit Sub
            End If
            If VouchList1.rows <= 1 Then
               Exit Sub
            End If
            
            If DoSave Then
'                MsgBox "����ɹ�"
            End If
            Call ZwTaskExec(goLogin, AuthEdit, 0)
            ExecRefresh
        Case "btnclose"
            If DoClose(True) Then
                
            End If
            ExecRefresh
        Case "btnopen"
            If DoClose(False) Then
                
            End If
            ExecRefresh
    End Select

End Sub

Private Function DoClose(bClose As Boolean) As Boolean
    On Error GoTo hErr
    Dim strSql As String
    Dim id As String
    Dim i As Long
    Dim BillType As String
    Dim table As String
    For i = 1 To VouchList1.rows - 1
        If VouchList1.TextMatrix(i, VouchList1.GetColIndex("selcol")) = "Y" Then
            id = VouchList1.TextMatrix(i, VouchList1.GetColIndex("id"))
            BillType = VouchList1.TextMatrix(i, VouchList1.GetColIndex("billtype"))
            If BillType = "�����嵥" Then
                table = "hzland_materialsdetail"
            ElseIf BillType = "�����嵥" Then
                table = "hzland_lzdetail"
            ElseIf BillType = "�����嵥" Then
                table = "hzland_configurationdetail"
            End If
            If bClose Then
                strSql = "update " & table & " set status='�ر�' where id='" & id & "'"
                g_Conn.Execute strSql
            Else
                strSql = "update " & table & " set status='' where id='" & id & "'"
                g_Conn.Execute strSql
            End If
        End If
    Next
    DoClose = True
    GoTo hFinish
hErr:
    MsgBox "�����쳣��" & Err.Description, vbCritical, "��ʾ"
hFinish:
    
End Function

Private Function DoSave() As Boolean
    On Error GoTo hErr
    Dim i As Long
    Dim bSelected As Boolean
    Dim doBom As New UF_Public_base.clsBom
    For i = 1 To VouchList1.rows - 1
        If VouchList1.TextMatrix(i, VouchList1.GetColIndex("selcol")) = "Y" Then
            bSelected = True
            Exit For
        End If
    Next
    If Not bSelected Then
        MsgBox "��ѡ������嵥���塣", vbInformation, "��ʾ"
        GoTo hFinish
    End If
    bSelected = False
'    For i = 1 To VouchList2.rows - 1
'        If VouchList2.TextMatrix(i, VouchList2.GetColIndex("selcol")) = "Y" Then
'            bSelected = True
'            Exit For
'        End If
'    Next
    If VouchList2.rows <= 1 Then
        MsgBox "�޲�����Ϣ��", vbInformation, "��ʾ"
        GoTo hFinish
    End If
    If MsgBox("�Ƿ�����BOM��", vbOKCancel, "��ʾ") = vbCancel Then
        GoTo hFinish
    End If
    doBom.Init g_Conn, g_oLogin
    doBom.DoCreateBom VouchList1
    DoSave = True
    GoTo hFinish
hErr:
    MsgBox "����BOM��������쳣��" & Err.Description, vbCritical, "��ʾ"
hFinish:
    Set doBom = Nothing
End Function

Private Function GetiBatch() As String
    Dim sSql As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Dim sbw As String
    Dim iBatch As Integer
    Dim sBatch As String
    Dim sCurDate As String
    On Error GoTo hErr
    If gaiz Then
        frmDate.Show vbModal
        sCurDate = frmDate.cmDate
        If sCurDate = "" Then sCurDate = g_oLogin.CurDate
    Else
        sCurDate = g_oLogin.CurDate
    End If
    sSql = "select iBatch from  EF_InScanDetail where dDate='" & sCurDate & "' "
    rs.Open sSql, gConn
    If Not rs.BOF And Not rs.EOF Then
        iBatch = CInt(Right(rs!iBatch, 5)) + 1
        
        For i = Len(iBatch) To 5
            sbw = sbw & "0"
        Next
        GetiBatch = Format(sCurDate, "YYYYMMDD") & sbw & iBatch
    Else
        GetiBatch = Format(sCurDate, "YYYYMMDD") & "00001"
    End If
    Exit Function
hErr:
    
End Function

Private Function DoDel() As Boolean
    Dim sKey As String
    Dim i As Integer
    Dim sSql As String
    On Error GoTo hErr
    sKey = ""
    For i = 1 To VouchList2.rows - 1
        If VouchList2.TextMatrix(i, VouchList2.GetColIndex("selcol")) = "Y" Then
            sKey = sKey & "'" & VouchList2.TextMatrix(i, VouchList2.GetColIndex("id")) & "',"
        End If
    Next
    If Len(sKey) > 0 Then
        sKey = Left(sKey, Len(sKey) - 1)
        sSql = "delete " & TblName & " where id in(" & sKey & ")"
        gConn.Execute sSql
    End If
    Exit Function
hErr:
    MsgBox "ɾ��ʧ��:" & Err.Description
End Function

Private Sub VouchList1_BeforeSendCommand(cmdType As U8VouchList.UFCommandType, pageSize As Long, PageCurrent As Long)
   
    Me.VouchList1.SetVchLstRst Nothing
    Me.VouchList1.FillMode = FillOverwrite
    
    
End Sub



Private Sub VouchList1_CopySelect(bAuther As Boolean)
    bAuther = ZwTaskExec(goLogin, AuthOut, 1, True)
End Sub

Private Sub VouchList1_DblClick()
'    ExecLink
End Sub

Private Sub VouchList1_FilterClick(fldsrv As Object)
    Set o_FilterObject = fldsrv
    ExecFilter
End Sub
'wangfb 11.0ToolbarǨ�� 2012-03-31
Private Function SetToolbarVisible()
    On Error Resume Next
    With Toolbar1
        '11.0�����ݲ�֧�ֽ�����뵥�б���������
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
        '.Buttons("Cancelconfirm").Visible = False  '���� ��Ϊ ����
        .Buttons("Open").Visible = False
        .Buttons("Close").Visible = False
        '.Buttons("Confirm").Visible = False        '���� ��Ϊ ���
        .Buttons("Preview").Visible = False
        .Buttons("PrintBatch").Visible = False
        .Buttons("Help").Visible = False
        
         '�ݲ�֧�ִ�ӡģ��
        .Buttons("DesignPT").Visible = False
        .Buttons("PrintTemplate").Visible = False
        
        'toolbar Ǩ�� ��Init_Toolbar���ȡ
        .Buttons("Print").ButtonMenus(sKey_Batchprint).Visible = False   '�ݲ�֧������
    End With
    VBA.Err.Clear
End Function
Private Sub SetToolbarForColumn()
On Error GoTo ErrHandle
    If objColset Is Nothing Then Exit Sub
    If objColset.IsSupportTotalTableMerge = True Then '֧������ϲ�
        If objColset.TotalTableMerge = True Then '��������ϲ�״̬
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

Public Function AppvouchAdd() As Boolean

    On Error GoTo Err_Handler

    Dim i        As Long

    Dim oDomHead As New DOMDocument

    Dim oDomBody As New DOMDocument
 
    Dim sqlstr   As String

    Dim isselcol As Boolean

    isselcol = False
 
    idtmp = ""
 
'    For i = 1 To VouchList1.rows - 1
'
'        If VouchList1.TextMatrix(i, VouchList1.GetColIndex("selcol")) = "Y" Then
'            isselcol = True
'
'            Exit For
'
'        End If
'
'    Next
 
'    If isselcol = False Then
'
'        '
'        MsgBox "��ѡ���б�����", vbInformation, "��ʾ"
'
'        AppvouchAdd = False
'
'        Exit Function
'
'    End If
 
'    For i = 1 To VouchList1.rows - 1
'
'        If VouchList1.TextMatrix(i, VouchList1.GetColIndex("selcol")) = "Y" Then
'
'            idtmp = idtmp & VouchList1.TextMatrix(i, VouchList1.GetColIndex("id")) & ","
'
'        End If
'
'    Next
'
'    idtmp = Mid$(idtmp, 1, Len(idtmp) - 1)
     
    '����
    Set oDomHead = New DOMDocument
    Set oDomBody = New DOMDocument
     
    gConn.BeginTrans

    If WriteSCBill(oDomHead, oDomBody, gConn, goLogin, ����Ʒ��ⵥ, "0411") Then
        ' ExecCreateVoucher(oDomHead, oDomBody, gConn, goLogin, �ɹ�, �ɹ��빺��) Then    '�Ƶ�����д����״̬
         AppvouchAdd = True
        gConn.CommitTrans
        
    Else
        gConn.RollbackTrans                   '�Ƶ�ʧ�ܻع�
         AppvouchAdd = False
    End If
   
    Exit Function
 
Err_Handler:
    AppvouchAdd = False
    
    
End Function

Private Sub VouchList1_SelectAll(ByVal Selected As Boolean, IsOverWrite As Boolean)
    If Selected Then
        SetWhere2 True
    Else
        strWhere2 = "1>2"
    End If
    Call InitConi2(strWhere2)
    Call FillList2
End Sub

Private Sub VouchList1_SelectClick(ByVal Selected As Boolean)
    SetWhere2 False
    Call InitConi2(strWhere2)
    Call FillList2
End Sub


Private Sub SetWhere2(SelectAll As Boolean)
    Dim i As Long
    Dim ids As String
    For i = 1 To VouchList1.rows - 1
        If SelectAll Then
            ids = ids & "'" & VouchList1.TextMatrix(i, VouchList1.GetColIndex("id")) & "',"
        Else
            If VouchList1.TextMatrix(i, VouchList1.GetColIndex("selcol")) = "Y" Then
                ids = ids & "'" & VouchList1.TextMatrix(i, VouchList1.GetColIndex("id")) & "',"
            End If
        End If
    Next
    If Len(ids) > 0 Then
        ids = Left(ids, Len(ids) - 1)
        strWhere2 = " _parentid in(" & ids & ")"
    Else
        strWhere2 = "1>2"
    End If
End Sub
