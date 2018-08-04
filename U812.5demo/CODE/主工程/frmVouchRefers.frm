VERSION 5.00
Object = "{005151D4-33F6-471B-B651-222D4E411832}#4.4#0"; "UFFormPartner.ocx"
Object = "{AF8BBBB7-94C6-4772-B826-624478C37D6A}#1.5#0"; "UFKEYHOOK.ocx"
Object = "{9FD12F62-6922-47E1-B1AC-3615BBD3D7A5}#1.0#0"; "UFLabel.ocx"
Object = "{0532C600-D183-40A1-802B-0E09F8DD709F}#1.0#0"; "ReferMakeVouch.ocx"
Begin VB.Form frmVouchRefers 
   ClientHeight    =   8595
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   11880
   Icon            =   "frmVouchRefers.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   Begin ReferMakeVouch.ctlReferMakeVouch ctlRMV 
      Height          =   9735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   17171
   End
   Begin VB.PictureBox picbottom 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1335
      ScaleWidth      =   13815
      TabIndex        =   1
      Top             =   9120
      Visible         =   0   'False
      Width           =   13815
      Begin UFLABELLib.UFLabel lblMsg 
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   4230
         _Version        =   65536
         _ExtentX        =   7461
         _ExtentY        =   317
         _StockProps     =   111
         Caption         =   "提示：按住Ctrl+鼠标为多选；按住Shift+鼠标为全选"
         BackStyle       =   0
      End
   End
   Begin UFKeyHook.UFKeyHookCtrl UFKeyHookCtrl1 
      Left            =   2640
      Top             =   3480
      _ExtentX        =   1905
      _ExtentY        =   529
   End
   Begin UFFormPartner.UFFrmCaption UFFrmCaptionMgr 
      Left            =   5280
      Top             =   3360
      _ExtentX        =   1508
      _ExtentY        =   873
      Caption         =   ""
      DebugFlag       =   0   'False
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
End
Attribute VB_Name = "frmVouchRefers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************
'  软件著作权：北京用友软件集团有限公司
'    系统名称：销售管理
'    模块名称：frmNewCZ
'    功能说明：单据参照
'    相关文档：
'        作者：zzg
'*****************************************************

Option Explicit
'Dim m_strCardNumber As String
'Dim moColSet As U8ColumnSet.clsColSet  'U8colset.clsColSet    '栏目设置对象
Dim strBusType As String       '业务类型rrrrrr

Public strCardNum As String
Dim strcInvoiceCompanyCode As String '二次参照时带入的发票上的客户编码
Dim strcInvoiceAbbCompany As String
Dim strVouchType As String
Private oDom As DOMDocument, oDomB As DOMDocument
Private bReferAgo As Boolean
'Dim RowCount  As Long
'Dim mstrcomp As String
'Dim mintstr As Integer
Dim strCode As String
Dim strcodedep As String
Dim strOppCode As String
Dim strcodeps As String
Private strbCredit As String '是否为立账单据
Private msInvoiceCompany As String '870 客户开票单位
''Dim strcodeSO As String
'Dim iTaxRate As Double
Dim cExch_Name As String
'Dim mbmoving As Boolean
''Dim bRefer As Boolean
Public bcancel As Boolean
Public bReturnVouch As Boolean      ''是否红字单据
Public cKeyCodeT As String
Public cKeyCodeB As String
Public dMin As Date
Public Domhead As DOMDocument, Dombody As DOMDocument  ''组织新的dom
Private strClientCode As String
Private strClientFather As String
Private strInvCode As String
Private sWhere As String
Private U8Refer  As New U8RefService.IService
Dim ReferSQLString As String
'Private clsShowRef As New USSAServer.clsShowRef
Dim strMakerAuth As String
Dim strVouchDate As String
Dim CheckBlueFun As Boolean
Dim mpageCount As Long
Dim mCurrentPage As Long

'Public Strrelationship As String
'Public StrTaskcode As String
'Public StrTaskname As String
'Public StrTasktypecode As String
'Public StrOutPutcode As String
'Public StrOutputname As String
'Public StrOutPuttype As String
'Public IntAutoid As Long
'Private clsVouchModel As New EFZZModel.clsVouchLoad

Public WithEvents clsReferVoucher As EFVoucherMo.clsAutoReferVouch  '' clsSaReferVoucher
Attribute clsReferVoucher.VB_VarHelpID = -1
'Dim m_SelectedStr As String


Private Function GetCol(strFieldName As String, bHead As Boolean) As Long
    Dim ele As IXMLDOMElement
    strFieldName = LCase(strFieldName)
    If bHead Then
        GetCol = ctlRMV.HeadList.GetColIndex(strFieldName)
    Else
        GetCol = ctlRMV.BodyList.GetColIndex(strFieldName)
    End If
    If Not ele Is Nothing Then
        GetCol = CLng(val(ele.getAttribute("iColPos")))
    End If
End Function


Public Function OpenFilter() As Boolean
    OpenFilter = Filter
End Function



Private Sub SelectAll(handle As Boolean, section As String)
    Dim i As Integer
    
    Dim bChange As Boolean
    Dim sErr As String
    If handle Then
        If UCase(section) = "T" Then
            If Me.ctlRMV.HeadList.Rows = 1 Then Exit Sub
            ctlRMV.RemoveBodyAll
            For i = 1 To Me.ctlRMV.HeadList.Rows - 1
                ctlRMV_onHeadSelecting i, bChange, sErr
            Next
        End If
    Else
        If UCase(section) = "T" Then
            ctlRMV.RemoveBodyAll
            clsReferVoucher.SelectedStr = ""
        End If
    End If
    ctlRMV.BodyList.RecordCount = ctlRMV.BodyList.GetGridBody().Rows - 2
End Sub

Private Sub SetVouchDom()
    On Error Resume Next
    Dim i As Integer, j As Integer
    Dim strFilter() As String, strFilter2() As String
    Dim intType As Integer
    Dim errMsg As String
    Dim strXml As String
    Dim domHeads As New DOMDocument
    
    Set domHeads = New DOMDocument
    strXml = "<?xml version='1.0' encoding='UTF-8'?><head>" & Chr(13)
    strXml = strXml + "<ddate>" & strVouchDate & "</ddate>"
    strXml = strXml + "<breturnflag>" & IIf(bReturnVouch, "1", "0") & "</breturnflag>"
    strXml = strXml + "</head>"
    domHeads.loadXML strXml
    bcancel = False
        
    Set Domhead = ctlRMV.GetHeadDom(True)
    Set Dombody = ctlRMV.GetBodyDom(True)
    
        ''''''组织Dom,需要重新写代码  2009/03/20
'    clsReferVoucher.RemoveHeadLines domHead, Dombody
'    If clsShowRef.GetRefHeadDom(CInt(iType), domHead, domHeads, errMsg) Then
'    Else
'        bcancel = True
'        MsgBox errMsg
'    End If
    Set domHeads = Nothing
End Sub



Private Sub comYes()
On Error Resume Next
    Dim i As Integer
    For i = 1 To Me.ctlRMV.HeadList.Rows - 1
        If Me.ctlRMV.HeadList.TextMatrix(i, GetCol("selcol", True)) <> "" Then
            GoTo SelectTrue
        End If
    Next
    MsgBox GetString("U8.SA.xsglsql_2.unselectedrow") '"没有选择表体行"
    Exit Sub
SelectTrue:
    bcancel = 0
    Form_KeyDown 13, 0
End Sub


Private Sub ctlRMV_BodyShiftSelect(ByVal lFromRow As Long, ByVal lToRow As Long, other As Variant)
    If ctlRMV.BodyList.Rows = 1 Then Exit Sub
End Sub

Private Sub ctlRMV_EdtAccepted()
    LoadMainDatas
End Sub

Private Sub ctlRMV_HeadBrowUser(RetValue As Variant, row As Long, col As Long)
On Error GoTo ErrHandle
    Exit Sub
ErrHandle:
    RetValue = ""
End Sub

Private Sub ctlRMV_HeadCellValueChanged(ByVal row As Integer, ByVal col As Integer, NewValue As Variant, OldValue As Variant, KeepFocus As Boolean)
    Dim lColInvoiceCompany As Long
    If Me.ctlRMV.HeadList.TextMatrix(row, GetCol("selcol", True)) = "Y" Then
        MsgBox GetString("U8.SA.xsglsql_2.frmVouchRefers.checkinput001"), vbOKOnly + vbInformation '不能修改选中的行！
        NewValue = OldValue
        KeepFocus = False
        Exit Sub
    End If
    lColInvoiceCompany = GetCol("cinvoicecompany", True)
    If lColInvoiceCompany = col Or col = GetCol("cinvoicecompanyabbname", True) Then
        If NewValue = "" Then
            MsgBox GetString("U8.SA.xsglsql_2.Error.InvoiceCorpIsNull"), vbOKOnly + vbInformation
            NewValue = OldValue
            KeepFocus = True
        Else
            If Not CheckInput(CStr(NewValue), row) Then
                NewValue = OldValue
                KeepFocus = True
            Else
                NewValue = ctlRMV.HeadList.TextMatrix(row, col)
            End If
        End If
    End If

    
End Sub
Private Function CheckInput(strInputValue As String, ByVal row As Integer) As Boolean
'    Dim clsAuth As New U8RowAuthsvr.clsRowAuth
    Dim strAuth As String
    Dim strFilter As String
    Dim rst As New ADODB.Recordset
    
'    clsSAWeb.clsAuth.Init DBConn.ConnectionString, m_Login.cUserId, False, "SA"
    rst.CursorLocation = adUseClient
    CheckInput = False
    If Not clsSAWeb.bAdmin Then
        If clsSAWeb.bAuth_Cus Then
            strAuth = clsSAWeb.clsAuth.getAuthString("CUSTOMER", , "W")
            If strAuth = "1=2" Then
                MsgBox GetString("U8.SA.xsglsql.01.frmbillvouch.00137"), vbExclamation 'zh-CN：没有客户权限！
                Set rst = Nothing
'                Set clsAuth = Nothing
                Exit Function
            ElseIf strAuth <> "" Then
                strFilter = " ccuscode in (select ccuscode from customer where iid in (" & strAuth & "))"
            End If
        End If
    End If
 
    rst.Open "select ccuscode,ccusabbname from customer where (ccuscode='" + strInputValue + "' or ccusabbname='" + strInputValue + "') and isnull(dEndDate,'9999-12-31')>'" + CStr(m_Login.CurDate) + "' and ccuscode in (select cinvoicecompany as ccuscode from sa_invoicecustomers where ccuscode =(select ccuscode from customer where ccusabbname='" + ctlRMV.HeadList.TextMatrix(row, GetCol("ccusabbname", True)) + "'))" & IIf(strAuth = "", "", " and ") & strFilter, DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rst.EOF Then
        MsgBox GetString("U8.SA.USSASERVER.clscommcheck.00410"), vbExclamation
        CheckInput = False
    Else
        Me.ctlRMV.HeadList.TextMatrix(row, GetCol("cinvoicecompany", True)) = rst.Fields("ccuscode")
        ctlRMV.HeadList.TextMatrix(row, GetCol("cinvoicecompanyabbname", True)) = rst.Fields("ccusabbname")
        CheckInput = True
    End If
    rst.Close
    Set rst = Nothing
'    Set clsAuth = Nothing
End Function

'add by renlb20090410
Private Sub ctlRMV_HeadDblClick()
'    Dim i As Integer
'    Dim bChange As Boolean
'    Dim sErr As String
'    Dim cusCode As String
'
'
'    With ctlRMV
'        If .HeadList.Rows = 1 Then Exit Sub
'
'        For i = 1 To .HeadList.Rows - 1
'            If cusCode <> "" And cusCode <> .HeadTextMatrix(i, "ccuscode") Then
'                ctlRMV.HeadList.TextMatrix(i, ctlRMV.GetHeadColIndex("selcol")) = ""
'            End If
'        Next i
'    End With
End Sub

Private Sub ctlRMV_HeadShiftSelect(ByVal lFromRow As Long, ByVal lToRow As Long, other As Variant)
    If ctlRMV.HeadList.Rows = 1 Then Exit Sub
    Dim i As Integer
    Dim bChange As Boolean
    Dim sErr As String
    Dim dom As DOMDocument
    Dim ele As IXMLDOMElement
    Set dom = other
    
    For i = lFromRow To lToRow
        Set ele = dom.documentElement.selectSingleNode("r[@irow='" + CStr(i) + "']")
        If Not ele Is Nothing Then
            If ele.Attributes.getNamedItem("oldsel").nodeValue = "" Then
                ctlRMV_onHeadSelecting i, bChange, sErr
            End If
        Else
            ctlRMV_onHeadSelecting i, bChange, sErr
        End If
    Next
    Set dom = Nothing
'    ctlRMV.BodyList.RecordCount = ctlRMV.BodyList.rows - 1
End Sub

Private Sub ButtonClick(ByVal sBtnKey As String, sErr As String)
Dim txtautoIquantity As String
Dim txtautoMoney As String
Dim strFilter As String

    Select Case sBtnKey
    Case "tlbTest"
'        frmAutoMatch.Show 1
'        If frmAutoMatch.bcancel = False Then
'            strClientCode = frmAutoMatch.strClientCode
'            strInvCode = frmAutoMatch.strInvCode
'            txtautoIquantity = frmAutoMatch.txtautoIquantity
'            txtautoMoney = frmAutoMatch.txtautoMoney
'            If strClientCode <> "" Then
'                strFilter = " ccuscode='" + strClientCode + "'"
'            End If
'            If strInvCode <> "" Then
'                strFilter = strFilter & IIf(strFilter = "", "", " and ") & " cinvcode='" + strInvCode + "'"
'            End If
'            Dim tmpSumStyle As Variant
'            tmpSumStyle = Me.ctlRMV.BodyList.SumStyle
'            Me.ctlRMV.BodyList.SumStyle = vlSumNone
'            LoadMainDatas strFilter
'
'            Me.ctlRMV.BodyList.SumStyle = tmpSumStyle
'            strClientCode = ""
'            strInvCode = ""
'        End If
    Case "tlbExit"
        Unload Me
    Case "tlbMakeVouch"
        comYes
    Case "tlbSel&tlbHead", "tlbSel"
        Call SelectAll(True, "T")
    Case "tlbSel&tlbBody"
        Call SelectAll(True, "B")
    Case "tlbUnSel&tlbHead", "tlbUnSel"
        Call SelectAll(False, "T")
'        SelectedStr = ""
    Case "tlbUnSel&tlbBody"
        Call SelectAll(False, "B")
    Case "tlbFilter&tlbBody"
    
    Case "tlbLM&tlbHead", "tlbLM"
        If clsReferVoucher.ColumnSet(True, ctlRMV) Then
            LoadMainDatas
        End If
    Case "tlbLM&tlbBody"
        If clsReferVoucher.ColumnSet(False, Me.ctlRMV) Then
            LoadMainDatas
        End If
    Case "tlbFilter&tlbHead", "tlbFilter"
        If Filter Then
            LoadMainDatas
        End If
    Case "tlbRefresh&tlbHead", "tlbRefresh", "tlbRefresh&tlbBody"
        LoadMainDatas
    Case "tlbFilterSetup&tlbBody"
    Case "tlbFilterSetup&tlbHead", "tlbFilterSetup"
        clsReferVoucher.SetFilter
    Case "tlbFirst"
        If mCurrentPage > 1 Then
            mCurrentPage = 1
            LoadMainDatas
        End If
    Case "tlbNext"
        If mCurrentPage <= mpageCount - 1 Then
            mCurrentPage = mCurrentPage + 1
            LoadMainDatas
        End If
    Case "tlbLast"
        If mCurrentPage > 0 And mCurrentPage < mpageCount Then
            mCurrentPage = mpageCount
            LoadMainDatas
        End If
    Case "tlbPrevious"
        If mCurrentPage > 1 Then
            mCurrentPage = mCurrentPage - 1
            LoadMainDatas
        End If
    Case "cmdAuto"
        MsgBox "cmdAuto"
    End Select
End Sub

Private Sub ctlRMV_onHeadSelecting(ByVal row As Long, bChange As Boolean, sErr As String)
    Dim i As Long
    Dim fltStr As String
    Dim sqdcode As String
    'add by renlb20090410
    
        Dim iRow As Long
    If ctlRMV.HeadTextMatrix(row, ctlRMV.GetHeadColIndex("selcol")) <> "" Then
        For iRow = 1 To ctlRMV.HeadList.RecordCount
            If row <> iRow And ctlRMV.HeadTextMatrix(iRow, ctlRMV.GetHeadColIndex("selcol")) <> "" Then
                ctlRMV.HeadTextMatrix(iRow, ctlRMV.GetHeadColIndex("selcol")) = ""
                Call OneHeadSelecting(iRow, False)
                clsReferVoucher.SetBodyData Me.ctlRMV, iRow, True
            End If
        Next
        'MsgBox ctlRMV.HeadList.row
    End If
    
    With ctlRMV
    For i = 1 To .HeadList.Rows - 1
        If sqdcode <> "" And sqdcode <> .HeadTextMatrix(row, ctlRMV.GetHeadColIndex("ccuscode")) Then
            .HeadTextMatrix(i, ctlRMV.GetHeadColIndex("selcol")) = ""
            Exit Sub
        End If
        If .HeadTextMatrix(i, ctlRMV.GetHeadColIndex("selcol")) <> "" And sqdcode = "" Then sqdcode = .HeadTextMatrix(i, ctlRMV.GetHeadColIndex("ccuscode"))
    Next i
    End With
    'add end-----
    
    Call OneHeadSelecting(row, bChange)
    If Not clsReferVoucher.blnSigleColumn Then
        clsReferVoucher.SetBodyData Me.ctlRMV, row, True
    End If
End Sub

Private Function OneHeadSelecting(ByVal row As Long, bChange As Boolean, Optional blnShowMsg As Boolean = True) As Boolean

    With Me.ctlRMV.HeadList
        If ctlRMV.HeadList.TextMatrix(row, ctlRMV.GetHeadColIndex("selcol")) = "" Then
            clsReferVoucher.RemoveSelected ctlRMV, row
            Exit Function
        End If
        OneHeadSelecting = True
        clsReferVoucher.SelectedStr = clsReferVoucher.SelectedStr & " " & CStr(row) & ","
    End With
    Exit Function
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        SetVouchDom
        Unload Me
    End If
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim strself As String
    Dim i As Integer
    '************************
    'u861多语Layout
'    Call MultiLangLoadForm(Me, "u8.sa.xsglsql." & Me.Name)
    '************************
    
    On Error Resume Next

    bcancel = True
    Me.HelpContextID = 10151180

    mCurrentPage = 1
    
    '设置显示表体为不显示

    dMin = CDate("1900-1-1")
    

    
'    If UCase(strCardNum) = "EFZZ0606" Then
'        Select Case Strrelationship
'            Case "销售订单"
'                clsReferVoucher.Init strCardNum, "efzz060601"
'            Case "采购请购单"
'                clsReferVoucher.Init strCardNum, "efzz060602"
'            Case "委外请购单"
'                clsReferVoucher.Init strCardNum, "efzz060602"
'            Case "生产订单"
'                clsReferVoucher.Init strCardNum, "efzz060603"
'            Case "子件发运单"
'
'        End Select
'    Else
'        clsReferVoucher.Init strCardNum, "efzz0402"
'    End If
    
    clsReferVoucher.InitReferVoucher Me.ctlRMV
    Me.ctlRMV.pageSize = clsReferVoucher.pageSize
    cKeyCodeT = clsReferVoucher.strMainKey
    cKeyCodeB = clsReferVoucher.strDetailKey
    LoadMainDatas
    strVouchDate = GetHeadItemValue(oDom, "ddate")
    Set Me.Icon = frmMain.Icon
    UFFrmCaptionMgr.Caption = GetString("U8.SO.VOUCH.copytabst.00716")
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    strClientCode = ""
    strClientFather = ""
    strInvCode = ""
    sWhere = ""
'    Set clsShowRef = Nothing
    clsReferVoucher.pageSize = Me.ctlRMV.pageSize
    clsReferVoucher.SavePageSize
    Set U8Refer = Nothing
    BReferAgain = False
'    Call U861ResEnd
    Me.ctlRMV.HeadList.GetGridBody().Dispose
    Me.ctlRMV.BodyList.GetGridBody().Dispose
End Sub

Private Function Filter() As Boolean
    Dim objFilter As New UFGeneralFilter.FilterSrv
    Dim objFltCb As New clsReportCallBack
    
    Dim cbustypeWhere As String
    Dim m_ckey As String
    Filter = False
    m_ckey = clsReferVoucher.StrFilterName
    objFltCb.StrReportName = m_ckey
    Set objFilter.BehaviorObject = objFltCb
    
'    Select Case UCase(strCardNum)
'        Case "EFZZ0606"
'            If Strrelationship = "销售订单" Then
'                Filter = objFilter.OpenFilter(m_Login, "", "efzz060601", "EF", "")
'            End If
'            If Strrelationship = "采购请购单" Or Strrelationship = "委外请购单" Then
'                Filter = objFilter.OpenFilter(m_Login, "", "efzz060602", "EF", "")
'            End If
'            If Strrelationship = "生产订单" Then
'                Filter = objFilter.OpenFilter(m_Login, "efzz060603", "", "EF", "")
'            End If
'            If Strrelationship = "子件发运单" Then
'                Filter = objFilter.OpenFilter(m_Login, "", m_ckey, "EF", "")
'            End If
'        Case Else
'    End Select
    
'    Filter = True
    Filter = objFilter.OpenFilter(m_Login, m_ckey, "", "", "")
    If Filter Then
'        sWhere = convertWhere(objFilter)
        sWhere = objFilter.GetSQLWhere()
        Filter = True
    End If
    Set objFilter = Nothing
    Set objFltCb = Nothing
End Function

''取得rst中的字段值，将null转换为0
Private Function GetRstVal(rst As ADODB.Recordset, FieldName As String) As Variant
    If IsNull(rst(FieldName)) = True Then
        If rst(FieldName).Type = adChar Or rst(FieldName).Type = adVarChar Or rst(FieldName).Type = adDate _
        Or rst(FieldName).Type = adDBDate Or rst(FieldName).Type = adDBTime Or rst(FieldName).Type = adDBTimeStamp _
        Or rst(FieldName).Type = adVarWChar Or rst(FieldName).Type = adLongVarChar Or rst(FieldName).Type = adLongVarWChar _
        Or rst(FieldName).Type = adWChar Or rst(FieldName).Type = adBSTR Then
            GetRstVal = ""
        Else
            GetRstVal = 0
        End If
    Else
        If rst(FieldName).Type = adBoolean Then
            GetRstVal = IIf(rst(FieldName), "是", "否")
        ElseIf rst(FieldName).Type = adDate Or rst(FieldName).Type = adDBDate Or rst(FieldName).Type = adDBTime Then
            GetRstVal = Format(rst(FieldName), "yyyy-mm-dd")
        Else
            GetRstVal = rst(FieldName)
        End If
    End If
    
End Function



Public Property Get BusType() As String
    BusType = strBusType
End Property

Public Property Let BusType(ByVal vNewValue As String)
    strBusType = vNewValue
End Property



Public Property Get VouchType() As String
    BusType = strVouchType
End Property

Public Property Let VouchType(ByVal vNewValue As String)
    strVouchType = vNewValue
End Property

Public Property Let VouchDOM(ByVal vNewValue As DOMDocument)
    Set oDom = vNewValue
End Property
Public Property Let BReferAgain(ByVal vNewValue As Boolean)
    bReferAgo = vNewValue
End Property
Public Property Let VouchDOMB(ByVal vNewValue As DOMDocument)
    Set oDomB = vNewValue
End Property


Private Sub ctlRMV_onButtonClick(ByVal sBtnKey As String, sErr As String)
    If ctlRMV.HeadList.ProtectUnload() <> 2 Then Exit Sub
    Call ButtonClick(sBtnKey, sErr)
End Sub
Private Sub Form_Resize()
On Error Resume Next
    Me.ctlRMV.Width = Me.ScaleWidth
    Me.ctlRMV.Height = Me.ScaleHeight
'    ctlRMV.HeadList.SetFocus
End Sub




Private Sub SetBottonState(pageCount As Long)
    With Me.ctlRMV
        If pageCount > 1 Then
             .Toolbar.buttons("tlbFirst").Enabled = True
             .Toolbar.buttons("tlbNext").Enabled = True
             .Toolbar.buttons("tlbLast").Enabled = True
             .Toolbar.buttons("tlbPrevious").Enabled = True
        Else
             .Toolbar.buttons("tlbFirst").Enabled = False
             .Toolbar.buttons("tlbNext").Enabled = False
             .Toolbar.buttons("tlbLast").Enabled = False
             .Toolbar.buttons("tlbPrevious").Enabled = False
        End If
        .UFToolbar.RefreshVisible
    End With
End Sub

Private Sub clsReferVoucher_AfterGetBodyDatas(lngRow As Long, domBodyDatas As MSXML2.DOMDocument, strError As String)
    Dim Domhead As DOMDocument
    Dim strXml As String
    Dim domHeads As New DOMDocument
    
    Set domHeads = New DOMDocument
    strXml = "<?xml version='1.0' encoding='UTF-8'?><head>" & Chr(13)
    strXml = strXml + "<ddate>" & strVouchDate & "</ddate>"
    strXml = strXml + "<breturnflag>" & IIf(bReturnVouch, "1", "0") & "</breturnflag>"
    strXml = strXml + "</head>"
    domHeads.loadXML strXml
    Set Domhead = ctlRMV.GetHeadLine(lngRow)
'    2009/03/20
'    clsShowRef.GetReferBodyDom CInt(iType), domHead, domBodyDatas, domHeads, strError
End Sub

'Public Property Get strCardNumber() As String
'    strCardNumber = m_strCardNumber
'End Property
'
'Public Property Let strCardNumber(ByVal vNewValue As String)
'    m_strCardNumber = vNewValue
'End Property

'Public Property Get SelectedStr() As String
'    SelectedStr = m_SelectedStr
'End Property
'
'Public Property Let SelectedStr(ByVal vNewValue As String)
'    m_SelectedStr = vNewValue
'End Property

Private Sub LoadMainDatas(Optional strOtherFilter As String)
    Dim eleTmp As IXMLDOMElement
    Dim sTmpWhere As String

    Me.ctlRMV.RemoveHeadAll
    
    Dim filterWhere As String '过滤条件的 filterWhere
    'clsReferVoucher.strFilter
    If strOtherFilter <> "" Then
        filterWhere = sWhere & IIf(sWhere = "", "", " and (") & strOtherFilter & ")"
    Else
'        If sWhere <> "" Then
            filterWhere = sWhere
'        Else
'            sWhere = clsReferVoucher.strFilter
'            filterWhere = clsReferVoucher.strFilter
'        End If
    End If
'    If m_Login.IsAdmin = False Then
'        strMakerAuth = clsShowRef.GetRefAuthStringForRef(, , CInt(iType), , True)
'        filterWhere = filterWhere & IIf(strMakerAuth = "", "", " and  (" & strMakerAuth & ")")
'    End If

    If filterWhere = "(1 = 1)" Then
        filterWhere = ""
    Else
        filterWhere = Replace(filterWhere, "(1=1)  and", "")
    End If
    clsReferVoucher.strFilter = filterWhere
    clsReferVoucher.SetHeadData Me.ctlRMV, ctlRMV.pageSize, mCurrentPage, mpageCount
    ctlRMV.RemoveBodyAll
    SetBottonState mpageCount
    ctlRMV.HeadList.SetFocus
End Sub


