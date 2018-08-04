VERSION 5.00
Object = "{005151D4-33F6-471B-B651-222D4E411832}#4.5#0"; "UFFormPartner.ocx"
Object = "{AF8BBBB7-94C6-4772-B826-624478C37D6A}#1.5#0"; "UFKEYHOOK.ocx"
Object = "{9FD12F62-6922-47E1-B1AC-3615BBD3D7A5}#1.0#0"; "UFLabel.ocx"
Object = "{0532C600-D183-40A1-802B-0E09F8DD709F}#1.6#0"; "ReferMakeVouch.ocx"
Begin VB.Form frmVouchRefers 
   ClientHeight    =   9015
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   13245
   Icon            =   "frmVouchRefers.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9015
   ScaleWidth      =   13245
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
         _ExtentY        =   423
         _StockProps     =   111
         Caption         =   "提示：按住Ctrl+鼠标为多选；按住Shift+鼠标为全选"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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

Dim mpageCount As Long
Dim mCurrentPage As Long

'存储参照数据
Public domHead As DOMDocument, domBody As DOMDocument      ''组织新的dom
Public clsReferVoucher As clsReferVoucher
Private sWhere As String
Public bcancel As Boolean

'窗体级变量-自定义
Dim strcode As String
Dim strcodedep As String
Dim strOppCode As String
Dim strcodeps As String
Dim cexch_name As String


'获取列的位置
Private Function GetCol(strFieldName As String, bHead As Boolean) As Long
    Dim ele As IXMLDOMElement
    strFieldName = LCase(strFieldName)
    If bHead Then
        GetCol = ctlRMV.HeadList.GetColIndex(strFieldName)
    Else
        GetCol = ctlRMV.BodyList.GetColIndex(strFieldName)
    End If
    If Not ele Is Nothing Then
        GetCol = CLng(Val(ele.getAttribute("iColPos")))
    End If
End Function

'参照
Public Function OpenFilter() As Boolean
    OpenFilter = Filter
End Function

Private Sub SelectAll(handle As Boolean, section As String)
    Dim i As Integer
    Dim bChange As Boolean
    Dim sErr As String
    If handle Then
        If UCase(section) = "T" Then
            If Me.ctlRMV.HeadList.rows = 1 Then Exit Sub
            ctlRMV.RemoveBodyAll
            For i = 1 To Me.ctlRMV.HeadList.rows - 1
                ctlRMV_onHeadSelecting i, bChange, sErr
            Next
        End If
    Else
        If UCase(section) = "T" Then
            ctlRMV.RemoveBodyAll
            clsReferVoucher.SelectedStr = ""
        End If
    End If
    ctlRMV.BodyList.RecordCount = ctlRMV.BodyList.GetGridBody().rows - 2
End Sub

'保存并整理dom
Private Sub SetVouchDom()
    On Error Resume Next

    Set domHead = ctlRMV.GetHeadDom(True)
    Set domBody = ctlRMV.GetBodyDom(True)
    clsReferVoucher.RemoveHeadLines domHead, domBody

End Sub

'确定
Private Sub comYes()

    On Error Resume Next

    Dim i      As Integer

    Dim j      As Long

    Dim isflag As Boolean

    isflag = False
    
    If Me.ctlRMV.HeadList.rows - 1 >= 2 Then

        For i = 1 To Me.ctlRMV.HeadList.rows - 2

            If isflag = False Then
                If ctlRMV.HeadList.TextMatrix(i, ctlRMV.GetHeadColIndex("selcol")) = "Y" Then
          
                    For j = 2 To Me.ctlRMV.HeadList.rows - 1
                   
                        If ctlRMV.HeadList.TextMatrix(j, ctlRMV.GetHeadColIndex("selcol")) = "Y" Then
                     
                            If ctlRMV.HeadList.TextMatrix(i, ctlRMV.GetHeadColIndex("ecustcode")) <> ctlRMV.HeadList.TextMatrix(j, ctlRMV.GetHeadColIndex("ecustcode")) Then
                       
                                isflag = True
                                MsgBox "选择了不同的客户编码"

                                Exit For

                            End If
                        End If
                   
                    Next

                End If

            Else

                Exit For
          
            End If
       
        Next
     End If
     If isflag = True Then
        For j = 0 To Me.ctlRMV.HeadList.rows - 1
            ctlRMV.HeadList.TextMatrix(j, ctlRMV.GetHeadColIndex("selcol")) = ""
          
         Next
           Exit Sub
      End If
         
    
    
        '    For i = 1 To Me.ctlRMV.BodyList.rows - 1
        If Me.ctlRMV.HeadList.TextMatrix("selcol") <> "" Then
            GoTo SelectTrue
        End If

        '    Next
        MsgBox " 没有选择数据", vbInformation, GetString("U8.DZ.JA.Res030")    '"没有选择表体行"

        Exit Sub

SelectTrue:
        bcancel = 0

        Form_KeyDown 13, 0
    End Sub

Private Sub ctlRMV_BodyDblClick()
    'Dim i As Long
    '
    '  For i = 1 To Me.ctlRMV.BodyList.rows - 1
    '    If i = 1 Then
    '      ctlRMV.BodyList.TextMatrix(i, ctlRMV.GetHeadColIndex("selcol")) = "Y"
    '    Else
    '     If ctlRMV.BodyList.TextMatrix(i, ctlRMV.GetHeadColIndex("cwhcode")) = ctlRMV.BodyList.TextMatrix(i, ctlRMV.GetHeadColIndex("cwhcode")) Then
    '        ctlRMV.BodyList.TextMatrix(i, ctlRMV.GetHeadColIndex("selcol")) = "Y"
    '       Else
    '       MsgBox "", vbInformation, "asd"
    '      End If
    '    End If
    '
    '   Next

End Sub

Private Sub ctlRMV_EdtAccepted()
    LoadMainDatas clsReferVoucher.OtherFilter
End Sub

Private Sub ButtonClick(ByVal sBtnKey As String, sErr As String)
    Dim txtautoiquantity As String
    Dim txtautoMoney As String
    Dim strFilter As String
    Dim i As Long
    Dim cwhcode As String
    cwhcode = ""

    Select Case sBtnKey

        Case "tlbExit"
            Unload Me
        Case "tlbMakeVouch"
            comYes
            '    Case "tlbSel&tlbHead", "tlbSel"
            '        Call SelectAll(True, "T")
        Case "tlbSel&tlbBody"
            'V11改用生单控件的控制规则方式
'            '检查表体是否有选择不同的仓库
'            'begin
'            If iSinvCZ = True Then
'                For i = 1 To Me.ctlRMV.BodyList.rows - 1
'                    If i = Me.ctlRMV.BodyList.rows - 1 Then
'                        cwhcode = ctlRMV.BodyList.TextMatrix(i, ctlRMV.GetBodyColIndex("cwhcode"))
'                    Else
'                        cwhcode = ctlRMV.BodyList.TextMatrix(i + 1, ctlRMV.GetBodyColIndex("cwhcode"))
'                    End If
'                    If cwhcode <> ctlRMV.BodyList.TextMatrix(i, ctlRMV.GetBodyColIndex("cwhcode")) Then
'                        MsgBox GetString("U8.DZ.JA.Res1870"), vbExclamation, GetString("U8.DZ.JA.Res030")
'                        ctlRMV.BodyList.AllNone
'                        Exit Sub
'                    End If
'
'                Next
'                '  ctlRMV.BodyList.TextMatrix(row, ctlRMV.GetBodyColIndex("selcol")) = "Y"
'            End If


            'end
            Call SelectAll(True, "B")
        Case "tlbUnSel&tlbHead", "tlbUnSel"
            Call SelectAll(False, "T")
            '        SelectedStr = ""
        Case "tlbUnSel&tlbBody"
            Call SelectAll(False, "B")
        Case "tlbFilter&tlbBody"
            OpenFilter
        Case "tlbLM&tlbHead", "tlbLM"
            If clsReferVoucher.ColumnSet(True, ctlRMV) Then
                LoadMainDatas clsReferVoucher.OtherFilter
            End If
        Case "tlbLM&tlbBody"
            If clsReferVoucher.ColumnSet(False, Me.ctlRMV) Then
                LoadMainDatas clsReferVoucher.OtherFilter
            End If
        Case "tlbFilter&tlbHead", "tlbFilter"
            If Filter Then
                LoadMainDatas clsReferVoucher.OtherFilter
            End If
        Case "tlbRefresh&tlbHead", "tlbRefresh", "tlbRefresh&tlbBody"
            LoadMainDatas clsReferVoucher.OtherFilter
        Case "tlbFilterSetup&tlbBody"
            '        Call objfltint.SetFilter("SARefVouchB" & iType, "SA", m_connLst, , False)
            '        objfltint.DeleteFilter
            '        Set objfltint = Nothing
        Case "tlbFilterSetup&tlbHead", "tlbFilterSetup"
            clsReferVoucher.SetFilter
        Case "tlbFirst"
            If mCurrentPage > 1 Then
                mCurrentPage = 1
                LoadMainDatas clsReferVoucher.OtherFilter
            End If
        Case "tlbNext"
            If mCurrentPage <= mpageCount - 1 Then
                mCurrentPage = mCurrentPage + 1
                LoadMainDatas clsReferVoucher.OtherFilter
            End If
        Case "tlbLast"
            If mCurrentPage > 0 And mCurrentPage < mpageCount Then
                mCurrentPage = mpageCount
                LoadMainDatas clsReferVoucher.OtherFilter
            End If
        Case "tlbPrevious"
            If mCurrentPage > 1 Then
                mCurrentPage = mCurrentPage - 1
                LoadMainDatas clsReferVoucher.OtherFilter
            End If
        Case "cmdAuto"
            MsgBox "cmdAuto"
    End Select
End Sub

Private Sub ctlRMV_onBodySelecting(ByVal row As Long, bChange As Boolean, sErr As String)
    Dim i As Long

    If ctlRMV.BodyList.TextMatrix(row, ctlRMV.GetHeadColIndex("selcol")) = "" Then
        clsReferVoucher.RemoveSelected ctlRMV, row
        Exit Sub
    End If

    'V11 改用生单控件的setrulestring控制
'    If iSinvCZ = True Then
'        For i = 1 To Me.ctlRMV.BodyList.rows - 1
'            If ctlRMV.BodyList.TextMatrix(i, ctlRMV.GetBodyColIndex("selcol")) = "Y" Then
'                If ctlRMV.BodyList.TextMatrix(row, ctlRMV.GetBodyColIndex("cwhcode")) <> ctlRMV.BodyList.TextMatrix(i, ctlRMV.GetBodyColIndex("cwhcode")) Then
'                    MsgBox GetString("U8.DZ.JA.Res1870"), vbExclamation, GetString("U8.DZ.JA.Res030")
'                    ctlRMV.BodyList.TextMatrix(row, ctlRMV.GetBodyColIndex("selcol")) = ""
'                    ' Exit For
'                    Exit Sub
'                End If
'            End If
'        Next
'        ctlRMV.BodyList.TextMatrix(row, ctlRMV.GetBodyColIndex("selcol")) = "Y"
'    End If
End Sub

Private Sub ctlRMV_onHeadSelecting(ByVal row As Long, bChange As Boolean, sErr As String)
    Dim i As Long
    Dim fltStr As String
    Call OneHeadSelecting(row, bChange)
    If Not clsReferVoucher.blnSigleColumn Then
        clsReferVoucher.SetBodyData Me.ctlRMV, row, True

    End If
End Sub


Private Function OneHeadSelecting(ByVal row As Long, bChange As Boolean, Optional blnShowMsg As Boolean = True) As Boolean
    Dim i As Long
    Dim sTemp As String
    Dim fltStr As String
    Dim lngPos1 As Long
    Dim varTmp As Variant

    With Me.ctlRMV.HeadList
        If ctlRMV.HeadList.TextMatrix(row, ctlRMV.GetHeadColIndex("selcol")) = "" Then
            clsReferVoucher.RemoveSelected ctlRMV, row
            Exit Function
        End If


        If clsReferVoucher.bSelectSingle Then

            For i = 1 To Me.ctlRMV.HeadList.rows - 1
                ctlRMV.HeadList.TextMatrix(i, ctlRMV.GetHeadColIndex("selcol")) = ""
            Next
            ctlRMV.HeadList.TextMatrix(row, ctlRMV.GetHeadColIndex("selcol")) = "Y"
            clsReferVoucher.SelectedStr = CStr(row) & ","
            Me.ctlRMV.RemoveBodyAll
            Exit Function
        End If


        OneHeadSelecting = True
        OneHeadSelecting = clsReferVoucher.CheckHeadSelecting(Me.ctlRMV)
        If Not OneHeadSelecting Then GoTo lbREF1:


        clsReferVoucher.SelectedStr = clsReferVoucher.SelectedStr & " " & CStr(row) & ","
        Exit Function
lbREF1:
        bChange = False
        ctlRMV.HeadList.TextMatrix(row, ctlRMV.GetHeadColIndex("selcol")) = ""
        OneHeadSelecting = False
        Exit Function
lbREF2:
        bChange = False
        ctlRMV.HeadList.TextMatrix(row, ctlRMV.GetHeadColIndex("selcol")) = ""
        OneHeadSelecting = False
    End With
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

    '************************
    'u861多语Layout
    '    Call MultiLangLoadForm(Me, "u8.sa.xsglsql." & Me.Name)
    '************************

    On Error Resume Next

    bcancel = True
    Me.HelpContextID = clsReferVoucher.HelpID              ' 10151180

    mCurrentPage = 1

    '设置显示表体为不显示
    '    Me.ctlRMV.UFShowBody.Visible = False

    clsReferVoucher.InitReferVoucher Me.ctlRMV

    Me.ctlRMV.pageSize = clsReferVoucher.pageSize


    LoadMainDatas clsReferVoucher.OtherFilter

    UFFrmCaptionMgr.Caption = clsReferVoucher.FrmCaption

    '    If clsReferVoucher.bSelectSingle Or clsReferVoucher.strCheckFlds <> "" Then
    '        Me.ctlRMV.Toolbar.Buttons("tlbSel").Visible = False
    '        Me.ctlRMV.Toolbar.Buttons("tlbUnSel").Visible = False
    '    Else
    '        Me.ctlRMV.Toolbar.Buttons("tlbSel").Visible = True
    '        Me.ctlRMV.Toolbar.Buttons("tlbUnSel").Visible = True
    '    End If
    Me.ctlRMV.Toolbar.Buttons("tlbSel").Visible = True
    Me.ctlRMV.Toolbar.Buttons("tlbUnSel").Visible = True

    Me.ctlRMV.Toolbar.Buttons("tlbSel").ButtonMenus("tlbHead").Visible = False
    Me.ctlRMV.Toolbar.Buttons("tlbUnSel").ButtonMenus("tlbHead").Visible = False
    Me.ctlRMV.Toolbar.Buttons("tlbSel").ButtonMenus("tlbHead").Enabled = False
    Me.ctlRMV.Toolbar.Buttons("tlbUnSel").ButtonMenus("tlbHead").Enabled = False

    'Me.ctlRMV.Toolbar.Buttons("tlbSel").Visible = False
    Me.ctlRMV.Toolbar.Buttons("tlbHelp").Visible = False

    'Me.ctlRMV.UFToolbar.
    Me.ctlRMV.UFToolbar.RefreshVisible
    Me.ctlRMV.UFToolbar.RefreshEnabled
    Me.ctlRMV.HeadList.EditLocked = clsReferVoucher.HeadEnabled
    Me.ctlRMV.BodyList.EditLocked = clsReferVoucher.BodyEnabled
    'V11改用生单控件的控制规则方式
    If iSinvCZ Then
        Me.ctlRMV.SetRulesString ("<rules>" & "<rule>" & "<head></head>" & "<body error='" + GetString("U8.DZ.JA.Res1870") + "'>" & "<column name='cwhcode'/>" & "</body>" & "</rule>" & "</rules>")
    End If
    SkinSE_Init Me.hWnd, True
    Skinse_SetStopChildSkinFlag Me.hWnd
    SkinSE_SetFrameTitleText Me.hWnd, StrPtr(Me.Caption) ' UFFrmCaptionMgr.Caption为赋值窗体的标题控件 如果没有用me.caption

End Sub

Private Sub Form_Unload(Cancel As Integer)


    On Error Resume Next

    sWhere = ""
    clsReferVoucher.pageSize = Me.ctlRMV.pageSize
End Sub

Private Function FirstFindString(string1 As String, string2 As String) As Boolean
    Dim i, iLen, jlen As Integer
    Dim C1, c2 As String
    Dim rec As New ADODB.Recordset

    rec.CursorLocation = adUseClient
    On Error Resume Next
    rec.Open "select cInvoiceCompany as cCusHeadCode from Customer where ccuscode=N'" & string1 & "'", g_Conn, adOpenForwardOnly, adLockReadOnly
    If Not (rec.EOF And rec.BOF) Then
        C1 = "" & rec("cCusHeadCode")
    End If
    rec.Close
    rec.Open "select cInvoiceCompany as cCusHeadCode from Customer where ccuscode=N'" & string2 & "'", g_Conn, adOpenForwardOnly, adLockReadOnly
    If Not (rec.EOF And rec.BOF) Then
        c2 = "" & rec("cCusHeadCode")
    End If
    If C1 <> c2 Or C1 = "" Or c2 = "" Then
        FirstFindString = 0
        Exit Function
    End If
    rec.Close
    Set rec = Nothing
    FirstFindString = True
End Function

Private Function Filter() As Boolean
    Dim objFilter As New UFGeneralFilter.FilterSrv
    'Dim objFltCb As New clsReportCallBack    回调时使用

    Dim cbustypeWhere As String
    Dim m_ckey As String
    Dim m_cSubID As String

    Filter = False

    'objFltCb.StrReportName = m_ckey         回调时使用
    'Set objFilter.BehaviorObject = objFltCb

    Filter = objFilter.OpenFilter(g_oLogin, "", clsReferVoucher.FilterKey, clsReferVoucher.FilterSubID, "")

    If Filter Then
        sWhere = clsReferVoucher.convertWhere(objFilter)
        Filter = True
    End If
    Set objFilter = Nothing
    ' Set objFltCb = Nothing
End Function

''取得rst中的字段值，将null转换为0
Private Function GetRstVal(rst As ADODB.Recordset, FieldName As String) As Variant
    If IsNull(rst(FieldName)) = True Then
        'If rst(FieldName).Type = adChar Or rst(FieldName).Type = adVarChar Or rst(FieldName).Type = adDate Or rst(FieldName).Type = adDBDate Or rst(FieldName).Type = adDBTime Or rst(FieldName).Type = adDBTimeStamp Then
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

Private Sub ctlRMV_onButtonClick(ByVal sBtnKey As String, sErr As String)
    If ctlRMV.HeadList.ProtectUnload() <> 2 Then Exit Sub
    Call ButtonClick(sBtnKey, sErr)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.ctlRMV.Width = Me.ScaleWidth
    Me.ctlRMV.Height = Me.ScaleHeight

End Sub

Private Sub SetBottonState(pageCount As Long)
    With Me.ctlRMV
        If pageCount > 1 Then
            .Toolbar.Buttons("tlbFirst").Enabled = True
            .Toolbar.Buttons("tlbNext").Enabled = True
            .Toolbar.Buttons("tlbLast").Enabled = True
            .Toolbar.Buttons("tlbPrevious").Enabled = True
        Else
            .Toolbar.Buttons("tlbFirst").Enabled = False
            .Toolbar.Buttons("tlbNext").Enabled = False
            .Toolbar.Buttons("tlbLast").Enabled = False
            .Toolbar.Buttons("tlbPrevious").Enabled = False
        End If
        .UFToolbar.RefreshVisible
    End With
End Sub


Private Sub LoadMainDatas(Optional strOtherFilter As String)
    Dim eleTmp As IXMLDOMElement
    Dim sTmpWhere As String


    Me.ctlRMV.RemoveHeadAll

    Dim filterWhere As String                              '过滤条件的 filterWhere


    If strOtherFilter <> "" Then
        filterWhere = sWhere & IIf(sWhere = "", "", " and (") & strOtherFilter & ")"
    Else
        filterWhere = sWhere
    End If

    '数据权限处理
    If g_oLogin.isAdmin = False Then
        filterWhere = filterWhere & IIf(clsReferVoucher.MakerAuth = "", "", " and  (" & clsReferVoucher.MakerAuth & ")")
    End If


    filterWhere = Replace(filterWhere, "(1=1) ", "")
    
    clsReferVoucher.strFilter = filterWhere
    clsReferVoucher.SetHeadData Me.ctlRMV, ctlRMV.pageSize, mCurrentPage, mpageCount


    ctlRMV.RemoveBodyAll
    SetBottonState mpageCount

End Sub

