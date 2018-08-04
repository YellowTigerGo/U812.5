VERSION 5.00
Object = "{9FD12F62-6922-47E1-B1AC-3615BBD3D7A5}#1.0#0"; "UFLabel.ocx"
Object = "{4C2F9AC0-6D40-468A-8389-518BB4F8C67D}#1.0#0"; "UFComboBox.ocx"
Object = "{E08B3B98-649C-46CD-A1AD-4A10DB106D57}#1.2#0"; "UFStatusBar.ocx"
Object = "{8C7C777D-4D83-4DE8-947E-098E2343A400}#1.0#0"; "CommandButton.ocx"
Object = "{D2B3369D-2E6C-45DE-A705-14481242A2BE}#1.10#0"; "UFMenu6U.ocx"
Object = "{AF8BBBB7-94C6-4772-B826-624478C37D6A}#1.5#0"; "UFKEYHOOK.ocx"
Object = "{005151D4-33F6-471B-B651-222D4E411832}#4.4#0"; "UFFormPartner.ocx"
Object = "{0532C600-D183-40A1-802B-0E09F8DD709F}#1.0#0"; "ReferMakeVouch.ocx"
Object = "{5E4640D0-A415-404B-A457-72980C429D2F}#10.25#0"; "U8RefEdit.ocx"
Object = "{A0C292A3-118E-11D2-AFDF-000021730160}#1.0#0"; "UFEDIT.OCX"
Begin VB.Form frmRefernew 
   Caption         =   "生单列表"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   10965
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin U8Ref.RefEdit refDate 
      Height          =   300
      Left            =   3645
      TabIndex        =   9
      Top             =   90
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   529
      BadStr          =   "<>'""|&,"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Property        =   5
      RefType         =   2
   End
   Begin U8Ref.RefEdit refWh 
      Height          =   300
      Left            =   1065
      TabIndex        =   8
      Top             =   90
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   529
      BadStr          =   "<>'""|&,"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RefType         =   1
   End
   Begin cPopMenu6.PopMenu PopMenuMgr 
      Left            =   4500
      Top             =   2775
      _ExtentX        =   1058
      _ExtentY        =   1058
      HighlightCheckedItems=   0   'False
      TickIconIndex   =   0
      HookInSubClassMenu=   0   'False
   End
   Begin UFCOMMANDBUTTONLib.UFCommandButton cmdDate 
      Height          =   255
      Left            =   4845
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   41
      Caption         =   ""
      UToolTipText    =   ""
      Cursor          =   30464
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Flat            =   0   'False
      Enabled         =   -1  'True
      Style           =   0   'False
      Value           =   0   'False
   End
   Begin UFStatusBar.UFStatusBarCtl UFStatusBarCtl1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   6255
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      SimpleStyle     =   0
   End
   Begin UFFormPartner.UFFrmCaption UFFrmCaption1 
      Left            =   3060
      Top             =   1095
      _ExtentX        =   1508
      _ExtentY        =   767
      Caption         =   "生单列表"
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
   Begin UFKeyHook.UFKeyHookCtrl UFKeyHookCtrl1 
      Left            =   4050
      Top             =   1140
      _ExtentX        =   1905
      _ExtentY        =   529
   End
   Begin ReferMakeVouch.ctlReferMakeVouch ctlReferMakeVouch1 
      Height          =   6405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   11298
   End
   Begin UFCOMMANDBUTTONLib.UFCommandButton cmdWh 
      Height          =   255
      Left            =   2265
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   41
      Caption         =   ""
      UToolTipText    =   ""
      Cursor          =   30464
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Flat            =   0   'False
      Enabled         =   -1  'True
      Style           =   0   'False
      Value           =   0   'False
   End
   Begin EDITLib.Edit txtWh 
      Height          =   300
      Left            =   1065
      TabIndex        =   3
      Top             =   90
      Visible         =   0   'False
      Width           =   1485
      _Version        =   65536
      _ExtentX        =   2619
      _ExtentY        =   529
      _StockProps     =   253
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
   End
   Begin UFLABELLib.UFLabel lblWh 
      Height          =   180
      Left            =   90
      TabIndex        =   4
      Top             =   150
      Visible         =   0   'False
      Width           =   900
      _Version        =   65536
      _ExtentX        =   1587
      _ExtentY        =   317
      _StockProps     =   111
      Caption         =   "入库仓库："
      Alignment       =   1
      AutoSize        =   -1  'True
   End
   Begin EDITLib.Edit txtDate 
      Height          =   300
      Left            =   3645
      TabIndex        =   6
      Top             =   105
      Visible         =   0   'False
      Width           =   1485
      _Version        =   65536
      _ExtentX        =   2619
      _ExtentY        =   529
      _StockProps     =   253
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
   End
   Begin UFLABELLib.UFLabel lblDate 
      Height          =   180
      Left            =   2655
      TabIndex        =   7
      Top             =   150
      Visible         =   0   'False
      Width           =   900
      _Version        =   65536
      _ExtentX        =   1587
      _ExtentY        =   317
      _StockProps     =   111
      Caption         =   "入库日期："
      Alignment       =   1
      AutoSize        =   -1  'True
   End
   Begin UFCOMBOBOXLib.UFComboBox CmbVTID 
      Height          =   300
      Left            =   5385
      TabIndex        =   10
      Top             =   90
      Visible         =   0   'False
      Width           =   2370
      _Version        =   65536
      _ExtentX        =   4180
      _ExtentY        =   529
      _StockProps     =   196
      Text            =   ""
      Style           =   2
      ForeColor       =   1996536096
   End
   Begin VB.Menu mnuBatch 
      Caption         =   "批号"
      Visible         =   0   'False
      Begin VB.Menu mnuBatchGenerate 
         Caption         =   "批号生成"
      End
   End
End
Attribute VB_Name = "frmRefernew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public oRef As Object
Public DomHead_Dest As DOMDocument      '目标单据head
Public DomBody_Dest As DOMDocument      '目标单据body
Public Head_Columnkey As String         '表头栏目key U
Public Body_Columnkey As String         '表体栏目key
Public Head_Source As String            '表头数据源 U
Public Body_Source As String            '表体数据源
Public Cardnumber_source As String
Public VIEWname As String               'GSP系统原始的数据源（包换条件）
Public head_key As String
Public body_key As String
Public inner_key As String
Public cVouchType As String
Public Defwhere_head As String
Public Defwhere_body As String
Public sFilterID_head As String
Public sFilterID_body As String


Public RefdomH As DOMDocument  '当前参照上半部分用户选择的数据
Public RefdomB As DOMDocument  '当前参照下半部分用户选择的数据 '


Public Source_Cardnum As String
Public Dest_Cardnum As String
 

Public DOM_sa_refervoucherconfig As DOMDocument
Public DOM_SA_ReferFillConfig As DOMDocument

'Public oDic As Scripting.Dictionary
Dim mLogin As U8Login.clsLogin
Dim DBconn As ADODB.Connection

Public moLogin As USCOMMON.Login
Public ClsBill As New USERPCO.VoucherCO

Public sReferUfts        As String
Public sReferMainV      As String

Dim isBody As Boolean

Dim sTemp As String
Dim STMsgTitle As String
'Public filterItf As New clsFilterInterface
'Public repFilter As clsReportFilter
Public Head_filter As New UFGeneralFilter.FilterSrv
Public Body_filter As New UFGeneralFilter.FilterSrv


 

Public sMustWhere  As String

Dim listFormat As DOMDocument

Dim iSotype As Long
Dim isosid  As Long

'Dim oEditableCol As Dictionary

Dim refType As String

'GSP用
Public lMainId As Long
Public sFromtype As String

Public isSimulate As Boolean

Dim cWhCode As String

'Dim oInvVOCache As Dictionary

Public VtId As String

Dim cbVTID As Object

Dim m_LastLocRow As Long '上一次定位行

Dim m_bodyCurRow As Long






'Public Sub Init(login As Object, Ref As Object, dic As Dictionary)
'
'End Sub

Public Sub LoadLMString(Optional ByVal section As VouchListPos = eListAll)
    Dim newColSet As New U8ColumnSet.clsColSet
    Dim ColSet As Object
    Dim oldColSet As Object

    Set oldColSet = CreateObject("U8Colset.clsColSet")
    Set ColSet = newColSet

    ColSet.Init mLogin.UfDbName, mLogin.cUserId

    ctlReferMakeVouch1.cUserId = mLogin.cUserId

    If section = eListHead Or section = eListAll Then
        ColSet.setColMode Head_Columnkey
'        oRef.ColSetOrderHead = ColSet.GetOrderString
'        oRef.ColSetSqlHead = ColSet.GetSqlString
        ctlReferMakeVouch1.HeadList.SumStyle = vlGridSum
        ctlReferMakeVouch1.HeadColSetXml = ColSet.getColInfo
        ctlReferMakeVouch1.HeadList.InitHead ctlReferMakeVouch1.HeadColSetXml
'        SetItemFormat ctlReferMakeVouch1.HeadList, oRef.ColSetKeyHead
        SetFormat ctlReferMakeVouch1.HeadList
    End If

    If section = eListBody Or section = eListAll Then
'        If oRef.ColSetKeyBody = "" Then Exit Sub

        ColSet.setColMode Body_Columnkey
'        oRef.ColSetOrderBody = ColSet.GetOrderString
'        oRef.ColSetSqlBody = ColSet.GetSqlString
        ctlReferMakeVouch1.BodyList.SumStyle = vlGridSum
        ctlReferMakeVouch1.BodyColSetXml = ColSet.getColInfo
'        ChangeEditState
        ctlReferMakeVouch1.BodyList.InitHead ctlReferMakeVouch1.BodyColSetXml
'        SetItemFormat ctlReferMakeVouch1.BodyList, oRef.ColSetKeyBody
        SetFormat ctlReferMakeVouch1.BodyList
    End If

    Set ColSet = Nothing
End Sub

Private Sub cmdDate_Click()
    Dim Calendar As Object
    Set Calendar = CreateObject("CalendarAPP.ICaleCom")
    txtDate.Text = Format(Calendar.Calendar(cmdDate.hWnd), "YYYY-MM-DD")
    If (txtDate.Visible And txtDate.Enabled) Then
        txtDate.SetFocus
    End If
End Sub

Private Sub cmdWh_Click()
    Dim cWhCode As String
    Dim cWhName As String

    cWhName = sRefWh(cWhCode)
    txtWh.Tag = cWhCode
    txtWh.Text = cWhName
    If txtWh.Visible And txtWh.Enabled Then
        txtWh.SetFocus
    End If

    m_bodyCurRow = ctlReferMakeVouch1.BodyList.row
'    FillWhName
End Sub

Private Sub ctlReferMakeVouch1_BodyBrowUser(RetValue As Variant, row As Long, Col As Long)
    Dim colName As String
    Dim cWhCode As String
    Dim Ref As UFReferC.UFReferClient
    Dim sErr As String
    Dim oDomBody As DOMDocument
    Dim attrNode As IXMLDOMNode
    Set oDomBody = ctlReferMakeVouch1.GetBodyLine(row)
    colName = LCase(ctlReferMakeVouch1.GetBodyColName(Col))
    On Error Resume Next
    If refType = "clssadispatchbatchrefer" Then




      ElseIf refType = "clspuarrivebatchrefer" Then
        If colName = "cwhname" Then
            RetValue = sRefWh(cWhCode)
            ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(row, ctlReferMakeVouch1.BodyList.GridColIndex("cwhcode")) = cWhCode
            ctlReferMakeVouch1.BodyTextMatrix(row, ctlReferMakeVouch1.BodyList.GridColIndex("cwhname")) = RetValue

          ElseIf colName Like "cfree*" Then
            RetValue = sRefFree(colName)

        End If
    End If
End Sub

Private Sub ctlReferMakeVouch1_BodyCellValueChanged(ByVal row As Integer, ByVal Col As Integer, newvalue As Variant, OldValue As Variant, KeepFocus As Boolean)
'    Dim r As Long
'    Dim c As Long
'    r = row
'    c = Col
'    m_bodyCurRow = r
'
'    If refType = "clssadispatchbatchrefer" Then
'        Call DataCheckOut(r, c, NewValue, OldValue, KeepFocus)
'        ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(r, c) = NewValue
'      ElseIf refType = "clspuorderbatchrefer" Then
'        Call DataCheckDdIn(r, c, NewValue, OldValue, KeepFocus)
'        ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(r, c) = NewValue
'      ElseIf refType = "clspuarrivebatchrefer" Then
'        Call DataCheckDhdIn(r, c, NewValue, OldValue, KeepFocus)
'        ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(r, c) = NewValue
'    End If
End Sub

Private Sub ctlReferMakeVouch1_BodyClick()
    On Error Resume Next
'    sTemp = ctlReferMakeVouch1.BodyTextMatrix(ctlReferMakeVouch1.BodyList.row, ctlReferMakeVouch1.BodyList.Col)
'    UFStatusBarCtl1.Panels(6).text = IIf(sTemp = "", GetResString("U8.SCM.ST.STReference.frmeqreference.00022"), sTemp)
'    UFStatusBarCtl1.Panels(7).text = "X:" & ctlReferMakeVouch1.BodyList.row & ";Y:" & ctlReferMakeVouch1.BodyList.Col
'
'    '将表体对应表头行相应单元格加亮显示
'    Dim m_LocRow As Long
'    m_LocRow = GetHeadRow(ctlReferMakeVouch1.BodyList.row)
'    If m_LastLocRow > 0 Then
'        ctlReferMakeVouch1.HeadList.Col = 0
'        ctlReferMakeVouch1.HeadList.row = m_LastLocRow
'        ctlReferMakeVouch1.HeadList.GetGridBody().CellBackColor = &HFFE3C6
'    End If
'    If m_LocRow > 0 Then
'        ctlReferMakeVouch1.HeadList.Col = 0
'        ctlReferMakeVouch1.HeadList.row = m_LocRow
'        ctlReferMakeVouch1.HeadList.GetGridBody().CellBackColor = &HFFFF00
'        ctlReferMakeVouch1.HeadList.TopRow = ctlReferMakeVouch1.HeadList.row
'    End If
'    m_LastLocRow = m_LocRow
'
'    Dim sVouchType As VouchType
'    If refType = "clssadispatchbatchrefer" Then
'        sVouchType = DispatchList
'      ElseIf refType = "clspuorderbatchrefer" Then
'        sVouchType = PurchaseOrder
'      ElseIf refType = "clspuarrivebatchrefer" Then
'        sVouchType = ArriveList
'    End If
'
'    If Len(ctlReferMakeVouch1.BodyTextMatrix(ctlReferMakeVouch1.BodyList.row, 0)) < 2 And ctlReferMakeVouch1.BodyList.RecordCount > 0 And (refType = "clssadispatchbatchrefer" Or refType = "clspuorderbatchrefer" Or refType = "clspuarrivebatchrefer") Then
'        sTemp = ""
'        If SetItemState(moLogin, sVouchType, ctlReferMakeVouch1.GetHeadLine(ctlReferMakeVouch1.HeadList.row), ctlReferMakeVouch1.GetBodyLine(ctlReferMakeVouch1.BodyList.row), ctlReferMakeVouch1.BodyList, sTemp) Then
'
'          Else
'            MsgBox sTemp, , LoadResST("U8.ST.USCONTROL.frmproduceorder.00594") 'zh-CN：系统提示
'            sTemp = ""
'        End If
'    End If
End Sub

Private Sub ctlReferMakeVouch1_BodyMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
'    If Button = 2 Then
'        Me.PopMenuMgr.Visible("mnuBatchGenerate") = ((refType = "clspuorderbatchrefer" Or refType = "clspuarrivebatchrefer") And ctlReferMakeVouch1.BodyList.RecordCount > 0)
'        Call Me.PopMenuMgr.ShowPopupMenu(ctlReferMakeVouch1.BodyList, "mnuBatch", X, Y)
'        Exit Sub
'    End If
End Sub

Private Sub ctlReferMakeVouch1_ButtonOnCommand(ByVal enumType As UFToolBarCtrl.ENUM_MENU_OR_BUTTON, ByVal cButtonId As String, ByVal cMenuId As String)
    If enumType = enumButton Then
        isBody = False
        ButtonClick cButtonId
      ElseIf enumType = enumMenu Then
        If LCase(cMenuId) = "tlbbody" Then
            isBody = True
          Else
            isBody = False
        End If
        ButtonClick cButtonId
    End If

'    If LCase(cButtonId) = "tlbfirst" Or LCase(cButtonId) = "tlbprevious" Or LCase(cButtonId) = "tlbnext" Or LCase(cButtonId) = "tlblast" Then
'        ctlReferMakeVouch1_EdtAccepted
'    End If
End Sub

Public Sub SetData(Optional headkey_val As String, Optional bAppend As Boolean = False, Optional Errstr As String)
    Dim ColSet As New U8ColumnSet.clsColSet
    Dim rds As ADODB.Recordset
    Dim Dom As DOMDocument
    Dim domtemp As New DOMDocument
    Dim strSql As String
    Dim head_key_vals As String
    
    Set rds = New ADODB.Recordset
    Set Dom = New DOMDocument
    
    If headkey_val <> "" Then GoTo Refreshbody
    
    If isBody = False Then '刷新表头数据
        ColSet.Init mLogin.UfDbName, mLogin.cUserId
        ColSet.setColMode Head_Columnkey
        
        Select Case cVouchType
            Case "005", "007", "015"
                head_key_vals = " select  distinct " & head_key & " from (" & VIEWname & ") KK  "
                If Not (Head_filter Is Nothing) Then
                    If Head_filter.GetSQLWhere <> "" Then
                        head_key_vals = head_key_vals & " where  " & Head_filter.GetSQLWhere
                    End If
                End If
            Case " "
        
        End Select
        strSql = "select  " & ColSet.GetSqlString & " from " & Head_Source & " where 1=1"
        If head_key_vals <> "" Then strSql = strSql & "  and " & head_key & " in ( " & head_key_vals & " )"
        If ColSet.GetOrderString <> "" Then strSql = strSql & " order by " & ColSet.GetOrderString
        If rds.State <> 0 Then rds.Close
        rds.CursorLocation = adUseClient
        rds.Open strSql, DBconn.ConnectionString, 3, 4
        rds.Save Dom, adPersistXML
        ctlReferMakeVouch1.SetHeadDom Dom
        GoTo Refreshbody
        
    ElseIf isBody = True Then  '刷新表体数据
    
Refreshbody:
        If Trim(Me.Body_Columnkey) = "" Then GoTo Onexit
        
        ColSet.Init mLogin.UfDbName, mLogin.cUserId
        ColSet.setColMode Me.Body_Columnkey
    
        Select Case cVouchType
            Case "ww"
            
            Case Else
                strSql = "select  " & ColSet.GetSqlString & " from " & Body_Source & " where 1=1 "
        End Select
    
        
        strSql = strSql & " and " & Me.head_key & "='" & headkey_val & "' "
        If Not (Body_filter Is Nothing) Then
            If Body_filter.GetSQLWhere <> "" Then
                strSql = strSql & " and " & Body_filter.GetSQLWhere
            End If
        End If
        If ColSet.GetOrderString <> "" Then
            strSql = strSql & " order by " & ColSet.GetOrderString
        End If
        If rds.State <> 0 Then rds.Close
        rds.CursorLocation = adUseClient
        rds.Open strSql, DBconn.ConnectionString, 3, 4
        rds.Save Dom, adPersistXML
        If bAppend Then
            ctlReferMakeVouch1.AppendBodyData rds, Errstr
        Else
            ctlReferMakeVouch1.SetBodyDom Dom
            ctlReferMakeVouch1.SelectBodyAll
        End If
    End If
    
Onexit:

End Sub

Private Function GetPageData(oRecordset As Recordset) As Recordset
    '    oRecordset.PageSize = ctlReferMakeVouch1.PageSize
    '    ctlReferMakeVouch1.PageCount = oRecordset.PageCount
    '    If ctlReferMakeVouch1.CurrentPage > ctlReferMakeVouch1.PageCount Then ctlReferMakeVouch1.CurrentPage = ctlReferMakeVouch1.PageCount
    '    RecordsetToPage oRecordset, ctlReferMakeVouch1.CurrentPage
    '    Set GetPageData = oRecordset
End Function

Private Sub SetAppendCondition()
    Dim sWhere As String
    sWhere = ""

    Dim i As Integer
    For i = 1 To ctlReferMakeVouch1.HeadList.rows - 1
        If ctlReferMakeVouch1.HeadList.TextMatrix(i, 0) = "Y" Then
            sWhere = sWhere & "'" & ctlReferMakeVouch1.HeadList.TextMatrix(i, IIf(ctlReferMakeVouch1.HeadList.GridColIndex(oRef.PKeyHead) = -1, ctlReferMakeVouch1.HeadList.GridColIndex("cpoid"), ctlReferMakeVouch1.HeadList.GridColIndex(oRef.PKeyHead))) & "',"
        End If
    Next
    If sWhere <> "" Then
        sWhere = Mid(sWhere, 1, Len(sWhere) - 1)
      Else
        sWhere = " 1=2 "
    End If

    oRef.AppendCondition = sWhere
End Sub

Private Sub ctlReferMakeVouch1_CheckClick(ByVal index As Integer, ByVal Name As String, ByVal caption As String, ByVal progID As String)
    If refType Like "clsmomaterailapp??" Or refType Like "clsmoorder??" Or refType Like "clsomorder??" Then
        ctlReferMakeVouch1.HeadList.SetFieldRevisable "iqty", CBool(ctlReferMakeVouch1.UFCheckBox(index).Value)
    End If
End Sub

Private Sub ctlReferMakeVouch1_EdtAccepted()
'    isBody = False
'    SetData False
End Sub

Private Sub ctlReferMakeVouch1_HeadClick()
    On Error Resume Next
    sTemp = ctlReferMakeVouch1.HeadList.TextMatrix(ctlReferMakeVouch1.HeadList.row, ctlReferMakeVouch1.HeadList.Col)
    UFStatusBarCtl1.Panels(6).Text = IIf(sTemp = "", GetResString("U8.SCM.ST.STReference.frmeqreference.00022"), sTemp)
    UFStatusBarCtl1.Panels(7).Text = "X:" & ctlReferMakeVouch1.HeadList.row & ";Y:" & ctlReferMakeVouch1.HeadList.Col
End Sub

Private Sub ctlReferMakeVouch1_HeadSelectClick(ByVal Selected As Boolean)
'Dim rd As New ADODB.Recordset
'MsgBox Selected

'    If oRef.ColSetKeyBody = "" Then Exit Sub
'    If refType = "clsmomaterailapppb" Or refType = "clsmomaterailappcl" Then
'        ctlReferMakeVouch1.SelectHeadAll
'    End If
'
'    isBody = True
''    SetData False
End Sub

Private Sub ctlReferMakeVouch1_HeadShiftSelect(ByVal lFromRow As Long, ByVal lToRow As Long, other As Variant)
    ctlReferMakeVouch1_HeadSelectClick True
End Sub

Private Sub ctlReferMakeVouch1_onHeadSelected(ByVal row As Long, sErr As String)
Dim key_val As String
Dim i As Long
    For i = 1 To ctlReferMakeVouch1.HeadList.rows - 1
        If row = i Then
            
        Else
            ctlReferMakeVouch1.HeadList.TextMatrix(i, 0) = ""
        End If
    Next
    key_val = ctlReferMakeVouch1.HeadTextMatrix(row, ctlReferMakeVouch1.GetHeadColIndex(Me.head_key))
'    Call SetData(key_val, True)
    Call SetData(key_val, False)
End Sub

Private Sub ctlReferMakeVouch1_ShowBody(ByVal Value As Integer)
'    If oRef.ColSetKeyBody = "" Then Exit Sub
'
'    isBody = True
'    SetData False
End Sub

Private Sub Form_Load()
    On Error Resume Next

'    Set Me.Icon = frmMain.Icon
'    Set oInvVOCache = New Dictionary
'
'    mnuBatchGenerate.Visible = True
'    Call Me.PopMenuMgr.SubClassMenu(Me)
'    Me.PopMenuMgr.caption("mnuBatchGenerate") = GetResString("U8.ST.V870.00057")
'
'    Set oEditableCol = New Dictionary
'    oEditableCol.Add "cbatch", ""
'    oEditableCol.Add "cinva_unit", ""
'    oEditableCol.Add "cinvouchcode", ""
'    oEditableCol.Add "cwhname", ""
'    oEditableCol.Add "dvdate", ""
'    oEditableCol.Add "imassdate", ""
'    oEditableCol.Add "dmadedate", ""
'    oEditableCol.Add "iinvexchrate", ""
'    oEditableCol.Add "cvmivencode", ""
'    oEditableCol.Add "cfree1", ""
'    oEditableCol.Add "cfree2", ""
'    oEditableCol.Add "cfree3", ""
'    oEditableCol.Add "cfree4", ""
'    oEditableCol.Add "cfree5", ""
'    oEditableCol.Add "cfree6", ""
'    oEditableCol.Add "cfree7", ""
'    oEditableCol.Add "cfree8", ""
'    oEditableCol.Add "cfree9", ""
'    oEditableCol.Add "cfree10", ""
'
'    oEditableCol.Add "fcurqty", ""
'    oEditableCol.Add "fcurnum", ""
'
'    With Me.UFStatusBarCtl1
'        .Panels.Add
'        .Panels.Add
'        .Panels.Add
'        .Panels.Add
'        .Panels.Add
'        .Panels.Add
'        .Panels.Add
'    End With
'
'    UFFrmCaption1.caption = oRef.caption
'    Me.caption = UFFrmCaption1.caption
'    Me.HelpContextID = oRef.HelpContextID
'
'    Dim control As DOMDocument
'    If oRef.ControlXmlString <> "" Then
'        Set control = New DOMDocument
'        control.loadXML oRef.ControlXmlString
'    End If
'    ctlReferMakeVouch1.Init False, Nothing, control, val(RegRead(C_SPLIT_MODE, "mvPageSize")), 0, 1, , stLogin.OldLogin
'
'    If oRef.IsSingleFilter Then ctlReferMakeVouch1.SetFilterOne
'
'    If oRef.ColSetKeyBody = "" Then ctlReferMakeVouch1.SetIsList
'
'    ctlReferMakeVouch1.SetRulesString LCase(oRef.RulesXmlString)
'
'    Set listFormat = oRef.ReferListFormat
'
'    If oRef.ColSetKeyBody <> "" Then
'        ctlReferMakeVouch1.UFShowBody.Value = 1
'    Else
'        ctlReferMakeVouch1.UFShowBody.Visible = False
'    End If
'
'    If refType = "clssadispatchbatchrefer" Then
'        ctlReferMakeVouch1.UFCheckBox(2).Value = val(IIf(RegRead(C_SPLIT_MODE, "dispatchform") = "", "1", RegRead(C_SPLIT_MODE, "dispatchform")))
'      ElseIf refType = "clssadispatchrefer" Then
'        ctlReferMakeVouch1.UFCheckBox(1).Value = val(IIf(RegRead(C_SPLIT_MODE, "dispatchform") = "", "1", RegRead(C_SPLIT_MODE, "dispatchform")))
'    End If
'
'    isBody = False
'    SetData
'
'    txtDate.text = moLogin.LoginDate
'    refDate.text = moLogin.LoginDate
'    If refType = "clssadispatchbatchrefer" Then
'        lblDate.Visible = True
'        txtDate.Visible = True
'        refDate.Visible = True
'        lblDate.caption = LoadResST("U8.SCM.ST.USCONTROL.frmCons.SSTab1.lblOutDate.Caption")
'        lblDate.UToolTipText = lblDate.caption
'
'        CmbVTID.Visible = True
'      ElseIf refType = "clspuorderbatchrefer" Or refType = "clspuarrivebatchrefer" Then
'        lblDate.Visible = True
'        txtDate.Visible = True
'        refDate.Visible = True
'        lblWh.Visible = True
'        txtWh.Visible = True
'        refWh.Visible = True
'        lblWh.caption = LoadResST("U8.SCM.ST.USCONTROL.frmStockOrder.lblWh.Caption")
'        lblWh.UToolTipText = lblWh.caption
'        lblDate.caption = LoadResST("U8.SCM.ST.USCONTROL.frmStockOrder.lblOutDate.Caption")
'        lblDate.UToolTipText = lblDate.caption
'
'        CmbVTID.Visible = True
'      ElseIf refType = "clsexsalesliprefer" Then
'        lblDate.Visible = True
'        txtDate.Visible = True
'        refDate.Visible = True
'        lblDate.caption = LoadResST("U8.SCM.ST.STReference.frmEQReference.Frame2.Label1.Caption")
'        lblDate.UToolTipText = lblDate.caption
'    End If
'
'    CmbVTID.Visible = True
'
'    If refType = "clssadispatchbatchrefer" Or refType = "clspuorderbatchrefer" Or refType = "clspuarrivebatchrefer" Then
'        Dim RecVTID As New ADODB.Recordset
'        If ClsBill.GetVT_ID(IIf(refType = "clssadispatchbatchrefer", "32", "01"), RecVTID, sTemp, 0, , False) = False Then
'            'Result:Row=638 Col=14  Content="取摸板号失败"  ID=4bdd6015-b28d-4cd5-944f-ad71a5eb2d87
'            MsgBox GetResString("U8.ST.USKCGLSQL.frmqc.01831")
'            Exit Sub
'        End If
'
'        If RecVTID.RecordCount = 0 Then
'            MsgBox GetResString("U8.ST.USKCGLSQL.modmain.03067")
'            'Result:Row=656 Col=25  Content="您没有模板使用权限!"   ID=09f3890c-6f9c-42e8-a3a1-c209ccd746e5
'            Exit Sub
'        End If
'
'        Dim i As Long
'        i = 0
'        Do While Not RecVTID.EOF
'            CmbVTID.AddItem RecVTID(1), i
'            CmbVTID.ItemData(i) = RecVTID(0)
'            RecVTID.MoveNext
'            i = i + 1
'            CmbVTID.ListIndex = 0
'        Loop
'    End If
'
'    If Not cbVTID Is Nothing Then
'        i = 0
'        Do While Not cbVTID.EOF
'            CmbVTID.AddItem cbVTID(1), i
'            CmbVTID.ItemData(i) = cbVTID(0)
'            cbVTID.MoveNext
'            i = i + 1
'            CmbVTID.ListIndex = 0
'        Loop
'    End If
'
'    If CmbVTID.ListCount = 0 Then CmbVTID.Visible = False
'
'    If refType = "clsmomaterailapppb" Or refType = "clsmomaterailappcl" Then
'        ctlReferMakeVouch1.UFShowBody.Visible = False
'    End If
'
'    Form_Resize
'
'    Dim sMetaXML As String
'    Dim sFilterSql As String
'
'    sMetaXML = ""
'    sFilterSql = moLogin.getAuthString("Warehouse", , "W")
'
'    If sFilterSql <> "" Then
'
'        sMetaXML = "<Ref><RefSet bAuth='1' /></Ref>"
'      Else
'        sMetaXML = "<Ref><RefSet bAuth='0' /></Ref>"
'    End If
'
'    refWh.Init oLogin, "Warehouse_AA", False, sMetaXML
'    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    ctlReferMakeVouch1.Top = 0
    ctlReferMakeVouch1.Left = 0
    ctlReferMakeVouch1.Width = Me.ScaleWidth
    ctlReferMakeVouch1.Height = Me.ScaleHeight - UFStatusBarCtl1.Height

    lblWh.ZOrder 0
    refWh.ZOrder 0
    lblDate.ZOrder 0
    refDate.ZOrder 0

    CmbVTID.ZOrder 0

    lblWh.BackColor = ctlReferMakeVouch1.HeadList.GetGridBody().BackColor
    txtWh.BorderStyle = 0
    lblDate.BackColor = lblWh.BackColor
    txtDate.BorderStyle = txtWh.BorderStyle

    Dim tmpTop As Single
    Dim tmpLeft As Single
    tmpTop = 620
    If ctlReferMakeVouch1.UFCheckBox.Count = 1 Then
        tmpLeft = ctlReferMakeVouch1.UFShowBody.Left
      Else
        tmpLeft = ctlReferMakeVouch1.UFCheckBox(1).Left
    End If

    If refType = "clsexsalesliprefer" Then tmpLeft = ctlReferMakeVouch1.UFShowBody.Left

    cmdDate.Left = tmpLeft - cmdDate.Width - 60
    lblDate.Left = cmdDate.Left - 2175
    txtDate.Left = cmdDate.Left - 1200
    lblDate.Top = tmpTop
    txtDate.Top = tmpTop - 60
    cmdDate.Top = tmpTop - 30

    cmdWh.Left = lblDate.Left - cmdWh.Width - 60
    lblWh.Left = cmdWh.Left - 2175
    txtWh.Left = cmdWh.Left - 1200
    lblWh.Top = tmpTop
    txtWh.Top = tmpTop - 60
    cmdWh.Top = tmpTop - 30

    refDate.Left = txtDate.Left
    refDate.Top = txtDate.Top
    refWh.Left = txtWh.Left
    refWh.Top = txtWh.Top

    CmbVTID.Top = refWh.Top + refWh.Height + 200
    CmbVTID.Left = Me.ScaleWidth - CmbVTID.Width - 1800 'IIf(refType = "clssadispatchbatchrefer", lblDate.Left - CmbVTID.width - 60, lblWh.Left - CmbVTID.width - 60)
End Sub

Private Sub mnuBatchGenerate_Click()
    With ctlReferMakeVouch1.BodyList

        Dim strXml As String
        Dim domxml As New DOMDocument
        Dim number As String
        Dim cInvCode As String
        Dim sErr As String

        cInvCode = ""
        strXml = "<data>" & vbCrLf
        strXml = strXml & "<row name='vchdate' value='' itype='0' bHave='true'/>" & vbCrLf
        strXml = strXml & "<row name='prdbatchnum' value='' itype='1' bHave='true'/>" & vbCrLf
        strXml = strXml & "<row name='department' value='' itype='2' bHave='true'/>" & vbCrLf
        strXml = strXml & "<row name='workcenter' value='' itype='3' bHave='true'/>" & vbCrLf
        strXml = strXml & "<row name='whname' value=''  itype='4' bHave='true'/>" & vbCrLf
        strXml = strXml & "<row name='businessclerk' value='' itype='5' bHave='true'/>" & vbCrLf
        strXml = strXml & "<row name='supplier' value='' itype='6' bHave='true'/>" & vbCrLf
        strXml = strXml & "<row name='cfree1' value='' itype='20' bHave='true'/>" & vbCrLf
        strXml = strXml & "<row name='cfree2' value='' itype='21' bHave='true'/>" & vbCrLf
        strXml = strXml & "<row name='cfree3' value='' itype='28' bHave='true'/>" & vbCrLf
        strXml = strXml & "<row name='cfree4' value='' itype='29' bHave='true'/>" & vbCrLf
        strXml = strXml & "<row name='cfree5' value='' itype='30' bHave='true'/>" & vbCrLf
        strXml = strXml & "<row name='cfree6' value='' itype='31' bHave='true'/>" & vbCrLf
        strXml = strXml & "<row name='cfree7' value='' itype='32' bHave='true'/>" & vbCrLf
        strXml = strXml & "<row name='cfree8' value='' itype='33' bHave='true'/>" & vbCrLf
        strXml = strXml & "<row name='cfree9' value='' itype='34' bHave='true'/>" & vbCrLf
        strXml = strXml & "<row name='cfree10' value='' itype='35' bHave='true'/>" & vbCrLf
        strXml = strXml & "</data>"
        domxml.loadXML strXml

        Dim rows As Long
        For rows = 1 To .rows - 1
            cInvCode = .TextMatrix(rows, .GridColIndex(LCase("cinvcode")))
            domxml.loadXML strXml

            domxml.documentElement.selectSingleNode("//data/row[@name='vchdate']").Attributes.getNamedItem("value").Text = ctlReferMakeVouch1.HeadList.TextMatrix(GetHeadRow(rows), ctlReferMakeVouch1.HeadList.GridColIndex(LCase(IIf(refType = "clspuorderbatchrefer", "dpodate", "ddate"))))
            domxml.documentElement.selectSingleNode("//data/row[@name='prdbatchnum']").Attributes.getNamedItem("value").Text = .TextMatrix(rows, .GridColIndex(LCase("cmolotcode")))
            domxml.documentElement.selectSingleNode("//data/row[@name='department']").Attributes.getNamedItem("value").Text = ctlReferMakeVouch1.HeadList.TextMatrix(GetHeadRow(rows), ctlReferMakeVouch1.HeadList.GridColIndex(LCase("cdepcode")))
            domxml.documentElement.selectSingleNode("//data/row[@name='workcenter']").Attributes.getNamedItem("value").Text = .TextMatrix(rows, .GridColIndex(LCase("cmworkcentercode")))
            domxml.documentElement.selectSingleNode("//data/row[@name='whname']").Attributes.getNamedItem("value").Text = .TextMatrix(rows, .GridColIndex(LCase("cwhcode")))
            domxml.documentElement.selectSingleNode("//data/row[@name='businessclerk']").Attributes.getNamedItem("value").Text = ctlReferMakeVouch1.HeadList.TextMatrix(GetHeadRow(rows), ctlReferMakeVouch1.HeadList.GridColIndex(LCase("cpersoncode")))
            domxml.documentElement.selectSingleNode("//data/row[@name='supplier']").Attributes.getNamedItem("value").Text = ctlReferMakeVouch1.HeadList.TextMatrix(GetHeadRow(rows), ctlReferMakeVouch1.HeadList.GridColIndex(LCase("cvencode")))
            domxml.documentElement.selectSingleNode("//data/row[@name='cfree1']").Attributes.getNamedItem("value").Text = GetFreeCode(.TextMatrix(rows, .GridColIndex(LCase("cfree1"))))
            domxml.documentElement.selectSingleNode("//data/row[@name='cfree2']").Attributes.getNamedItem("value").Text = GetFreeCode(.TextMatrix(rows, .GridColIndex(LCase("cfree2"))))
            domxml.documentElement.selectSingleNode("//data/row[@name='cfree3']").Attributes.getNamedItem("value").Text = GetFreeCode(.TextMatrix(rows, .GridColIndex(LCase("cfree3"))))
            domxml.documentElement.selectSingleNode("//data/row[@name='cfree4']").Attributes.getNamedItem("value").Text = GetFreeCode(.TextMatrix(rows, .GridColIndex(LCase("cfree4"))))
            domxml.documentElement.selectSingleNode("//data/row[@name='cfree5']").Attributes.getNamedItem("value").Text = GetFreeCode(.TextMatrix(rows, .GridColIndex(LCase("cfree5"))))
            domxml.documentElement.selectSingleNode("//data/row[@name='cfree6']").Attributes.getNamedItem("value").Text = GetFreeCode(.TextMatrix(rows, .GridColIndex(LCase("cfree6"))))
            domxml.documentElement.selectSingleNode("//data/row[@name='cfree7']").Attributes.getNamedItem("value").Text = GetFreeCode(.TextMatrix(rows, .GridColIndex(LCase("cfree7"))))
            domxml.documentElement.selectSingleNode("//data/row[@name='cfree8']").Attributes.getNamedItem("value").Text = GetFreeCode(.TextMatrix(rows, .GridColIndex(LCase("cfree8"))))
            domxml.documentElement.selectSingleNode("//data/row[@name='cfree9']").Attributes.getNamedItem("value").Text = GetFreeCode(.TextMatrix(rows, .GridColIndex(LCase("cfree9"))))
            domxml.documentElement.selectSingleNode("//data/row[@name='cfree10']").Attributes.getNamedItem("value").Text = GetFreeCode(.TextMatrix(rows, .GridColIndex(LCase("cfree10"))))

            If Trim(cInvCode) <> "" And Trim(.TextMatrix(rows, .GridColIndex(LCase("cbatch")))) = "" Then
                Dim moInventoryPst As InventoryPst
                Dim moInventory As USERPVO.Inventory

                Set moInventoryPst = New InventoryPst
                moInventoryPst.Login = mLogin
                moInventoryPst.Load cInvCode, moInventory, , , "R"

                If moInventory.IsBatch Then
                    number = GetBatchNO(cInvCode, domxml, sErr, True)
                    .TextMatrix(rows, .GridColIndex(LCase("cbatch"))) = number

                    Dim Rs As New ADODB.Recordset
                    Rs.CursorLocation = adUseClient
                    Rs.Open "select dvdate,dmdate from currentstock where cbatch='" & number & "' and cwhcode = '" & .TextMatrix(rows, .GridColIndex(LCase("cwhcode"))) & "' and cinvcode = '" & cInvCode & "'", DBconn, adOpenDynamic, adLockOptimistic
                    If Not Rs.EOF And Not Rs.BOF Then
                        If Not IsNull(Rs.Fields("dvdate")) Then .TextMatrix(rows, .GridColIndex(LCase("dvdate"))) = Rs.Fields("dvdate")
                        If Not IsNull(Rs.Fields("dmdate")) Then .TextMatrix(rows, .GridColIndex(LCase("dmadedate"))) = Rs.Fields("dmdate")
                    End If

                    Rs.Close
                    Set Rs = Nothing

                End If
            End If
        Next
    End With
End Sub

Private Sub PopMenuMgr_MenuClick(sMenuKey As String)
    Select Case sMenuKey
      Case "mnuBatchGenerate"
        mnuBatchGenerate_Click
    End Select
End Sub

Private Sub refWh_AfterBrowse(RstClass As ADODB.Recordset, RstGrid As ADODB.Recordset, sXml As String)
    m_bodyCurRow = ctlReferMakeVouch1.BodyList.row
'    FillWhName
End Sub

Private Sub refWh_LostFocus()
    '    FillWhName
End Sub

Private Sub txtDate_GotFocus()
    cmdDate.Visible = True
End Sub

Private Sub txtDate_LostFocus()
    If LCase(ActiveControl.Name) <> "cmddate" Then
        cmdDate.Visible = False
    End If
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)
    If txtDate.Text <> "" And Not IsDate(txtDate.Text) Then
        MsgBox GetResString("U8.SCM.ST.STReference.mdlreference.00074"), vbApplicationModal + vbCritical + vbOKOnly, STMsgTitle
        '         txtDate.SelStart = 0
        '         txtDate.SelLength = Len(txtDate.text)
        txtDate.Text = ""
        Cancel = True
    End If
End Sub

Private Sub txtWh_GotFocus()
    cmdWh.Visible = True
End Sub

Private Sub txtWh_LostFocus()
    If LCase(ActiveControl.Name) <> "cmdwh" Then
        cmdWh.Visible = False
    End If

    m_bodyCurRow = ctlReferMakeVouch1.BodyList.row
'    FillWhName
End Sub

Private Sub UFKeyHookCtrl1_ContainerKeyDown(KeyCode As Integer, ByVal Shift As Integer)
    Dim j As Long
    Dim oRefSelect As RefSelect
    Dim strReturn As String
    Dim strReturnErr As String
    Dim strMsg As String

    Select Case KeyCode
      Case vbKeyEscape
        ButtonClick "tlbExit"
      Case vbKeyF4
        If Shift = vbCtrlMask Then
            Unload Me
            Exit Sub
        End If
      Case vbKeyP
        If Shift = vbCtrlMask Then
            ButtonClick "tlbPrint"
            KeyCode = 0
        End If
      Case vbKeyF3
        If Shift = vbCtrlMask Then
            ButtonClick "tlbLocal"
            KeyCode = 0
          Else
            ButtonClick "tlbFilter"
            KeyCode = 0
        End If
      Case vbKeyReturn
        '           If refType <> "clssadispatchbatchrefer" And refType <> "clspuorderbatchrefer" And refType <> "clspuarrivebatchrefer" Then
        '              Exit Sub
        '           End If
        '
        '           If LCase(ActiveControl.Name) = LCase("ctlReferMakeVouch1") And ctlReferMakeVouch1.BodyList.row >= 1 Then
        '              ctlReferMakeVouch1_BodyCellValueChanged ctlReferMakeVouch1.BodyList.row, ctlReferMakeVouch1.BodyList.col, ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(ctlReferMakeVouch1.BodyList.row, ctlReferMakeVouch1.BodyList.col), "", False
        '           End If

      Case vbKeyB
        If refType <> "clssadispatchbatchrefer" And refType <> "clspuorderbatchrefer" And refType <> "clspuarrivebatchrefer" Then
            Exit Sub
        End If

        If Shift = vbCtrlMask Then
            If ctlReferMakeVouch1.BodyList.RecordCount < 1 Then
                Beep
              Else
                If refType = "clssadispatchbatchrefer" Then
                    sRefBatchOut BodyColData("cbatch"), True, False
                  Else
                    sRefBatchIn BodyColData("cbatch"), True, False
                End If
            End If
            KeyCode = 0
        End If
      Case vbKeyE
        If refType <> "clssadispatchbatchrefer" And refType <> "clspuorderbatchrefer" And refType <> "clspuarrivebatchrefer" Then
            Exit Sub
        End If

        If Shift = vbCtrlMask Then
            If ctlReferMakeVouch1.BodyList.RecordCount < 1 Then
                Beep
              Else
                '全部自动指定批次
                Set oRefSelect = New RefSelect
                oRefSelect.CreateAndDropTmpCurrentStock mLogin, True
                For j = 1 To ctlReferMakeVouch1.BodyList.rows - 1
                    strReturnErr = GetResString("U8.ST.Default.00186") '"批量指定批次"
                    '先清除当前批次
                    ctlReferMakeVouch1.BodyTextMatrix(j, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cbatch"))) = ""
                    strReturn = ""
                    If refType = "clssadispatchbatchrefer" Then
                        sRefBatchOut ctlReferMakeVouch1.BodyTextMatrix(j, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cbatch"))), True, True, strReturn, strReturnErr
                      Else
                        sRefBatchIn ctlReferMakeVouch1.BodyTextMatrix(j, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cbatch"))), True, True, strReturn, strReturnErr
                    End If
                    If strReturnErr <> "" And strReturnErr <> GetResString("U8.ST.Default.00186") Then
                        strMsg = strMsg & strReturnErr & Chr(13) & Chr(10)
                    End If
                Next j
                oRefSelect.CreateAndDropTmpCurrentStock mLogin, False
                Set oRefSelect = Nothing
                If strMsg <> "" Then
                    MsgBox strMsg, vbOKOnly, LoadResST("U8.ST.USCONTROL.frmstockorder.00365") 'zh-CN：批量批次指定
                End If
            End If
            KeyCode = 0
        End If
      Case vbKeyO
        If refType <> "clssadispatchbatchrefer" And refType <> "clspuorderbatchrefer" And refType <> "clspuarrivebatchrefer" Then
            Exit Sub
        End If

        If Shift = vbCtrlMask Then
            If ctlReferMakeVouch1.BodyList.RecordCount < 1 Then
                Beep
              Else

                '全部自动指定入库单据号
                Set oRefSelect = New RefSelect
                oRefSelect.CreateAndDropTmpRdrecord mLogin, True
                For j = 1 To ctlReferMakeVouch1.BodyList.rows - 1
                    strReturnErr = GetResString("U8.ST.Default.00187") '"批量指定入库单据号"
                    '先清除当前入库单据号
                    'ctlReferMakeVouch1.BodyTextMatrix(j, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("ccode"))) = ""
                    strReturn = ""
                    If refType = "clssadispatchbatchrefer" Then
                        ctlReferMakeVouch1.BodyTextMatrix(j, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("ccode"))) = ""
                        sRefTackOut ctlReferMakeVouch1.BodyTextMatrix(j, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("ccode"))), True, True, strReturn, strReturnErr
                      Else
                        ctlReferMakeVouch1.BodyTextMatrix(j, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cInVouchCode"))) = ""
                        sRefTackIn ctlReferMakeVouch1.BodyTextMatrix(j, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cInVouchCode"))), True, True, strReturn, strReturnErr
                    End If
                    If strReturnErr <> "" And strReturnErr <> GetResString("U8.ST.Default.00187") Then
                        strMsg = strMsg & strReturnErr & Chr(13) & Chr(10)
                    End If
                Next j
                oRefSelect.CreateAndDropTmpRdrecord mLogin, False
                Set oRefSelect = Nothing
                If strMsg <> "" Then
                    MsgBox strMsg, vbOKOnly, LoadResST("U8.ST.USCONTROL.frmstockorder.00370") 'zh-CN：批量入库单据号指定
                End If
            End If
            KeyCode = 0
        End If
      Case vbKeyQ
        If refType <> "clssadispatchbatchrefer" And refType <> "clspuorderbatchrefer" And refType <> "clspuarrivebatchrefer" Then
            Exit Sub
        End If

        If Shift = vbCtrlMask Then
            If ctlReferMakeVouch1.BodyList.RecordCount < 1 Then
                Beep
              Else
                If refType = "clssadispatchbatchrefer" Then
                    sRefTackOut BodyColData("ccode"), True, False
                  Else
                    sRefTackIn BodyColData("cInVouchCode"), True, False
                End If
            End If
            KeyCode = 0
        End If

    End Select
End Sub

Public Function filter(Optional sError As String) As Boolean
    Set Head_filter = New UFGeneralFilter.FilterSrv
'    Set Head_filter.BehaviorObject = New clsFilterCallBack '  oFilterCallBack
    Set Body_filter = New UFGeneralFilter.FilterSrv
'    Set Body_filter.BehaviorObject = New clsFilterCallBack '  oFilterCallBack
    
    If (isBody = False) And Trim(sFilterID_head) <> "" Then '表头过滤
        If Head_filter.OpenFilter(mLogin, sFilterID_head, "", "", sError) = True Then
            filter = True
        Else
            Set Head_filter = Nothing
            filter = False
        End If
    ElseIf (isBody = True) And Trim(sFilterID_body) <> "" Then '表体过滤
        If Body_filter.OpenFilter(mLogin, sFilterID_body, "", "", sError) = True Then
            filter = True
        Else
            Set Body_filter = Nothing
            filter = False
        End If
    End If
    

Exit_Handler:

End Function

Private Sub ButtonClick(ByVal cButtonId As String)
    Dim sSubID As String
    Dim sError As Variant
    Dim cls_UI As Cls_UI_interface
    Dim strErr As String
    Dim Suc As Boolean
    
    
    sTemp = ""
    Select Case LCase(cButtonId)
      Case "tlbfiltersetup"
'        Set filterItf = New UFGeneralFilter.FilterSrv
'        If Not isBody Then
'            '                filterItf.SetFilter oRef.FitlerNameHead, "ST", stLogin.AccountConnection, , True
'            If refType = "clsmoplanrefer" Then '==gaojwadd==
'                filterItf.OpenFilterConfig stLogin.OldLogin, "", oRef.FitlerNameHead, "ST", sError
'            Else
'                filterItf.OpenFilterConfig stLogin.OldLogin, "", oRef.FitlerNameHead, "MOAT", sError
'            End If
'          Else
'            '                filterItf.SetFilter oRef.FitlerNameBody, "ST", stLogin.AccountConnection, , True
'            If refType = "clsmoplanrefer" Then '==gaojwadd==
'                 filterItf.OpenFilterConfig stLogin.OldLogin, "", oRef.FitlerNameBody, "ST", sError
'            Else
'                 filterItf.OpenFilterConfig stLogin.OldLogin, "", oRef.FitlerNameBody, "MOAT", sError
'            End If
'        End If
'        Set filterItf = Nothing
      Case "tlbrefresh"
        Screen.MousePointer = vbHourglass
        '            DoEvents

        SetData

        Screen.MousePointer = vbDefault
      Case "tlbsel"
        If Not isBody Then
            ctlReferMakeVouch1_HeadSelectClick True
        End If
      Case "tlbunsel"
        If Not isBody Then
                            ctlReferMakeVouch1_HeadSelectClick False
'            ctlReferMakeVouch1_HeadSelectClick True
        End If
      Case "tlblm"
                    If ctlReferMakeVouch1.Mode = enmRefreshBody Then
                        SetData
                    End If
                    If ctlReferMakeVouch1.Mode = enmRefreshHead Then
                        LoadLMString IIf(Not isBody, eListHead, eListBody)
                    End If

        If ctlReferMakeVouch1.Mode <> enmCancel Then
        
        End If
      Case "tlbfilter"
        Dim retFilter As Variant
        retFilter = filter()

'        If Not IsNull(retFilter) Then
'            oRef.WhereAll = retFilter & IIf(sMustWhere <> "", " and ", "") & sMustWhere
'            If Not isBody Then
'                oRef.WhereHead = retFilter & IIf(sMustWhere <> "", " and ", "") & sMustWhere
'              Else
'                oRef.WhereBody = retFilter & IIf(sMustWhere <> "", " and ", "") & sMustWhere
'            End If

            '                isBody = False
            SetData
'        End If
      Case "tlbmakevouch"  '确定
            Set cls_UI = New Cls_UI_interface
            '参照上半部分选中的数据
            Set RefdomH = Me.ctlReferMakeVouch1.GetHeadDom(True)
            '参照下半部分选中的数据
            Set RefdomB = Me.ctlReferMakeVouch1.GetBodyDom(True)
            cls_UI.Full_Voucher Source_Cardnum, Dest_Cardnum, RefdomH, RefdomB, DomHead_Dest, DomBody_Dest, strErr, Suc
            If Suc Then
                Me.Hide
            Else
                 If Auto_MakeVouch Then
                    Me.Hide
                End If
            End If
        
      Case "tlbexit"
'        oRef.ExitRefer oDic

        Unload Me
      Case Else

    End Select
    If Trim(sTemp) <> "" And sTemp <> vbCrLf Then
        MsgBox sTemp, vbOKOnly + vbInformation, GetResString("U8.SCM.ST.STReference.mdlreference.00073")
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If refType = "clssadispatchbatchrefer" Then
'        Call RegWrite(C_SPLIT_MODE, "dispatchform", CInt(ctlReferMakeVouch1.GetControlValue("chkUpdate")))
'      ElseIf refType = "clssadispatchrefer" Then '销售发货单单张生单
'        Call RegWrite(C_SPLIT_MODE, "dispatchform", CInt(ctlReferMakeVouch1.GetControlValue("chk")))
'    End If
'
'    Set oInvVOCache = Nothing
'    Set cbVTID = Nothing
'
'    On Error Resume Next
''    oRef.ExitRefer oDic
'
'    Dim resLog As Object
'    Set resLog = CreateObject("MultiLangPkg.ResLog")
'    resLog.Unload
End Sub





'
'实现将参照下半部分用户选择的数据自动填中到目标单据中
' RefdomB 对象中的 "t_*"部分的数据直接填中到目标单据的表体数据中
' RefdomB 对象中的 "b_*"部分的数据直接填中到目标单据的表体数据中
'
Private Function Auto_MakeVouch(Optional strErr As String) As Boolean
    Dim strSql As String
    Dim rds As New ADODB.Recordset
    Dim tmpdom As New DOMDocument
    Dim ele As IXMLDOMElement
    Dim Attributes As IXMLDOMAttribute
On Error GoTo OnErrexit

'处理表头填充
For Each Attributes In DomHead_Dest.selectSingleNode("//z:row").Attributes
    Debug.Print Attributes.nodeName & "    " & Attributes.Value
    If Left(LCase(Attributes.nodeName), 2) = "t_" Then
        If GetHeadItemValue(RefdomB, Attributes.nodeName) <> "" Then
            Attributes.Value = GetHeadItemValue(RefdomB, Attributes.nodeName)
        End If
    Else
        If GetHeadItemValue(RefdomB, "t_" & Attributes.nodeName) <> "" Then
            Attributes.Value = GetHeadItemValue(RefdomB, "t_" & Attributes.nodeName)
        End If
    End If
Next
DomHead_Dest.Save "c:\DomHead_Dest.xml"
DomBody_Dest.Save "c:\DomBody_Dest.xml"
RefdomB.Save "c:\RefdomB.xml"
'清除当前单据标题数据
If Not (DomBody_Dest.selectSingleNode("//rs:data") Is Nothing) Then
    DomBody_Dest.selectSingleNode("//xml").removeChild DomBody_Dest.selectSingleNode("//rs:data")
End If
'将当前的参照的数据强制 追加到目标单据上表体DOM 上
 If DomBody_Dest.selectSingleNode("//rs:data") Is Nothing Then
    DomBody_Dest.selectSingleNode("//xml").appendChild RefdomB.selectSingleNode("//rs:data").cloneNode(True)
 End If
'格式化表体DOM数据
FormatDom DomBody_Dest, tmpdom, "A"
Set DomBody_Dest = tmpdom '
For Each ele In DomBody_Dest.selectNodes("//z:row")
'循环表体数据
    For Each Attributes In ele.Attributes
        '循环每一行数据中的字段
        Debug.Print Attributes.nodeName & "    " & Attributes.Value
        
        If Left(LCase(Attributes.nodeName), 2) = "b_" Then
            Debug.Print Attributes.nodeName & "    " & Attributes.Value
            
        ElseIf Left(LCase(Attributes.nodeName), 2) = "t_" Then
            ele.removeAttributeNode Attributes
            
        Else
            If GetHeadItemValue(RefdomB, "b_" & Attributes.nodeName) <> "" Then
                Attributes.Value = GetHeadItemValue(RefdomB, "b_" & Attributes.nodeName)
            End If
        End If
    Next
Next
Auto_MakeVouch = True
Exit Function
OnErrexit:
    Auto_MakeVouch = False
    strErr = Err.Description
End Function


Private Function CreateNode(ByVal node_name As String, Optional ByVal node_value As String = "") As IXMLDOMNode
Dim new_node As IXMLDOMNode
Dim xml_document As DOMDocument
Set xml_document = New DOMDocument
Dim parent As IXMLDOMNode
Set parent = xml_document.createElement("Values")

' Create the new node.
Set CreateNode = parent.ownerDocument.createElement(node_name).cloneNode(True)
CreateNode.Text = node_value
' Set the node's text value.
'parent.text = node_value

' Add the node to the parent.
'parent.appendChild new_node
End Function


Private Sub SetFormat(VouchList As Object)
    On Error Resume Next
    Dim Col As IXMLDOMNode
    For Each Col In listFormat.documentElement.selectNodes("//formats/col")
        VouchList.SetFormatString Col.Attributes.getNamedItem("name").Text, Col.Attributes.getNamedItem("format").Text
    Next

    If refType = "clseqworkapprefer" Then VouchList.SetFormatString "fCurNum", moLogin.Account.FormatQuanDecString
End Sub

'格式化单据列表栏目
Private Sub setItemFormat(ByRef VouchList As Object, ByVal sVouchType As String)
    '设置格式化串
    Dim rec                 As ADODB.Recordset
    Dim sFormatQty             As String
    Dim sFormatNum             As String
    Dim sFormatCost            As String
    Dim sFormatExc             As String
    Const m_sDatefmt As String = "YYYY-MM-DD"
    sFormatQty = moLogin.Account.FormatQuanDecString
    sFormatNum = moLogin.Account.FormatNumDecString
    sFormatCost = moLogin.Account.FormatPriceDecString
    sFormatExc = moLogin.Account.FormatExchDecString
    'lliang-860-Develop-Bug-165166
    Set rec = New ADODB.Recordset
    'rec.open "SELECT AA_COLUMNDIC.CFLD,VOUCHERITEMS.FIELDTYPE,isnull(VOUCHERITEMS.NUMPOINT,0) as NUMPOINT,isnull(VOUCHERITEMS.FORMATDATA,0)AS FORMATDATA FROM AA_COLUMNDIC INNER JOIN VOUCHERITEMS ON AA_COLUMNDIC.CKEY=VOUCHERITEMS.CARDNUM AND AA_COLUMNDIC.CFLD=VOUCHERITEMS.FIELDNAME WHERE VOUCHERITEMS.FIELDTYPE IN (3,4,5) AND VOUCHERITEMS.VT_ID=" & Trim(lngVt_ID), moLogin.AccountConnection, adOpenStatic, adLockReadOnly, adCmdText
    rec.Open "select distinct cfld,ccaption from aa_columndic_base where LocaleID=N'zh-CN' and ckey=N'" & sVouchType & "'", moLogin.AccountConnection, adOpenStatic, adLockReadOnly, adCmdText
    Do While Not rec.EOF
        'Result:Row=4823        Col=38  Content="*数量*"        ID=e4409c8c-cf07-4592-9bf7-be5548a50937
        If (rec("ccaption").Value Like "*数量*") Then
            VouchList.SetFormatString rec("cfld").Value, sFormatQty
            GoTo EndLoop
        End If
        'Result:Row=4828        Col=38  Content="*件数*"        ID=1f085f52-afcd-415a-9ced-e448152dd774
        If (rec("ccaption").Value Like "*件数*") Then
            VouchList.SetFormatString rec("cfld").Value, sFormatNum
            GoTo EndLoop
        End If
        'Result:Row=4833        Col=38  Content="*单价*"        ID=c0de8587-6ff0-4976-a3b2-eae960be3ded
        If (rec("ccaption").Value Like "*单价*") Then
            VouchList.SetFormatString rec("cfld").Value, sFormatCost
            GoTo EndLoop
        End If
        'Result:Row=4838        Col=38  Content="*率"   ID=ff150902-29a6-49b1-b4d5-698ad3b44d96
        If (rec("ccaption").Value Like "*率") Then ' Or rec("ccaption").value Like "*率") Then
            VouchList.SetFormatString rec("cfld").Value, sFormatExc
            GoTo EndLoop
        End If
        'Result:Row=4843        Col=38  Content="*额"   ID=6217ff3c-453c-4a72-8791-d74bf1b813c2
        'Result:Row=4844        Col=73  Content="*费"   ID=698b74a8-da50-44d9-892c-2401e720749b
        If (rec("ccaption").Value Like "*额" Or rec("ccaption").Value Like "*费") Then  'Or rec("ccaption").value Like "*额") Then
            VouchList.SetFormatString rec("cfld").Value, "#,##0.00"
            GoTo EndLoop
        End If
        'Result:Row=4849        Col=38  Content="*日期*"        ID=2ff9e858-fd0b-4c15-8aae-a89c54ceb062
        If (rec("ccaption").Value Like "*日期*") Then
            VouchList.SetFormatString rec("cfld").Value, m_sDatefmt
            GoTo EndLoop
        End If

EndLoop:
        rec.MoveNext
    Loop
    rec.Close
    Set rec = Nothing
    VouchList.DoFormat
End Sub

Private Sub ChangeEditState()
    If refType = "clssadispatchrefer" Or refType = "clspuorderrefer" Or refType = "clspuarriverefer" Or refType = "clsomorderrefer" Or refType = "clsomarriverefer" Then
        ctlReferMakeVouch1.BodyColSetXml = Replace(ctlReferMakeVouch1.BodyColSetXml, "CanModify=""1""", "CanModify=""0""")
        ctlReferMakeVouch1.BodyColSetXml = Replace(ctlReferMakeVouch1.BodyColSetXml, "ReferType=""1""", "ReferType=""""")
        ctlReferMakeVouch1.BodyColSetXml = Replace(ctlReferMakeVouch1.BodyColSetXml, "ReferType=""2""", "ReferType=""""")
        ctlReferMakeVouch1.BodyColSetXml = Replace(ctlReferMakeVouch1.BodyColSetXml, "ReferType=""3""", "ReferType=""""")
    End If
End Sub

Private Function sRefWh(Optional ByRef cWhCode As String) As String
    Dim strFilterSQL As String
    strFilterSQL = moLogin.GetAuthString("Warehouse", , "W")

    Dim obj As New U8RefService.IService
    Dim sMetaXML As String
    sMetaXML = "<Ref><RefSet bAuth='0' /></Ref>"
    obj.RefID = "Warehouse_AA"
    obj.Mode = RefModes.modeRefing
    obj.Web = False
    obj.MetaXML = sMetaXML

    If strFilterSQL <> "" Then
        strFilterSQL = "#FN[cWhCode] IN (" & strFilterSQL & ")"
    End If

    obj.FilterSQL = strFilterSQL

    Dim retRstClass As ADODB.Recordset, retRstGrid As ADODB.Recordset
    Dim sErrMsg As String
    If obj.ShowRef(mLogin, retRstClass, retRstGrid, sErrMsg) = False Then
        MsgBox sErrMsg
      Else
        If Not (retRstGrid Is Nothing) Then
            sRefWh = retRstGrid.Fields("cWhName").Value
            cWhCode = retRstGrid.Fields("cWhCode").Value
        End If
    End If
    Set obj = Nothing
End Function

Private Function sRefFree(Optional cFree As String, Optional cLike As String) As String
    '自由项参照
    Dim ssql As String
    Dim Ref As UFReferC.UFReferClient
    Set Ref = New UFReferC.UFReferClient
    Call Ref.SetLogin(moLogin.OldLogin)
    ClsBill.RefUserdef cFree, cLike, ssql, False
    Ref.SetReferDisplayMode enuGrid
    If Ref.StrRefInit(moLogin.OldLogin, True, "", ssql, LoadResST("U8.ST.USCONTROL.frmstockorder.00387")) = True Then 'zh-CN：代码,值,条码
        Ref.Show
    End If
    If Not (Ref.recmx Is Nothing) Then
        sRefFree = Ref.recmx(1)
    End If
End Function

Private Function sRefTackOut(ByRef sBatch As String, Optional ByVal bAutoRef As Boolean = False, Optional bAllBatch As Boolean = False, Optional ByRef sReturn As String, Optional ByRef sReturnErr As String) As String
    Dim oBatchPst As BatchPst
    Dim oRefSelect As RefSelect
    Dim sFree As Collection
    Dim sInVouchList As String
    Dim sFreeName As String
    Dim bFg As Boolean
    Dim nRows As Long
    Dim i As Long
    Dim iCurRow As Long
    iCurRow = ctlReferMakeVouch1.BodyList.row
    Dim oTmpCO As USERPCO.VoucherCO
    Set sFree = New Collection
    Set oRefSelect = New RefSelect
    For i = 1 To 10
        sFree.Add BodyColData("cfree" & i)
    Next
    Set oBatchPst = New BatchPst
    Set oTmpCO = ClsBill 'New USERPCO.voucherCo
    sTemp = ""
    If True Then 'oTmpCO.IniLogin(moLogin.OldLogin, sTemp) Then
        oBatchPst.Login = oTmpCO.Login
        'If val2(oLstSup.TextMatrix(oLstSup.row, oLstSup.cols - 8)) < 0 Then
        If val2(BodyColData("iquantity")) < 0 And bAutoRef Then
            MsgBox LoadResST("U8.ST.USCONTROL.frmcons.00748") 'zh-CN：红字发货单不支持自动指定功能，请手工指定。
            Exit Function
            'oBatchPst.BackList sInVouchList, oTmpRs("cWhCode"), oTmpRs("cInvCode"), sFree(1), sFree(2), sFree(3), sFree(4), _
                sFree(5), sFree(6), sFree(7), sFree(8), sFree(9), sFree(10), vFieldVal(oTmpRs("cCode")), vFieldVal(oTmpRs("cbatch"))
          Else
            isosid = val2(BodyColData("iSoSid"))
            If isosid <> 0 Then
                iSotype = 1
              Else
                iSotype = 0
            End If
            If val2(BodyColData("iquantity")) < 0 Then
                oBatchPst.BackList sInVouchList, BodyColData("cWhCode"), BodyColData("cInvCode"), sFree(1), sFree(2), sFree(3), sFree(4), _
                                   sFree(5), sFree(6), sFree(7), sFree(8), sFree(9), sFree(10), BodyColData("cinvouchcode"), BodyColData("cbatch"), , iSotype, isosid, BodyColData("cvmivencode")
              Else
                oBatchPst.List sInVouchList, BodyColData("cWhCode"), BodyColData("cInvCode"), sFree(1), sFree(2), sFree(3), sFree(4), _
                               sFree(5), sFree(6), sFree(7), sFree(8), sFree(9), sFree(10), BodyColData("cinvouchcode"), , , BodyColData("cbatch"), iSotype, isosid, BodyColData("cvmivencode")
            End If
        End If

        '如果全部自动指定入库单号，存储本次分配数量,同时将读现存量的数据改写。
        If bAllBatch Then
            sInVouchList = oRefSelect.GetAllRDSSQL(sInVouchList)
        End If

        If val2(BodyColData("iquantity")) < 0 Then
            If bAutoRef Then
                oRefSelect.AutoRefer moLogin.OldLogin, BodyColData("cWhCode"), BodyColData("cInvCode"), sFree, val2(BodyColData("fcurqty")), _
                                     val2(Replace(BodyColData("fcurnum"), ",", "")), _
                                     val2(Replace(BodyColData("iInvExchRate"), ",", "")), refoutvouch, sInVouchList, True, , , sReturnErr, , , BodyColData("cvmivencode")
              Else
                oRefSelect.Refer moLogin.OldLogin, BodyColData("cWhCode"), BodyColData("cInvCode"), sFree, val2(BodyColData("fcurqty")), _
                                 val2(Replace(BodyColData("fcurnum"), ",", "")), _
                                 val2(Replace(BodyColData("iInvExchRate"), ",", "")), refoutvouch, sInVouchList, True, , , , , BodyColData("cvmivencode")
            End If
          Else
            If bAutoRef Then
                oRefSelect.AutoRefer moLogin.OldLogin, BodyColData("cWhCode"), BodyColData("cInvCode"), sFree, val2(BodyColData("fcurqty")), _
                                     val2(Replace(BodyColData("fcurnum"), ",", "")), _
                                     val2(Replace(BodyColData("iInvExchRate"), ",", "")), RefInVouch, sInVouchList, , , , sReturnErr, , , BodyColData("cvmivencode")
              Else
                oRefSelect.Refer moLogin.OldLogin, BodyColData("cWhCode"), BodyColData("cInvCode"), sFree, val2(BodyColData("fcurqty")), _
                                 val2(Replace(BodyColData("fcurnum"), ",", "")), _
                                 val2(Replace(BodyColData("iInvExchRate"), ",", "")), RefInVouch, sInVouchList, , , , , , BodyColData("cvmivencode")
            End If
        End If

        If bAllBatch And sReturnErr <> "" Then
            sReturnErr = LoadResST("U8.ST.USCONTROL.frmstockorder.00490") & BodyColData("cInvCode") & sReturnErr 'zh-CN：存货
        End If

        Dim recRef As New ADODB.Recordset
        If Not oRefSelect.ReturnData Is Nothing Then
            ctlReferMakeVouch1.BodyTextMatrix(iCurRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iQuantity"))) = val2(BodyColData("iquantity"))
            ctlReferMakeVouch1.BodyTextMatrix(iCurRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iNum"))) = val2(BodyColData("inum"))

            Set recRef = oRefSelect.ReturnData
            Do While Not recRef.EOF
                For i = 0 To recRef.Fields.Count - 1
                    sFreeName = SetNull(recRef.Fields(i).Properties("BASECOLUMNNAME"), "")
                    If LCase(Mid(sFreeName, 1, 5)) = "cfree" Then
                        BodyColData(sFreeName) = recRef.Fields(i).Value
                      Else
                        Exit For
                    End If
                Next
                If bFg = False Then
                    bFg = True
                  Else
                    ctlReferMakeVouch1.AddBodyLine iCurRow, ctlReferMakeVouch1.GetBodyLine(iCurRow)

                    For i = 0 To ctlReferMakeVouch1.BodyList.Cols - 1
                        ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(iCurRow, i) = ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(iCurRow + 1, i)
                        If i > 0 And Not ctlReferMakeVouch1.BodyList.FieldItem(i + 1).CanRevisable Then
                            ctlReferMakeVouch1.DisableBodyTextMatrix True, LCase(ctlReferMakeVouch1.GetBodyColName(i)), iCurRow
                        End If
                    Next
                    ctlReferMakeVouch1.DisableBodyTextMatrix True, LCase("cwhname"), iCurRow

                End If
                ctlReferMakeVouch1.BodyTextMatrix(iCurRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cbatch"))) = SetNull(recRef.Fields("批号"), "")
                BodyColData("dvdate") = SetNull(recRef.Fields("失效日期"), "")
                If val2(BodyColData("fnqty")) < 0 Then
                    BodyColData("fcurqty") = Format(recRef.Fields("退回数量"), moLogin.Account.FormatQuanDecString)
                    BodyColData("fcurnum") = Format(SetNull(recRef.Fields("退回件数"), ""), moLogin.Account.FormatNumDecString)
                  Else
                    BodyColData("fcurqty") = Format(recRef.Fields("出库数量"), moLogin.Account.FormatQuanDecString)
                    BodyColData("fcurnum") = Format(SetNull(recRef.Fields("出库件数"), ""), moLogin.Account.FormatNumDecString)
                End If

                ctlReferMakeVouch1.BodyTextMatrix(iCurRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cinvouchcode"))) = SetNull(recRef.Fields("入库单号"), "")
                BodyColData("ccode") = SetNull(recRef.Fields("入库单号"), "")

                sRefTackOut = SetNull(recRef.Fields("入库单号"), "")

                BodyColData("iBatch") = recRef.Fields("入库系统编号") '返回入库Id
                BodyColData("iMassDate") = SetNull(recRef.Fields("保质期"), "")
                BodyColData("dMadeDate") = SetNull(recRef.Fields("生产日期"), "")

                If val2(BodyColData("iQuantity")) < 0 Then

                    '因为iCorId可能另有他用，所以在此借用AutoId来保存[出库系统编号]
                    BodyColData("AutoId") = recRef.Fields("出库系统编号")
                End If

                '如果自动整批指定入库单据号，则更新临时表数据
                If bAllBatch Then
                    'by dcb-dev-119752-2004-8-25
                    oRefSelect.UpdateTmpRdrecords mLogin, IIf(IsNull(recRef.Fields("出库数量")), 0, recRef.Fields("出库数量")), IIf(IsNull(recRef.Fields("出库件数")), 0, recRef.Fields("出库件数")), recRef.Fields("入库系统编号")
                End If

                '                If bFg Then oTmpRs.Update
                recRef.MoveNext
            Loop
        End If
    End If
End Function

Private Function sRefBatchOut(ByRef sBatch As String, Optional ByVal bAutoRef As Boolean = False, Optional ByVal bAllBatch As Boolean = False, Optional sReturn As String, Optional ByRef sReturnErr As String) As String
    Dim recRef As New ADODB.Recordset
    Dim oStockPst As StockPst
    Dim oRefSelect As RefSelect
    Dim sFree As Collection
    Dim sInVouchList As String
    Dim sFreeName As String
    Dim bFg As Boolean
    Dim i As Long
    Dim oTmpCO As USERPCO.VoucherCO
    Set sFree = New Collection
    Set oRefSelect = New RefSelect
    For i = 1 To 10
        sFree.Add BodyColData("cfree" & i)
    Next
    Set oStockPst = New StockPst
    Set oTmpCO = ClsBill 'New USERPCO.voucherCo
    sTemp = ""
    Dim iCurRow As Long
    iCurRow = ctlReferMakeVouch1.BodyList.row
    If True Then 'oTmpCO.IniLogin(moLogin.OldLogin, sTemp) Then
        If val2(BodyColData("iquantity")) < 0 Then
            Exit Function    '红字发货单不支持自动参照批号
        End If
        oStockPst.Login = oTmpCO.Login
        isosid = val2(BodyColData("iSoSid"))
        If isosid <> 0 Then
            iSotype = 1
          Else
            iSotype = 0
        End If
        oStockPst.BatchList sInVouchList, BodyColData("cWhCode"), BodyColData("cInvCode"), sFree(1), sFree(2), sFree(3), sFree(4), _
                            sFree(5), sFree(6), sFree(7), sFree(8), sFree(9), sFree(10), BodyColData("cBatch"), , iSotype, isosid, BodyColData("cvmivencode")

        '批量指定批次
        If bAllBatch Then
            sInVouchList = oRefSelect.GetAllBSQL(sInVouchList)
        End If

        If bAutoRef Then
            oRefSelect.AutoRefer moLogin.OldLogin, BodyColData("cWhCode"), BodyColData("cInvCode"), sFree, val2(BodyColData("fcurqty")), val2(Replace(BodyColData("fcurnum"), ",", "")), val2(Replace(BodyColData("iinvexchrate"), ",", "")), RefBatch, sInVouchList, , , , sReturnErr, , , BodyColData("cvmivencode")
            If bAllBatch And sReturnErr <> "" Then
                sReturnErr = LoadResST("U8.ST.USCONTROL.frmstockorder.00490") & BodyColData("cInvCode") & sReturnErr 'zh-CN：存货
            End If
          Else
            oRefSelect.Refer moLogin.OldLogin, BodyColData("cWhCode"), BodyColData("cInvCode"), sFree, val2(BodyColData("fcurqty")), val2(Replace(BodyColData("fcurnum"), ",", "")), val2(Replace(BodyColData("iinvexchrate"), ",", "")), RefBatch, sInVouchList, , , , , , BodyColData("cvmivencode")
        End If

        If Not oRefSelect.ReturnData Is Nothing Then
            ctlReferMakeVouch1.BodyTextMatrix(iCurRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iQuantity"))) = val2(BodyColData("iquantity"))
            ctlReferMakeVouch1.BodyTextMatrix(iCurRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iNum"))) = val2(BodyColData("inum"))

            Set recRef = oRefSelect.ReturnData
            Do While Not recRef.EOF
                For i = 0 To recRef.Fields.Count - 1
                    sFreeName = IIf(IsNull(recRef.Fields(i).Properties("BASECOLUMNNAME")), "", recRef.Fields(i).Properties("BASECOLUMNNAME"))
                    If LCase(Mid(sFreeName, 1, 5)) = "cfree" Then
                        BodyColData(sFreeName) = recRef.Fields(i).Value
                      Else
                        Exit For
                    End If
                Next
                If bFg = False Then
                    bFg = True
                  Else
                    ctlReferMakeVouch1.AddBodyLine iCurRow, ctlReferMakeVouch1.GetBodyLine(iCurRow)

                    For i = 0 To ctlReferMakeVouch1.BodyList.Cols - 1
                        ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(iCurRow, i) = ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(iCurRow + 1, i)
                        If i > 0 And Not ctlReferMakeVouch1.BodyList.FieldItem(i + 1).CanRevisable Then
                            ctlReferMakeVouch1.DisableBodyTextMatrix True, LCase(ctlReferMakeVouch1.GetBodyColName(i)), iCurRow
                        End If
                    Next
                    ctlReferMakeVouch1.DisableBodyTextMatrix True, LCase("cwhname"), iCurRow

                End If
                ctlReferMakeVouch1.BodyTextMatrix(iCurRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cBatch"))) = SetNull(recRef.Fields("批号"), "")

                sRefBatchOut = SetNull(recRef.Fields("批号"), "")

                If val2(BodyColData("fnqty")) < 0 Then
                    BodyColData("fcurqty") = Format(recRef.Fields("退回数量"), moLogin.Account.FormatQuanDecString)
                    BodyColData("fcurnum") = Format(SetNull(recRef.Fields("退回件数"), ""), moLogin.Account.FormatNumDecString)
                  Else
                    BodyColData("fcurqty") = Format(recRef.Fields("出库数量"), moLogin.Account.FormatQuanDecString)
                    BodyColData("fcurnum") = Format(SetNull(recRef.Fields("出库件数"), ""), moLogin.Account.FormatNumDecString)
                End If

                BodyColData("dVdate") = SetNull(recRef.Fields("失效日期"), "")
                BodyColData("imassdate") = SetNull(recRef.Fields("保质期"), "")
                BodyColData("dmadedate") = SetNull(recRef.Fields("生产日期"), "")

                '如果自动整批指定，则更新临时表数据
                If bAllBatch Then
                    isosid = val2(BodyColData("iSOSID"))
                    If isosid <> 0 Then
                        iSotype = 1
                      Else
                        iSotype = 0
                    End If
                    oRefSelect.UpdateTmpCurrentStock mLogin, recRef.Fields("出库数量"), recRef.Fields("出库件数"), oRefSelect.GetAutoID(mLogin, BodyColData("cWhCode"), BodyColData("cInvCode"), sFree(1), sFree(2), sFree(3), sFree(4), sFree(5), sFree(6), sFree(7), sFree(8), sFree(9), sFree(10), recRef.Fields("批号"), iSotype, CStr(isosid))
                End If
                '同步92664问题
                Call ctlReferMakeVouch1_BodyCellValueChanged(iCurRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cBatch")), sRefBatchOut, sRefBatchOut, False)

                recRef.MoveNext
            Loop
        End If
    End If
End Function

Private Function sRefUnit(ByVal colName As String, ByVal OldValue As String) As String
    Dim moInventory As USERPVO.Inventory
    Dim ClsInventoryCO As USERPCO.InventoryCO
    Set moInventory = New USERPVO.Inventory
    Set ClsInventoryCO = New USERPCO.InventoryCO
    ClsInventoryCO.Login = moLogin
    sTemp = ""
    If ClsInventoryCO.Load(BodyColData("cinvcode"), moInventory, sTemp, True) Then
        Dim obj As New U8RefService.IService
        Dim sMetaXML As String
        sMetaXML = "<Ref><RefSet bAuth='0' /></Ref>"
        obj.RefID = "ComputationUnit_AA"
        obj.Mode = RefModes.modeRefing
        obj.Web = False
        obj.MetaXML = sMetaXML
        If NoBlank(CStr(OldValue)) Then
            obj.FilterSQL = "cComUnitCode" & " like '%" & OldValue & "%'" & " or " & "cComUnitName" & " like '%" & OldValue & "%'"
        End If

        Dim retRstClass As ADODB.Recordset, retRstGrid As ADODB.Recordset
        Dim sErrMsg As String
        If obj.ShowRef(mLogin, retRstClass, retRstGrid, sErrMsg) = False Then
            MsgBox sErrMsg
          Else
            If Not (retRstGrid Is Nothing) Then
                sRefUnit = retRstGrid.Fields("cComUnitName").Value
            End If
        End If
        Set obj = Nothing

    End If
End Function

Private Function sRefBatchIn(ByRef sBatch As String, Optional ByVal bAutoRef As Boolean = False, Optional bAllBatch As Boolean = False, Optional sReturn As String, Optional ByRef sReturnErr As String) As String
    Dim bFg As Boolean
    Dim recRef As New ADODB.Recordset
    Dim oStockPst As StockPst
    Dim oRefSelect As RefSelect
    Dim sFree As Collection
    Dim sInVouchList As String
    Dim sFreeName As String
    Dim i As Long
    Dim oTmpCO As USERPCO.VoucherCO
    Set sFree = New Collection
    Set oRefSelect = New RefSelect
    Dim iCurRow As Long
    iCurRow = ctlReferMakeVouch1.BodyList.row
    If val2(BodyColData("fcurqty")) > 0 Then
        Exit Function   '蓝字到货单不自动参照批号
    End If
    If SetNull(BodyColData("cWhCode"), "") = "" Then
        If bAllBatch = True Then
            sReturnErr = BodyColData("cInvCode") & LoadResST("U8.ST.USCONTROL.frmstockorder.00477") 'zh-CN：仓库不能为空
            Exit Function
          Else
            MsgBox LoadResST("U8.ST.USCONTROL.frmstockorder.00477"), vbOKOnly + vbCritical, LoadResST("U8.ST.USCONTROL.frmproduceorder.00594") 'zh-CN：仓库不能为空 'zh-CN：系统提示
            sRefBatchIn = ""
            Exit Function
        End If
    End If
    For i = 1 To 10
        sFree.Add BodyColData("cfree" & i)
    Next
    Set oStockPst = New StockPst
    Set oTmpCO = ClsBill 'New USERPCO.voucherCo
    If True Then 'oTmpCO.IniLogin(moLogin.OldLogin, sTemp) Then
        oStockPst.Login = oTmpCO.Login
        oStockPst.BatchList sInVouchList, BodyColData("cWhCode"), BodyColData("cInvCode"), sFree(1), sFree(2), sFree(3), sFree(4), _
                            sFree(5), sFree(6), sFree(7), sFree(8), sFree(9), sFree(10), BodyColData("cBatch"), , val2(BodyColData("SoType")), val2(BodyColData("SoDID")), IIf(GetHeadBusType(iCurRow) = "代管采购", ctlReferMakeVouch1.HeadList.TextMatrix(GetHeadRow(iCurRow), ctlReferMakeVouch1.HeadList.GridColIndex(LCase("cvencode"))), "")

        '批量指定批次
        If bAllBatch Then
            sInVouchList = oRefSelect.GetAllBSQL(sInVouchList)
        End If

        If bAutoRef Then
            oRefSelect.AutoRefer moLogin.OldLogin, BodyColData("cWhCode"), BodyColData("cInvCode"), sFree, val2(BodyColData("fcurqty")), val2(Replace(BodyColData("fcurnum"), ",", "")), val2(Replace(BodyColData("iinvexchrate"), ",", "")), RefBatch, sInVouchList, IIf(val2(BodyColData("fcurqty")) < 0, True, False), , , sReturnErr, , , IIf(GetHeadBusType(iCurRow) = "代管采购", ctlReferMakeVouch1.HeadList.TextMatrix(GetHeadRow(iCurRow), ctlReferMakeVouch1.HeadList.GridColIndex(LCase("cvencode"))), "")
          Else
            oRefSelect.Refer moLogin.OldLogin, BodyColData("cWhCode"), BodyColData("cInvCode"), sFree, val2(BodyColData("fcurqty")), val2(Replace(BodyColData("fcurnum"), ",", "")), val2(Replace(BodyColData("iinvexchrate"), ",", "")), RefBatch, sInVouchList, IIf(val2(BodyColData("fcurqty")) < 0, True, False), , , , , IIf(GetHeadBusType(iCurRow) = "代管采购", ctlReferMakeVouch1.HeadList.TextMatrix(GetHeadRow(iCurRow), ctlReferMakeVouch1.HeadList.GridColIndex(LCase("cvencode"))), "")
        End If

        '        oTmpRs("iQuantity") = oLstSup.TextMatrix(oLstSup.row, oLstSup.cols - 8)
        '        oTmpRs("iNum") = oLstSup.TextMatrix(oLstSup.row, oLstSup.cols - 7)

        '多行指定批次给出错误提示
        If sReturnErr <> "" Then sReturnErr = BodyColData("cInvCode") & sReturnErr

        If Not oRefSelect.ReturnData Is Nothing Then
            ctlReferMakeVouch1.BodyTextMatrix(iCurRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iQuantity"))) = val2(BodyColData("iquantity"))
            ctlReferMakeVouch1.BodyTextMatrix(iCurRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iNum"))) = val2(BodyColData("inum"))

            '           oTmpRsbak("iQuantity") = val2(oLstSup.TextMatrix(oLstSup.row, oLstSup.cols - 8))
            '           oTmpRsbak("iNum") = val2(oLstSup.TextMatrix(oLstSup.row, oLstSup.cols - 7))
            Set recRef = oRefSelect.ReturnData
            Do While Not recRef.EOF
                For i = 0 To recRef.Fields.Count - 1
                    sFreeName = IIf(IsNull(recRef.Fields(i).Properties("BASECOLUMNNAME")), "", recRef.Fields(i).Properties("BASECOLUMNNAME"))
                    If LCase(Mid(sFreeName, 1, 5)) = "cfree" Then
                        BodyColData(sFreeName) = recRef.Fields(i).Value
                      Else
                        Exit For
                    End If
                Next
                If bFg = False Then
                    bFg = True
                  Else
                    ctlReferMakeVouch1.AddBodyLine iCurRow, ctlReferMakeVouch1.GetBodyLine(iCurRow)
                    '                oTmpRs.AddNew
                    '                For i = 0 To oTmpRsbak.fields.count - 1
                    '                    oTmpRs(i).value = oTmpRsbak(i).value
                    '                Next
                    '                iCurRow = iCurRow + 1
                    '                ctlReferMakeVouch1.BodyList.row = iCurRow

                    For i = 0 To ctlReferMakeVouch1.BodyList.Cols - 1
                        ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(iCurRow, i) = ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(iCurRow + 1, i)
                        'If ctlReferMakeVouch1.BodyList.GetGridBody().Cell(6, iCurRow + 1, i, iCurRow + 1, i) <> 0 Then
                        If i > 0 And Not ctlReferMakeVouch1.BodyList.FieldItem(i + 1).CanRevisable Then
                            ctlReferMakeVouch1.DisableBodyTextMatrix True, LCase(ctlReferMakeVouch1.GetBodyColName(i)), iCurRow
                        End If
                    Next

                End If
                ctlReferMakeVouch1.BodyTextMatrix(iCurRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cBatch"))) = SetNull(recRef.Fields("批号"), "")

                sRefBatchIn = SetNull(recRef.Fields("批号"), "")

                '              oTmpRs("iQuantity") = recRef.Fields("出库数量")
                '              oTmpRs("iNum") = recRef.Fields("出库件数")
                BodyColData("fcurqty") = Format(recRef.Fields("出库数量"), moLogin.Account.FormatQuanDecString)
                BodyColData("fcurnum") = Format(SetNull(recRef.Fields("出库件数"), ""), moLogin.Account.FormatNumDecString)

                BodyColData("dVdate") = SetNull(recRef.Fields("失效日期"), "")
                BodyColData("imassdate") = SetNull(recRef.Fields("保质期"), "")
                '              If LCase(oLstSup.name) = "lstdhdbody" Then
                '                  oTmpRs("dpdate") = recRef.Fields("生产日期")
                '              Else
                BodyColData("dmadedate") = SetNull(recRef.Fields("生产日期"), "")
                '              End If

                '如果自动整批指定，则更新临时表数据
                If bAllBatch Then
                    oRefSelect.UpdateTmpCurrentStock moLogin.OldLogin, (-recRef.Fields("出库数量")), (-recRef.Fields("出库件数")), oRefSelect.GetAutoID(moLogin.OldLogin, BodyColData("cWhCode"), BodyColData("cInvCode"), sFree(1), sFree(2), sFree(3), sFree(4), sFree(5), sFree(6), sFree(7), sFree(8), sFree(9), sFree(10), recRef.Fields("批号"))
                End If

                recRef.MoveNext
            Loop
            '            If oTmpRs.RecordCount = 1 Then
            '                iniSupperGrid oTmpRs, oLstSup, True, , , recRef, True
            '            Else
            '                sReturn = "M"
            '                oLstSup.Redraw = False
            '                oLstSup.RowHeight(oLstSup.row) = 0
            '                iniSupperGrid oTmpRs, oLstSup, True, , True, recRef
            '                SendKeys "{ESC}"
            '                oLstSup.Redraw = True
            '            End If
        End If
    End If
End Function

Private Function sRefTackIn(ByRef sBatch As String, Optional ByVal bAutoRef As Boolean = False, Optional ByVal bAllBatch As Boolean = False, Optional ByRef sReturn As String, Optional ByRef sReturnErr As String) As String
    Dim oBatchPst As BatchPst
    Dim oRefSelect As RefSelect
    Dim sFree As Collection
    Dim bFg As Boolean
    Dim sInVouchList As String
    Dim sFreeName As String
    Dim i As Long
    Dim oTmpCO As USERPCO.VoucherCO
    Dim iCurRow As Long
    iCurRow = ctlReferMakeVouch1.BodyList.row
    If val2(BodyColData("fcurqty")) < 0 Then
        Set sFree = New Collection
        Set oRefSelect = New RefSelect

        If SetNull(BodyColData("cWhCode"), "") = "" Then
            If bAllBatch = True Then
                sReturnErr = BodyColData("cInvCode") & LoadResST("U8.ST.USCONTROL.frmstockorder.00477") 'zh-CN：仓库不能为空
                Exit Function
              Else
                MsgBox LoadResST("U8.ST.USCONTROL.frmstockorder.00477"), vbOKOnly + vbExclamation, LoadResST("U8.ST.USCONTROL.frmproduceorder.00594") 'zh-CN：仓库不能为空 'zh-CN：系统提示
                sRefTackIn = ""
                Exit Function
            End If
        End If

        For i = 1 To 10
            sFree.Add BodyColData("cfree" & i)
        Next
        Set oBatchPst = New BatchPst
        Set oTmpCO = ClsBill 'New USERPCO.voucherCo
        If True Then 'oTmpCO.IniLogin(moLogin.OldLogin, sTemp) Then
            oBatchPst.Login = oTmpCO.Login
            oBatchPst.List sInVouchList, BodyColData("cWhCode"), BodyColData("cInvCode"), sFree(1), sFree(2), sFree(3), sFree(4), _
                           sFree(5), sFree(6), sFree(7), sFree(8), sFree(9), sFree(10), BodyColData("cInvouchcode"), , , BodyColData("cbatch"), val2(BodyColData("SoType")), val2(BodyColData("SoDID")), IIf(GetHeadBusType(iCurRow) = "代管采购", ctlReferMakeVouch1.HeadList.TextMatrix(GetHeadRow(iCurRow), ctlReferMakeVouch1.HeadList.GridColIndex(LCase("cvencode"))), "")

            '如果全部自动指定入库单号，存储本次分配数量,同时将读现存量的数据改写。
            If bAllBatch Then
                sInVouchList = oRefSelect.GetAllRDSSQL(sInVouchList)
            End If

            If bAutoRef Then
                oRefSelect.AutoRefer moLogin.OldLogin, BodyColData("cWhCode"), BodyColData("cInvCode"), sFree, val2(BodyColData("fcurqty")), _
                                     val2(Replace(BodyColData("fcurnum"), ",", "")), _
                                     val2(Replace(BodyColData("iInvExchRate"), ",", "")), RefInVouch, sInVouchList, True, , , sReturnErr, , , IIf(GetHeadBusType(iCurRow) = "代管采购", ctlReferMakeVouch1.HeadList.TextMatrix(GetHeadRow(iCurRow), ctlReferMakeVouch1.HeadList.GridColIndex(LCase("cvencode"))), "")
              Else
                oRefSelect.Refer moLogin.OldLogin, BodyColData("cWhCode"), BodyColData("cInvCode"), sFree, val2(BodyColData("fcurqty")), _
                                 val2(Replace(BodyColData("fcurnum"), ",", "")), _
                                 val2(Replace(BodyColData("iInvExchRate"), ",", "")), RefInVouch, sInVouchList, True, , , , , IIf(GetHeadBusType(iCurRow) = "代管采购", ctlReferMakeVouch1.HeadList.TextMatrix(GetHeadRow(iCurRow), ctlReferMakeVouch1.HeadList.GridColIndex(LCase("cvencode"))), "")
            End If
            Dim recRef As ADODB.Recordset
            Set recRef = oRefSelect.ReturnData

            If bAllBatch And sReturnErr <> "" Then
                sReturnErr = "存货" & BodyColData("cInvCode") & sReturnErr
            End If

            If Not recRef Is Nothing Then
                ctlReferMakeVouch1.BodyTextMatrix(iCurRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iQuantity"))) = val2(BodyColData("iquantity"))
                ctlReferMakeVouch1.BodyTextMatrix(iCurRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iNum"))) = val2(BodyColData("inum"))

                '                oTmpRsbak("iQuantity") = val2(oLstSup.TextMatrix(oLstSup.row, oLstSup.cols - 8))
                '                oTmpRsbak("iNum") = val2(oLstSup.TextMatrix(oLstSup.row, oLstSup.cols - 7))
                Do While Not recRef.EOF
                    For i = 0 To recRef.Fields.Count - 1
                        sFreeName = SetNull(recRef.Fields(i).Properties("BASECOLUMNNAME"), "")
                        If LCase(Mid(sFreeName, 1, 5)) = "cfree" Then
                            BodyColData(sFreeName) = recRef.Fields(i).Value
                          Else
                            Exit For
                        End If
                    Next
                    If bFg = False Then
                        bFg = True
                      Else
                        ctlReferMakeVouch1.AddBodyLine iCurRow, ctlReferMakeVouch1.GetBodyLine(iCurRow)
                        '                        oTmpRs.AddNew
                        '                        For i = 0 To oTmpRsbak.fields.count - 1
                        '                            oTmpRs(i).value = oTmpRsbak(i)
                        '                        Next
                        '                        iCurRow = iCurRow + 1
                        '                        ctlReferMakeVouch1.BodyList.row = iCurRow

                        For i = 0 To ctlReferMakeVouch1.BodyList.Cols - 1
                            ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(iCurRow, i) = ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(iCurRow + 1, i)
                            'If ctlReferMakeVouch1.BodyList.GetGridBody().Cell(6, iCurRow + 1, i, iCurRow + 1, i) <> 0 Then
                            If i > 0 And Not ctlReferMakeVouch1.BodyList.FieldItem(i + 1).CanRevisable Then
                                ctlReferMakeVouch1.DisableBodyTextMatrix True, LCase(ctlReferMakeVouch1.GetBodyColName(i)), iCurRow
                            End If
                        Next

                    End If
                    ctlReferMakeVouch1.BodyTextMatrix(iCurRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cbatch"))) = SetNull(recRef.Fields("批号"), "")
                    BodyColData("dvdate") = SetNull(recRef.Fields("失效日期"), "")
                    '            oTmpRs("iquantity") = recRef.Fields("出库数量")
                    '            oTmpRs("inum") = recRef.Fields("出库件数")
                    BodyColData("fcurqty") = Format(recRef.Fields("出库数量"), moLogin.Account.FormatQuanDecString)
                    BodyColData("fcurnum") = Format(SetNull(recRef.Fields("出库件数"), ""), moLogin.Account.FormatNumDecString)

                    ctlReferMakeVouch1.BodyTextMatrix(iCurRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cInVouchCode"))) = SetNull(recRef.Fields("入库单号"), "")

                    sRefTackIn = SetNull(recRef.Fields("入库单号"), "")

                    BodyColData("iCorId") = recRef.Fields("入库系统编号")  '返回入库Id
                    '            oTmpRs("cVencode") = recRef.Fields("供货单位编码")  '返回供货单位
                    '            oTmpRs("cvenname") = recRef.Fields("供货单位")  '返回供货单位
                    BodyColData("iMassDate") = SetNull(recRef.Fields("保质期"), "")
                    'oTmpRs("dPDate") = recRef.Fields("生产日期")
                    BodyColData("dMadeDate") = SetNull(recRef.Fields("生产日期"), "")
                    '                    If bFg Then oTmpRs.Update
                    '更新数据

                    '如果自动整批指定入库单据号，则更新临时表数据
                    If bAllBatch Then
                        'by dcb-dev-119752-2004-8-25
                        oRefSelect.UpdateTmpRdrecords moLogin.OldLogin, (-1) * IIf(IsNull(recRef.Fields("出库数量")), 0, recRef.Fields("出库数量")), (-1) * IIf(IsNull(recRef.Fields("出库件数")), 0, recRef.Fields("出库件数")), recRef.Fields("入库系统编号")

                    End If

                    recRef.MoveNext
                Loop
                '                If oTmpRs.RecordCount = 1 Then
                '                    iniSupperGrid oTmpRs, oLstSup, True, , , recRef, True
                '                    oLstSup.ProtectUnload
                '                Else
                '                    sReturnErr = "M"
                '                    oLstSup.Redraw = False
                '                    oLstSup.RowHeight(oLstSup.row) = 0
                '                    iniSupperGrid oTmpRs, oLstSup, True, , True, recRef
                '                    SendKeys "{ESC}"
                '                    oLstSup.Redraw = True
                '                End If
            End If
        End If
    End If
End Function

Private Function SetItemState(ByVal moLogin As USCOMMON.Login, ByVal sVouchType As VouchType, ByVal oRsHead As DOMDocument, ByVal oRsBody As DOMDocument, _
                              ByRef oSupGrid As Object, ByRef sErrMsg As String) As Boolean
    On Error GoTo Error_General_Handler:

    'Call TraceIn(PROC_SIG)
    Dim oCollection As Collection
    Dim domBody As DOMDocument
    Dim domHead As DOMDocument
    Dim lOldcol As Long
    Dim oTmpCO As USERPCO.VoucherCO
    Dim sTempStr As String
    With oSupGrid
        Set oCollection = New Collection
        Set domBody = oRsBody
        Set domHead = oRsHead
        Set oTmpCO = ClsBill 'New USERPCO.voucherCo
        If True Then 'oTmpCO.IniLogin(moLogin.OldLogin, sTempStr) Then
            Call oTmpCO.CheckBodyState(sVouchType, nInsert, .row, "cinvcode", domBody, sTempStr, domHead, oCollection, IIf(val2(BodyColData("fcurqty")) < 0, True, False))
            Dim nIndex As Long
            Dim lCol As Long
            Dim sBodyName As String
            '            .Redraw = False
            For nIndex = 1 To oCollection.Count
                sBodyName = oCollection(nIndex).Name
                '              If LCase(oSupGrid.Name) = "lstfhdbody" And sBodyName = "cinvouchcode" Then
                '                sBodyName = "ccode"
                '              End If
                lCol = ctlReferMakeVouch1.BodyList.GridColIndex(LCase(sBodyName))
                lOldcol = .Col
                '              .col = lCol
                '              If lCol > 0 And .CellBackColor <> &HD7D0CD Then 'RGB(192, 192, 192) Then
                If oCollection(nIndex).Enabled Then
                    If LCase(sBodyName) Like "cfree*" Or LCase(sBodyName) = "cbatch" Then
                        If BodyColData("cwhname") = "" Then
                            ctlReferMakeVouch1.DisableBodyTextMatrix False, LCase(sBodyName), .row 'editable
                        End If
                      Else
                        If ctlReferMakeVouch1.BodyList.GetGridBody().Cell(6, .row, ctlReferMakeVouch1.BodyList.GridColIndex(LCase(sBodyName)), .row, ctlReferMakeVouch1.BodyList.GridColIndex(LCase(sBodyName))) = 0 Then
                            ctlReferMakeVouch1.DisableBodyTextMatrix False, LCase(sBodyName), .row
                        End If
                    End If
                  Else
                    ctlReferMakeVouch1.DisableBodyTextMatrix True, LCase(sBodyName), .row 'RGB(192, 192, 192) readonly
                End If
                '              End If
                '              .col = lOldcol
            Next
            '=======================================================================================================
            '为了支持根据采购订单退货的功能，采购订单跟据本次入库数量的正负来重新设置跟踪性存货的对应“入库单号”的编辑状态值。
            '李亮 2003-03-25
            If LCase(TypeName(oRef)) = "clspuorderbatchrefer" Then
                sBodyName = "cinvouchcode"
                If val2(BodyColData("fcurqty")) >= 0 Then
                    ctlReferMakeVouch1.DisableBodyTextMatrix True, LCase(sBodyName), .row 'RGB(192, 192, 192)
                  Else
                    If oCollection(sBodyName).Enabled Then
                        ctlReferMakeVouch1.DisableBodyTextMatrix False, LCase(sBodyName), .row
                      Else
                        ctlReferMakeVouch1.DisableBodyTextMatrix True, LCase(sBodyName), .row 'RGB(192, 192, 192)
                    End If
                End If
                '为了支持根据采购订单退货的功能，采购订单跟据本次入库数量的正负来重新设置跟踪性存货的对应“批号”的编辑状态值。
                sBodyName = "cbatch"
                '              If val2(BodyColData("fcurqty")) >= 0 Then
                '                  .SetCellBackColor .row, ctlReferMakeVouch1.BodyList.GridColIndex(lcase(sBodyName)), RGB(192, 192, 192)
                '              Else
                If val2(BodyColData("fcurqty")) < 0 Then
                    If oCollection(sBodyName).Enabled Then
                        ctlReferMakeVouch1.DisableBodyTextMatrix False, LCase(sBodyName), .row
                      Else
                        ctlReferMakeVouch1.DisableBodyTextMatrix True, LCase(sBodyName), .row 'RGB(192, 192, 192)
                    End If
                End If
            End If

            ctlReferMakeVouch1.DisableBodyTextMatrix True, "iquantity", .row
            ctlReferMakeVouch1.DisableBodyTextMatrix True, "inum", .row
            ctlReferMakeVouch1.DisableBodyTextMatrix False, "fcurqty", .row

            ctlReferMakeVouch1.DisableBodyTextMatrix (BodyColData("igrouptype") = 0), "fcurnum", .row

            '=======================================================================================================
            '            .Redraw = True
          Else
            sErrMsg = sTempStr
        End If
    End With
    'Call TraceOut(PROC_SIG)
ExitFunction:
    If sErrMsg <> "" Then
        SetItemState = False
      Else
        SetItemState = True
    End If

Exit Function

Error_General_Handler:
    SetItemState = False
    sErrMsg = USCOMMON.ErrObject.GetErrorString
    GoTo ExitFunction
End Function

Private Property Get BodyColData(ByVal colName As String) As String
    BodyColData = ctlReferMakeVouch1.BodyTextMatrix(ctlReferMakeVouch1.BodyList.row, ctlReferMakeVouch1.BodyList.GridColIndex(LCase(colName)))
End Property

Private Property Let BodyColData(ByVal colName As String, ByVal vNewValue As String)
    '    ctlReferMakeVouch1.BodyTextMatrix(ctlReferMakeVouch1.BodyList.row, ctlReferMakeVouch1.BodyList.GridColIndex(LCase(colName))) = vNewValue
    ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(ctlReferMakeVouch1.BodyList.row, ctlReferMakeVouch1.BodyList.GridColIndex(LCase(colName))) = vNewValue
End Property

Private Function val2(ByVal vVal) As Double
    Select Case vVal
      Case Null
        val2 = 0
      Case "", "-"
        val2 = 0
      Case Else
        val2 = CDbl(vVal)
    End Select
End Function

Private Function SetNull(vValue As Variant, vDefault As Variant) As Variant
    If IsNull(vValue) Then
        SetNull = vDefault
      Else
        SetNull = vValue
    End If
End Function

Private Function BodyValueToRow(ByVal keyValue As String) As Long
    Dim i As Integer
    For i = 1 To ctlReferMakeVouch1.BodyList.rows - 1
        If ctlReferMakeVouch1.BodyList.TextMatrix(i, 0) = "Y" Then
            If ctlReferMakeVouch1.BodyList.TextMatrix(i, ctlReferMakeVouch1.BodyList.GridColIndex(oRef.PKeyReturn)) = keyValue Then
                BodyValueToRow = i
                Exit For
            End If
        End If
    Next
End Function

Private Sub DataCheckOut(R As Long, C As Long, newvalue As Variant, OldValue As Variant, KeepFocus As Boolean)
    Dim lQuan As Double
    Dim cOldValue1 As Double
    Dim sErrStr As String
    Dim oRec As Recordset
    'LLIANG-860-Develop-Bug-122172
    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00167") Or ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00138") Then 'zh-CN：存货代码 'zh-CN：规格型号
        newvalue = OldValue
        Exit Sub
    End If
    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00172") Or ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00173") Then 'zh-CN：生产日期 'zh-CN：失效日期
        If newvalue <> "" And Not IsDate(newvalue) Then
            MsgBox LoadResST("U8.ST.USCONTROL.modulectl.00258"), vbApplicationModal + vbCritical + vbOKOnly, LoadResST("U8.ST.USCONTROL.clschackfilter.00034") 'zh-CN：日期格式不正确，必须按格式YYYY-MM-DD输入 'zh-CN：提示
            newvalue = OldValue
            Exit Sub
        End If
    End If
    '    'LLIANG-860-Develop-Bug-63672
    '    If c < oRec.fields.count Then
    '    If oRec(c - 1).Type = adDate Or oRec(c - 1).Type = adDBDate Or oRec(c - 1).Type = adDBTime Or oRec(c - 1).Type = adDBTimeStamp Then
    '        If NewValue <> "" And Not IsDate(NewValue) Then
    '             MsgBox LoadResST("U8.ST.USCONTROL.modulectl.00258"), vbApplicationModal + vbCritical + vbOKOnly, LoadResST("U8.ST.USCONTROL.clschackfilter.00034") 'zh-CN：日期格式不正确，必须按格式YYYY-MM-DD输入 'zh-CN：提示
    '             NewValue = OldValue
    '             Exit Sub
    '         End If
    '    End If
    '    End If
    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00241") Then 'zh-CN：本次出库件数
        If ctlReferMakeVouch1.BodyList.TextMatrix(R, C) = "-" Then
            MsgBox LoadResST("U8.ST.USCONTROL.frmstockorder.00398"), vbApplicationModal + vbCritical + vbOKOnly, LoadResST("U8.ST.USCONTROL.clschackfilter.00034") 'zh-CN：数字输入错误,请修改 'zh-CN：提示
            newvalue = OldValue
            Exit Sub
        End If
        lQuan = val2(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fcurqty"))))
    End If
    '=======================================================
    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00240") Then 'zh-CN：本次出库数量
        If ctlReferMakeVouch1.BodyList.TextMatrix(R, C) = "-" Then
            MsgBox LoadResST("U8.ST.USCONTROL.frmstockorder.00398"), vbApplicationModal + vbCritical + vbOKOnly, LoadResST("U8.ST.USCONTROL.clschackfilter.00034") 'zh-CN：数字输入错误,请修改 'zh-CN：提示
            newvalue = OldValue
            Exit Sub
        End If
        lQuan = val2(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fcurnum"))))
    End If
    '=======================================================
    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00241") Then 'zh-CN：本次出库件数
        'LLIANG-860-Develop-Bug-178327
        If val2(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iNum")))) * val2(newvalue) < 0 Then
            MsgBox LoadResSTWithArg("U8.ST.USCONTROL.frmcons.00710", Array(IIf(val2(newvalue) > 0, LoadResST("U8.ST.USCONTROL.frmstockorder.00451"), LoadResST("U8.ST.USCONTROL.frmstockorder.00452")))), , LoadResST("U8.ST.USCONTROL.frmproduceorder.00594") 'Para zh-CN：本次出库件数不能为{0} 'zh-CN：正数 'zh-CN：负数 'zh-CN：系统提示
            ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fcurqty"))) = lQuan
            newvalue = OldValue
            Exit Sub
        End If
    End If
    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00240") Then 'zh-CN：本次出库数量
        If val2(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iQuantity")))) * val2(newvalue) < 0 Then
            MsgBox LoadResSTWithArg("U8.ST.USCONTROL.frmcons.00715", Array(IIf(val2(newvalue) > 0, LoadResST("U8.ST.USCONTROL.frmstockorder.00451"), LoadResST("U8.ST.USCONTROL.frmstockorder.00452")))), , LoadResST("U8.ST.USCONTROL.frmproduceorder.00594") 'Para zh-CN：本次出库数量不能为{0} 'zh-CN：正数 'zh-CN：负数 'zh-CN：系统提示
            newvalue = OldValue
            Exit Sub
        End If
    End If

    '=================================================================================
    '修改原因:如果修改了[入库单号]，则把旧的[入库单IDs]清空以便重新参与校验和赋值
    '修改范围:库存参照销售发货单生成销售出库单；
    '修改日期:2003-05-09
    '修改人  :李亮
    Dim lOldVouchCode As String
    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00182") And OldValue <> "" Then 'zh-CN：入库单号
        lOldVouchCode = ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iBatch")))
        ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iBatch"))) = ""
    End If

    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00182") Then 'zh-CN：入库单号
        ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("ccode"))) = newvalue
    End If
    '=================================================================================

    Dim oVO As New USERPVO.DispatchList
    Dim oStates As New Collection
    '    oRs.Filter = 0
    oVO.Head = DomToRecordSet(ctlReferMakeVouch1.GetHeadDom(False))
    '    oRec.Filter = oRec(0).Name & "=null"
    oVO.Body = DomToRecordSet(ctlReferMakeVouch1.GetBodyLine(R))
    oVO.Body.Fields("iQuantity") = oVO.Body.Fields("fcurqty")
    oVO.Body.Fields("iNum") = oVO.Body.Fields("fcurnum")
    '    FillRst lstFhdBody, oVO.Body, r
    'oRec.Filter = 0
    Set oRec = oVO.Body
    sErrStr = ""
    If ReBodyItems(moLogin, oVO, oRec, R, C, ctlReferMakeVouch1.BodyList, sErrStr, oStates) Then
        newvalue = ctlReferMakeVouch1.BodyList.TextMatrix(R, C)

        If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00176") Then 'zh-CN：库存单位
            newvalue = vFieldVal(oVO.Body("cinva_unit"))
        End If
        If C < oVO.Body.Fields.Count Then
            If LCase(Left(oVO.Body(C).Name, 5)) = "cfree" Then
                newvalue = vFieldVal(oVO.Body(C))
            End If
        End If
      Else
        MsgBox sErrStr, , LoadResST("U8.ST.USCONTROL.frmproduceorder.00594") 'zh-CN：系统提示
        sErrStr = ""
        newvalue = OldValue
        If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00182") And OldValue <> "" Then 'zh-CN：入库单号
            ctlReferMakeVouch1.BodyList.TextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iBatch"))) = lOldVouchCode
        End If
    End If
    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00171") Then 'zh-CN：批号
        If oStates.Count > 0 Then
            If oStates("cbatch").Refrence Then
                ctlReferMakeVouch1_BodyBrowUser newvalue, R, C
            End If
        End If
    End If
    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00182") Then           '出入库追踪自动参照 'zh-CN：入库单号
        If oStates.Count > 0 Then
            If oStates("cinvouchcode").Refrence Then
                ctlReferMakeVouch1_BodyBrowUser newvalue, R, C
                Exit Sub
            End If
        End If
    End If
    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00240") Or ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00241") Then 'zh-CN：本次出库数量 'zh-CN：本次出库件数
        Dim oTmpCO As USERPCO.VoucherCO
        Set oTmpCO = ClsBill 'New USERPCO.voucherCo
        sErrStr = ""
        If True Then 'oTmpCO.IniLogin(moLogin.OldLogin, sErrStr) Then
            '循环处理本次出库数量
            Dim lTempQuan As Double
            Dim i As Long
            For i = 1 To ctlReferMakeVouch1.BodyList.rows - 1
                If ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("Autoid"))) = ctlReferMakeVouch1.BodyTextMatrix(i, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("Autoid"))) Then
                    '                    If lstFhdBody.RowHeight(i) <> 0 Then
                    'lTempQuan = lTempQuan + val2(ctlReferMakeVouch1.BodyList.TextMatrix(i, lstFhdBody.cols - 1))
                    lTempQuan = lTempQuan + val2(ctlReferMakeVouch1.BodyTextMatrix(i, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fcurqty"))))
                    '                    End If
                End If
            Next
            If oTmpCO.Login.Account.OverDispOut Then
                '                oRec.Filter = "AutoID=" & ctlReferMakeVouch1.BodyList.TextMatrix(r, sAutoId)
                cOldValue1 = val2(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iQuantity"))))
                If val2(cOldValue1) <> 0 Then
                    If (val2(lTempQuan) + val2(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fOutQuantity")))) - cOldValue1) / cOldValue1 > val2(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fOutExcess")))) Then
                        MsgBox LoadResST("U8.ST.USCONTROL.frmcons.00727"), , LoadResST("U8.ST.USCONTROL.frmproduceorder.00594") 'zh-CN：本次出库数量不能大于出库上限范围！ 'zh-CN：系统提示
                        If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00241") Then 'zh-CN：本次出库件数
                            ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fcurqty"))) = lQuan
                        End If
                        '=======================================================
                        If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00240") Then 'zh-CN：本次出库数量
                            ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fcurnum"))) = lQuan
                        End If
                        '=======================================================
                        newvalue = OldValue
                    End If
                End If
              Else
                'If Abs(val2(lTempQuan)) > Abs(val2(ctlReferMakeVouch1.BodyList.TextMatrix(R, C - 3))) Then
                If Abs(val2(lTempQuan)) > Abs(val2(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fnqty"))))) Then
                    MsgBox LoadResST("U8.ST.USCONTROL.frmcons.00731"), , LoadResST("U8.ST.USCONTROL.frmproduceorder.00594") 'zh-CN：本次出库数量不能大于未出库数量 'zh-CN：系统提示
                    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00241") Then 'zh-CN：本次出库件数
                        ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fcurqty"))) = lQuan
                    End If
                    '=======================================================
                    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00240") Then 'zh-CN：本次出库数量
                        ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fcurnum"))) = lQuan
                    End If
                    '=======================================================
                    newvalue = OldValue
                End If
            End If
          Else
            MsgBox sErrStr, vbOKOnly + vbInformation, LoadResST("U8.ST.USCONTROL.frmproduceorder.00594") 'zh-CN：系统提示
            sErrStr = ""
        End If
    End If

    If InStr(1, ctlReferMakeVouch1.BodyList.TextMatrix(0, C), LoadResST("U8.ST.USCONTROL.modulectl.00235")) > 0 Then 'zh-CN：数量
        newvalue = Format(newvalue, "##0." & String(moLogin.Account.QuanDecDgt, "0"))
    End If
    If InStr(1, ctlReferMakeVouch1.BodyList.TextMatrix(0, C), LoadResST("U8.ST.USCONTROL.modulectl.00236")) Then 'zh-CN：件数
        newvalue = Format(newvalue, "##0." & String(moLogin.Account.NumDecDgt, "0"))
    End If
    If InStr(1, ctlReferMakeVouch1.BodyList.TextMatrix(0, C), LoadResST("U8.ST.USCONTROL.modulectl.00175")) Then 'zh-CN：换算率
        newvalue = Format(newvalue, "##0." & String(moLogin.Account.ExchRateDecDgt, "0"))
    End If
End Sub

Private Sub DataCheckDdIn(R As Long, C As Long, newvalue As Variant, OldValue As Variant, KeepFocus As Boolean)
    Dim lQunt As Double
    Dim sBusType As String
    Dim sErrStr As String
    Dim oRec As Recordset
    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00172") Or ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00173") Then 'zh-CN：生产日期 'zh-CN：失效日期
        If newvalue <> "" And Not IsDate(newvalue) Then
            MsgBox LoadResST("U8.ST.USCONTROL.modulectl.00258"), vbApplicationModal + vbCritical + vbOKOnly, LoadResST("U8.ST.USCONTROL.clschackfilter.00034") 'zh-CN：日期格式不正确，必须按格式YYYY-MM-DD输入 'zh-CN：提示
            newvalue = OldValue
            Exit Sub
          Else
            If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00172") Then
                'If ClsAccount.AlarmVDate Then
                If DateDiff("d", moLogin.OldLogin.CurDate, DateAdd(MassChEn(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cMassUnit")))), val2(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("imassdate")))), CDate(newvalue))) < 0 Then
                    'Result:Row=3114        Col=36  Content="该存货已失效，是否继续？"      ID=0c6c580a-7389-44ab-9a12-051d7822dbb6
                    If MsgBox(GetResString("U8.ST.USKCGLSQL.clsbasebillhandler.01435"), vbYesNo + vbQuestion) = vbNo Then
                        newvalue = OldValue
                        Exit Sub
                    End If
                End If
                'End If
              Else
                'If ClsAccount.AlarmVDate Then
                If DateDiff("d", moLogin.OldLogin.CurDate, newvalue) < 0 Then
                    'Result:Row=3132        Col=36  Content="该存货已失效，是否继续？"      ID=207e6692-6486-439c-96a5-f214cac0dab3
                    If MsgBox(GetResString("U8.ST.USKCGLSQL.clsbasebillhandler.01435"), vbYesNo + vbQuestion) = vbNo Then
                        newvalue = OldValue
                        Exit Sub
                    End If
                End If
                'End If
            End If
        End If
    End If
    '    'LLIANG-860-Develop-Bug-60608
    '    If c < oRec.fields.count Then
    '    If oRec(c - 1).Type = adDate Or oRec(c - 1).Type = adDBDate Or oRec(c - 1).Type = adDBTime Or oRec(c - 1).Type = adDBTimeStamp Then
    '        If NewValue <> "" And Not IsDate(NewValue) Then
    '             MsgBox LoadResST("U8.ST.USCONTROL.modulectl.00258"), vbApplicationModal + vbCritical + vbOKOnly, LoadResST("U8.ST.USCONTROL.clschackfilter.00034") 'zh-CN：日期格式不正确，必须按格式YYYY-MM-DD输入 'zh-CN：提示
    '             NewValue = OldValue
    '             Exit Sub
    '         End If
    '    End If
    '    End If

    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00165") Then 'zh-CN：仓库
        If ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cinvouchcode"))) <> "" Then
            MsgBox LoadResST("U8.ST.USCONTROL.frmstockorder.00395"), vbCritical, LoadResST("U8.ST.USCONTROL.frmproduceorder.00594") 'zh-CN：已经指定了对应入库单号,不能修改仓库！ 'zh-CN：系统提示
            newvalue = OldValue
        End If

        If Not bCheckWh(newvalue, GetHeadBusType(R)) Then
            newvalue = ""
            ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cwhname"))) = ""
            ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cwhcode"))) = ""
          Else
            ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cwhcode"))) = cWhCode
            ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cwhname"))) = newvalue
        End If
    End If
    Dim lOldQuan As String
    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00245") Then 'zh-CN：本次入库件数
        If ctlReferMakeVouch1.BodyList.TextMatrix(R, C) = "-" Then
            MsgBox LoadResST("U8.ST.USCONTROL.frmstockorder.00398"), vbApplicationModal + vbCritical + vbOKOnly, LoadResST("U8.ST.USCONTROL.clschackfilter.00034") 'zh-CN：数字输入错误,请修改 'zh-CN：提示
            newvalue = OldValue
            Exit Sub
        End If
        lOldQuan = ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fcurqty")))
    End If
    '==========================================================
    '修改原因:(U851)因为固定换算率存货允许输入[数量]
    '修改范围:库存参照采购订单生成采购入库单；库存参照销售发货单生成销售出库单；
    '修改日期:2003-04-15
    '修改人  :李亮
    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00244") Then 'zh-CN：本次入库数量
        If ctlReferMakeVouch1.BodyList.TextMatrix(R, C) = "-" Then
            MsgBox LoadResST("U8.ST.USCONTROL.frmstockorder.00398"), vbApplicationModal + vbCritical + vbOKOnly, LoadResST("U8.ST.USCONTROL.clschackfilter.00034") 'zh-CN：数字输入错误,请修改 'zh-CN：提示
            newvalue = OldValue
            Exit Sub
        End If
        lOldQuan = ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fcurnum")))
        If val2(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("igrouptype")))) = 2 And (val2(newvalue) * val2(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fcurnum")))) < 0) Then
            MsgBox LoadResSTWithArg("U8.ST.USCONTROL.frmstockorder.00403", Array(IIf(Sgn(val2(newvalue)) < 0, "负", "正"))) 'Para zh-CN：必须先修改[件数]为{0}数然后修改[数量]!
            newvalue = OldValue
            Exit Sub
        End If
    End If
    '==========================================================

    '=================================================================================
    '修改原因:如果修改了[入库单号]，则把旧的[入库单IDs]清空以便重新参与校验和赋值
    '修改范围:库存参照采购订单生成采购入库单；
    '修改日期:2003-05-09
    '修改人  :李亮
    Dim lOldVouchCode As String
    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00182") And OldValue <> "" Then 'zh-CN：入库单号
        lOldVouchCode = ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iCorId")))
        ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iCorId"))) = ""
    End If
    '=================================================================================
    Dim oVO As New USERPVO.PurchaseOrder
    Dim oStates As New Collection
    '    oRs.Filter = 0
    oVO.Head = DomToRecordSet(ctlReferMakeVouch1.GetHeadDom(False))
    '    oRec.Filter = oRec(0).Name & "=null"
    oVO.Body = DomToRecordSet(ctlReferMakeVouch1.GetBodyLine(R))
    oVO.Body.Fields("iQuantity") = oVO.Body.Fields("fcurqty")
    oVO.Body.Fields("iNum") = oVO.Body.Fields("fcurnum")
    '    FillRst lstDdBody, oVO.Body, r
    Set oRec = oVO.Body
    'oRec.Filter = 0
    If ReBodyItems(moLogin, oVO, oRec, R, C, ctlReferMakeVouch1.BodyList, sErrStr, oStates) Then
        newvalue = ctlReferMakeVouch1.BodyList.TextMatrix(R, C)

        If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00176") Then 'zh-CN：库存单位
            newvalue = vFieldVal(oVO.Body("cinva_unit"))
        End If
        If C < oVO.Body.Fields.Count Then
            If LCase(Left(oVO.Body(C).Name, 5)) = "cfree" Then
                newvalue = vFieldVal(oVO.Body(C))
            End If
        End If
      Else
        MsgBox sErrStr
        sErrStr = ""
        newvalue = OldValue
        If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00182") And OldValue <> "" Then 'zh-CN：入库单号
            ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iCorId"))) = lOldVouchCode
        End If
        sErrStr = ""
    End If

    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00182") Then           '出入库追踪自动参照 'zh-CN：入库单号
        If oStates.Count > 0 Then
            If oStates("cinvouchcode").Refrence Then
                ctlReferMakeVouch1_BodyBrowUser newvalue, R, C
                Exit Sub
            End If
        End If
    End If

    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00244") Or ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00245") Then 'zh-CN：本次入库数量 'zh-CN：本次入库件数

        'If val2(OldValue) * val2(NewValue) < 0 Then

        '=============================================================================================================================
        'u851需求: 本次入库数量和本次入库件数中可以输入负数
        '          |本次入库数量（件数）|≤订单的累计入库数量（件数）
        If val2(newvalue) < 0 Then
            '            MsgBox "本次入库数量不能为负数"
            '            NewValue = OldValue
            If (ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00244") And (Abs(newvalue) > val2(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iReceivedQTY")))))) Or _
               (ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00245") And (Abs(newvalue) > val2(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iReceivedNUM")))))) Then 'zh-CN：本次入库件数
                MsgBox LoadResST("U8.ST.USCONTROL.frmstockorder.00415") 'zh-CN：本次退货数量/件数不能大于订单的累计入库数量/件数。
                newvalue = OldValue
                '==========================================================
                If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00245") Then 'zh-CN：本次入库件数
                    ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fcurqty"))) = lOldQuan
                End If
                If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00244") Then 'zh-CN：本次入库数量
                    ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fcurnum"))) = lOldQuan
                End If
            End If
        End If
        '=============================================================================================================================
        Dim oTmpCO As USERPCO.VoucherCO
        Set oTmpCO = ClsBill 'New USERPCO.voucherCo
        If True Then 'oTmpCO.IniLogin(moLogin.OldLogin, sErrStr) Then
            If oTmpCO.Login.Account.OverOrder Then
                '                oRec.Filter = "ID=" & ctlReferMakeVouch1.BodyList.TextMatrix(r, sAutoId)
                '循环处理本次入库数量
                '                For i = 1 To lstDdBody.Rows - 1
                '                    if
                '                    lQunt = lQunt + 1
                '                Next
                If (val2(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iReceivedQTY")))) + val2(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fcurqty")))) - val2(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iQuantity"))))) / val2(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iQuantity")))) > Getfinexcess(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cinvcode")))) Then
                    MsgBox LoadResST("U8.ST.USCONTROL.frmstockorder.00418") 'zh-CN：本次入库数量不能大于入库上限范围！
                    newvalue = OldValue

                    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00245") Then 'zh-CN：本次入库件数
                        ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fcurqty"))) = lOldQuan
                    End If
                    '==========================================================
                    '修改原因:(U851)因为固定换算率存货允许输入[数量]
                    '修改范围:库存参照采购订单生成采购入库单；库存参照销售发货单生成销售出库单；
                    '修改日期:2003-04-15
                    '修改人  :李亮
                    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00244") Then 'zh-CN：本次入库数量
                        ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fcurnum"))) = lOldQuan
                    End If
                    '==========================================================
                End If
              Else

                'by dcb 2004-3-24 修改46688问题 取消ABS
                If FormatToDouble((val2(newvalue)), moLogin.Account.FormatQuanDecString) > FormatToDouble(Abs(val2(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fcurqty"))))), moLogin.Account.FormatQuanDecString) Then
                    MsgBox LoadResST("U8.ST.USCONTROL.frmstockorder.00421") 'zh-CN：本次入库数量不能大于未入库数量
                    newvalue = OldValue

                    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00245") Then 'zh-CN：本次入库件数
                        ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fcurqty"))) = lOldQuan
                    End If
                    '==========================================================
                    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00244") Then 'zh-CN：本次入库数量
                        ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fcurnum"))) = lOldQuan
                    End If
                    '==========================================================
                End If
            End If
          Else
            MsgBox sErrStr, vbOKOnly + vbInformation, LoadResST("U8.ST.USCONTROL.frmproduceorder.00594") 'zh-CN：系统提示
            sErrStr = ""
        End If
    End If
    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00172") And newvalue <> "" Then 'zh-CN：生产日期
        ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("dvdate"))) = DateAdd(MassChEn(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cMassUnit")))), val2(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("imassdate")))), CDate(newvalue))
    End If

    If InStr(1, ctlReferMakeVouch1.BodyList.TextMatrix(0, C), LoadResST("U8.ST.USCONTROL.modulectl.00235")) > 0 Then 'zh-CN：数量
        newvalue = Format(newvalue, "##0." & String(moLogin.Account.QuanDecDgt, "0"))
    End If
    If InStr(1, ctlReferMakeVouch1.BodyList.TextMatrix(0, C), LoadResST("U8.ST.USCONTROL.modulectl.00236")) Then 'zh-CN：件数
        newvalue = Format(newvalue, "##0." & String(moLogin.Account.NumDecDgt, "0"))
    End If
    If InStr(1, ctlReferMakeVouch1.BodyList.TextMatrix(0, C), LoadResST("U8.ST.USCONTROL.modulectl.00175")) Then 'zh-CN：换算率
        newvalue = Format(newvalue, "##0." & String(moLogin.Account.ExchRateDecDgt, "0"))
    End If
End Sub

Private Sub DataCheckDhdIn(R As Long, C As Long, newvalue As Variant, OldValue As Variant, KeepFocus As Boolean)
    Dim sErrStr As String
    Dim oRec As Recordset
    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00172") Or ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00173") Then 'zh-CN：生产日期 'zh-CN：失效日期
        If newvalue <> "" And Not IsDate(newvalue) Then
            MsgBox LoadResST("U8.ST.USCONTROL.modulectl.00258"), vbApplicationModal + vbCritical + vbOKOnly, LoadResST("U8.ST.USCONTROL.clschackfilter.00034") 'zh-CN：日期格式不正确，必须按格式YYYY-MM-DD输入 'zh-CN：提示
            newvalue = OldValue
            Exit Sub
          Else
            If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00172") Then
                'If ClsAccount.AlarmVDate Then
                If DateDiff("d", moLogin.OldLogin.CurDate, DateAdd(MassChEn(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cMassUnit")))), val2(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("imassdate")))), CDate(newvalue))) < 0 Then
                    'Result:Row=3114        Col=36  Content="该存货已失效，是否继续？"      ID=0c6c580a-7389-44ab-9a12-051d7822dbb6
                    If MsgBox(GetResString("U8.ST.USKCGLSQL.clsbasebillhandler.01435"), vbYesNo + vbQuestion) = vbNo Then
                        newvalue = OldValue
                        Exit Sub
                    End If
                End If
                'End If
              Else
                'If ClsAccount.AlarmVDate Then
                If DateDiff("d", moLogin.OldLogin.CurDate, newvalue) < 0 Then
                    'Result:Row=3132        Col=36  Content="该存货已失效，是否继续？"      ID=207e6692-6486-439c-96a5-f214cac0dab3
                    If MsgBox(GetResString("U8.ST.USKCGLSQL.clsbasebillhandler.01435"), vbYesNo + vbQuestion) = vbNo Then
                        newvalue = OldValue
                        Exit Sub
                    End If
                End If
                'End If
            End If
        End If
    End If
    '    'LLIANG-860-Develop-Bug-60608
    '    If c < oRec.fields.count Then
    '    If oRec(c - 1).Type = adDate Or oRec(c - 1).Type = adDBDate Or oRec(c - 1).Type = adDBTime Or oRec(c - 1).Type = adDBTimeStamp Then
    '        If NewValue <> "" And Not IsDate(NewValue) Then
    '             MsgBox LoadResST("U8.ST.USCONTROL.modulectl.00258"), vbApplicationModal + vbCritical + vbOKOnly, LoadResST("U8.ST.USCONTROL.clschackfilter.00034") 'zh-CN：日期格式不正确，必须按格式YYYY-MM-DD输入 'zh-CN：提示
    '             NewValue = OldValue
    '             Exit Sub
    '         End If
    '    End If
    '    End If

    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00165") And ctlReferMakeVouch1.BodyList.TextMatrix(R, C) <> "" Then 'zh-CN：仓库
        If ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cinvouchcode"))) <> "" Then
            MsgBox LoadResST("U8.ST.USCONTROL.frmstockorder.00395"), vbCritical, LoadResST("U8.ST.USCONTROL.frmproduceorder.00594") 'zh-CN：已经指定了对应入库单号,不能修改仓库！ 'zh-CN：系统提示
            newvalue = OldValue
        End If

        If bCheckWh(newvalue, GetHeadBusType(R)) Then
            ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cwhcode"))) = cWhCode
            ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cwhname"))) = newvalue
          Else
            newvalue = ""
            ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cwhcode"))) = ""
            ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cwhname"))) = newvalue
        End If
      ElseIf ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00165") And ctlReferMakeVouch1.BodyList.TextMatrix(R, C) = "" Then 'zh-CN：仓库
        ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cwhcode"))) = ""
    End If
    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00244") Then 'zh-CN：本次入库数量
        If ctlReferMakeVouch1.BodyList.TextMatrix(R, C) = "-" Then
            MsgBox LoadResST("U8.ST.USCONTROL.frmstockorder.00398"), vbApplicationModal + vbCritical + vbOKOnly, LoadResST("U8.ST.USCONTROL.clschackfilter.00034") 'zh-CN：数字输入错误,请修改 'zh-CN：提示
            newvalue = OldValue
            Exit Sub
        End If
        'LLIANG-860-Develop-Bug-178327
        If val2(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iQuantity")))) * val2(newvalue) < 0 Then
            'MsgBox "本次入库数量不能为" & IIf(val2(NewValue) > 0, "正数", "负数"), , LoadResST("U8.ST.USCONTROL.frmproduceorder.00594") 'zh-CN：系统提示
            MsgBox GetResString("U8.ST.Default.00189") & IIf(val2(newvalue) > 0, GetResString("U8.ST.Default.00191"), GetResString("U8.ST.Default.00190")), , LoadResST("U8.ST.USCONTROL.frmproduceorder.00594") 'zh-CN：系统提示
            newvalue = OldValue
            Exit Sub
        End If
        If FormatToDouble(Abs(val2(newvalue)), moLogin.Account.FormatQuanDecString) > FormatToDouble(Abs(val2(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fcurqty"))))), moLogin.Account.FormatQuanDecString) Then
            MsgBox LoadResST("U8.ST.USCONTROL.frmstockorder.00421") 'zh-CN：本次入库数量不能大于未入库数量
            newvalue = OldValue
            Exit Sub
        End If
    End If
    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00245") Then 'zh-CN：本次入库件数
        If ctlReferMakeVouch1.BodyList.TextMatrix(R, C) = "-" Then
            MsgBox LoadResST("U8.ST.USCONTROL.frmstockorder.00398"), vbApplicationModal + vbCritical + vbOKOnly, LoadResST("U8.ST.USCONTROL.clschackfilter.00034") 'zh-CN：数字输入错误,请修改 'zh-CN：提示
            newvalue = OldValue
            Exit Sub
        End If
        'LLIANG-860-Develop-Bug-178327
        If val2(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iNum")))) * val2(newvalue) < 0 Then
            'MsgBox "本次入库件数不能为" & IIf(val2(NewValue) > 0, "正数", "负数"), , LoadResST("U8.ST.USCONTROL.frmproduceorder.00594") 'zh-CN：系统提示
            MsgBox GetResString("U8.ST.Default.00192") & IIf(val2(newvalue) > 0, GetResString("U8.ST.Default.00191"), GetResString("U8.ST.Default.00190")), , LoadResST("U8.ST.USCONTROL.frmproduceorder.00594") 'zh-CN：系统提示
            newvalue = OldValue
            Exit Sub
        End If
        If FormatToDouble(Abs(val2(newvalue)), moLogin.Account.FormatQuanDecString) > FormatToDouble(Abs(val2(ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fcurqty"))))), moLogin.Account.FormatQuanDecString) Then
            'LLIANG-860-Develop-Bug-60565
            MsgBox LoadResST("U8.ST.USCONTROL.frmstockorder.00462") 'zh-CN：本次入库件数不能大于未入库件数
            newvalue = OldValue
            Exit Sub
        End If
    End If
    Dim oVO As New USERPVO.ArriveList
    Dim oStates As New Collection
    '    oRs.Filter = 0
    oVO.Head = DomToRecordSet(ctlReferMakeVouch1.GetHeadDom(False))
    '    oRec.Filter = oRec(0).Name & "=null"
    oVO.Body = DomToRecordSet(ctlReferMakeVouch1.GetBodyLine(R))
    oVO.Body.Fields("iQuantity") = oVO.Body.Fields("fcurqty")
    oVO.Body.Fields("iNum") = oVO.Body.Fields("fcurnum")
    '    FillRst lstDhdBody, oVO.Body, r
    Set oRec = oVO.Body
    'oRec.Filter = 0
    If ReBodyItems(moLogin, oVO, oRec, R, C, ctlReferMakeVouch1.BodyList, sErrStr, oStates) Then
        newvalue = ctlReferMakeVouch1.BodyList.TextMatrix(R, C)

        If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00176") Then 'zh-CN：库存单位
            newvalue = vFieldVal(oVO.Body("cinva_unit"))
        End If
      Else
        MsgBox sErrStr
        sErrStr = ""
        newvalue = OldValue
    End If
    'lliang add 2002-12-14 根据到货退回单生单，存货为出库跟踪入库，在生单界面选择入库单号时，如果该入库单上有两条记录，则没有进行定位
    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00171") Then 'zh-CN：批号
        If oStates.Count > 0 Then
            If oStates("cbatch").Refrence Then
                ctlReferMakeVouch1_BodyBrowUser newvalue, R, C
            End If
        End If
    End If
    If ctlReferMakeVouch1.BodyList.TextMatrix(0, C) = LoadResST("U8.ST.USCONTROL.modulectl.00182") Then           '出入库追踪自动参照 'zh-CN：入库单号
        If oStates.Count > 0 Then
            If oStates("cinvouchcode").Refrence Then
                ctlReferMakeVouch1_BodyBrowUser newvalue, R, C
            End If
        End If
    End If

    If InStr(1, ctlReferMakeVouch1.BodyList.TextMatrix(0, C), LoadResST("U8.ST.USCONTROL.modulectl.00235")) > 0 Then 'zh-CN：数量
        newvalue = Format(newvalue, "##0." & String(moLogin.Account.QuanDecDgt, "0"))
    End If
    If InStr(1, ctlReferMakeVouch1.BodyList.TextMatrix(0, C), LoadResST("U8.ST.USCONTROL.modulectl.00236")) Then 'zh-CN：件数
        newvalue = Format(newvalue, "##0." & String(moLogin.Account.NumDecDgt, "0"))
    End If
    If InStr(1, ctlReferMakeVouch1.BodyList.TextMatrix(0, C), LoadResST("U8.ST.USCONTROL.modulectl.00175")) Then 'zh-CN：换算率
        newvalue = Format(newvalue, "##0." & String(moLogin.Account.ExchRateDecDgt, "0"))
    End If
End Sub

Private Function ReBodyItems(ByVal moLogin As USCOMMON.Login, ByVal oVO As Object, ByVal oOldBody As Recordset, _
                             ByVal R As Long, ByVal C As Long, ByRef oSupGrid As Object, ByRef sErrMsg As String, _
                             Optional ByRef oStates As Collection) As Boolean
    On Error GoTo Error_General_Handler:

    'Call TraceIn(PROC_SIG)
    Dim i As Long
    Dim oTmpBO As Object
    Dim sTempStr As String
    With oSupGrid
        Set oTmpBO = CreateObject("USERPBO.VoucherBO")
        oTmpBO.Login = moLogin
        '        If c <= oVO.Body.fields.count Then
        '            sTempStr = oVO.Body(c - 1).Name
        '        ElseIf c = .cols - 1 Then
        '            sTempStr = "inum"
        '        ElseIf c = .cols - 2 Then
        '            sTempStr = "iquantity"
        '        End If
        sTempStr = LCase(ctlReferMakeVouch1.GetBodyColName(C))
        If sTempStr = "fcurqty" Then sTempStr = "iquantity"
        If sTempStr = "fcurnum" Then sTempStr = "inum"
        If oTmpBO.CheckBody(nUpdate, oVO, sTempStr, oStates, oOldBody) Then
            Dim stItemName  As String
            Dim stItemValue As Variant
            Dim nRow As Long
            nRow = R
            For i = 0 To oVO.Body.Fields.Count - 1
                stItemName = LCase(oVO.Body(i).Name)
                stItemValue = oVO.Body(i).Value
                Select Case LCase(stItemName)
                  Case "iquantity"
                    ctlReferMakeVouch1.BodyTextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fcurqty"))) = Format(stItemValue, moLogin.Account.FormatQuanDecString)
                  Case "inum"
                    ctlReferMakeVouch1.BodyTextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fcurnum"))) = Format(stItemValue, moLogin.Account.FormatNumDecString)
                  Case "iinvexchrate"
                    ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iinvexchrate"))) = Format(stItemValue, moLogin.Account.FormatExchDecString)
                    '固定换算率存货，参照采购销售单据窗体，修改库存单位后应计算件数，（应入库、未入库、已入库），现在未算。

                  Case "cwhcode", "cwhname", "selcol"
                    '仓库就不赋值了!
                    '                    Case "cinva_unit"
                    '                        .TextMatrix(nRow, LookUpArray("cinva_unit", oVO.body, .cols)) = vFieldVal(oVO.body("cinva_unit"))
                  Case Else
                    If IsNull(stItemValue) Then
                        ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase(stItemName))) = ""
                      Else
                        If stItemValue <> ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase(stItemName))) And LCase(stItemName) <> "fcurqty" And LCase(stItemName) <> "fcurnum" Then
                            ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase(stItemName))) = stItemValue
                        End If
                    End If
                End Select
            Next

            'lliang add 2002-12-17 固定换算率存货，修改[库存单位]后[换算率]变化；应该反算各件数
            Dim iInvExchRate As Double
            iInvExchRate = val2(ctlReferMakeVouch1.BodyTextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iinvexchrate"))))
            'lliang-860-Develop-Bug-57681 增加:判断是否为固定换算率存货
            If iInvExchRate <> 0 And val2(ctlReferMakeVouch1.BodyTextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("igrouptype")))) = 1 Then
                ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fnnum"))) = Format(val2(ctlReferMakeVouch1.BodyTextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fnqty")))) / iInvExchRate, moLogin.Account.FormatNumDecString)
                ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("foutnum"))) = Format(val2(ctlReferMakeVouch1.BodyTextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("foutquantity")))) / iInvExchRate, moLogin.Account.FormatNumDecString)
                ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("inum"))) = Format(val2(ctlReferMakeVouch1.BodyTextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iquantity")))) / iInvExchRate, moLogin.Account.FormatNumDecString)

                ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fcurnum"))) = Format(val2(ctlReferMakeVouch1.BodyTextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fcurqty")))) / iInvExchRate, moLogin.Account.FormatNumDecString)
            End If
            ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fnnum"))) = Format(val2(ctlReferMakeVouch1.BodyTextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fnnum")))), moLogin.Account.FormatNumDecString)
            ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fnqty"))) = Format(val2(ctlReferMakeVouch1.BodyTextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("fnqty")))), moLogin.Account.FormatQuanDecString)
            ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("foutnum"))) = Format(val2(ctlReferMakeVouch1.BodyTextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("foutnum")))), moLogin.Account.FormatNumDecString)
            ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("foutquantity"))) = Format(val2(ctlReferMakeVouch1.BodyTextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("foutquantity")))), moLogin.Account.FormatQuanDecString)
            ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("inum"))) = Format(val2(ctlReferMakeVouch1.BodyTextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("inum")))), moLogin.Account.FormatNumDecString)
            ctlReferMakeVouch1.BodyList.GetGridBody().TextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iquantity"))) = Format(val2(ctlReferMakeVouch1.BodyTextMatrix(nRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("iquantity")))), moLogin.Account.FormatQuanDecString)

            '------------------------------------------------------------------------------
          Else
            sErrMsg = USCOMMON.ErrObject.GetErrorString
        End If
    End With
ExitFunction:
    If sErrMsg <> "" Then
        ReBodyItems = False
      Else
        ReBodyItems = True
    End If

Exit Function

Error_General_Handler:
    ReBodyItems = False
    sErrMsg = USCOMMON.ErrObject.GetErrorString
    GoTo ExitFunction
End Function

Private Function GetHeadRow(ByVal R As Long, Optional ByVal Value As String = "", Optional ByVal isHead As Boolean = False) As Long
    Dim i As Long
    Dim sKey As String
    GetHeadRow = 0

    If refType = "clspuorderbatchrefer" Then  '订单
        If Not isHead Then
            sKey = ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("poid")))
          Else
            sKey = Value
        End If

        For i = 1 To ctlReferMakeVouch1.HeadList.rows - 1
            If sKey = ctlReferMakeVouch1.HeadTextMatrix(i, ctlReferMakeVouch1.HeadList.GridColIndex(LCase("poid"))) Then
                GetHeadRow = i
                Exit Function
            End If
        Next
      ElseIf refType = "clspuarrivebatchrefer" Then '到货单
        If Not isHead Then
            sKey = ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("id")))
          Else
            sKey = Value
        End If

        For i = 1 To ctlReferMakeVouch1.HeadList.rows - 1
            If sKey = ctlReferMakeVouch1.HeadTextMatrix(i, ctlReferMakeVouch1.HeadList.GridColIndex(LCase("id"))) Then
                GetHeadRow = i
                Exit Function
            End If
        Next
      ElseIf refType = "clssadispatchbatchrefer" Then
        If Not isHead Then
            sKey = ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("dlid")))
          Else
            sKey = Value
        End If

        For i = 1 To ctlReferMakeVouch1.HeadList.rows - 1
            If sKey = ctlReferMakeVouch1.HeadTextMatrix(i, ctlReferMakeVouch1.HeadList.GridColIndex(LCase("dlid"))) Then
                GetHeadRow = i
                Exit Function
            End If
        Next

      Else

        If oRef.ColHJoinB <> "" Then
            sKey = ctlReferMakeVouch1.BodyTextMatrix(R, ctlReferMakeVouch1.BodyList.GridColIndex(LCase(Split(oRef.ColHJoinB, "|")(1))))

            For i = 1 To ctlReferMakeVouch1.HeadList.rows - 1
                If sKey = ctlReferMakeVouch1.HeadTextMatrix(i, ctlReferMakeVouch1.HeadList.GridColIndex(LCase(Split(oRef.ColHJoinB, "|")(0)))) Then
                    GetHeadRow = i
                    Exit Function
                End If
            Next
        End If
    End If
End Function
Private Function GetHeadBusType(ByVal R As Long) As String
    Dim i As Long
    GetHeadBusType = ""
    i = GetHeadRow(R)
    GetHeadBusType = ctlReferMakeVouch1.HeadTextMatrix(i, ctlReferMakeVouch1.HeadList.GridColIndex(LCase("cbustype")))
End Function

Private Function bCheckWh(ByRef sWh As Variant, Optional ByVal sBusType As String = "") As Boolean
    Dim sInvCode As String
    Dim sErrStr As String
    Dim oWh As USERPVO.WareHouse
    Dim oWHPst As WarehousePst
    If sWh <> "" Then
        Set oWHPst = New WarehousePst
        Set oWh = New USERPVO.WareHouse
        oWHPst.Login = moLogin
        bCheckWh = True
        If oWHPst.Load(sWh, oWh, True) Then
            '检查仓库是否合法
            If sBusType <> "" Then
                If (sBusType = "代管采购" And Not oWh.bVMI) Or (sBusType <> "代管采购" And oWh.bVMI) Then
                    MsgBox LoadResST("U8.ST.V870.00162"), vbExclamation, LoadResST("U8.ST.USCONTROL.frmproduceorder.00594") 'zh-CN：仓库不合法！ 'zh-CN：系统提示
                    '                   txtWh.SetFocus
                    refWh.SetFocus
                    sWh = ""
                    cWhCode = ""
                    bCheckWh = False
                    Exit Function
                End If
            End If

            sWh = oWh.Name
            cWhCode = oWh.Id
            bCheckWh = True
          Else
            MsgBox LoadResST("U8.ST.USCONTROL.frmstockorder.00473"), vbExclamation, LoadResST("U8.ST.USCONTROL.frmproduceorder.00594") 'zh-CN：仓库不合法！ 'zh-CN：系统提示
            '            txtWh.SetFocus
            refWh.SetFocus
            USCOMMON.ErrObject.GetErrorString
            sWh = ""
            cWhCode = ""
            bCheckWh = False
            Exit Function
        End If
        Dim oTmpCO As USERPCO.VoucherCO
        Set oTmpCO = ClsBill 'New USERPCO.voucherCo
        If True Then 'oTmpCO.IniLogin(moLogin.OldLogin, sErrStr) Then
            If oTmpCO.Login.Account.CheckWareInv Then
                sInvCode = ctlReferMakeVouch1.BodyTextMatrix(m_bodyCurRow, ctlReferMakeVouch1.BodyList.GridColIndex(LCase("cinvcode")))
                If oWHPst.Load2(oWh.Id, oWh, sInvCode) Then
                    sWh = oWh.Name
                    bCheckWh = True
                  Else
                    If MsgBox(LoadResSTWithArg("U8.ST.USCONTROL.frmstockorder.00475", Array(sWh)), vbYesNo + vbQuestion + vbDefaultButton1, LoadResST("U8.ST.USCONTROL.frmproduceorder.00594")) = vbNo Then
                        USCOMMON.ErrObject.GetErrorString
                        sWh = ""
                        cWhCode = ""
                        bCheckWh = False
                    End If
                End If
            End If
        End If
    End If
End Function



Private Sub RemoveBlankLine(VouchList As Object)
    On Error Resume Next

    Dim i As Long
    Dim j As Long

    Dim strTemp As String
    Dim bHasBlankRow As Boolean

    For i = 1 To VouchList.rows - 1
        bHasBlankRow = True

        For j = 0 To VouchList.Cols - 1
            strTemp = VouchList.TextMatrix(i, j)
            If Len(Trim(strTemp)) <> 0 Then
                bHasBlankRow = False
                Exit For
            End If
        Next j

        If bHasBlankRow Then
            VouchList.RemoveItem i
            i = i - 1
        End If
    Next i
End Sub

Public Sub InitVTID(ocbVTID As Object)
    Dim i As Long
    '    For i = 0 To cbVTID.ListCount - 1
    '        CmbVTID.AddItem cbVTID.list(i), i
    '        CmbVTID.ItemData(i) = cbVTID.ItemData(i)
    '        CmbVTID.ListIndex = 0
    '    Next

    i = 0
    '    Do While Not cbVTID.EOF
    '        CmbVTID.AddItem cbVTID(1), i
    '        CmbVTID.ItemData(i) = cbVTID(0)
    '        cbVTID.MoveNext
    '        i = i + 1
    '        CmbVTID.ListIndex = 0
    '    Loop

    Set cbVTID = ocbVTID
End Sub

Private Function Getfinexcess(ByVal sInvCode As String) As Double

    Getfinexcess = 0

    If Trim(sInvCode) = "" Then Exit Function

    On Error GoTo ErrHandler
    Dim oRs As New ADODB.Recordset
    oRs.CursorLocation = adUseClient

    oRs.Open "select isnull(finexcess,0)  as finexcess from  inventory where cinvcode='" & sInvCode & "'", moLogin.AccountConnection, adOpenStatic, adLockReadOnly
    If oRs.EOF Or oRs.BOF Then
        Getfinexcess = 0
      Else
        Getfinexcess = val2(oRs.Fields("finexcess"))
    End If
    oRs.Close
    Set oRs = Nothing

Exit Function

ErrHandler:

Exit Function

End Function





'自由项名称转化为对照码
Public Function GetFreeCode(sFreeName As String) As String
    Dim ssql As String
    Dim Rs As New ADODB.Recordset
    
    Rs.CursorLocation = adUseClient
    ssql = "select calias from userdefine where cvalue = '" & sFreeName & "'"
    Rs.Open ssql, moLogin.AccountConnection, adOpenDynamic, adLockOptimistic
    If Not Rs.EOF And Not Rs.BOF Then
       GetFreeCode = IIf(IsNull(Rs.Fields("calias")), "", Rs.Fields("calias"))
    End If
    
    Rs.Close
    Set Rs = Nothing
End Function




'按批号规则生成批号
Public Function GetBatchNO(cInvCode As String, DomValue As DOMDocument, ByRef sErr As String, Optional ByVal bSave As Boolean = True) As String
    Dim objGetbatch As Object
    Set objGetbatch = CreateObject("lotnumberserver.clslotnumberserver")
    objGetbatch.Init moLogin.OldLogin
    GetBatchNO = objGetbatch.getBatchNumber(cInvCode, DomValue, sErr, bSave)
End Function






Public Sub FormatDom(SourceDom As DOMDocument, DistDom As DOMDocument, Optional editprop As String = "")
Dim element As IXMLDOMElement
Dim ele_head As IXMLDOMElement
Dim ele_body As IXMLDOMElement
Dim nd  As IXMLDOMNode
Dim tempnd As IXMLDOMNode
Dim ndheadlist As IXMLDOMNodeList
Dim ndbodylist As IXMLDOMNodeList
 
DistDom.loadXML SourceDom.xml
Dim Filedname As String
'格式部分
 Set ndheadlist = SourceDom.selectNodes("//s:Schema/s:ElementType/s:AttributeType")
 
 '数据部分
 
 
 Set ndbodylist = DistDom.selectNodes("//rs:data/z:row")
 
 For Each ele_body In ndbodylist
    For Each ele_head In ndheadlist
        Filedname = ele_head.Attributes.getNamedItem("name").nodeValue
        If ele_body.Attributes.getNamedItem(Filedname) Is Nothing Then
            '若没有当前元素，就增加当前元素
                ele_body.setAttribute Filedname, ""
 
        End If
            
            
            Select Case ele_head.lastChild.Attributes.getNamedItem("dt:type").nodeValue
            
            Case "number", "float", "boolean"
                If UCase(ele_body.Attributes.getNamedItem(Filedname).nodeValue) = UCase("false") Then
                    ele_body.setAttribute Filedname, 0
                End If
            Case Else
            
                If UCase(ele_body.Attributes.getNamedItem(Filedname).nodeValue) = UCase("否") Then
                    ele_body.setAttribute Filedname, 0
                End If
 
            End Select
       
        
        
'         Debug.Print Filedname & "=" & ele_head.selectSingleNode("//s:datatype").Attributes.getNamedItem("dt:type").nodeValue
        
        
        
        
        
    Next
    If editprop <> "" Then
        ele_body.setAttribute "editprop", editprop
    End If
Next
End Sub




Function Init(m_login As U8Login.clsLogin, DBcon As ADODB.Connection, Frm As frmRefernew, Dest_Domh As DOMDocument, Dest_Domb As DOMDocument, Source_Cardnumber As String, Dest_Cardnumber As String) As Boolean
    Dim domControl As New DOMDocument
    Dim DomRefSet  As New DOMDocument
    Dim rds As New ADODB.Recordset
    Dim keys_table As String
    Dim strSql As String
    Dim strErr As String
    Dim tmpdom As New DOMDocument
    
    Set mLogin = m_login
    Set DBconn = DBcon
    ClsBill.IniLogin mLogin, strErr
    Set moLogin = ClsBill.Login
    Set DOM_sa_refervoucherconfig = New DOMDocument
    Set DOM_SA_ReferFillConfig = New DOMDocument
    
    Source_Cardnum = Source_Cardnumber
    Dest_Cardnum = Dest_Cardnumber
    
    
    If rds.State <> 0 Then rds.Close
    strSql = "select * from sa_refervoucherconfig where cardnum='" & Dest_Cardnumber & "' and referkey='" & Source_Cardnumber & "' "
    rds.Open strSql, DBconn.ConnectionString, 3, 4
    rds.Save DOM_sa_refervoucherconfig, adPersistXML
    
    If rds.State <> 0 Then rds.Close
    strSql = "select * from SA_ReferFillConfig where cardnumber='" & Dest_Cardnumber & "' and refername='" & Source_Cardnumber & "' "
    rds.Open strSql, DBconn.ConnectionString, 3, 4
    rds.Save DOM_SA_ReferFillConfig, adPersistXML
     
        Frm.Defwhere_head = GetHeadItemValue(DOM_sa_refervoucherconfig, "defaultfilter") '  IIf(IsNull(rds.Fields("defaultfilter").Value), "", rds.Fields("defaultfilter").Value) '得到默认过滤条件
'        .Cardnumber_source = rds.Fields("cothertype").Value
'        .VIEWname = GetOtherSystemSql(.Cardnumber_source, False)  'GSP默认视图
'        If Trim(Frm.Defwhere_head) <> "" Then Frm.VIEWname = .VIEWname & "  and  " & .Defwhere_head
        
        
        Frm.Head_Columnkey = GetHeadItemValue(DOM_sa_refervoucherconfig, "maincolumnkey")   '发货单参照表头栏目     IIf(IsNull(rds.Fields("maincolumnkey").Value), "", rds.Fields("maincolumnkey").Value) '
        Frm.Body_Columnkey = GetHeadItemValue(DOM_sa_refervoucherconfig, "detailcolumnkey") ' 发货单参照表体栏目    IIf(IsNull(rds.Fields("detailcolumnkey").Value), "", rds.Fields("detailcolumnkey").Value)  '
        Frm.head_key = GetHeadItemValue(DOM_sa_refervoucherconfig, "mainuniquekey")         '表头关键字             IIf(IsNull(rds.Fields("mainuniquekey").Value), "", rds.Fields("mainuniquekey").Value)       '
        Frm.body_key = GetHeadItemValue(DOM_sa_refervoucherconfig, "detailuniquekey")       ' 表体关键字            IIf(IsNull(rds.Fields("detailuniquekey").Value), "", rds.Fields("detailuniquekey").Value)   '
        Frm.inner_key = GetHeadItemValue(DOM_sa_refervoucherconfig, "mainuniquekey")        ' 关联关键字            IIf(IsNull(rds.Fields("mainuniquekey").Value), "", rds.Fields("mainuniquekey").Value)      '
        Frm.Head_Source = GetHeadItemValue(DOM_sa_refervoucherconfig, "maindatasource")     ' 参照表头视图          IIf(IsNull(rds.Fields("maindatasource").Value), "", rds.Fields("maindatasource").Value) '
        Frm.Body_Source = GetHeadItemValue(DOM_sa_refervoucherconfig, "detaildatasource")   ' 参照表体视图          IIf(IsNull(rds.Fields("detaildatasource").Value), "", rds.Fields("detaildatasource").Value)  '
        Frm.sFilterID_head = GetHeadItemValue(DOM_sa_refervoucherconfig, "filtername")      ' 表头过滤界面id        IIf(IsNull(rds.Fields("filtername").Value), "", rds.Fields("filtername").Value)  '
        
        Frm.sFilterID_body = "" '表体过滤界面id
'        keys_table = " select distinct " & .head_key & "  from (" & .VIEWname & ") TT "
'        .Head_Source = "( select * from " & .Head_Source & " where " & .head_key & " in( " & keys_table & " )) YY "
'        .cVouchType = Source_Cardnumber




    
    With Frm
    
        FormatDom Dest_Domh, tmpdom, "A"
        Set Dest_Domh = tmpdom.cloneNode(True)
        
        FormatDom Dest_Domb, tmpdom, "A"
        Set Dest_Domb = tmpdom.cloneNode(True)
        
        Set .DomHead_Dest = Dest_Domh
        Set .DomBody_Dest = Dest_Domb
        
        domControl.appendChild domControl.createElement("root")
        Frm.ctlReferMakeVouch1.Init False, Nothing, domControl, Frm.ctlReferMakeVouch1.PageSize, Frm.ctlReferMakeVouch1.PageCount, 1, , mLogin
        Frm.ctlReferMakeVouch1.SetKey Frm.Head_Columnkey, Frm.Body_Columnkey
        If Trim(.Body_Columnkey) = "" Then
            '设置只有表头部分的参照 ahzzd
            .ctlReferMakeVouch1.SetIsList
        End If
        DomRefSet.loadXML "<Sets><Set s_headkeyname='" & .head_key & "' s_bbodykeyname='" & .body_key & "' s_joinkeyname='" & .inner_key & "' bappendbody='true' bheaddisableshiftsel='0' bbodydisableshiftsel='0' /></Sets>"
        Frm.ctlReferMakeVouch1.DomRefSet = DomRefSet.cloneNode(True)
  End With
End Function





Public Function GetHeadItemValue(ByVal domHead As DOMDocument, ByVal sKey As String) As String
    sKey = LCase(sKey)
    If Not domHead.selectSingleNode("//z:row").Attributes.getNamedItem(sKey) Is Nothing Then
        GetHeadItemValue = domHead.selectSingleNode("//z:row").Attributes.getNamedItem(sKey).nodeValue
    Else
        GetHeadItemValue = ""
    End If
End Function



Public Function GetBodyItemValue(ByVal domBody As DOMDocument, ByVal sKey As String, ByVal R As Long) As String
    sKey = LCase(sKey)
    If Not domBody.selectNodes("//z:row").Item(R).Attributes.getNamedItem(sKey) Is Nothing Then
        GetBodyItemValue = domBody.selectNodes("//z:row").Item(R).Attributes.getNamedItem(sKey).nodeValue
    Else
        GetBodyItemValue = ""
    End If
End Function


