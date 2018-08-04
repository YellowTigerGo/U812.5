VERSION 5.00
Object = "{005151D4-33F6-471B-B651-222D4E411832}#4.4#0"; "UFFormPartner.ocx"
Object = "{AF8BBBB7-94C6-4772-B826-624478C37D6A}#1.5#0"; "UFKEYHOOK.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.OCX"
Object = "{8C7C777D-4D83-4DE8-947E-098E2343A400}#1.0#0"; "CommandButton.ocx"
Object = "{A98B9C82-88D8-4F94-91BB-F2289111C59C}#1.0#0"; "UFCheckBox.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStockMsg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "系统信息"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9690
   Icon            =   "frmStockMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleMode       =   0  'User
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin UFFormPartner.UFFrmCaption UFFrmCaptionMgr 
      Left            =   1080
      Top             =   5520
      _ExtentX        =   2778
      _ExtentY        =   450
      Caption         =   "系统信息"
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
   Begin TabDlg.SSTab stbMsg 
      Height          =   5332
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   9393
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "零出库控制"
      TabPicture(0)   =   "frmStockMsg.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "flgZero"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkZeroOut"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "最高最低库存控制"
      TabPicture(1)   =   "frmStockMsg.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "flgTopLow"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "存货可用量"
      TabPicture(2)   =   "frmStockMsg.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "flgCheck"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "项目超预算控制"
      TabPicture(3)   =   "frmStockMsg.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "flgItemCheck"
      Tab(3).ControlCount=   1
      Begin UFCHECKBOXLib.UFCheckBox chkZeroOut 
         Height          =   225
         Left            =   7740
         TabIndex        =   7
         Top             =   30
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   397
         _StockProps     =   15
         Caption         =   "是否按可用量出库"
         ForeColor       =   0
         ForeColor       =   0
         BorderStyle     =   0
         ReadyState      =   0
         Picture         =   "frmStockMsg.frx":007C
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flgItemCheck 
         Height          =   4815
         Left            =   -74940
         TabIndex        =   6
         Top             =   450
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   8493
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flgTopLow 
         Height          =   4860
         Left            =   -74940
         TabIndex        =   4
         Top             =   345
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   8573
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flgZero 
         Height          =   4845
         Left            =   45
         TabIndex        =   3
         Top             =   360
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   8546
         _Version        =   393216
         AllowUserResizing=   1
         BandDisplay     =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flgCheck 
         Height          =   4845
         Left            =   -74880
         TabIndex        =   5
         Top             =   360
         Width           =   9510
         _ExtentX        =   16775
         _ExtentY        =   8546
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin UFCOMMANDBUTTONLib.UFCommandButton cmdNo 
      Height          =   354
      Left            =   8486
      TabIndex        =   1
      Top             =   5514
      Width           =   1068
      _Version        =   65536
      _ExtentX        =   1884
      _ExtentY        =   624
      _StockProps     =   41
      Caption         =   "取消"
      UToolTipText    =   ""
      Cursor          =   828
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
   Begin UFCOMMANDBUTTONLib.UFCommandButton cmdOk 
      Height          =   354
      Left            =   7283
      TabIndex        =   0
      Top             =   5514
      Width           =   1068
      _Version        =   65536
      _ExtentX        =   1884
      _ExtentY        =   624
      _StockProps     =   41
      Caption         =   "确定"
      UToolTipText    =   ""
      Cursor          =   0
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
   Begin UFKeyHook.UFKeyHookCtrl UFKeyHookCtrlPre 
      Left            =   0
      Top             =   630
      _ExtentX        =   1905
      _ExtentY        =   529
   End
End
Attribute VB_Name = "frmStockMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private moDomMsg As DOMDocument
Private moSelect As VbMsgBoxResult
Private moZeroOut As Boolean
Private msVouchType As String

Private Sub chkZeroOut_Click()
    moZeroOut = chkZeroOut.Value

    'U860材料出库单保存或限额领料单、配比出库单生单时出现可用量控制的提示框后，可由用户选择"是否按可用量出库"，选择是则将可用量不足存货的出库数量（件数）自动更新为可用量并保存单据（可用量指保存单据时可用量控制界面的可用量），可用量为0时自动删除对应的存货记录；选择否处理与851相同。
    Dim oZeroNode As IXMLDOMElement
    Set oZeroNode = moDomMsg.selectSingleNode("//zeroout")
    If LCase(Trim(oZeroNode.getAttribute("okonly"))) = 1 And chkZeroOut.Value Then
        moSelect = vbYes
        Unload Me
    End If
End Sub

Private Sub cmdNo_Click()
    moSelect = vbNo
    Unload Me
End Sub

Private Sub cmdOK_Click()
'
    moSelect = vbYes
    Unload Me
End Sub

Private Sub Form_Load()
'U861
    Call LayOutSvr.InitFormLayout(Me, "U8.SCM.ST.KCGLSQL." & Me.Name)
    SetUniformStyle flgCheck
    SetUniformStyle flgItemCheck
    SetUniformStyle flgTopLow
    SetUniformStyle flgZero
    InitButton
    InitCaption
    InitGrid
    moSelect = vbNo
    moZeroOut = False
End Sub

Public Property Get Message() As DOMDocument
    Set Message = moDomMsg
End Property

Public Property Let Message(ByVal vNewValue As DOMDocument)
    Set moDomMsg = vNewValue
End Property

Private Function InitButton()
'added by wjmin
    Dim iZeroOkonly As Variant
    Dim oZeroNode As IXMLDOMElement
    Set oZeroNode = moDomMsg.selectSingleNode("//zeroout")
    iZeroOkonly = oZeroNode.getAttribute("okonly")
    If IsNull(iZeroOkonly) Then iZeroOkonly = 0

    Dim oItemCheckNode As IXMLDOMElement
    Dim iItemCheckOkonly As Variant
    Set oItemCheckNode = moDomMsg.selectSingleNode("//itemcheck")
    If oItemCheckNode Is Nothing Then
        iItemCheckOkonly = 0
    Else
        iItemCheckOkonly = oItemCheckNode.getAttribute("okonly")
    End If

    If IsNull(iItemCheckOkonly) Then iItemCheckOkonly = 0

    If iZeroOkonly = 1 Or iItemCheckOkonly = 1 Then
        cmdOk.Visible = False
        'Result:Row=215 Col=22  Content="确认"  ID=f176d04c-52f9-45ee-a3e0-b01d02ce4955
        cmdNo.Caption = GetResString("U8.ST.USKCGLSQL.frmstockmsg.00834")
        'Result:Row=217 Col=29  Content="存货可用量"    ID=a4f793a0-2165-4528-a73e-9ec730b0287d
        ' stbMsg.TabCaption(2) = GetResString("U8.ST.USKCGLSQL.frmstockmsg.00835")
        '--------------------------------
        'chkZeroOut.visible = False
        chkZeroOut.Visible = True
    Else
        'Result:Row=223 Col=29  Content="以下存货可用量不足"    ID=5154dd91-bbca-4a46-9a29-d2f4bd945c45
        ' stbMsg.TabCaption(2) = GetResString("U8.ST.USKCGLSQL.frmstockmsg.00836")
    End If


    If Not moDomMsg.selectSingleNode("//itemcheck/w:row") Is Nothing Then
        If moDomMsg.selectSingleNode("//z:row") Is Nothing Then
            cmdOk.Visible = False
            'Result:Row=231 Col=23  Content="确认"  ID=4d6fd6bd-0dca-4b85-b55b-b0d9fe9cc7cf
            cmdNo.Caption = GetResString("U8.ST.USKCGLSQL.frmstockmsg.00834")
        End If
        'Result:Row=234 Col=29  Content="项目超预算警告"        ID=cc2bd084-fbfc-4593-a0c1-57328b8736b5
        ' stbMsg.TabCaption(3) = GetResString("U8.ST.USKCGLSQL.frmstockmsg.00838")
    Else
        'Result:Row=237 Col=29  Content="项目超预算控制"        ID=131bd823-3a03-4090-82d3-73d2578ed4e2
        'stbMsg.TabCaption(3) = GetResString("U8.ST.USKCGLSQL.frmstockmsg.00839")
    End If

    'U860 材料出库单保存或限额领料单、配比出库单生单时出现可用量控制的提示框后，可由用户选择"是否按可用量出库"
    If Not (Val(vouchtype) = MaterialOut Or Val(vouchtype) = VouchRef Or Val(vouchtype) = MaterialFd) Then
        chkZeroOut.Visible = False
    End If
    'removed by wjmin

    '    Dim oAttr As IXMLDOMAttribute
    '    Dim oZeroNode As IXMLDOMElement
    '    Set oZeroNode = moDomMsg.selectSingleNode("//zeroout")
    '    Set oAttr = oZeroNode.Attributes.getNamedItem("okonly")
    '
    '    Dim sOkOnly As String
    '    If oAttr Is Nothing Then
    '       sOkOnly = 0
    '    Else
    '       sOkOnly = oZeroNode.getAttribute("okonly")
    '    End If
    '    If sOkOnly = 1 Then
    '       cmdOK.visible = False
    '       cmdNo.caption = "确认"
    '
    '       stbMsg.TabCaption(2) = "存货可用量"
    '    Else
    '       stbMsg.TabCaption(2) = "以下存货可用量不足"
    '    End If
End Function

Private Function SetFormat(oAttribute As IXMLDOMAttribute) As String
    Dim sName As String
    Dim sValue As String
    sName = LCase(oAttribute.nodename)
    Select Case sName
    Case "iquantity", "icurquantity", "iforquantity", "iavaquantity", "topstock", "lowstock", "moretopstock", "lesslowstock"
        sValue = Format(oAttribute.nodeValue, mologin.Account.FormatQuanDecString)
    Case "inum", "icurnum", "iavanum", "ifornum"
        sValue = Format(oAttribute.nodeValue, mologin.Account.FormatNumDecString)
    Case "iinvexchrate"
        sValue = Format(oAttribute.nodeValue, mologin.Account.FormatExchDecString)
    Case Else
        sValue = oAttribute.nodeValue
    End Select
    SetFormat = sValue

End Function
Private Function InitCaption()

    Dim oZeroNode As IXMLDOMElement
    Dim oTopLowNode As IXMLDOMElement
    Dim oCheckNode As IXMLDOMElement
    'added by wjmin 03-05-09
    Dim oItemCheckNode As IXMLDOMElement
    'the end of the added section


    Dim oNodeList As IXMLDOMNodeList
    Dim oNode As IXMLDOMElement
    Dim oAttr As IXMLDOMAttribute

    flgTopLow.Cols = 1
    flgTopLow.rows = 2
    flgZero.rows = 2
    flgZero.Cols = 1
    flgTopLow.ColWidth(0) = 300
    flgZero.ColWidth(0) = 300
    flgCheck.ColWidth(0) = 300

    'added by wjmin 03-05-09
    flgItemCheck.Cols = 1
    flgItemCheck.rows = 2
    flgItemCheck.ColWidth(0) = 300
    'the end of the added section


    Dim j As Long
    j = 1
    Set oZeroNode = moDomMsg.selectSingleNode("//zeroout")
    If oZeroNode Is Nothing Then
        stbMsg.TabVisible(0) = False
    Else
        Set oNode = oZeroNode.selectSingleNode("//zerocaption")
        flgZero.Cols = oNode.Attributes.Length + 1
        For Each oAttr In oNode.Attributes
            flgZero.TextMatrix(0, j) = oAttr.nodeValue

            j = j + 1
        Next
    End If
    j = 1
    Set oTopLowNode = moDomMsg.selectSingleNode("//toplow")
    If oTopLowNode Is Nothing Then
        stbMsg.TabVisible(1) = False
    Else
        Set oNode = oTopLowNode.selectSingleNode("//toplowcaption")
        flgTopLow.Cols = oNode.Attributes.Length + 1
        For Each oAttr In oNode.Attributes
            flgTopLow.TextMatrix(0, j) = oAttr.nodeValue
            j = j + 1
        Next
    End If
    j = 1
    Set oCheckNode = moDomMsg.selectSingleNode("//check")
    If oCheckNode Is Nothing Then
        stbMsg.TabVisible(2) = False
    Else
        Set oNode = oCheckNode.selectSingleNode("//checkcaption")
        flgCheck.Cols = oNode.Attributes.Length + 1
        For Each oAttr In oNode.Attributes
            flgCheck.TextMatrix(0, j) = oAttr.nodeValue
            j = j + 1
        Next
    End If

    'added by wjmi 03-05-09
    j = 1
    Set oItemCheckNode = moDomMsg.selectSingleNode("//itemcheck")
    If oItemCheckNode Is Nothing Then
        stbMsg.TabVisible(3) = False
    Else
        Set oNode = oItemCheckNode.selectSingleNode("//itemcheckcaption")
        flgItemCheck.Cols = oNode.Attributes.Length + 1
        For Each oAttr In oNode.Attributes
            flgItemCheck.TextMatrix(0, j) = oAttr.nodeValue
            j = j + 1
        Next
    End If
    'the end of the added section
End Function

Private Function InitGrid()
    Dim bCodeRow As Boolean
    Dim bAddRow As Boolean
    Dim oZeroNode As IXMLDOMElement
    Dim oTopLowNode As IXMLDOMElement
    Dim oCheckNode As IXMLDOMElement
    Dim oItemCheckNode As IXMLDOMElement    'added by wjmin

    Dim oNodeList As IXMLDOMNodeList
    Dim oNode As IXMLDOMElement
    Dim oAttr As IXMLDOMAttribute
    flgTopLow.ColWidth(0) = 300
    flgZero.ColWidth(0) = 300
    flgCheck.ColWidth(0) = 300
    Dim i As Long
    Dim j As Long
    Dim K As Long
    i = 1
    Set oZeroNode = moDomMsg.selectSingleNode("//zeroout")
    If oZeroNode Is Nothing Then
        stbMsg.TabVisible(0) = False
    Else
        Set oNodeList = oZeroNode.selectNodes("z:row")
        If oNodeList.Length > 0 Then

            flgZero.rows = oNodeList.Length + 1

            '            flgZero.cols = oNodeList(0).Attributes.length + 1
            K = 0
            For Each oNode In oNodeList
                j = 1
                If oNode.Attributes(0).nodename = "code" Then
                    If bCodeRow Then
                        bAddRow = False
                        If flgZero.rows > 1 Then
                            flgZero.rows = flgZero.rows - 1
                        Else

                        End If
                    Else
                        bAddRow = True

                    End If
                    bCodeRow = True
                Else
                    bCodeRow = False
                    bAddRow = True
                End If
                If bAddRow Then
                    '                    If flgZero.cols < oNode.Attributes.length + 1 Then flgZero.cols = oNode.Attributes.length + 1

                    For Each oAttr In oNode.Attributes

                        If flgZero.Cols > j Then
                            flgZero.TextMatrix(i, j) = SetFormat(oAttr)
                            j = j + 1
                        End If


                    Next
                    i = i + 1
                End If
                If oNode.Attributes.Length > K Then
                    K = oNode.Attributes.Length
                End If
            Next
            flgZero.Cols = K + 1
            '
            'Result:Row=436 Col=75  Content="子表ID"        ID=02daed59-ee28-420d-9fd2-bd6125122fae
            chkZeroOut.Enabled = (flgZero.TextMatrix(0, flgZero.Cols - 1) = GetResString("U8.ST.USERPCO.moduleco.00022"))    'GetResString("U8.ST.USKCGLSQL.frmstockmsg.00840"))
        Else
            stbMsg.TabVisible(0) = False
            '
        End If
    End If

    i = 1
    Set oTopLowNode = moDomMsg.selectSingleNode("//toplow")
    If oTopLowNode Is Nothing Then
        stbMsg.TabVisible(1) = False
    Else
        bCodeRow = False
        Set oNodeList = oTopLowNode.selectNodes("z:row")
        If oNodeList.Length > 0 Then
            flgTopLow.rows = oNodeList.Length + 1
            flgTopLow.Cols = oNodeList(0).Attributes.Length + 1
            For Each oNode In oNodeList
                j = 1
                If oNode.Attributes(0).nodename = "code" Then
                    If bCodeRow Then
                        bAddRow = False
                        If flgTopLow.rows > 1 Then
                            flgTopLow.rows = flgTopLow.rows - 1
                        End If
                    Else
                        bAddRow = True
                    End If
                    bCodeRow = True
                Else
                    bCodeRow = False
                    bAddRow = True
                End If
                If bAddRow Then
                    For Each oAttr In oNode.Attributes
                        flgTopLow.TextMatrix(i, j) = SetFormat(oAttr)
                        j = j + 1
                    Next
                    i = i + 1
                End If
            Next
        Else
            stbMsg.TabVisible(1) = False
        End If
    End If

    i = 1
    Set oCheckNode = moDomMsg.selectSingleNode("//check")
    If oCheckNode Is Nothing Then
        stbMsg.TabVisible(2) = False
    Else
        bCodeRow = False
        Set oNodeList = oCheckNode.selectNodes("z:row")
        If oNodeList.Length > 0 Then
            flgCheck.rows = oNodeList.Length + 1
            flgCheck.Cols = oNodeList(0).Attributes.Length + 1
            For Each oNode In oNodeList
                j = 1
                If oNode.Attributes(0).nodename = "code" Then
                    If bCodeRow Then
                        bAddRow = False
                        If flgCheck.rows > 1 Then
                            flgCheck.rows = flgCheck.rows - 1
                        End If
                    Else
                        bAddRow = True
                    End If
                    bCodeRow = True
                Else
                    bCodeRow = False
                    bAddRow = True
                End If
                If bAddRow Then
                    For Each oAttr In oNode.Attributes
                        If oNode.nodename = "iavaquantity" Then
                            If CDbl(oAttr.nodeValue) < 0 Then
                                flgCheck.CellBackColor = vbRed
                            End If
                        End If
                        flgCheck.TextMatrix(i, j) = SetFormat(oAttr)
                        j = j + 1
                    Next
                    i = i + 1
                End If
            Next
        Else
            stbMsg.TabVisible(2) = False
        End If
    End If

    'added by wjmin
    Dim nodename As String
    If Not moDomMsg.selectSingleNode("//z:row") Is Nothing Then
        nodename = "z:row"
    Else
        nodename = "w:row"
    End If

    i = 1
    Set oItemCheckNode = moDomMsg.selectSingleNode("//itemcheck")
    If oItemCheckNode Is Nothing Then
        stbMsg.TabVisible(3) = False
    Else
        bCodeRow = False
        Set oNodeList = oItemCheckNode.selectNodes(nodename)
        If oNodeList.Length > 0 Then
            flgItemCheck.rows = oNodeList.Length + 1
            flgItemCheck.Cols = oNodeList(0).Attributes.Length + 1
            For Each oNode In oNodeList
                j = 1
                If oNode.Attributes(0).nodename = "code" Then
                    If bCodeRow Then
                        bAddRow = False
                        If flgItemCheck.rows > 1 Then
                            flgItemCheck.rows = flgItemCheck.rows - 1
                        End If
                    Else
                        bAddRow = True
                    End If
                    bCodeRow = True
                Else
                    bCodeRow = False
                    bAddRow = True
                End If
                If bAddRow Then
                    For Each oAttr In oNode.Attributes
                        flgItemCheck.TextMatrix(i, j) = SetFormat(oAttr)
                        j = j + 1
                    Next
                    i = i + 1
                End If
            Next
        Else
            stbMsg.TabVisible(3) = False
        End If
    End If



    If flgTopLow.Cols > 1 Then
        flgTopLow.ColWidth(1) = Me.TextWidth("aaaaaaaaaaaaaaaaaaaaaa")
    End If
    If flgZero.Cols > 1 Then
        flgZero.ColWidth(1) = Me.TextWidth("aaaaaaaaaaaaaaaaaaaaaa")
    End If
    If flgCheck.Cols > 1 Then
        flgCheck.ColWidth(1) = Me.TextWidth("aaaaaaaaaaaaaaaaaaaaaa")
    End If

    'added
    If flgItemCheck.Cols > 1 Then
        flgItemCheck.ColWidth(1) = Me.TextWidth("aaaaaaaaaaaaaaaaaaaaaa")
    End If


End Function

Public Property Get Result() As VbMsgBoxResult
    Result = moSelect
End Property

Public Property Let Result(ByVal vNewValue As VbMsgBoxResult)
    moSelect = vNewValue
End Property


Public Property Get ZeroOut() As Boolean
    ZeroOut = moZeroOut
End Property

Public Property Let ZeroOut(ByVal vNewValue As Boolean)
    moZeroOut = vNewValue
End Property


Public Property Get vouchtype() As String
    vouchtype = msVouchType
End Property

Public Property Let vouchtype(ByVal vNewValue As String)
    msVouchType = vNewValue
End Property

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim resLog As Object
    Set resLog = CreateObject("MultiLangPkg.ResLog")
    resLog.Unload
End Sub
