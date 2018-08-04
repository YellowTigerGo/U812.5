VERSION 5.00
Object = "{005151D4-33F6-471B-B651-222D4E411832}#4.4#0"; "UFFormPartner.ocx"
Object = "{AF8BBBB7-94C6-4772-B826-624478C37D6A}#1.5#0"; "UFKEYHOOK.ocx"
Object = "{9FD12F62-6922-47E1-B1AC-3615BBD3D7A5}#1.0#0"; "UFLabel.ocx"
Object = "{A0C292A3-118E-11D2-AFDF-000021730160}#1.0#0"; "UFEDIT.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.OCX"
Object = "{51388549-C886-4FD6-AE5F-8AA28C63CE94}#1.0#0"; "PrintControl.ocx"
Object = "{72EA6AED-E68F-4497-BCB2-A5DB5A9ECB0A}#1.0#0"; "UFListBox.ocx"
Object = "{4C2F9AC0-6D40-468A-8389-518BB4F8C67D}#1.0#0"; "UFComboBox.ocx"
Object = "{8C7C777D-4D83-4DE8-947E-098E2343A400}#1.0#0"; "CommandButton.ocx"
Begin VB.Form frmCheckCredit 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "信用检查"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   9510
   StartUpPosition =   1  '所有者中心
   Begin UFKeyHook.UFKeyHookCtrl UFKeyHookCtrl1 
      Left            =   2640
      Top             =   3000
      _ExtentX        =   1905
      _ExtentY        =   529
   End
   Begin UFFormPartner.UFFrmCaption UFFrmCaptionMgr 
      Left            =   840
      Top             =   3600
      _ExtentX        =   1085
      _ExtentY        =   238
      Caption         =   "信用检查"
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
   Begin PRINTCONTROLLib.PrintControl Printer 
      Height          =   240
      Left            =   45
      TabIndex        =   9
      Top             =   3105
      Visible         =   0   'False
      Width           =   285
      _Version        =   65536
      _ExtentX        =   503
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin EDITLib.Edit Textpass 
      Height          =   315
      Left            =   4125
      TabIndex        =   4
      Top             =   3255
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   556
      _StockProps     =   253
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      PasswordChar    =   "*"
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   3480
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin UFCOMMANDBUTTONLib.UFCommandButton cmdExp 
      Height          =   300
      Left            =   8220
      TabIndex        =   8
      ToolTipText     =   "保存到文件"
      Top             =   3270
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1984
      _ExtentY        =   529
      _StockProps     =   41
      BackColor       =   -2147483633
      Caption         =   "保存"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   1140
      Left            =   735
      TabIndex        =   7
      Top             =   525
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   2011
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   16769990
      BackColorSel    =   10446406
      ForeColorSel    =   16777215
      BackColorBkg    =   16777215
      GridColor       =   11447982
      GridColorFixed  =   8947848
      AllowUserResizing=   3
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin UFLISTBOXLib.UFListBox ListCredit 
      Height          =   2910
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9255
      _Version        =   65536
      _ExtentX        =   16325
      _ExtentY        =   5133
      _StockProps     =   207
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      Appearance      =   0
   End
   Begin UFCOMBOBOXLib.UFComboBox Combo1 
      Height          =   2040
      Left            =   1485
      TabIndex        =   2
      Top             =   3240
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   529
      _StockProps     =   196
      Appearance      =   1
      Text            =   ""
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
   End
   Begin UFCOMMANDBUTTONLib.UFCommandButton CmdCancel 
      Height          =   300
      Left            =   6960
      TabIndex        =   1
      ToolTipText     =   "取 消"
      Top             =   3255
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1984
      _ExtentY        =   529
      _StockProps     =   41
      BackColor       =   -2147483633
      Caption         =   "取 消"
      UToolTipText    =   ""
      Cursor          =   17925
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
   Begin UFCOMMANDBUTTONLib.UFCommandButton CmdYes 
      Height          =   300
      Left            =   5700
      TabIndex        =   0
      ToolTipText     =   "确 定"
      Top             =   3270
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1984
      _ExtentY        =   529
      _StockProps     =   41
      BackColor       =   -2147483633
      Caption         =   "确 定"
      UToolTipText    =   ""
      Cursor          =   17925
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
   Begin UFLABELLib.UFLabel Label2 
      Height          =   255
      Left            =   3285
      TabIndex        =   6
      Top             =   3300
      Width           =   930
      _Version        =   65536
      _ExtentX        =   1640
      _ExtentY        =   450
      _StockProps     =   111
      Caption         =   "输入密码"
      BackColor       =   -2147483633
      BackStyle       =   0
   End
   Begin UFLABELLib.UFLabel Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3300
      Width           =   1425
      _Version        =   65536
      _ExtentX        =   2514
      _ExtentY        =   450
      _StockProps     =   111
      Caption         =   "选择信用审核人"
      BackColor       =   -2147483633
      BackStyle       =   0
   End
End
Attribute VB_Name = "frmCheckCredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cCheckName As String, cCheckPass As String
Attribute cCheckPass.VB_VarUserMemId = 1073938432
Private bcancel As Boolean
Attribute bcancel.VB_VarUserMemId = 1073938434
Private CheCkBalance As Double, CheckDate As Integer
Attribute CheCkBalance.VB_VarUserMemId = 1073938435
Attribute CheckDate.VB_VarUserMemId = 1073938435
Private cPass() As String
Attribute cPass.VB_VarUserMemId = 1073938437
Private iJsq As Integer
Attribute iJsq.VB_VarUserMemId = 1073938438
Private CheckMode As Integer  '0信用 1现存量
Attribute CheckMode.VB_VarUserMemId = 1073938439
Public SaveAfterOk As Boolean    '点击确定后继续保存
Attribute SaveAfterOk.VB_VarUserMemId = 1073938440

'销售系统选项
Public myinfo As Object
Attribute myinfo.VB_VarUserMemId = 1073938441

Public Property Get cCheckerName() As String
    cCheckerName = cCheckName
End Property
Public Property Get cCheckerPass() As String
    cCheckerPass = cCheckPass
End Property
Public Property Get bCanceled() As Boolean
    bCanceled = bcancel
End Property
''显示信用检查窗口
''传递参数:mDom :信用检查返回的dom,errMsg:错误返回信息
''
Public Function CheckShow(mDom As DOMDocument, errMsg As String, Optional iMode = 0) As Boolean
    Dim ele As IXMLDOMElement, ele1 As IXMLDOMElement
    Dim i As Long, j As Long, row As Long, Col As Long
    Dim rstTmp As New ADODB.Recordset
    Dim sName As String
    Dim isOkOnly As Integer    'iMode=2
    '如果 dom中存在okonly则检查不通过 cmdyes执灰 cancel显示 保存显示 isOkOnly>0
    ''如果 dom中不存在okonly则检查通过 cmdyes显示 cancel显示 保存显示 isOkOnly=0
    On Error GoTo DoErr

    bcancel = True
    CheCkBalance = 0
    CheckDate = 0
    ''处理额度
    ListCredit.Clear
    CheckMode = iMode
    Me.Grid.Visible = False
    Me.ListCredit.Visible = True
    Select Case iMode
    Case 0    ''信用
        For i = 0 To mDom.selectNodes("//error/credit").Length - 1
            ListCredit.AddItem mDom.selectNodes("//error/credit").Item(i).Attributes.getNamedItem("cname").nodeValue + Format(Val(mDom.selectNodes("//error/credit").Item(i).Attributes.getNamedItem("unpay").nodeValue), "#,##0.00")
            'ListCredit.ToolTipText = mDom.selectNodes("//error/credit").item(i).Attributes.getNamedItem("cname").nodeValue + str(Round(Val(mDom.selectNodes("//error/credit").item(i).Attributes.getNamedItem("unpay").nodeValue), 2))
        Next i
        ''ListCredit.AddItem mDom.selectNodes("//error/credithj").item(i).Attributes.getNamedItem("cname").nodeValue + str(Round(Val(mDom.selectNodes("//error/credithj").item(i).Attributes.getNamedItem("unpay").nodeValue), 2))
        For i = 0 To mDom.selectNodes("//error/credithj").Length - 1
            If CheCkBalance < CDbl(Round(Val(mDom.selectNodes("//error/credithj").Item(i).Attributes.getNamedItem("unpay").nodeValue), 2)) Then
                CheCkBalance = CDbl(Round(Val(mDom.selectNodes("//error/credithj").Item(i).Attributes.getNamedItem("unpay").nodeValue), 2))
            End If
        Next i

        'Set ele = mDom.selectNodes("//error/creditdatehj")
        For i = 0 To mDom.selectNodes("//error/creditdate").Length - 1
            ListCredit.AddItem mDom.selectNodes("//error/creditdate").Item(i).Attributes.getNamedItem("cname").nodeValue + mDom.selectNodes("//error/creditdate").Item(i).Attributes.getNamedItem("unpay").nodeValue
        Next i
        'ListCredit.AddItem mDom.selectNodes("//error/creditdatehj").item(i).Attributes.getNamedItem("cname").nodeValue + mDom.selectNodes("//error/creditdatehj").item(i).Attributes.getNamedItem("unpay").nodeValue
        For i = 0 To mDom.selectNodes("//error/creditdatehj").Length - 1
            If CheckDate < Val(mDom.selectNodes("//error/creditdatehj").Item(i).Attributes.getNamedItem("unpay").nodeValue) Then
                CheckDate = Val(mDom.selectNodes("//error/creditdatehj").Item(i).Attributes.getNamedItem("unpay").nodeValue)
            End If
        Next i

        If myinfo.bCrCtrWShow = True Then
            With rstTmp
                .ActiveConnection = dbconn
                .CursorLocation = adUseClient
                .Open "select cPersonName,cPass from SA_CreditCheck where (isnull(iCreLine,0)>=" & CheCkBalance & " or isnull(iCreLine,0)=0) and (isnull(iCreDate,0)>=" & CheckDate & " or isnull(iCreDate,0)=0) order by cPersonName"
            End With
            If Not rstTmp.EOF Then
                ReDim cPass(rstTmp.RecordCount, 1)
                i = 0
                Do While Not rstTmp.EOF
                    Combo1.AddItem rstTmp(0), i
                    cPass(i, 0) = rstTmp(0)
                    cPass(i, 1) = rstTmp(1)
                    rstTmp.MoveNext
                    i = i + 1
                Loop
                Combo1.ListIndex = 0
            Else
                CmdYes.Enabled = False
                ListCredit.AddItem "没有审核人可以复核此单据"
            End If
            jsq = 0
            rstTmp.Close
            Set rstTmp = Nothing
        Else
            Me.Label1.Visible = False
            Me.Label2.Visible = False
            Me.Textpass.Visible = False
            Me.Combo1.Visible = False
        End If
        'Me.Caption = GetString("U8.SA.xsglsql.01.frmcheckcredit.00458") 'zh-CN：信用检查不通过：
        UFFrmCaptionMgr.Caption = GetString("U8.DZ.JA.Res850")
    Case 1  ''可用量
        UFFrmCaptionMgr.Caption = GetString("U8.DZ.JA.Res860")
        For i = 0 To mDom.selectNodes("//可用量检查不过/存货").Length - 1
            ListCredit.AddItem GetStringPara("U8.SA.xsglsql.01.frmcheckcredit.00461", mDom.selectNodes("//可用量检查不过/存货").Item(i).Attributes.getNamedItem("存货编码").nodeValue, mDom.selectNodes("//可用量检查不过/存货").Item(i).Attributes.getNamedItem("存货名称").nodeValue, _
                                             str(gcSales.FourFive(Val(mDom.selectNodes("//可用量检查不过/存货").Item(i).Attributes.getNamedItem("可用量").nodeValue), myinfo.cMQBit)), str(gcSales.FourFive(Val(mDom.selectNodes("//可用量检查不过/存货").Item(i).Attributes.getNamedItem("订货数量").nodeValue), myinfo.cMQBit)))    '编码为 {0}的存货{1}的可用量为{2}, 本单据上此存货数量合计{3}
        Next i
        Me.Label1.Visible = False
        Me.Label2.Visible = False
        Me.Textpass.Visible = False
        Me.Combo1.Visible = False
    Case 2
        UFFrmCaptionMgr.Caption = GetString("U8.DZ.JA.Res870")
        Me.ListCredit.Visible = False
        Me.Grid.Visible = True
        Me.Grid.Redraw = False
        Set ele = mDom.selectSingleNode("//rs:data/zeroout")
        ''填充
        If ele.Attributes.getNamedItem("okonly") Is Nothing Then
            isOkOnly = 0
        Else
            isOkOnly = 1
        End If
        With Grid
            row = ele.selectNodes("//zeroout/z:row").Length
            Col = ele.selectSingleNode("zerocaption").Attributes.Length
            .rows = row + 1
            .Cols = Col + 1
            For j = 0 To Col
                .ColWidth(j) = 1000
                For i = 0 To row
                    If j = 0 Then
                        If i = 0 Then
                            .TextMatrix(i, j) = ""
                        Else
                            .TextMatrix(i, j) = str(i)
                        End If
                    Else
                        If i = 0 Then
                            sName = ele.selectSingleNode("zerocaption").Attributes.Item(j - 1).nodename
                            .TextMatrix(i, j) = ele.selectSingleNode("zerocaption").Attributes.getNamedItem(sName).nodeValue
                        Else
                            If Not ele.selectNodes("//zeroout/z:row").Item(i - 1).Attributes.getNamedItem(sName) Is Nothing Then
                                Select Case LCase(sName)
                                Case "inum", "icurnum", "iavanum", "ifornum"
                                    .TextMatrix(i, j) = gcSales.FourFive(ele.selectNodes("//zeroout/z:row").Item(i - 1).Attributes.getNamedItem(sName).nodeValue, myinfo.cPieceBit)
                                Case "iquantity", "icurquantity", "iavaquantity", "iforquantity"
                                    .TextMatrix(i, j) = gcSales.FourFive(ele.selectNodes("//zeroout/z:row").Item(i - 1).Attributes.getNamedItem(sName).nodeValue, myinfo.cMQBit)
                                Case Else
                                    .TextMatrix(i, j) = ele.selectNodes("//zeroout/z:row").Item(i - 1).Attributes.getNamedItem(sName).nodeValue
                                End Select
                                '                                    If LCase(sName) = "bokonly" Then
                                '                                        isOkOnly = isOkOnly + 1
                                '                                    End If
                            End If
                        End If

                    End If

                Next i
            Next j
            .Width = Me.ListCredit.Width
            .Height = Me.ListCredit.Height
            .Left = Me.ListCredit.Left
            .Top = Me.ListCredit.Top
            .Redraw = True
        End With
        Me.Label1.Visible = False
        Me.Label2.Visible = False
        Me.Textpass.Visible = False
        Me.Combo1.Visible = False
        Me.CmdCancel.Visible = True
        If isOkOnly > 0 Then
            Me.CmdYes.Enabled = False
        ElseIf isOkOnly = 0 Then
            Me.CmdYes.Enabled = True
        End If
    Case 3  ''最低售价
        UFFrmCaptionMgr.Caption = GetString("U8.DZ.JA.Res880")
        For i = 0 To mDom.selectNodes("//err/list").Length - 1
            ListCredit.AddItem mDom.selectNodes("//err/list").Item(i).Attributes.getNamedItem("cvalue").nodeValue
        Next i
        Me.Label1.Visible = False
        Me.Label2.Visible = True
        Me.Label2.Caption = GetString("U8.DZ.JA.Res890")
        Me.Label2.Width = 1200
        Label2.Left = Label2.Left - 500
        Me.Textpass.Visible = True
        Me.Combo1.Visible = False

    End Select
    Me.Show vbModal
    CheckShow = True
    Exit Function
DoErr:
    CheckShow = False
    errMsg = Err.Description
    Set rstTmp = Nothing
End Function

Private Sub cmdCancel_Click()
    SaveAfterOk = False
    bcancel = True
    Unload Me
End Sub

Private Sub cmdExp_Click()
    On Error Resume Next
    Dim str As String, strContent As String, iLen As Long, sSpace As String
    Dim iRow As Long, iCol As Long

    If Grid.Visible = True Then
        PrintGrid Me.Printer, "out", Grid, "", , , 1
        Exit Sub
    End If

    comDlg.Filter = "*.txt|*.txt|*.*|*.*"
    comDlg.ShowSave
    str = comDlg.FileName
    If str = "" Then Exit Sub
    If Dir(str) <> "" Then
        If MsgBox(GetString("U8.DZ.JA.Res900"), vbOKCancel + vbQuestion, Me.Caption) = vbCancel Then    'zh-CN：文件已经存在，是否覆盖？
            Exit Sub
        Else
            Kill str
        End If
    End If
    Open str For Output As #1

    If ListCredit.Visible = True Then
        For iRow = 0 To ListCredit.ListCount - 1
            Write #1, ListCredit.List(iRow)
        Next
    End If
    Close #1

End Sub

Private Sub PrintGrid(objPrinter As PrintControl, strMode As String, GrdEdit As Object, strTitle As String, Optional str1 As String, Optional str2 As String, Optional lngCheckCol As Long, Optional sumGrid As Object)
    Dim e As Long
    Dim strError As Variant


    On Error GoTo ErrHandle

    objPrinter.EnableSave = True

    e = objPrinter.ExportToFile(0, "10", "254", "", "")

    If e = 0 Then
        MsgBox GetString("U8.DZ.JA.Res910"), vbOKOnly, GetString("U8.DZ.JA.Res030")
    Else
        If e <> 3999 Then
            objPrinter.GetErrorMessage e, strError
            MsgBox strError, vbInformation, GetString("U8.DZ.JA.Res030")
        End If
    End If

    Exit Sub
ErrHandle:
    MsgBox Err.Description, vbInformation, , GetString("U8.DZ.JA.Res030")
End Sub


Private Sub cmdYes_Click()
    Dim i As Integer, j As Integer

    On Error GoTo DoErr
    Select Case CheckMode
    Case 0
        If myinfo.bCrCtrWShow = True And Combo1.ListCount > 0 Then
            iJsq = iJsq + 1
            j = -1
            For i = 0 To Combo1.ListCount
                If cPass(i, 0) = Combo1.Text Then
                    j = i
                End If
            Next i
            If j = -1 Then
                ReDim varArgs(0)
                varArgs(0) = Combo1.Text
                MsgBox GetStringPara("U8.DZ.JA.Res920", varArgs(0)), GetString("U8.DZ.JA.Res030")
                '                MsgBox "没找到审核人" & Combo1.Text & ",请检查!", GetString("U8.DZ.JA.Res030")
                Exit Sub
            Else
                If LCase(Textpass.Text) = LCase(cPass(j, 1)) Then
                    cCheckName = cPass(j, 0)
                    cCheckPass = cPass(j, 1)
                    bcancel = False
                    Unload Me
                Else
                    If iJsq < 3 Then
                        ReDim varArgs(0)
                        varArgs(0) = CStr(3 - iJsq)
                        MsgBox GetStringPara("U8.DZ.JA.Res930", varArgs(0)), GetString("U8.DZ.JA.Res030")
                        '               MsgBox "密码不正确,请重新输入!(还有" & CStr(3 - iJsq) & "次机会)", GetString("U8.DZ.JA.Res030")
                    Else
                        bcancel = True
                        Unload Me
                    End If
                End If
            End If
        Else
            cCheckName = ""
            cCheckPass = "ufsoft"
            bcancel = False
            Unload Me
        End If
    Case 1
        bcancel = False
        Unload Me
    Case 2
        bcancel = False
        SaveAfterOk = True
        Unload Me
    Case 3
        iJsq = iJsq + 1
        If LCase(Textpass.Text) = LCase(myinfo.cLowPricePwd) Then
            bcancel = False
            Unload Me
        Else
            If iJsq < 3 Then
                ReDim varArgs(0)
                varArgs(0) = CStr(3 - iJsq)
                MsgBox GetStringPara("U8.DZ.JA.Res930", varArgs(0)), GetString("U8.DZ.JA.Res030")
                '               MsgBox "密码不正确,请重新输入!(还有" & CStr(3 - iJsq) & "次机会)", GetString("U8.DZ.JA.Res030")
            Else
                bcancel = True
                Unload Me
            End If
        End If
    End Select
    Exit Sub
DoErr:
    MsgBox Err.Description, vbInformation, GetString("U8.DZ.JA.Res030")
End Sub



Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call cmdYes_Click
    End If
End Sub


Private Sub Form_Load()
'u861多语Layout
'   Call MultiLangLoadForm(Me, "u8.sa.xsglsql." & Me.Name)
    iJsq = 0
    InitGrdCol Me.Grid
End Sub

Private Sub InitGrdCol(Grid As Object)
    With Grid
        .BackColor = &HFFFFFF
        .BackColorBkg = &HFFFFFF
        .BackColorFixed = &HFFE3C6
        .BackColorSel = &H9F6646
        .ForeColorSel = &HFFFFFF
        .GridColor = &HAEAEAE
        .GridColorFixed = &H888888
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Call U861ResEnd
End Sub

Private Sub Textpass_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.CmdYes.Enabled = True Then
            Call cmdYes_Click
        End If
    End If
End Sub



Private Sub UFKeyHookCtrl1_ContainerKeyUp(KeyCode As Integer, ByVal Shift As Integer)
    Call Form_KeyUp(KeyCode, Shift)
End Sub
