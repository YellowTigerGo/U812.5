VERSION 5.00
Object = "{A0C292A3-118E-11D2-AFDF-000021730160}#1.0#0"; "UFEDIT.OCX"
Object = "{005151D4-33F6-471B-B651-222D4E411832}#4.4#0"; "UFFormPartner.ocx"
Object = "{4C2F9AC0-6D40-468A-8389-518BB4F8C67D}#1.0#0"; "UFComboBox.ocx"
Object = "{9FD12F62-6922-47E1-B1AC-3615BBD3D7A5}#1.0#0"; "UFLabel.ocx"
Object = "{BF022F1C-E440-4790-987F-252926B9B602}#5.1#0"; "UFFrames.ocx"
Object = "{A98B9C82-88D8-4F94-91BB-F2289111C59C}#1.0#0"; "UFCheckBox.ocx"
Object = "{8C7C777D-4D83-4DE8-947E-098E2343A400}#1.0#0"; "CommandButton.ocx"
Object = "{D5646CCD-3DEF-4356-8564-4C2AB79D21E9}#2.2#0"; "UFRadio.ocx"
Begin VB.Form BillPrn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "打印"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7095
   ControlBox      =   0   'False
   Icon            =   "BillPrn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   7095
   StartUpPosition =   2  '屏幕中心
   Begin UFCHECKBOXLib.UFCheckBox chkFD3 
      Height          =   210
      Left            =   195
      TabIndex        =   18
      Top             =   2880
      Visible         =   0   'False
      Width           =   2445
      _Version        =   65536
      _ExtentX        =   4313
      _ExtentY        =   370
      _StockProps     =   15
      Caption         =   "按发货单分单打印"
      ForeColor       =   0
      ForeColor       =   0
      BorderStyle     =   0
      ReadyState      =   0
      Picture         =   "BillPrn.frx":000C
   End
   Begin UFFormPartner.UFFrmCaption UFFrmCaptionMgr 
      Left            =   600
      Top             =   2760
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "打印"
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
   Begin UFCHECKBOXLib.UFCheckBox chkFirst 
      Height          =   210
      Left            =   4515
      TabIndex        =   12
      Top             =   2505
      Width           =   2445
      _Version        =   65536
      _ExtentX        =   4313
      _ExtentY        =   370
      _StockProps     =   15
      Caption         =   "下次打印不再弹出此窗口"
      ForeColor       =   0
      ForeColor       =   0
      BorderStyle     =   0
      ReadyState      =   0
      Picture         =   "BillPrn.frx":0028
   End
   Begin UFCHECKBOXLib.UFCheckBox chkFD2 
      Height          =   270
      Left            =   2790
      TabIndex        =   11
      Top             =   2475
      Width           =   1590
      _Version        =   65536
      _ExtentX        =   2805
      _ExtentY        =   476
      _StockProps     =   15
      Caption         =   "按存货大类分单"
      ForeColor       =   0
      ForeColor       =   0
      BorderStyle     =   0
      ReadyState      =   0
      Picture         =   "BillPrn.frx":0044
   End
   Begin UFCHECKBOXLib.UFCheckBox chkFD1 
      Height          =   195
      Left            =   1440
      TabIndex        =   10
      Top             =   2520
      Width           =   1290
      _Version        =   65536
      _ExtentX        =   2275
      _ExtentY        =   344
      _StockProps     =   15
      Caption         =   "按仓库分单"
      ForeColor       =   0
      ForeColor       =   0
      BorderStyle     =   0
      ReadyState      =   0
      Picture         =   "BillPrn.frx":0060
   End
   Begin UFCOMMANDBUTTONLib.UFCommandButton Command2 
      Height          =   300
      Left            =   5850
      TabIndex        =   9
      ToolTipText     =   "取消"
      Top             =   2850
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1984
      _ExtentY        =   529
      _StockProps     =   41
      Caption         =   "取消"
      UToolTipText    =   ""
      Cursor          =   1673
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
   Begin UFCHECKBOXLib.UFCheckBox chkDiscount 
      Height          =   195
      Left            =   195
      TabIndex        =   8
      Top             =   2520
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1984
      _ExtentY        =   344
      _StockProps     =   15
      Caption         =   "折扣横打"
      ForeColor       =   0
      ForeColor       =   0
      BorderStyle     =   0
      ReadyState      =   0
      Picture         =   "BillPrn.frx":007C
   End
   Begin UFFrames.UFFrame framePrn 
      Height          =   2340
      Left            =   120
      TabIndex        =   1
      Top             =   75
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   4128
      Caption         =   "90099879"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin UFRadioLib.UFRadio op2 
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   1365
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   318
         Style           =   0
         Caption         =   "全部汇总打印"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   0
         Appearance      =   1
         BackColor       =   -2147483633
         DisabledPicture =   "BillPrn.frx":0098
         DownPicture     =   "BillPrn.frx":00B4
         Enabled         =   -1  'True
         ForeColor       =   -2147483630
         MaskColor       =   12632256
         MouseIcon       =   "BillPrn.frx":00D0
         MousePointer    =   0
         Picture         =   "BillPrn.frx":00EC
         OLEDropMode     =   0
         RightToLeft     =   0   'False
         UseMaskColor    =   0   'False
         Value           =   0   'False
      End
      Begin EDITLib.Edit AXCase1 
         Height          =   350
         Left            =   1575
         TabIndex        =   16
         Top             =   1280
         Width           =   3375
         _Version        =   65536
         _ExtentX        =   5953
         _ExtentY        =   617
         _StockProps     =   253
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
      End
      Begin UFCOMBOBOXLib.UFComboBox CmbChoice 
         Height          =   2040
         Left            =   1575
         TabIndex        =   14
         Top             =   1935
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
         _ExtentY        =   529
         _StockProps     =   196
         Appearance      =   1
         Text            =   ""
         Style           =   2
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
      End
      Begin UFCHECKBOXLib.UFCheckBox chkBatch 
         Height          =   210
         Left            =   2145
         TabIndex        =   13
         Top             =   660
         Width           =   1155
         _Version        =   65536
         _ExtentX        =   2037
         _ExtentY        =   370
         _StockProps     =   15
         Caption         =   "考虑批号"
         ForeColor       =   0
         ForeColor       =   0
         BorderStyle     =   0
         ReadyState      =   0
         Picture         =   "BillPrn.frx":0108
      End
      Begin UFCHECKBOXLib.UFCheckBox Check1 
         Height          =   225
         Left            =   4515
         TabIndex        =   7
         Top             =   270
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   397
         _StockProps     =   15
         Caption         =   "劳务不需要数量单价"
         ForeColor       =   0
         ForeColor       =   0
         BorderStyle     =   0
         ReadyState      =   0
         Picture         =   "BillPrn.frx":0124
      End
      Begin UFCHECKBOXLib.UFCheckBox Chk2 
         Height          =   240
         Left            =   2145
         TabIndex        =   6
         Top             =   270
         Width           =   1830
         _Version        =   65536
         _ExtentX        =   3228
         _ExtentY        =   423
         _StockProps     =   15
         Caption         =   "按销货清单打印"
         ForeColor       =   0
         ForeColor       =   0
         BorderStyle     =   0
         ReadyState      =   0
         Picture         =   "BillPrn.frx":0140
      End
      Begin UFRadioLib.UFRadio Op4 
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   990
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   397
         Style           =   0
         Caption         =   "按货物所属分类相同的汇总打印"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   0
         Appearance      =   1
         BackColor       =   -2147483633
         DisabledPicture =   "BillPrn.frx":015C
         DownPicture     =   "BillPrn.frx":0178
         Enabled         =   -1  'True
         ForeColor       =   -2147483630
         MaskColor       =   12632256
         MouseIcon       =   "BillPrn.frx":0194
         MousePointer    =   0
         Picture         =   "BillPrn.frx":01B0
         OLEDropMode     =   0
         RightToLeft     =   0   'False
         UseMaskColor    =   0   'False
         Value           =   0   'False
      End
      Begin UFCHECKBOXLib.UFCheckBox Chk1 
         Height          =   270
         Left            =   5025
         TabIndex        =   4
         Top             =   780
         Width           =   1392
         _Version        =   65536
         _ExtentX        =   2455
         _ExtentY        =   476
         _StockProps     =   15
         Caption         =   "数量需要合计"
         ForeColor       =   0
         ForeColor       =   0
         BorderStyle     =   0
         ReadyState      =   0
         Picture         =   "BillPrn.frx":01CC
      End
      Begin UFRadioLib.UFRadio Op3 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   636
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         Style           =   0
         Caption         =   "存货+单价相同合并"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   0
         Appearance      =   1
         BackColor       =   -2147483633
         DisabledPicture =   "BillPrn.frx":01E8
         DownPicture     =   "BillPrn.frx":0204
         Enabled         =   -1  'True
         ForeColor       =   -2147483630
         MaskColor       =   12632256
         MouseIcon       =   "BillPrn.frx":0220
         MousePointer    =   0
         Picture         =   "BillPrn.frx":023C
         OLEDropMode     =   0
         RightToLeft     =   0   'False
         UseMaskColor    =   0   'False
         Value           =   0   'False
      End
      Begin UFRadioLib.UFRadio op1 
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   318
         Style           =   0
         Caption         =   "打印清单"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   0
         Appearance      =   1
         BackColor       =   -2147483633
         DisabledPicture =   "BillPrn.frx":0258
         DownPicture     =   "BillPrn.frx":0274
         Enabled         =   -1  'True
         ForeColor       =   -2147483630
         MaskColor       =   12632256
         MouseIcon       =   "BillPrn.frx":0290
         MousePointer    =   0
         Picture         =   "BillPrn.frx":02AC
         OLEDropMode     =   0
         RightToLeft     =   0   'False
         UseMaskColor    =   0   'False
         Value           =   -1  'True
      End
      Begin UFLABELLib.UFLabel Label1 
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1965
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   111
         Caption         =   "请选择打印模版"
         BackColor       =   16777215
         BackStyle       =   0
      End
      Begin VB.Line Line2 
         Visible         =   0   'False
         X1              =   4455
         X2              =   4815
         Y1              =   1440
         Y2              =   915
      End
      Begin VB.Line Line1 
         Visible         =   0   'False
         X1              =   4410
         X2              =   4818
         Y1              =   750
         Y2              =   918
      End
   End
   Begin UFCOMMANDBUTTONLib.UFCommandButton Command1 
      Height          =   300
      Left            =   4620
      TabIndex        =   0
      ToolTipText     =   "确定"
      Top             =   2850
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1984
      _ExtentY        =   529
      _StockProps     =   41
      Caption         =   "确定"
      UToolTipText    =   ""
      Cursor          =   -1
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
End
Attribute VB_Name = "BillPrn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bLY As Boolean
Public bCancelExit As Boolean
Public bDiscount As Boolean
Public iType As Long
Public bFirst As Boolean
Dim m_blnBatch As Boolean
Dim m_VTID As Long
Public strCardNum As String
Public strVouchtype As String
Dim domVtid As New DOMDocument




'2001-2-8 by liutao
'---------------------
Private Sub Chk2_Click()
    If Chk2.value = 1 Then
        chkDiscount.Enabled = False
        chkDiscount.value = 0
    Else
        chkDiscount.Enabled = True
    End If
End Sub
'---------------------

Private Sub chkFirst_Click()
    If chkFirst.value <> vbUnchecked Then
        bFirst = True
    Else
        bFirst = False
    End If
End Sub


Private Sub Command1_Click()
    On Error Resume Next
  
    If op1 Then BillPrnSet = 1
    If op2 Then BillPrnSet = 2
    If Op3 Then BillPrnSet = 3
    If Op4 Then BillPrnSet = 4
    If op1 And Chk2 Then BillPrnSet = 11
    If Op3 And Chk1 Then BillPrnSet = 31
    If Op4 And Chk1 Then BillPrnSet = 41
    '2001-3-9 by liutao 汇总打印时，也需要有数量是否合计的选项
    '------------------
    If op2 And Chk1 Then BillPrnSet = 51
    '------------------
    'wxy add on 20051026 增加发票 按照 发货单分单打印74010
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If strVouchtype <> "27" And strVouchtype <> "26" Then
        If Not chkFD1 And Not chkFD2 Then VouchPrnFD = 0
        If chkFD1 And Not chkFD2 Then VouchPrnFD = 1
        If Not chkFD1 And chkFD2 Then VouchPrnFD = 2
        If chkFD1 And chkFD2 Then VouchPrnFD = 3
    Else
        If (Not chkFD1 And Not chkFD2) And Not chkFD3 Then VouchPrnFD = 0
        If chkFD1 And Not chkFD2 And Not chkFD3 Then VouchPrnFD = 1
        If Not chkFD1 And chkFD2 And Not chkFD3 Then VouchPrnFD = 2
        If chkFD1 And chkFD2 And Not chkFD3 Then VouchPrnFD = 3
        
        If (Not chkFD1 And Not chkFD2) And chkFD3 Then VouchPrnFD = 4
        If chkFD1 And Not chkFD2 And chkFD3 Then VouchPrnFD = 5
        If Not chkFD1 And chkFD2 And chkFD3 Then VouchPrnFD = 6
        If chkFD1 And chkFD2 And chkFD3 Then VouchPrnFD = 7
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    bLY = Check1.value
  
    BillInvName = AXCase1.Text
    bCancelExit = False
    bDiscount = chkDiscount.value
    blnBatch = CBool(chkBatch.value)
    Dim lngPos1 As Long
    Dim strValue As String
    strValue = LTrim(CmbChoice.Text)
    lngPos1 = InStr(1, strValue, " ")
    If lngPos1 > 0 Then
        VTID = val(Mid(strValue, 1, lngPos1))
    Else
         VTID = val(strValue)
    End If
    strCardNum = GetCardNumByVtid(CStr(VTID))
    Unload Me
End Sub

Private Sub Command2_Click()
    bFirst = False

    Unload Me
End Sub

Private Sub Form_Load()
    Dim recTemp As New ADODB.Recordset
    Dim strsql As String
'    Dim clsAuth As New U8RowAuthsvr.clsRowAuth
    Dim strDJAuth As String
    ''Dim i  As Integer
   
    On Error Resume Next
        
    '************************
    'u861多语Layout
'    Call MultiLangLoadForm(Me, "u8.sa.xsglsql." & Me.Name)
    '************************
    
'    If clsAuth.Init(DBConn.ConnectionString, m_Login.cUserId) = False Then
'        MsgBox GetString("U8.SA.xsglsql.billprn.00618"), vbExclamation 'zh-CN：权限初始化失败
'        bCancelExit = True
'        Unload Me
'        Exit Sub
'    End If
    If strVouchtype = "27" Or strVouchtype = "26" Then
        chkFD3.Visible = True
    End If
        
    strDJAuth = clsSAWeb.clsAuth.getAuthString("DJMB", , "R")

    
    CmbChoice.Clear
    recTemp.CursorLocation = adUseClient
    If strDJAuth = "1=2" Then
        MsgBox GetString("U8.SA.xsglsql.billprn.00619"), vbExclamation 'zh-CN：您没有打印模版权限
        bCancelExit = True
        Unload Me
        Exit Sub
    Else
        If InStr(1, LCase(strCardNum), LCase("VT_CardNumber")) > 0 Then
            strsql = "SELECT VT_ID, VT_Name, VT_CardNumber, VT_TemplateMode,VT_CardNumber From VoucherTemplates " & _
                    " WHERE " & strCardNum & _
                    " AND VT_TemplateMode=1" & _
                    IIf(strDJAuth = "", "", " AND VT_ID in (" & strDJAuth & ")")
        Else
            strsql = "SELECT VT_ID, VT_Name, VT_CardNumber, VT_TemplateMode,VT_CardNumber From VoucherTemplates " & _
                    " WHERE VT_CardNumber='" & strCardNum & "'" & _
                    " AND VT_TemplateMode=1" & _
                    IIf(strDJAuth = "", "", " AND VT_ID in (" & strDJAuth & ")")
        End If
    End If
    recTemp.Open ConvertSQLString(strsql), DBConn, adOpenForwardOnly, adLockReadOnly
    recTemp.Save domVtid, adPersistXML
    
    If recTemp.RecordCount = 0 Then
       MsgBox GetString("U8.SA.xsglsql.billprn.00620"), vbInformation 'zh-CN：没有当前单据所对应的打印模板！
        bCancelExit = True
       Unload Me
       Exit Sub
    End If
'    CmbChoice.Text = recTemp!VT_ID & Space(20) & recTemp!VT_Name
'    CmbChoice.AddItem   U8PubGetVTID.
    
    Do While Not recTemp.EOF
        CmbChoice.AddItem recTemp!VT_ID & Space(20) & recTemp!VT_Name
        If recTemp("vt_id") = VTID Then
            CmbChoice.ListIndex = CmbChoice.ListCount - 1
'            strCardNum = GetCardNumByVtid(CStr(VTID))
        End If
        recTemp.MoveNext
    Loop
    If CmbChoice.ListIndex = -1 Then CmbChoice.ListIndex = 0
    If iType = 1 Then Chk2.Visible = False
    AXCase1.Text = ""
    AXCase1.Enabled = False
    Chk1.Enabled = False
    bCancelExit = True
    bDiscount = False
    bFirst = False
    chkBatch.Visible = False
    Dim ctl As Control
    If Not (strVouchtype = "05" Or strVouchtype = "06" Or strVouchtype = "26" Or strVouchtype = "27" Or strVouchtype = "28" Or strVouchtype = "29") Then
        For Each ctl In Me.Controls
            ctl.Enabled = False
        Next
        Label1.Enabled = True
        CmbChoice.Enabled = True
        Command1.Enabled = True
        Command2.Enabled = True
        framePrn.Enabled = True
    End If
    
'    Set Command1.Picture = frmMain.imgBmp.ListImages(86).Picture
'    Set Command2.Picture = frmMain.imgBmp.ListImages(87).Picture
    

    
End Sub
Private Function GetCardNumByVtid(strVtid As String) As String
    Dim ele As IXMLDOMElement
    Set ele = domVtid.selectSingleNode("//z:row[@VT_ID='" + strVtid + "']")
    GetCardNumByVtid = ""
    If Not ele Is Nothing Then
        GetCardNumByVtid = ele.Attributes.getNamedItem("VT_CardNumber").nodeValue
    End If
End Function
Private Sub Form_Unload(Cancel As Integer)
'    Call U861ResEnd
    Set domVtid = Nothing
End Sub

Private Sub op1_Click()
    On Error Resume Next
        AXCase1.Text = ""
        AXCase1.Enabled = False
        Chk2.Enabled = True
        Chk1.Enabled = False
        chkFD2.Enabled = True
        chkFD1.Enabled = True
        chkFD3.Enabled = True
End Sub

Private Sub op2_Click()
    On Error Resume Next
  
    AXCase1.Enabled = True
''    Chk1.Enabled = False
    Chk1.Enabled = True
  
    Chk2.value = 0
    Chk2.Enabled = False
    chkFD2.Enabled = False
    chkFD1.Enabled = False
    chkFD3.Enabled = False
End Sub

Private Sub op3_Click()
    On Error Resume Next
  
    AXCase1.Text = ""
    AXCase1.Enabled = False
    Chk1.Enabled = True
  
    Chk2.value = 0
    Chk2.Enabled = False
    If Op3.value Then
        chkBatch.Visible = True
    Else
        chkBatch.Visible = False
    End If

    chkFD2.Enabled = False
    chkFD1.Enabled = False
    chkFD3.Enabled = False
End Sub

Private Sub op4_Click()
    On Error Resume Next
  
    Chk2.value = 0
    Chk2.Enabled = False
  
    If myinfo.bStockClass Then
        AXCase1.Text = ""
        AXCase1.Enabled = False
        Chk1.Enabled = True
    Else
        MsgBox GetString("U8.SA.xsglsql.billprn.00621"), vbExclamation 'zh-CN：存货不分级！
        Op3.value = True
    End If
    chkFD2.Enabled = False
    chkFD1.Enabled = False
    chkFD3.Enabled = False
End Sub


Public Property Get blnBatch() As Boolean
    blnBatch = m_blnBatch
End Property

Public Property Let blnBatch(ByVal vNewValue As Boolean)
    m_blnBatch = vNewValue
End Property

Public Property Get VTID() As Long
    VTID = m_VTID
End Property

Public Property Let VTID(ByVal vNewValue As Long)
    m_VTID = vNewValue
End Property

