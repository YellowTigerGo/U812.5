VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{51388549-C886-4FD6-AE5F-8AA28C63CE94}#1.0#0"; "PrintControl.ocx"
Object = "{201FB79B-5556-47A4-AD9C-A46BA0C45A44}#6.25#0"; "UFToolBarCtrl.ocx"
Begin VB.Form frmType 
   Caption         =   "��׼����"
   ClientHeight    =   5895
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8865
   Icon            =   "frmTYPE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   8865
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar STBTimer 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   15
      Top             =   5205
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin PRINTCONTROLLib.PrintControl Prn 
      Height          =   405
      Left            =   930
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   465
      _Version        =   65536
      _ExtentX        =   820
      _ExtentY        =   714
      _StockProps     =   0
      EnableSave      =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   820
      Left            =   0
      ScaleHeight     =   805
      ScaleMode       =   0  'User
      ScaleWidth      =   8865
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   8895
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GSP��׼����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3420
         TabIndex        =   6
         Top             =   210
         Width           =   1650
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         Visible         =   0   'False
         X1              =   1470
         X2              =   2910
         Y1              =   394.906
         Y2              =   394.906
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         Visible         =   0   'False
         X1              =   5790
         X2              =   7230
         Y1              =   394.906
         Y2              =   394.906
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4245
      Left            =   4170
      ScaleHeight     =   4215
      ScaleWidth      =   4695
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1230
      Width           =   4725
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   2130
         TabIndex        =   2
         Top             =   1410
         Width           =   255
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   2130
         TabIndex        =   1
         Top             =   945
         Width           =   2385
      End
      Begin VB.TextBox txtCode 
         Height          =   315
         Left            =   2130
         TabIndex        =   0
         Top             =   330
         Width           =   2385
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   2130
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1935
         Width           =   255
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "�Ƿ���ĩ�ڵ�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   14
         Top             =   1575
         Width           =   1080
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "��׼��������"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   480
         TabIndex        =   13
         Top             =   1005
         Width           =   1080
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "�Ƿ�ϵͳĬ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   12
         Top             =   2100
         Width           =   1080
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "��׼�������"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   480
         TabIndex        =   11
         Top             =   390
         Width           =   1080
      End
   End
   Begin MSComctlLib.StatusBar Stb 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   4
      Top             =   5550
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgClassImageList 
      Left            =   180
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Tlb 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   7
      Top             =   0
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   635
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin MSComctlLib.TreeView TREE1 
      Height          =   4245
      Left            =   0
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1230
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   7488
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin UFToolBarCtrl.UFToolbar CTBCtrl1 
      Height          =   240
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   2655
      _ExtentX        =   4683
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
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuSetup 
         Caption         =   "����(&U)"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "Ԥ��(&V)"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOutput 
         Caption         =   "���(&S)"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuOperation 
      Caption         =   "����(&O)"
      Begin VB.Menu mnuAdd 
         Caption         =   "����(&A)"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuModify 
         Caption         =   "�޸�(&M)"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBack 
         Caption         =   "����(&B)"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "����(&S)"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "����(&W)"
      Visible         =   0   'False
      Begin VB.Menu mnuTBL 
         Caption         =   "�ı���ť(&T)"
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuH 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelp 
         Caption         =   "����"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------
' �� �� ��: frmType.frm
'
' �� ��: GSP����ģ�洰��
'
' �� ��: ������    ʱ  ��:     2002
' ��������������� ��Ȩ���� Copyright(c) 2002
'--------------------------------------------



Option Explicit
'<�¼�������

'���ش���
Event Load()
'Toolbar����¼���sKey=toolbar.buttons().key
Event ButtonClick(sKey As String)
'Tree�ڵ����¼�
Event TreeNodeClick(ByVal Node As MSComctlLib.Node)
'Tree�ڵ�չ���¼�
Event TreeNodeCollapse(ByVal Node As MSComctlLib.Node)
'textbox�ؼ�KeyPress�¼�;txt��ǰ����textbox,asc��ǰ����ascii,lLen�������볤��,objName��ǰtextbox����
Event TxtKeyPress(txt As TextBox, ByRef asc As Integer, lLen As Long, objName As String)
'textbox�ؼ�KeyUp�¼�;txt��ǰ����textbox
Event TxtKeyUp(txt As TextBox)
'textbox�ؼ�Change�¼�;txt��ǰ����textbox,lLen�������볤��
Event TxtChange(txt As TextBox, lLen As Long)
'textbox�ؼ�GetFocus�¼�;txt��ǰ����textbox
Event TxtGetFocus(txt As Control)
'��ǰ�ؼ���ȥ�����¼���objΪ��ǰ����Ŀؼ�
Event LostFocus(obj As Control)
'�����ӡ����
Event SettingChanged(ByVal varLocalSettings As Variant, ByVal varModuleSettings As Variant)
'���Checkox�¼�
Event Check()
'�����FormKey�¼�
Event FormKey(KeyCode As Integer, Shift As Integer)
'�����QuryUnload�¼�
Event FrmQuryExit(ByRef Cancel As Integer)
'/�¼�����>

Private Sub Check2_Click()
    RaiseEvent Check
End Sub

Private Sub CTBCtrl1_OnCommand(ByVal enumType As UFToolBarCtrl.ENUM_MENU_OR_BUTTON, ByVal cButtonId As String, ByVal cMenuId As String)
    RaiseEvent ButtonClick(cButtonId)
End Sub

''Private Sub CTBCtrl1_OnCommand(ByVal enumType As prjTBCtrl.ENUM_MENU_OR_BUTTON, ByVal cButtonId As String, ByVal cMenuId As String)
''    RaiseEvent ButtonClick(cButtonId)
''End Sub

'-----------------------------------------------------------
'���ܣ����弤��ʱ�������沼��
'
'����:
'
'����:
'
'-----------------------------------------------------------
Private Sub Form_Activate()
    Me.KeyPreview = True
'    LoadResPicture 120, vbResIcon
End Sub
'-----------------------------------------------------------
'���ܣ����ͼ���ʼ���Ĺ���
'
'����:
'
'����:
'
'-----------------------------------------------------------
Public Sub InitImageList()
    On Error GoTo Err_Info:
    imgClassImageList.ListImages.Add 1, "Add", LoadResPicture(101, vbResIcon)
    imgClassImageList.ListImages.Add 2, "Delete", LoadResPicture(102, vbResIcon)
    imgClassImageList.ListImages.Add 3, "Modify", LoadResPicture(103, vbResIcon)
    imgClassImageList.ListImages.Add 4, "Back", LoadResPicture(104, vbResIcon)
    imgClassImageList.ListImages.Add 5, "SaveRs", LoadResPicture(105, vbResIcon)
    imgClassImageList.ListImages.Add 6, "Seek", LoadResPicture(106, vbResIcon)
    imgClassImageList.ListImages.Add 7, "Exit", LoadResPicture(107, vbResIcon)
    imgClassImageList.ListImages.Add 8, "Help", LoadResPicture(108, vbResIcon)
    imgClassImageList.ListImages.Add 9, "Refresh", LoadResPicture(109, vbResIcon)
    imgClassImageList.ListImages.Add 10, "Print", LoadResPicture(110, vbResIcon)
    imgClassImageList.ListImages.Add 11, "Preview", LoadResPicture(111, vbResIcon)
    imgClassImageList.ListImages.Add 12, "SaveFile", LoadResPicture(112, vbResIcon)
    imgClassImageList.ListImages.Add 13, "SetUp", LoadResPicture(113, vbResIcon)
    Exit Sub
Err_Info:
    Debug.Print "FrmInvenCls_InitImageList_Error"
End Sub

'-----------------------------------------------------------
'���ܣ���ɹ�������ʼ���Ĺ���
'
'����:
'
'����:
'
'-----------------------------------------------------------
Public Sub InitToolBar()
    On Error GoTo Err_Info
    Tlb.ImageList = imgClassImageList
'    Tlb.TextAlignment = tbrTextAlignRight
    With Tlb.Buttons
        .Add 1, "SetUp", , tbrDefault, "SetUp"
        .Add 2, "Print", , tbrDefault, "Print"
        .Add 3, "Preview", , tbrDefault, "Preview"
        .Add 4, "SaveFile", , tbrDefault, "SaveFile"
        .Add 5, "btnSep1", , tbrSeparator
        .Add 6, "Add", , tbrDefault, "Add"
        .Add 7, "Modify", , tbrDefault, "Modify"
        .Add 8, "Delete", , tbrDefault, "Delete"
        .Add 9, "btnSep2", , tbrSeparator
        .Add 10, "Back", , tbrDefault, "Back"
        .Add 11, "SaveRs", , tbrDefault, "SaveRs"
        .Add 12, "btnSep3", , tbrSeparator
        .Add 13, "Seek", , tbrDefault, "Seek"
        .Add 14, "Refresh", , tbrDefault, "Refresh"
        .Add 15, "btnSep4", , tbrSeparator
        .Add 16, "Help", "����", tbrDefault, "Help"
        .Add 17, "Exit", , tbrDefault, "Exit"
    End With
    Tlb.Buttons("SetUp").ToolTipText = "��ӡ����"
    Tlb.Buttons("SetUp").Caption = "����"
    Tlb.Buttons("Refresh").ToolTipText = "ˢ��"
    Tlb.Buttons("Refresh").Caption = "ˢ��"
    Tlb.Buttons("Print").ToolTipText = "��ӡ"
    Tlb.Buttons("Print").Caption = "��ӡ"
    Tlb.Buttons("Preview").ToolTipText = "��ӡԤ��"
    Tlb.Buttons("Preview").Caption = "Ԥ��"
    Tlb.Buttons("SaveFile").ToolTipText = "���"
    Tlb.Buttons("SaveFile").Caption = "���"
    Tlb.Buttons("Add").ToolTipText = LoadResString(4000)
    Tlb.Buttons("Add").Caption = LoadResString(4000)
    Tlb.Buttons("Delete").ToolTipText = LoadResString(4005)
    Tlb.Buttons("Delete").Caption = LoadResString(4005)
    Tlb.Buttons("Modify").ToolTipText = LoadResString(4010)
    Tlb.Buttons("Modify").Caption = LoadResString(4010)
    Tlb.Buttons("Back").Caption = LoadResString(5001)
    Tlb.Buttons("Back").ToolTipText = LoadResString(5002)
    Tlb.Buttons("SaveRs").ToolTipText = LoadResString(4020)
    Tlb.Buttons("SaveRs").Caption = LoadResString(4020)
    Tlb.Buttons("Seek").ToolTipText = LoadResString(4025)
    Tlb.Buttons("Seek").Visible = False
    Tlb.Buttons("Exit").Caption = LoadResString(5003)
    Tlb.Buttons("Help").Caption = LoadResString(5005)
    
    'new portal �޸�3 ����button.tag
    Tlb.Buttons("SetUp").Tag = CreatePortalToolbarTag("setting", "IDEAL", "PortalToolbar")
    Tlb.Buttons("Refresh").Tag = CreatePortalToolbarTag("Refresh", "ICOMMON", "PortalToolbar")
    Tlb.Buttons("Print").Tag = CreatePortalToolbarTag("Print", "ICOMMON", "PortalToolbar")
    Tlb.Buttons("Preview").Tag = CreatePortalToolbarTag("Print Preview", "ICOMMON", "PortalToolbar")
    Tlb.Buttons("SaveFile").Tag = CreatePortalToolbarTag("Save", "ICOMMON", "PortalToolbar")
    Tlb.Buttons("Add").Tag = CreatePortalToolbarTag("Add", "IEDIT", "PortalToolbar")
    Tlb.Buttons("Delete").Tag = CreatePortalToolbarTag("Delete", "IEDIT", "PortalToolbar")
    Tlb.Buttons("Modify").Tag = CreatePortalToolbarTag("Modify", "IEDIT", "PortalToolbar")
    Tlb.Buttons("Back").Tag = CreatePortalToolbarTag("Cancel", "ICOMMON", "PortalToolbar")
    Tlb.Buttons("SaveRs").Tag = CreatePortalToolbarTag("save as", "ICOMMON", "PortalToolbar")
    Tlb.Buttons("Seek").Tag = CreatePortalToolbarTag("report for check", "ISEARCH", "PortalToolbar")
    Tlb.Buttons("Exit").Tag = CreatePortalToolbarTag("Exit", "ICOMMON", "PortalToolbar")
    Tlb.Buttons("Help").Tag = CreatePortalToolbarTag("Help", "ICOMMON", "PortalToolbar")
    
    SetTooltip Tlb
    Tlb.Buttons("Back").Enabled = False
    Tlb.Buttons("SaveRs").Enabled = False
    ''��ʼ��CTBCtrl1
    If TBLStyle <> TBLNormal Then
        CTBCtrl1.SetToolbar Tlb
        CTBCtrl1.SetDisplayStyle 2
        CTBCtrl1.RefreshVisible
        Tlb.Visible = False
        CTBCtrl1.Visible = True
        CTBCtrl1.ZOrder 0
    Else
        Tlb.Visible = True
        CTBCtrl1.Visible = False
        Tlb.ZOrder 0
    End If
    Exit Sub
Err_Info:
    Debug.Print "FrmInvenCls_InitToolBar_Error"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> 112 Then
        RaiseEvent FormKey(KeyCode, Shift)
'    Else
'        Me.Picture2.SetFocus
    End If
End Sub

Private Sub Form_Load()
    'new portal �޸�2
'    Set CTBCtrl1.Business = g_business
    
    InitImageList
    InitToolBar
    RaiseEvent Load
    CTBCtrl1.BackColor = Picture1.BackColor
    App.HelpFile = App.path & "\�����������.chm"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    RaiseEvent ButtonClick("Exit")
    RaiseEvent FrmQuryExit(Cancel)
    If Cancel = 1 Then Exit Sub
End Sub

Private Sub Form_Resize()
 Resize Me
End Sub



Private Sub mnuAdd_Click()
    RaiseEvent ButtonClick("Add")
End Sub

Private Sub mnuBack_Click()
    RaiseEvent ButtonClick("Back")
End Sub

Private Sub mnuDelete_Click()
    RaiseEvent ButtonClick("Delete")
End Sub

Private Sub mnuExit_Click()
    RaiseEvent ButtonClick("Exit")
End Sub

Private Sub mnuHelp_Click()
    SendMessage Me.hwnd, WM_KEYDOWN, VK_F1, 0
End Sub

Private Sub mnuModify_Click()
    RaiseEvent ButtonClick("Modify")
End Sub

Private Sub mnuOutput_Click()
    RaiseEvent ButtonClick("SaveFile")
End Sub

Private Sub mnuPreview_Click()
    RaiseEvent ButtonClick("Preview")
End Sub

Private Sub mnuPrint_Click()
    RaiseEvent ButtonClick("Print")
End Sub

Private Sub mnuRefresh_Click()
    RaiseEvent ButtonClick("Refresh")
End Sub

Private Sub mnuSave_Click()
    RaiseEvent ButtonClick("SaveRs")
End Sub

Private Sub mnuSetup_Click()
    RaiseEvent ButtonClick("SetUp")
End Sub

Private Sub mnuStatus_Click()
    RaiseEvent ButtonClick("Status")
End Sub

Private Sub mnuTBL_Click()
    RaiseEvent ButtonClick("TBL")
End Sub

Private Sub Prn_SettingChanged(ByVal varLocalSettings As Variant, ByVal varModuleSettings As Variant)
    RaiseEvent SettingChanged(varLocalSettings, varModuleSettings)
End Sub



Private Sub Tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
    RaiseEvent ButtonClick(Button.Key)
End Sub

Private Sub TREE1_Collapse(ByVal Node As MSComctlLib.Node)
    RaiseEvent TreeNodeCollapse(Node)
End Sub

Private Sub TREE1_NodeClick(ByVal Node As MSComctlLib.Node)
    RaiseEvent TreeNodeClick(Node)
End Sub

Private Sub txtCode_Change()
    RaiseEvent TxtChange(txtCode, 20)
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    RaiseEvent TxtKeyPress(txtCode, KeyAscii, 20, "txtCode")
End Sub

Private Sub txtCode_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent TxtKeyUp(txtCode)
End Sub

Private Sub txtCode_LostFocus()
    RaiseEvent LostFocus(txtCode)
End Sub

Private Sub txtName_Change()
    RaiseEvent TxtChange(txtName, 50)
End Sub

Private Sub txtName_GotFocus()
    RaiseEvent TxtGetFocus(txtName)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    RaiseEvent TxtKeyPress(txtName, KeyAscii, 50, "txtName")
End Sub

Private Sub txtName_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent TxtKeyUp(txtName)
End Sub

Private Sub txtName_LostFocus()
    RaiseEvent LostFocus(txtName)
End Sub

'new portal�޸�4
Public Sub ExitForm(Cancel As Integer, UnloadMode As Integer)
    
'    RaiseEvent ButtonClick("Exit")
    RaiseEvent FrmQuryExit(Cancel)
    If Cancel = 1 Then
        Exit Sub
    Else
        Unload Me
    End If
End Sub


