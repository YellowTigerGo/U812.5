VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{A0C292A3-118E-11D2-AFDF-000021730160}#1.0#0"; "UFEDIT.OCX"
Object = "{9ADF72AD-DDA9-11D1-9D4B-000021006D51}#1.31#0"; "UFSpGrid.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5AD81966-3173-4597-A32E-4F4620DA3B57}#3.8#0"; "U8TBCtl.ocx"
Object = "{201FB79B-5556-47A4-AD9C-A46BA0C45A44}#6.25#0"; "UFToolBarCtrl.ocx"
Object = "{5E4640D0-A415-404B-A457-72980C429D2F}#10.25#0"; "U8RefEdit.ocx"
Begin VB.Form frmZD1 
   BackColor       =   &H80000005&
   Caption         =   "总帐制单"
   ClientHeight    =   5655
   ClientLeft      =   3075
   ClientTop       =   3630
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   8175
   WindowState     =   2  'Maximized
   Begin MsSuperGrid.SuperGrid Grid 
      Height          =   1725
      Left            =   420
      TabIndex        =   11
      Top             =   3270
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   3043
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EditBorderStyle =   0
      Redraw          =   1
      MouseIcon       =   "frmZD1.frx":0000
      ForeColorSel    =   -2147483634
      ForeColorFixed  =   -2147483630
      FixedCols       =   0
      BackColorSel    =   -2147483635
      BackColorFixed  =   -2147483633
      AllowBigSelection=   0   'False
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7350
      Top             =   3060
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6720
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   360
      ScaleHeight     =   1035
      ScaleWidth      =   7365
      TabIndex        =   0
      Top             =   720
      Width           =   7365
      Begin U8Ref.RefEdit txtDate 
         Height          =   300
         Left            =   3630
         TabIndex        =   13
         Top             =   690
         Width           =   1815
         _ExtentX        =   3201
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
      Begin VB.ComboBox cboSign 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1080
         TabIndex        =   9
         Top             =   690
         Width           =   1545
      End
      Begin VB.CommandButton cmdViewCal 
         Height          =   300
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   690
         Visible         =   0   'False
         Width           =   300
      End
      Begin EDITLib.Edit txtDate1 
         Height          =   300
         Left            =   3630
         TabIndex        =   3
         Top             =   690
         Visible         =   0   'False
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   529
         _StockProps     =   253
         ForeColor       =   0
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         Caption         =   "共 0 条"
         Height          =   180
         Left            =   6405
         TabIndex        =   6
         Top             =   750
         Width           =   630
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         Caption         =   "制单"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3120
         TabIndex        =   4
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         Caption         =   "制单日期"
         Height          =   180
         Left            =   2820
         TabIndex        =   2
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         Caption         =   "凭证类别"
         Height          =   180
         Left            =   180
         TabIndex        =   1
         Top             =   720
         Width           =   720
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   979
      ButtonWidth     =   820
      ButtonHeight    =   926
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "全选"
            Key             =   "All"
            Description     =   "All"
            Object.ToolTipText     =   "全选"
            Object.Tag             =   "All"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "全消"
            Key             =   "None"
            Description     =   "None"
            Object.ToolTipText     =   "全消"
            Object.Tag             =   "None"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "合并"
            Key             =   "Add"
            Description     =   "Add"
            Object.ToolTipText     =   "合并"
            Object.Tag             =   "Add"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查询"
            Key             =   "Find"
            Description     =   "Find"
            Object.ToolTipText     =   "查询"
            Object.Tag             =   "Find"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "制单"
            Key             =   "ZD"
            Description     =   "ZD"
            Object.ToolTipText     =   "制单"
            Object.Tag             =   "ZD"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "单据"
            Key             =   "DJ"
            Description     =   "DJ"
            Object.ToolTipText     =   "单据"
            Object.Tag             =   "DJ"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "摘要"
            Key             =   "digest"
            Description     =   "digest"
            Object.ToolTipText     =   "摘要"
            Object.Tag             =   "digest"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "凭证"
            Key             =   "Auto"
            Description     =   "Auto"
            Object.ToolTipText     =   "凭证"
            Object.Tag             =   "Auto"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "标记"
            Key             =   "Mark"
            Description     =   "Mark"
            Object.ToolTipText     =   "标记"
            Object.Tag             =   "Mark"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "删除"
            Key             =   "Cancel"
            Object.ToolTipText     =   "删除"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin UFToolBarCtrl.UFToolbar UFToolbar1 
      Height          =   240
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   7935
      _ExtentX        =   13996
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
   Begin prjTBCtrl.CTBCtrl CTBCtrl1 
      Height          =   660
      Left            =   2280
      Top             =   2310
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   1164
   End
   Begin VB.Label UFFrmCaptionMgr 
      Caption         =   "制单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6840
      TabIndex        =   10
      Top             =   3900
      Width           =   960
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   6840
      TabIndex        =   8
      Top             =   3600
      Width           =   975
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   2760
      Top             =   1440
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
End
Attribute VB_Name = "frmZD1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strPz As String, pStyle As String
Public StrPz1 As String, pStyle1 As String
Public blnPZSearch As Boolean
Public sMsgTitle As String
Private WithEvents ARPZ As ZzPz.clsPZ
Attribute ARPZ.VB_VarHelpID = -1

'by lg070314 增加U870支持
Private m_Cancel As Integer
Private m_UnloadMode As Integer
 
Dim vfd As Object
Dim sGuid As String
Private WithEvents m_mht As UFPortalProxyMessage.IMessageHandler
Attribute m_mht.VB_VarHelpID = -1

Public Property Let strsguid(ByVal Str As String)
    sGuid = Str
End Property

Public Property Get strsguid() As String
   strsguid = sGuid
End Property

Public Property Let Object_vfd(ByVal Obj As Object)
    Set vfd = Obj
End Property

Public Property Get Object_vfd() As Object
   Set Object_vfd = vfd
End Property

'by lg070314 增加U870支持
'修改3 每个窗体都需要这个方法。Cancel与UnloadMode的参数的含义与QueryUnload的参数相同
'请在此方法中调用窗体Exit(退出)方法，并将设置窗体Unload事件参数(如Cancel)的值同时传给此方法的参数
Public Sub ExitForm(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    
    Cancel = m_Cancel
    UnloadMode = m_UnloadMode
End Sub



Private Sub beginZD()
    Dim CtmpZDList As clsZDList, CTmpZD As clsZD
    Dim dTmp As String
    Dim dTmp1 As String
    Dim bHide As Boolean
    Dim bRet As Boolean
    Dim i As Long
    
    Dim sVchCode As String
    
    If cboSign.ListCount = 0 Then
        Msg "请先设置凭证类别。", vbInformation
        Exit Sub
    End If
        
    If DateCheck(txtDate.Text) = "" Then
        Msg "制单日期非法！", vbInformation
        txtDate.SetFocus
        Exit Sub
    End If
    dTmp = DateCheck(txtDate.Text)
        
    With Grid
        For i = 1 To .Rows - 1
            If .RowHeight(i) > 0 And Trim(.TextMatrix(i, 0)) <> "" Then
                dTmp1 = DateCheck(.TextMatrix(i, 3))
                If dTmp < dTmp1 Then
                    Msg "制单日期不能小于单据日期！", vbCritical
                    txtDate.SetFocus
                    Exit Sub
                End If
            End If
'            If Trim(.TextMatrix(i, 0)) <> "" And Trim(.TextMatrix(i, 13)) = "" Then
'
'
'                 Msg "第" & i & "行 科目不能为空，请检查！"
'                 Exit Sub
'            End If
        Next i
    End With
            
    Dim itotal As Double
    itotal = 0
    Set CtmpZDList = New clsZDList
         
        With Grid
            sVchCode = ""
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, 0)) = "-" Then .TextMatrix(i, 0) = ""
                
                If .RowHeight(i) > 0 And Trim(.TextMatrix(i, 0)) <> "" Then
                    .row = i
                    .col = 2
                    'by ahzzd 060825
                    'If .CellForeColor(i, 2) = &H8000000F Then
                    If .CellForeColor = &H8000000F Then
                        bHide = True
                        .TextMatrix(i, 0) = ""
                    Else
                        itotal = itotal + val(.TextMatrix(i, 12) & "")
                        sVchCode = sVchCode & IIf(sVchCode = "", "", ",") & "'" & VBA.Replace(.TextMatrix(i, 2), "'", "''") & "'"
                    End If
                End If
            Next i
            AddPzValue CtmpZDList, sVchCode, dTmp
        End With
        If itotal = 0 Then
            Msg "所选单据金额合计为0，不能制单！", vbExclamation:       GoTo noerrExit
        End If
        If CtmpZDList.Count > 0 Then
            bRet = AP_ZD(CtmpZDList)
            For i = 1 To Grid.Rows - 1
                Grid.TextMatrix(i, 0) = ""
            Next
        Else
            If bHide Then
                Msg "隐藏单据不需要制单！", vbExclamation
            Else
              If frmZdCX.bShowForm = False Then
                  Msg "单据已生成凭证不能再生成凭证了！", vbExclamation
              Else
                  Msg "请选择要制单的单据！", vbExclamation
              End If
            End If
        End If
noerrExit:
        Set CtmpZDList = Nothing
        Screen.MousePointer = vbDefault
End Sub

Private Sub SelectAll()
    Dim Count   As Long
    For Count = 1 To Grid.Rows - 1
        Grid.TextMatrix(Count, 0) = "1"
    Next Count
End Sub

Public Sub StrZDSingle()
    SelectAll
    beginZD
End Sub


Private Sub cboSign_click()
    If cboSign.ListCount = 0 Then Exit Sub
    Dim i As Long
    With Grid
        For i = 1 To .Rows - 1
            .TextMatrix(i, 1) = cboSign.Text
        Next
    End With
End Sub

 

Private Sub grid_BeforeEdit(Cancel As Boolean, sReturnText As String)
    If Grid.col = 0 Then Cancel = True
End Sub



'Private Sub UFToolbar1_OnCommand(ByVal enumType As prjTBCtrl.ENUM_MENU_OR_BUTTON, ByVal cButtonId As String, ByVal cMenuId As String)
'    Dim Button As MSComctlLib.Button
'    Set Button = Toolbar1.buttons(cButtonId)
'    Toolbar1_ButtonClick Button
'End Sub

Private Sub UFToolbar1_OnCommand(ByVal enumType As UFToolBarCtrl.ENUM_MENU_OR_BUTTON, ByVal cButtonId As String, ByVal cmenuid As String)
    OnButtonClick IIf(enumType = enumButton, cButtonId, cmenuid)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Call WriteResourceLog
End Sub

'在多语环境(繁体，web)下会导致多语控件乱码的情况。利用ufcommandbutton的event实现
Private Sub UFKeyHookCtrl1_ContainerKeyUp(KeyCode As Integer, ByVal Shift As Integer)
    Dim TmpButton As MSComctlLib.Button
    
    Select Case KeyCode
        Case vbKeyF3
            If Shift = 2 Then
                Exit Sub
            Else
                Set TmpButton = Toolbar1.buttons("filtersetting")
            End If
        Case Else
            Exit Sub
    End Select
    If TmpButton.Enabled = False Or TmpButton.Visible = False Then Exit Sub
'    Call Toolbar1_ButtonClick(TmpButton)
    Set TmpButton = Nothing
End Sub
 

Private Sub Form_Load()
    Dim i As Long, j As Long
    Dim cSql As String, cCond As String
'    Dim rst As New UfDbKit.UfRecordset
    Dim rst As New ADODB.Recordset
   ' Grid.FormatString = "^  选择标志  |^ 凭 证 类 别  |^ 业 务 类 型  |^  业 务 描 述  |^  处 理 号  |^   日    期  |^  原 币 金 额  |^  本 币 金 额  |^  汇  率  |^  使 用 部 门  |^  使 用 人  |^  项 目 大 类  |^  项  目   |^  科  目  |^  ID   " 'GetResStringNoParam("U8.CW.APAR.ARAPMain.frmZD1_SuperGrid1_FormatString") '"^  选择标志  |^  凭证类别  |^  业务类型  |^  处理类型  |^  处 理 号  |^  单据类型编码  |^  单据类型  |^  单 据 号  |^  日    期  |^  客户名称  |^  部    门  |^  业 务 员  |^  金    额  |^对应单据类型|^对应单据号|^排序列"
    'by lg070314增加U870菜单融合功能
    ''''''''''''''''''''''''''''''''''''''
 
    Me.HelpContextID = 20090407
    
    sGuid = CreateGUID()
    On Error Resume Next
    If Not (g_business Is Nothing) Then
         Set vfd = g_business.CreateFormEnv(sGuid, frmZD1)
    End If
    
    If Not g_business Is Nothing Then
        Set Me.UFToolbar1.Business = g_business
    End If
    Call RegisterMessage
    
    'by增加 870功菜单功能  zzd 0324
    AddButtons
    'Grid.FormatString = "^  选择标志  |^ 凭 证 类 别  |^ 业 务 类 型  |^  业 务 描 述  |^  处 理 号  |^   日    期  |^  人民币金额  |^  美元金额  |^  汇  率  |^  使 用 部 门  |^  使 用 人  |^  项 目 大 类  |^  项  目   |^  科  目  |^  ID   " 'GetResStringNoParam("U8.CW.APAR.ARAPMain.frmZD1_SuperGrid1_FormatString") '"^  选择标志  |^  凭证类别  |^  业务类型  |^  处理类型  |^  处 理 号  |^  单据类型编码  |^  单据类型  |^  单 据 号  |^  日    期  |^  客户名称  |^  部    门  |^  业 务 员  |^  金    额  |^对应单据类型|^对应单据号|^排序列"
    Grid.FormatString = "^  选择标志  |^ 凭 证 类 别  |^单据编号|^单据日期|^期间|^附单据数|^单据类型|^摘要|^图书编码|^图书名称|^版次|^印次|^总金额"
    DoForm Me, 2
    Call menu_refurbish
    txtDate.Property = EditDate
    txtDate.Text = Format(m_Login.CurDate, "yyyy-mm-dd")
    With Grid
        .Rows = 1: .cols = 14
        .ColAlignment(0) = 4
        .ColAlignment(1) = 4
        .ColAlignment(2) = 4
        .ColAlignment(3) = 4
        .ColAlignment(4) = 4
        .ColAlignment(5) = 1
        .ColAlignment(6) = flexAlignRightCenter
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(8) = flexAlignRightCenter
        .ColAlignment(9) = 1
        .ColAlignment(10) = 1
        .ColAlignment(11) = 1
        .ColAlignment(12) = flexAlignRightCenter

        .colwidth(0) = 900
        .colwidth(1) = 1200
        .colwidth(2) = 1200
        .colwidth(3) = 1200
        .colwidth(4) = 1200
        .colwidth(5) = 1200
        .colwidth(6) = 1200
        .colwidth(7) = 1300
        .colwidth(8) = 1200
        .colwidth(9) = 1200  '
        .colwidth(10) = 1200 '
        .colwidth(11) = 1200 '
        .colwidth(12) = 1200 '
        .colwidth(13) = 0 '
        
        .SetColProperty 0, 4, BrowNull, EditLng, , ""

        For i = 0 To .cols - 1
            .FixedAlignment(i) = 4
        Next i
    End With
    '装入凭证类别
    cboSign.Clear
    If rst.State <> 0 Then rst.Close
    rst.Open "select * from dsign order by iSignSeq", DBConn, adOpenStatic, adLockReadOnly
    If rst.EOF And rst.BOF Then
        rst.Close
        Set rst = Nothing
    Else
        With rst
            .MoveFirst
            Do While Not .EOF
                cboSign.AddItem !ctext
                .MoveNext
            Loop
            .Close
        End With
       cboSign.ListIndex = 0
    End If
    
    Set rst = Nothing
    DoGrid Me.Grid
    Picture1.BackColor = &HFFFFFF
    Picture1.BorderStyle = 0
    If rst.State <> 0 Then rst.Close
    Set rst = Nothing
    
    

    ChangeOneFormTbr Me, Me.Toolbar1, Me.UFToolbar1
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'     clsTask.TaskEnd oAcc.SysId + "0402"
'     If oAcc.SysId = "AP" Then Call clsTask.TaskEnd("AP0403")
End Sub

Public Sub Form_Resize()
    On Error Resume Next
'    frmZD1.Label1.Caption = frmZD1_Label1_Caption
'    frmZD1.Caption = Me_Caption
    Me.UFToolbar1.Top = 0
    Me.UFToolbar1.Width = Me.ScaleWidth
    With Picture1
        .Top = Toolbar1.Height + Toolbar1.Top - 10
        .Left = 0
        .Width = Me.Width
    End With
    
    With Grid
        .Top = Picture1.Top + Picture1.Height
        .Left = Screen.TwipsPerPixelX * 8
        .Width = Me.Width - Screen.TwipsPerPixelY * 25
        If Me.Width > 3000 Then
            .Height = Me.Height - Grid.Top - Screen.TwipsPerPixelY * 40
        End If
    End With
End Sub

Private Sub Picture1_Resize()
    On Error Resume Next
'    lblTitle.Top = (Picture1.Height - lblTitle.Height) / 2
    lblTitle.Left = (Picture1.ScaleWidth - lblTitle.Width) / 2
    
    Call AdjustTypePosition         '调整查询项目内容後相关的位置变化
    Call AdjustPeriodPosition       '调整制单日期内容後相关的位置变化
    Call AdjustLabel1Position       '调整条数信息後相关的位置变化
End Sub

Private Sub lblType_Change()
    '凭证类别
    Call AdjustTypePosition         '调整查询项目内容後相关的位置变化
End Sub

Private Sub lblDate_Change()
    '制单日期
    Call AdjustPeriodPosition       '调整制单日期内容後相关的位置变化
End Sub

Private Sub Label1_Change()
    'Label1 共 0 条
    Call AdjustLabel1Position       '调整条数信息後相关的位置变化
End Sub

'调整查询项目内容後相关的位置变化
Private Sub AdjustTypePosition()
    '凭证类别
    lblType.Left = Screen.TwipsPerPixelX * 8
    cboSign.Left = lblType.Left + lblType.Width + Screen.TwipsPerPixelX * 40  'lblType.Width计算有误，需要调整
End Sub

'调整制单日期内容後相关的位置变化
Private Sub AdjustPeriodPosition()
    '制单日期
    lblDate.Left = lblTitle.Left
    txtDate.Left = lblDate.Left + lblDate.Width + Screen.TwipsPerPixelX * 8
    cmdViewCal.Left = txtDate.Left + txtDate.Width + Screen.TwipsPerPixelX * 8 - 300
    cmdViewCal.Width = Screen.TwipsPerPixelX * 30 '300
    'by ahzzd
    cmdViewCal.Width = 300
End Sub

'调整条数信息後相关的位置变化
Private Sub AdjustLabel1Position()
    'Twip模式下，Picture1.ScaleWidth > Picture1.Width
    'Label1 共 0 条
    Label1.Left = Picture1.Left + Picture1.ScaleWidth - Label1.Width - Screen.TwipsPerPixelX * 30
End Sub

Private Sub Grid_Click()
    Static nSort    As Long
    Static bDesc     As Boolean
    
    If Grid.MouseRow = 0 Then
        If Grid.MouseCol = nSort Then
            bDesc = Not bDesc
        Else
            bDesc = False
        End If
        If Grid.MouseCol = 12 Then
            Grid.col = 14
            Grid.ColSel = 15
            If bDesc Then
                Grid.Sort = 4
            Else
                Grid.Sort = 3
            End If
        Else
            Grid.ColSel = Grid.MouseCol
            If bDesc Then
                Grid.Sort = 6
            Else
                Grid.Sort = 5
            End If
        End If
        nSort = Grid.MouseCol
    End If
End Sub

Private Sub Grid_DblClick()
    Dim i As Long
    Dim iMaxNum As Long
    iMaxNum = 0
    If Grid.col <> 0 Then
       Grid.ReadOnly = True
    Else
       Grid.ReadOnly = False
    End If
    If Grid.row < 1 Then Exit Sub
    If Trim(Grid.TextMatrix(Grid.row, 0)) = "" Then
        With Grid
            For i = Grid.FixedRows To Grid.Rows - 1
                If val(Grid.TextMatrix(i, 0)) > iMaxNum Then iMaxNum = val(Grid.TextMatrix(i, 0))
            Next i
        End With
        Grid.TextMatrix(Grid.row, 0) = iMaxNum + 1
    Else
        Grid.TextMatrix(Grid.row, 0) = ""
    End If

End Sub

'Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Private Sub OnButtonClick(strButtonKey As String)
    On Error GoTo errMsg
    Dim Count   As Long, i As Long, j As Long, k As Long
    Dim StrZd   As String
    Dim sMsg    As String
    Dim dTmp As Date, dTmp1 As Date
    Dim bRet As Boolean
    Dim CtmpZDList As clsZDList, CTmpZD As clsZD
    Dim iOldRow As Long, iOldCol As Long
    Dim bHide As Boolean
    Dim rstTemp4 As UfRecordset
    Dim rstTemp1 As UfRecordset
    Dim rsttemp As UfRecordset
    Dim rstTemp5 As UfRecordset
    Dim rstTemp6 As UfRecordset
    Dim rstTemp7 As UfRecordset
    Grid.ProtectUnload
'    Select Case Button.key
    Select Case strButtonKey
        Case "SelectAll" '全选
            For Count = 1 To Grid.Rows - 1
                Grid.TextMatrix(Count, 0) = Trim(Str(Count))
            Next Count
'            MsgBox Grid.rows & Grid.cols
        Case "UnSelectAll" '取消
            For Count = 1 To Grid.Rows - 1
                Grid.TextMatrix(Count, 0) = ""
            Next Count
        Case "consolidation" '合并
            For Count = 1 To Grid.Rows - 1
                Grid.TextMatrix(Count, 0) = "1"
            Next Count
        Case "filter" '过滤
            frmZdCX.blnPZSearch = blnPZSearch
            frmZdCX.Show 1
            
        Case "save_voucher"        '制单
            beginZD
            Call FillZd(Me.pStyle, Me.strPz, "", "")
        Case "show_voucher"
            If Grid.TextMatrix(Grid.row, 2) = "" Then
                MsgBox "请选择单据"
            End If
            frmMain.MenuClick "EFFYGL040201", "EFFYGL04020101", Grid.TextMatrix(Grid.row, 13), 0
            
        Case "digest"
        Case "help"
            Me.SetFocus
            SendKeys "{F1}"
 
        
    End Select
'     MsgBox "2222  Ok    " & Count
    Exit Sub
errMsg:
    On Error Resume Next
    Msg "制单过程出现异常", vbExclamation
    gcAccount.dbSales.Rollback
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF2 Then cmdViewCal.value = True
End Sub

Public Sub menu_refurbish()
    If blnPZSearch = True Then
'        Me.Caption = "凭证查询"
        Me.Toolbar1.buttons("SelectAll").Visible = False   '全选
        Me.Toolbar1.buttons("UnSelectAll").Visible = False   '取消
        Me.Toolbar1.buttons("consolidation").Visible = False    '合并add
        Me.Toolbar1.buttons("save_voucher").Visible = False     '制单zd
 
 
        Me.Toolbar1.buttons("find_voucher").Visible = True     '查询凭证 Auto
        Me.Toolbar1.buttons("del_voucher").Visible = True   '删除凭证Cancel
    Else
'        Me.Caption = "凭证制单"
        Me.Toolbar1.buttons("SelectAll").Visible = True   '全选
        Me.Toolbar1.buttons("UnSelectAll").Visible = True   '取消
        Me.Toolbar1.buttons("consolidation").Visible = True   '合并
        Me.Toolbar1.buttons("save_voucher").Visible = True     '制单zd
 
 
        Me.Toolbar1.buttons("find_voucher").Visible = False    '查询凭证 Auto
        Me.Toolbar1.buttons("del_voucher").Visible = False   '删除凭证Cancel
        
'        全选all  取消none  合并add  查询find
'
'    制单zd   单据dj    删除凭证Cancel   查询凭证 Auto
    End If
        Me.UFToolbar1.RefreshVisible
End Sub

 Private Sub RegisterMessage()
    Set m_mht = New UFPortalProxyMessage.IMessageHandler
    m_mht.MessageType = "DocAuditComplete"
    If Not g_business Is Nothing Then
        Call g_business.RegisterMessageHandler(m_mht)
    End If
End Sub

 
Private Sub AddButtons()
    Dim btnX As MSComctlLib.Button
    With Toolbar1.buttons
        .Clear
        '全选
        Set btnX = .Add(, "SelectAll", strSelectAll, tbrDefault)
'            btnX.image = 314
        btnX.ToolTipText = strSelectAll
        btnX.Description = btnX.ToolTipText
        btnX.Tag = "Select all"
 
        '取消
         Set btnX = .Add(, "UnSelectAll", strUnSelectAll, tbrDefault)
'          btnX.image = 314
         btnX.ToolTipText = strUnSelectAll
         btnX.Description = btnX.ToolTipText
         btnX.Tag = "Select none"
        
        '合并
         Set btnX = .Add(, "consolidation", "合并", tbrDefault)
'          btnX.image = 314
         btnX.ToolTipText = "合并"
         btnX.Description = btnX.ToolTipText
         btnX.Tag = "consolidation"
         
         
         '过滤
         Set btnX = .Add(, "filter", strFilter, tbrDefault)
'          btnX.image = 314
         btnX.ToolTipText = strFilter
         btnX.Description = btnX.ToolTipText
         btnX.Tag = "filter"
         
        '制单
         Set btnX = .Add(, "save_voucher", "制单", tbrDefault)
'          btnX.image = 314
         btnX.ToolTipText = "制单"
         btnX.Description = btnX.ToolTipText
         btnX.Tag = "creating"
         
        '凭证
         Set btnX = .Add(, "find_voucher", "凭证", tbrDefault)
'          btnX.image = 314
         btnX.ToolTipText = "凭证"
         btnX.Description = btnX.ToolTipText
         btnX.Tag = "sum"
         
        '单据
         Set btnX = .Add(, "show_voucher", "单据", tbrDefault)
'          btnX.image = 314
         btnX.ToolTipText = "单据"
         btnX.Description = btnX.ToolTipText
         btnX.Tag = "bill"
         
        '删除
         Set btnX = .Add(, "del_voucher", "删除", tbrDefault)
'          btnX.image = 314
         btnX.ToolTipText = "删除"
         btnX.Description = btnX.ToolTipText
         btnX.Tag = "delete"
         
             
         '刷新
          Set btnX = .Add(, "refresh", strRefresh, tbrDefault)
'          btnX.image = 314
         btnX.ToolTipText = strRefresh
         btnX.Description = btnX.ToolTipText
         btnX.Tag = "refresh"
         
        '帮助
         Set btnX = .Add(, "help", strHelp, tbrDefault)
'          btnX.image = 314
         btnX.ToolTipText = strHelp
         btnX.Description = btnX.ToolTipText
         btnX.Tag = "help"
        
         '退出
'                 Set btnX = .Add(, "strExit", strExit, tbrDefault)
''          btnX.image = 314
'         btnX.ToolTipText = strExit
'         btnX.Description = btnX.ToolTipText
'         btnX.Tag = "strExit"
 
    End With
    Call InitToolbarTag(Me.Toolbar1)
End Sub

Private Sub AddPzValue(CtmpZDList As clsZDList, ByVal sVchCode As String, ByVal dTmp As String)
    Dim CTmpZD As clsZD
    Dim strsql As String: Dim rs As Object
    Dim iInid As Long
    
    iInid = 1
    strsql = "select * from EFFYGL_v_Pcostbudget where ccode in (" & IIf(sVchCode = "", "''", sVchCode) & ") order by ccode"
    Set rs = DBConn.Execute(strsql)
    Do While Not rs.EOF
        Set CTmpZD = New clsZD
        With CTmpZD
            .cPH = "1"
            .cPzlb = PzlbNameToCode(Grid.TextMatrix(1, 1))                                                '凭证类别字
            .dBillDate = dTmp                                 '凭证日期
            .cProcSign = rs("cbillsign") & ""
            .cProcStyle = rs("cbillsign") & ""                                    '处理方式
            .cCode = getAccinformation("EF", "EFFYGL_FyygdcCode")
            '-----------------------------------------------------------------------------------------------
            .md = val(rs("je") & "")                        '借方金额
            .mc = 0                                                      '贷方金额
            .nfrat = "1"                                                 '汇率
            .cExch_Name = "人民币"                                       '币种
            .cPerson_id = rs("cpersoncode") & ""              '业务员编码
            .cname = rs("cpersonname") & ""                    '业务员名称
            .cbill = m_Login.cUserName
             .nc_s = 1
             .nd_s = 0 '制单人
            .cdept_id = rs("cdepcode") & ""                    '部门编码
'            .ccus_id = voucher.headerText("strcvencode")                   '客户编码
            .csup_id = rs("cvencode") & ""                '供应商编码
            .citem_id = "" ' voucher.bodyText(i, "strItemID")                  '项目编码
            .cItem_Class = "" ' voucher.bodyText(i, "strItemClassID")             '项目大类编码
            .cDigest = rs("cdigest") & ""    '摘要=部门+借款人+用途+自定义项1
            .idoc = val(rs("ibillnum") & "")
'                            .cDefine1 = voucher.headerText("cDefine1")                 '自定义项1
'                            .cDefine2 = voucher.headerText("cDefine2")                    '自定义项2
'                            .cDefine3 = voucher.headerText("cDefine3")                     '自定义项3
'                            .cDefine4 = voucher.headerText("cDefine4")                    '自定义项4
'                            .cDefine5 = voucher.headerText("cDefine5")                   '自定义项5
'                            .cDefine6 = voucher.headerText("cDefine6")                  '自定义项6
'                            .cDefine7 = voucher.headerText("cDefine7")                    '自定义项7
'                            .cDefine8 = voucher.headerText("cDefine8")                   '自定义项8
'                            .cDefine9 = voucher.headerText("cDefine9")                    '自定义项9
'                            .cDefine10 = voucher.headerText("cDefine10")                    '自定义项10
'                            .cDefine11 = voucher.headerText("cDefine11")                   '自定义项11
'                            .cDefine12 = voucher.headerText("cDefine12")                   '自定义项12
'                            .cDefine13 = voucher.headerText("cDefine13")                    '自定义项13
'                            .cDefine14 = voucher.headerText("cDefine14")                  '自定义项14
'                            .cDefine15 = voucher.headerText("cDefine15")                   '自定义项15
'                            .cDefine16 = voucher.headerText("cDefine16")                 '自定义项16
            .inid = iInid:  iInid = iInid + 1
            .ibillno_id = 0
            .cCancelNo = VBA.Replace(sVchCode, "'", "") ' rs("ccode") & ""
            .cmergeno = getcMergenobyCode(rs("ccode") & "")
        End With
        CtmpZDList.Add CTmpZD
        Set CTmpZD = Nothing
        rs.MoveNext
    Loop
        
    strsql = "select * from EFFYGL_v_Pcostbudgetlist where ccode in (" & IIf(sVchCode = "", "''", sVchCode) & ") order by ccode,autoid"
    Set rs = DBConn.Execute(strsql)
    Do While Not rs.EOF
        Set CTmpZD = New clsZD
        With CTmpZD                                        '//ID号,用于找到原单据或记录
            .cPH = "1"
            .cPzlb = PzlbNameToCode(Grid.TextMatrix(1, 1))                                                '凭证类别字
            .dBillDate = dTmp                                 '凭证日期
            .cProcSign = rs("cbillsign") & ""
            .cProcStyle = rs("cbillsign") & ""                                    '处理方式
            .dCode = getCcodebyeleMentCode(rs("celementcode") & "")
            .md = 0                                        '借方金额
            .mc = val(rs("imoney") & "")                    '贷方金额
            .nfrat = "1"                                                 '汇率
            .cExch_Name = "人民币"                                       '币种
            .cPerson_id = "" ' voucher.headerText("strLoanID")              '业务员编码
            .cname = "" 'voucher.headerText("strLoan")                   '业务员名称
            .cbill = m_Login.cUserName
             .nc_s = 1
             .nd_s = 0 '制单人
            .cdept_id = "" ' voucher.headerText("strDepartID")                    '部门编码
            .ccus_id = rs("ccuscode") & ""                               '客户编码
            .csup_id = rs("detailcvencode") & ""                               '供应商编码
            .citem_id = "" ' voucher.bodyText(i, "strItemID")                  '项目编码
            .cItem_Class = "" ' voucher.bodyText(i, "strItemClassID")             '项目大类编码
            .cDigest = IIf(rs("cremark") & "" = "", "费用预估单", rs("cremark") & "") '摘要=部门+借款人+用途+自定义项1
'                            .cDefine1 = voucher.headerText("cDefine1")                 '自定义项1
'                            .cDefine2 = voucher.headerText("cDefine2")                    '自定义项2
'                            .cDefine3 = voucher.headerText("cDefine3")                     '自定义项3
'                            .cDefine4 = voucher.headerText("cDefine4")                    '自定义项4
'                            .cDefine5 = voucher.headerText("cDefine5")                   '自定义项5
'                            .cDefine6 = voucher.headerText("cDefine6")                  '自定义项6
'                            .cDefine7 = voucher.headerText("cDefine7")                    '自定义项7
'                            .cDefine8 = voucher.headerText("cDefine8")                   '自定义项8
'                            .cDefine9 = voucher.headerText("cDefine9")                    '自定义项9
'                            .cDefine10 = voucher.headerText("cDefine10")                    '自定义项10
'                            .cDefine11 = voucher.headerText("cDefine11")                   '自定义项11
'                            .cDefine12 = voucher.headerText("cDefine12")                   '自定义项12
'                            .cDefine13 = voucher.headerText("cDefine13")                    '自定义项13
'                            .cDefine14 = voucher.headerText("cDefine14")                  '自定义项14
'                            .cDefine15 = voucher.headerText("cDefine15")                   '自定义项15
'                            .cDefine16 = voucher.headerText("cDefine16")                 '自定义项16
            .inid = iInid:  iInid = iInid + 1
            .ibillno_id = 0
            .cCancelNo = VBA.Replace(sVchCode, "'", "") 'rs("ccode") & ""
            .cmergeno = getcMergenobyCode(rs("ccode") & "")
            If isUnit(.cmergeno) = False Then
                .cmergeno = .cmergeno & "." & Right(rs("autoid") & "", 5)
            End If
        End With
        CtmpZDList.Add CTmpZD
        Set CTmpZD = Nothing
        rs.MoveNext
    Loop
    
    rs.Close: Set rs = Nothing
End Sub

Private Function getCcodebyeleMentCode(ByVal eleMentCode As String) As String
    Dim strsql As String: Dim rs As Object
    
    strsql = "SELECT * FROM EFFYGL_ElementCCodeOption WHERE celementcode='" & VBA.Replace(eleMentCode, "'", "''") & "'"
    Set rs = DBConn.Execute(strsql)
    If Not rs.EOF Then
        getCcodebyeleMentCode = rs("ccode") & ""
    Else
        getCcodebyeleMentCode = ""
    End If
    rs.Close: Set rs = Nothing
End Function

Private Function getcMergenobyCode(ByVal cCode As String) As String
    Dim iRow As Long
    
    For iRow = 1 To Grid.Rows - 1
        If LCase(Grid.TextMatrix(iRow, 2)) = LCase(cCode) Then
            getcMergenobyCode = Grid.TextMatrix(iRow, 0): Exit Function
        End If
    Next
    getcMergenobyCode = 1
End Function

Private Function isUnit(ByVal cID As String) As Boolean
    Dim iRow As Long
    Dim i As Long
    i = 0
    For iRow = 1 To Grid.Rows - 1
        If LCase(Grid.TextMatrix(iRow, 0)) = LCase(cID) Then
            i = i + 1
        End If
    Next
    If i > 1 Then
        isUnit = True
    Else
        isUnit = False
    End If
End Function

Private Sub tmpdelsub()
    Dim CTmpZD As clsZD
    Dim CtmpZDList As clsZDList
    Dim i As Integer
    Dim dTmp As String
                        Set CTmpZD = New clsZD
                        With CTmpZD
'                            .cPH = Grid.TextMatrix(i, 0)                     ' "FA01" & Grid.TextMatrix(i, 14) 'Grid.TextMatrix(i, 0)
'                            .cPzlb = PzlbNameToCode(Grid.TextMatrix(i, 1))
'                            .cProcSign = Grid.TextMatrix(i, 6)                                       '//业务类型
'                            .cProcStyle = "" 'Grid.TextMatrix(i, 3)                                      '//处理方式
'                            .CVouchId = Grid.TextMatrix(i, 2)                                        '//单据号
'                            .dBillDate = dTmp ' Grid.TextMatrix(i, 5)                                        '//单据日期
'                            .md = IIf(Len(Trim(Grid.TextMatrix(i, 12))) = 0, 0, Grid.TextMatrix(i, 12)) '//本币金额
'                            .mc = IIf(Len(Trim(Grid.TextMatrix(i, 12))) = 0, 0, Grid.TextMatrix(i, 12))    '//本币金额
'                            .md_f = IIf(Len(Trim(Grid.TextMatrix(i, 12))) = 0, 0, Grid.TextMatrix(i, 12))  '//原币金额
'                            .mc_f = IIf(Len(Trim(Grid.TextMatrix(i, 12))) = 0, 0, Grid.TextMatrix(i, 12))  '//原币金额
'                            .nfrat = 1 ' IIf(Len(Trim(Grid.TextMatrix(i, 8))) = 0, 0, Grid.TextMatrix(i, 8)) '//汇率
'                            .cDwCode = "" ' Grid.TextMatrix(i, 9)                                             '//部门编码
'                            .cdept_id = "" ' Grid.TextMatrix(i, 9)                                            '//部门编码
'                            .cCode = "" 'Grid.TextMatrix(i, 13)                                             '//科目编码
'                            .cCancelNo = Grid.TextMatrix(i, 2)                                         '//ID号,用于找到原单据或记录

                            .cPH = Grid.TextMatrix(i, 0)
                            .cPzlb = PzlbNameToCode(Grid.TextMatrix(i, 1))                                                '凭证类别字
                            .dBillDate = dTmp                                 '凭证日期
                            .cProcSign = Grid.TextMatrix(i, 6)
                            .cProcStyle = Grid.TextMatrix(i, 6)                                    '处理方式
'
'                    ' 借方科目,取支出范围科目设置--------------------------------------------------------------------------
'                            strsqlcode = "select * from vwkjiokm  where stroperattype='" & voucher.headerText("strTypeCode") & "'"
'                            Set rscode = New ADODB.Recordset
'                            rscode.CursorLocation = adUseClient
'                            If rscode.State = 1 Then rscode.Close
'                            rscode.Open strsqlcode, DBConn, adOpenForwardOnly, adLockOptimistic
'                            If rscode.RecordCount > 0 Then
'                              .cCode = rscode.Fields("strcode")
'                            Else
                              .cCode = ""
'                            End If
                            '-----------------------------------------------------------------------------------------------
                            .md = Grid.TextMatrix(i, 12)                        '借方金额
                            .mc = 0                                                      '贷方金额
                            .nfrat = "1"                                                 '汇率
                            .cExch_Name = "人民币"                                       '币种
                            .cPerson_id = "" ' voucher.headerText("strLoanID")              '业务员编码
                            .cname = "" 'voucher.headerText("strLoan")                   '业务员名称
                            .cbill = m_Login.cUserName
                             .nc_s = 1
                             .nd_s = 0 '制单人
                            .cdept_id = "" ' voucher.headerText("strDepartID")                    '部门编码
'                            .ccus_id = Voucher.headerText("strcvencode")                   '客户编码
'                            .csup_id = Voucher.headerText("strcvencode")                   '供应商编码
                            .citem_id = "" ' voucher.bodyText(i, "strItemID")                  '项目编码
                            .cItem_Class = "" ' voucher.bodyText(i, "strItemClassID")             '项目大类编码
                            .cDigest = "testjf" ' voucher.headerText("strDepart") & voucher.headerText("strLoan") & voucher.headerText("strUsed") & voucher.headerText("cDefine1") '摘要=部门+借款人+用途+自定义项1
'                            .cDefine1 = voucher.headerText("cDefine1")                 '自定义项1
'                            .cDefine2 = voucher.headerText("cDefine2")                    '自定义项2
'                            .cDefine3 = voucher.headerText("cDefine3")                     '自定义项3
'                            .cDefine4 = voucher.headerText("cDefine4")                    '自定义项4
'                            .cDefine5 = voucher.headerText("cDefine5")                   '自定义项5
'                            .cDefine6 = voucher.headerText("cDefine6")                  '自定义项6
'                            .cDefine7 = voucher.headerText("cDefine7")                    '自定义项7
'                            .cDefine8 = voucher.headerText("cDefine8")                   '自定义项8
'                            .cDefine9 = voucher.headerText("cDefine9")                    '自定义项9
'                            .cDefine10 = voucher.headerText("cDefine10")                    '自定义项10
'                            .cDefine11 = voucher.headerText("cDefine11")                   '自定义项11
'                            .cDefine12 = voucher.headerText("cDefine12")                   '自定义项12
'                            .cDefine13 = voucher.headerText("cDefine13")                    '自定义项13
'                            .cDefine14 = voucher.headerText("cDefine14")                  '自定义项14
'                            .cDefine15 = voucher.headerText("cDefine15")                   '自定义项15
'                            .cDefine16 = voucher.headerText("cDefine16")                 '自定义项16
                            .inid = 1
                            .ibillno_id = 0
                            .cCancelNo = Grid.TextMatrix(i, 2)
                        End With
                        CtmpZDList.Add CTmpZD
                        Set CTmpZD = Nothing
                        
                        Set CTmpZD = New clsZD
                        With CTmpZD
'                            .cPH = Grid.TextMatrix(i, 0)                     ' "FA01" & Grid.TextMatrix(i, 14) 'Grid.TextMatrix(i, 0)
'                            .cPzlb = PzlbNameToCode(Grid.TextMatrix(i, 1))
'                            .cProcSign = Grid.TextMatrix(i, 6)                                       '//业务类型
'                            .cProcStyle = "" 'Grid.TextMatrix(i, 3)                                      '//处理方式
'                            .CVouchId = Grid.TextMatrix(i, 2)                                        '//单据号
'                            .dBillDate = dTmp ' Grid.TextMatrix(i, 5)                                        '//单据日期
'                            .md = IIf(Len(Trim(Grid.TextMatrix(i, 12))) = 0, 0, Grid.TextMatrix(i, 12)) '//本币金额
'                            .mc = IIf(Len(Trim(Grid.TextMatrix(i, 12))) = 0, 0, Grid.TextMatrix(i, 12))    '//本币金额
'                            .md_f = IIf(Len(Trim(Grid.TextMatrix(i, 12))) = 0, 0, Grid.TextMatrix(i, 12))  '//原币金额
'                            .mc_f = IIf(Len(Trim(Grid.TextMatrix(i, 12))) = 0, 0, Grid.TextMatrix(i, 12))  '//原币金额
'                            .nfrat = 1 ' IIf(Len(Trim(Grid.TextMatrix(i, 8))) = 0, 0, Grid.TextMatrix(i, 8)) '//汇率
'                            .cDwCode = "" ' Grid.TextMatrix(i, 9)                                             '//部门编码
'                            .cdept_id = "" ' Grid.TextMatrix(i, 9)                                            '//部门编码
'                            .cCode = "" 'Grid.TextMatrix(i, 13)                                             '//科目编码
'                            .cCancelNo = Grid.TextMatrix(i, 2)                                         '//ID号,用于找到原单据或记录

                            .cPH = Grid.TextMatrix(i, 0)
                            .cPzlb = PzlbNameToCode(Grid.TextMatrix(i, 1))                                                '凭证类别字
                            .dBillDate = dTmp                                 '凭证日期
                            .cProcSign = Grid.TextMatrix(i, 6)
                            .cProcStyle = Grid.TextMatrix(i, 6)                                    '处理方式
'
'                    ' 借方科目,取支出范围科目设置--------------------------------------------------------------------------
'                            strsqlcode = "select * from vwkjiokm  where stroperattype='" & voucher.headerText("strTypeCode") & "'"
'                            Set rscode = New ADODB.Recordset
'                            rscode.CursorLocation = adUseClient
'                            If rscode.State = 1 Then rscode.Close
'                            rscode.Open strsqlcode, DBConn, adOpenForwardOnly, adLockOptimistic
'                            If rscode.RecordCount > 0 Then
'                              .cCode = rscode.Fields("strcode")
'                            Else
                              .cCode = " "
'                            End If
                            '-----------------------------------------------------------------------------------------------
                            .md = 0                                        '借方金额
                            .mc = Grid.TextMatrix(i, 12)                     '贷方金额
                            .nfrat = "1"                                                 '汇率
                            .cExch_Name = "人民币"                                       '币种
                            .cPerson_id = "" ' voucher.headerText("strLoanID")              '业务员编码
                            .cname = "" 'voucher.headerText("strLoan")                   '业务员名称
                            .cbill = m_Login.cUserName
                             .nc_s = 1
                             .nd_s = 0 '制单人
                            .cdept_id = "" ' voucher.headerText("strDepartID")                    '部门编码
'                            .ccus_id = Voucher.headerText("strcvencode")                   '客户编码
'                            .csup_id = Voucher.headerText("strcvencode")                   '供应商编码
                            .citem_id = "" ' voucher.bodyText(i, "strItemID")                  '项目编码
                            .cItem_Class = "" ' voucher.bodyText(i, "strItemClassID")             '项目大类编码
                            .cDigest = "testdf" ' voucher.headerText("strDepart") & voucher.headerText("strLoan") & voucher.headerText("strUsed") & voucher.headerText("cDefine1") '摘要=部门+借款人+用途+自定义项1
'                            .cDefine1 = voucher.headerText("cDefine1")                 '自定义项1
'                            .cDefine2 = voucher.headerText("cDefine2")                    '自定义项2
'                            .cDefine3 = voucher.headerText("cDefine3")                     '自定义项3
'                            .cDefine4 = voucher.headerText("cDefine4")                    '自定义项4
'                            .cDefine5 = voucher.headerText("cDefine5")                   '自定义项5
'                            .cDefine6 = voucher.headerText("cDefine6")                  '自定义项6
'                            .cDefine7 = voucher.headerText("cDefine7")                    '自定义项7
'                            .cDefine8 = voucher.headerText("cDefine8")                   '自定义项8
'                            .cDefine9 = voucher.headerText("cDefine9")                    '自定义项9
'                            .cDefine10 = voucher.headerText("cDefine10")                    '自定义项10
'                            .cDefine11 = voucher.headerText("cDefine11")                   '自定义项11
'                            .cDefine12 = voucher.headerText("cDefine12")                   '自定义项12
'                            .cDefine13 = voucher.headerText("cDefine13")                    '自定义项13
'                            .cDefine14 = voucher.headerText("cDefine14")                  '自定义项14
'                            .cDefine15 = voucher.headerText("cDefine15")                   '自定义项15
'                            .cDefine16 = voucher.headerText("cDefine16")                 '自定义项16
                            .inid = 2
                            .ibillno_id = 0
                            .cCancelNo = Grid.TextMatrix(i, 2)
                        End With
                        CtmpZDList.Add CTmpZD
                        Set CTmpZD = Nothing

End Sub

