VERSION 5.00
Object = "{A0C292A3-118E-11D2-AFDF-000021730160}#1.0#0"; "UFEDIT.OCX"
Object = "{5E4640D0-A415-404B-A457-72980C429D2F}#10.25#0"; "U8RefEdit.ocx"
Begin VB.Form frmZdCX 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "制单查询"
   ClientHeight    =   4860
   ClientLeft      =   945
   ClientTop       =   3000
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   420
      Left            =   3000
      TabIndex        =   48
      Top             =   3960
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   420
      Left            =   585
      TabIndex        =   47
      Top             =   3960
      Width           =   1635
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   34
      Top             =   240
      Width           =   4935
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3585
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4995
      Begin U8Ref.RefEdit edtDate2 
         Height          =   300
         Left            =   3405
         TabIndex        =   50
         Top             =   2940
         Width           =   1455
         _ExtentX        =   2566
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
      Begin U8Ref.RefEdit edtDate1 
         Height          =   300
         Left            =   1005
         TabIndex        =   49
         Top             =   2940
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.CommandButton cmdContractType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   2280
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CommandButton cmdContractID1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1320
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CommandButton cmdContractID2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1320
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.ComboBox cboCheckMan 
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3420
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1980
         Width           =   1485
      End
      Begin VB.ComboBox cboSaleType 
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2340
         Width           =   1455
      End
      Begin VB.ComboBox cboOperator 
         Height          =   300
         Left            =   3420
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2340
         Width           =   1485
      End
      Begin VB.CommandButton CmdSStyle 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1980
         Visible         =   0   'False
         Width           =   300
      End
      Begin EDITLib.Edit EdtSStyle 
         Height          =   300
         Left            =   1020
         TabIndex        =   11
         Top             =   1980
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   529
         _StockProps     =   253
         ForeColor       =   0
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin VB.ComboBox cmbBZ 
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3420
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   180
         Width           =   1455
      End
      Begin VB.CommandButton cmdPsn 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   540
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CommandButton cmdDept 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   540
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CommandButton cmdDate2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2940
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CommandButton cmdDate1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2940
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CommandButton cmdDw 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   180
         Visible         =   0   'False
         Width           =   300
      End
      Begin EDITLib.Edit edtPsn 
         Height          =   300
         Left            =   3420
         TabIndex        =   4
         Top             =   540
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   529
         _StockProps     =   253
         ForeColor       =   0
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin EDITLib.Edit edtDept 
         Height          =   300
         Left            =   1020
         TabIndex        =   3
         Top             =   540
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   529
         _StockProps     =   253
         ForeColor       =   0
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin EDITLib.Edit edtAmount2 
         Height          =   300
         Left            =   3420
         TabIndex        =   10
         Top             =   1620
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   529
         _StockProps     =   253
         ForeColor       =   0
         BackColor       =   16777215
         Appearance      =   1
         Property        =   4
         NumPoint        =   2
      End
      Begin EDITLib.Edit edtAmount1 
         Height          =   300
         Left            =   1020
         TabIndex        =   9
         Top             =   1620
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   529
         _StockProps     =   253
         ForeColor       =   0
         BackColor       =   16777215
         Appearance      =   1
         Property        =   4
         NumPoint        =   2
      End
      Begin EDITLib.Edit edtNum2 
         Height          =   300
         Left            =   3420
         TabIndex        =   6
         Top             =   900
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   529
         _StockProps     =   253
         ForeColor       =   0
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin EDITLib.Edit edtDate21 
         Height          =   300
         Left            =   3420
         TabIndex        =   8
         Top             =   2940
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   529
         _StockProps     =   253
         ForeColor       =   0
         BackColor       =   16777215
         Appearance      =   1
         Property        =   5
         MaxLength       =   30
      End
      Begin EDITLib.Edit edtDate11 
         Height          =   300
         Left            =   1020
         TabIndex        =   7
         Top             =   2940
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   529
         _StockProps     =   253
         ForeColor       =   0
         BackColor       =   16777215
         Appearance      =   1
         Property        =   5
         MaxLength       =   30
      End
      Begin EDITLib.Edit edtNum1 
         Height          =   300
         Left            =   1020
         TabIndex        =   5
         Top             =   900
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   529
         _StockProps     =   253
         ForeColor       =   0
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin EDITLib.Edit edtDw 
         Height          =   300
         Left            =   1020
         TabIndex        =   1
         Top             =   180
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   529
         _StockProps     =   253
         ForeColor       =   0
         BackColor       =   16777215
         Appearance      =   1
         MaxLength       =   150
      End
      Begin EDITLib.Edit txtOrderCode2 
         Height          =   300
         Left            =   3360
         TabIndex        =   16
         Top             =   2040
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   529
         _StockProps     =   253
         ForeColor       =   0
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin EDITLib.Edit txtOrderCode1 
         Height          =   300
         Left            =   1080
         TabIndex        =   15
         Top             =   2160
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   529
         _StockProps     =   253
         ForeColor       =   0
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin EDITLib.Edit txtContractType 
         Height          =   300
         Left            =   2160
         TabIndex        =   17
         Top             =   1920
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   529
         _StockProps     =   253
         ForeColor       =   0
         BackColor       =   16777215
         Appearance      =   1
         MaxLength       =   100
      End
      Begin EDITLib.Edit txtContractID1 
         Height          =   300
         Left            =   1020
         TabIndex        =   18
         Top             =   1320
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   529
         _StockProps     =   253
         ForeColor       =   0
         BackColor       =   16777215
         Appearance      =   1
         MaxLength       =   100
      End
      Begin EDITLib.Edit txtContractID2 
         Height          =   300
         Left            =   3420
         TabIndex        =   19
         Top             =   1320
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   529
         _StockProps     =   253
         ForeColor       =   0
         BackColor       =   16777215
         Appearance      =   1
         MaxLength       =   100
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "合同类型"
         Height          =   180
         Left            =   2400
         TabIndex        =   44
         Top             =   1920
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "订单号"
         Height          =   180
         Left            =   360
         TabIndex        =   40
         Top             =   2160
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "审核人"
         Height          =   180
         Left            =   2670
         TabIndex        =   39
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "销售类型"
         Height          =   180
         Left            =   180
         TabIndex        =   38
         Top             =   2400
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "制单人"
         Height          =   180
         Left            =   2670
         TabIndex        =   37
         Top             =   2400
         Width           =   540
      End
      Begin VB.Label LblSStyle 
         AutoSize        =   -1  'True
         Caption         =   "结算方式"
         Height          =   180
         Left            =   180
         TabIndex        =   35
         Top             =   2010
         Width           =   720
      End
      Begin VB.Label lblBZ 
         AutoSize        =   -1  'True
         Caption         =   "币种"
         Height          =   180
         Left            =   2670
         TabIndex        =   33
         Top             =   195
         Width           =   360
      End
      Begin VB.Label lblPsn 
         AutoSize        =   -1  'True
         Caption         =   "业务员"
         Height          =   180
         Left            =   2670
         TabIndex        =   27
         Top             =   555
         Width           =   540
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         Caption         =   "部门"
         Height          =   180
         Left            =   180
         TabIndex        =   26
         Top             =   570
         Width           =   360
      End
      Begin VB.Label lblAmount 
         AutoSize        =   -1  'True
         Caption         =   "金额"
         Height          =   180
         Left            =   180
         TabIndex        =   25
         Top             =   1650
         Width           =   360
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         Caption         =   "单据日期"
         Height          =   180
         Left            =   180
         TabIndex        =   24
         Top             =   2955
         Width           =   720
      End
      Begin VB.Label lblNum 
         AutoSize        =   -1  'True
         Caption         =   "单据号"
         Height          =   180
         Left            =   180
         TabIndex        =   23
         Top             =   915
         Width           =   540
      End
      Begin VB.Label lblDw 
         AutoSize        =   -1  'True
         Caption         =   "客户"
         Height          =   180
         Left            =   180
         TabIndex        =   22
         Top             =   195
         Width           =   360
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "合同号"
         Height          =   180
         Left            =   180
         TabIndex        =   45
         Top             =   1380
         Width           =   540
      End
      Begin VB.Line Line1 
         X1              =   2670
         X2              =   3210
         Y1              =   1043
         Y2              =   1043
      End
      Begin VB.Line Line2 
         X1              =   2670
         X2              =   3210
         Y1              =   3090
         Y2              =   3090
      End
      Begin VB.Line Line3 
         X1              =   2670
         X2              =   3210
         Y1              =   1763
         Y2              =   1763
      End
      Begin VB.Line Line4 
         X1              =   2760
         X2              =   3300
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line5 
         X1              =   2670
         X2              =   3210
         Y1              =   1470
         Y2              =   1470
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3000
      TabIndex        =   21
      Top             =   3960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   600
      TabIndex        =   20
      Top             =   3960
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label UFFrmCaptionMgr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "制单查询"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   2640
      TabIndex        =   46
      Top             =   2160
      Visible         =   0   'False
      Width           =   1200
   End
End
Attribute VB_Name = "frmZdCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bShowForm As Boolean
Public cWhere As String
Public cType As String
Public blnPZSearch As Boolean
Dim frmXSDDD As Form

Public Sub ZDSingle(tmptype As String, cSQlWhere As String, bShow As Boolean)
    bShowForm = bShow
    cType = tmptype
    cWhere = cSQlWhere
    cmdOK_Click
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

 

Private Sub cmdDate1_LostFocus()
    If Me.ActiveControl Is edtDate1 Then Exit Sub
    edtDate1.Width = 1450
    cmdDate1.Visible = False
End Sub

 

Private Sub cmdDate2_LostFocus()
    If Me.ActiveControl Is edtDate2 Then Exit Sub
    edtDate2.Width = 1450
    cmdDate2.Visible = False
End Sub

Private Sub cmdDept_LostFocus()
    If Me.ActiveControl Is edtDept Then Exit Sub
    edtDept.Width = 1450
    cmdDept.Visible = False
End Sub

Private Sub cmdDw_LostFocus()
    If Me.ActiveControl Is edtDw Then Exit Sub
    edtDw.Width = 1450
    cmdDw.Visible = False
End Sub

Private Sub cmdOK_Click()
    Dim FpAndYsd As Boolean
    Dim StrZd As String, Zd As UfRecordset, Title As String
    Dim Str As String, Count As Long
    Dim i As Integer, Clear As Boolean
    Dim iNum    As Byte
    Dim oAuth       As New U8RowAuthsvr.clsRowAuth
    Dim sAuthCust$, sAuthVend$, sAuthDept$, sAuthCMType$
    Dim rs  As New ADODB.Recordset
    Dim cDwCode$, cDepCode$, cPersonCode$, cSSCode$
    Dim wherestr As String
    
    If Trim(edtDate1.Text) <> "" And Trim(edtDate2.Text) <> "" Then
        If CDate(edtDate1.Text) > CDate(edtDate2.Text) Then
            Msg "日期范围输入错误！", vbExclamation
            edtDate1.SetFocus
            Exit Sub
        End If
    End If
    FpAndYsd = False
    Clear = False
    frmZD1.blnPZSearch = blnPZSearch
'    frmZD1_Label1_Caption = "共 0 条"
    frmZD1.Grid.Rows = 1
'    frmZD1.Grid.colwidth(15) = 0

    '取制单类型
    For i = 0 To List1.ListCount - 1
        If (List1.Selected(i) = True) Then                                             '
'            iNum = iNum + 1

            Select Case List1.List(i)
                Case "费用预估单"
                  cType = "FYGL"
                  iNum = iNum + 1
                Case "资产减少"
                    cType = "INC"
                    iNum = iNum + 1
            End Select
        End If
    Next i
    If cType = "" Then
        Msg "请选择单据！", vbExclamation
        Exit Sub
    End If
    wherestr = ""
    Select Case UCase(cType)
        Case "FYGL"
            If edtDate1.Text <> "" And edtDate2.Text <> "" Then
                wherestr = " and (ddate >= '" & edtDate1.Text & "' and ddate <= '" & edtDate2.Text & "' )"
            ElseIf Trim(edtDate1.Text) <> "" Then
                wherestr = " and (ddate >= '" & CDate(edtDate1.Text) & "')"
            ElseIf Trim(edtDate2.Text) <> "" Then
                wherestr = " and (ddate <= '" & CDate(edtDate2.Text) & "')"
            End If
            
            StrZd = "select * from EFFYGL_v_Pcostbudget where IsNULL(cVerifier,'')<>''  AND ISNULL(cCloser,'')='' and bbuild<>1 " & wherestr & cWhere
    End Select
    Me.Hide
    frmZD1.menu_refurbish
    frmZD1.StrPz1 = ""
    frmZD1.pStyle1 = ""
    frmZD1.strPz = StrZd
    frmZD1.pStyle = cType
    FpAndYsd = True
    Select Case m_Login.cSub_Id
        Case "AP"
            Title = "应付制单"
        Case "AR"
            Title = "应收制单"
        Case "GL"
            Title = "总帐制单"
        Case "FA"
            If blnPZSearch = False Then
                Title = "预算制单"
            Else
                Title = "凭证查询"
            End If
    End Select
    
    Title = "用友软件"
    
    Dim sGuid As String
    sGuid = CreateGUID()
            If bShowForm Then
                If g_business Is Nothing Then
                     frmZD1.Show
                Else
            '        InitToolbarTag Me.tbrvoucher
                    Call g_business.ShowForm(frmZD1, "FA", sGuid, False, True, frmZD1.Object_vfd)
 
                    
'                    frmZD1.Caption = Me_Caption
    '                        Set .VouchList.PortalBusinessObject = g_business
    '                        VouchList.PortalBizGUID = sguid
                End If
'                frmZD1.Show
            End If
            DoEvents
            If Not FpAndYsd Then
                frmZD1_Label1_Caption = "共 0 条" 'GetResStringNoParam("U8.CW.APAR.ARAPMain.Total_Qua_Zero")
            End If
        frmZD1_Label1_Caption = "共" & CStr(frmZD1.Grid.Rows - 1) & "条"
'    Next i
    If iNum = 0 Then
            If Trim(cType) = "" Then
                Msg "请至少选择一种制单类型！", vbCritical
                Exit Sub
            End If
    ElseIf iNum > 1 Then
    Else
        frmZD1.lblTitle.Caption = Title
    End If
    
    Call FillZd(cType, StrZd, edtAmount1.Text, edtAmount2.Text)
    With frmZD1.Grid
        For Count = 1 To .Rows - 1
            .TextMatrix(Count, 1) = frmZD1.cboSign.Text  'cSign
        Next Count
    End With
    frmZD1.Label1.Caption = frmZD1_Label1_Caption
    Unload Me
    Exit Sub
Wrong:
'    ClsTask.TaskEnd oAcc.Sysid & "0402"
    Msg "查询过程出现异常，请您重新输入条件", vbExclamation
End Sub

Private Sub cmdPsn_LostFocus()
    If Me.ActiveControl Is edtPsn Then Exit Sub
    edtPsn.Width = 1450
    cmdPsn.Visible = False
End Sub

Private Sub Command1_Click()
cmdOK_Click
End Sub

Private Sub Command2_Click()
 Unload Me
End Sub


Private Sub edtDate1_GotFocus()
'    edtDate1.Width = 1150
'    cmdDate1.Visible = True
End Sub

Private Sub edtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF2 Then Call cmdDate1_Click
End Sub

Private Sub edtDate1_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then Call SetNextTab(edtDate1, Me)
End Sub

Private Sub edtDate2_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF2 Then Call cmdDate2_Click
End Sub

Private Sub edtDate2_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then Call SetNextTab(edtDate2, Me)
End Sub

Private Sub EdtDept_GotFocus()
    edtDept.Width = 1150
    cmdDept.Visible = True
End Sub


Private Sub EdtDept_LostFocus()
    If Me.ActiveControl Is cmdDept Then Exit Sub
    If Me.ActiveControl Is cmdCancel Then Exit Sub
    edtDept.Width = 1450
    cmdDept.Visible = False
End Sub

Private Sub edtDw_GotFocus()
    edtDw.Width = 1150
    cmdDw.Visible = True
End Sub


Private Sub edtDw_LostFocus()
    If Me.ActiveControl Is cmdDw Then Exit Sub
    If Me.ActiveControl Is cmdCancel Then Exit Sub
    edtDw.Width = 1450
    cmdDw.Visible = False
End Sub
  
 
Private Sub Form_Load()
    On Error Resume Next
    Me.Icon = frmMain.Icon
    DoForm Me, 3
    Call change_caption
    cmdOk.Caption = "确定"
    Me.cmdCancel.Caption = "取消"
    List1.Clear
    With List1
        .AddItem "费用预估单"
'        .AddItem "资产减少"
'        .AddItem "资产变动"
        .Selected(0) = True
    End With
    
    edtDate2.Text = Format(m_Login.CurDate, "yyyy-mm-dd")
    cboOperator.Clear
    cboOperator.AddItem ""
    cboCheckMan.Clear
    cboCheckMan.AddItem ""
End Sub

 
Public Sub change_caption()
    If Me.blnPZSearch = True Then
        frmZdCX.Caption = "凭证查询条件"
    Else
        frmZdCX.Caption = "单据制单查询"
    End If
End Sub
