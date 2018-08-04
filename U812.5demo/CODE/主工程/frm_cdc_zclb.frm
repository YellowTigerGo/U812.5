VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F47AFD17-0D3E-11D3-8772-00002100F7B3}#1.1#0"; "US_Case.ocx"
Begin VB.Form frm_cdc_zclx 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "资产类型"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   7095
   StartUpPosition =   2  '屏幕中心
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   5295
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   9340
      _Version        =   393216
      FixedCols       =   0
   End
   Begin VB.TextBox sassetname_Text 
      Height          =   300
      Left            =   4200
      TabIndex        =   7
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox stypenum_Text 
      Height          =   300
      Left            =   1440
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
   Begin AX_Case.AXCase AXCase1 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BackColor       =   16777215
      MaxLength       =   0
      RootDesc        =   ""
      BadString       =   "<>'""|&,"
      ListIndex       =   0
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确认"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   6240
      Width           =   1215
   End
   Begin VB.ComboBox Comb_zclx 
      Height          =   300
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "资产名称"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "国标分类码"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frm_cdc_zclx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim row As Long
Private Sub Command1_Click()
    If Comb_zclx.ListIndex = -1 Then
        MsgBox "你没有选择单据模板，请选择模板类型", vbQuestion, "提示"
    Else
        Me.Hide
    End If
End Sub

Private Sub AXCase1_ButtonClick()
    Dim clsRefer As New UFReferC.UFReferClient
    Dim strGrid As String
    If Trim(AXCase1.Text) <> "" Then
        strGrid = "select snum as 资产类别编码,sname 资产类别名称,showid as 显示模板,printid as 打印模板号 from fa_AssetTypes where lchilds=0 and (sname like '%" & Trim(AXCase1.Text) & "%' or  snum like '%" & Trim(AXCase1.Text) & "%')  order by snum "
    Else
        strGrid = "select snum as 资产类别编码,sname 资产类别名称,showid as 显示模板,printid as 打印模板号 from fa_AssetTypes where lchilds=0 order by snum"
    End If
    If clsRefer.StrRefInit_SetColWidth(m_login, False, "", strGrid, "资产类别编码,资产类别名称,显示模板号,打印模板号", "1500,4000,1000,1000") = False Then Exit Sub
    clsRefer.Show
    If Not clsRefer.recMx Is Nothing Then
        AXCase1.Text = clsRefer.recMx.Fields("资产类别名称")
        strTemplate = clsRefer.recMx.Fields("显示模板") & ""
        stypenum = clsRefer.recMx.Fields("资产类别编码") & ""
        sAssetName = clsRefer.recMx.Fields("资产类别名称") & ""
    Else
        strTemplate = ""
        stypenum = ""
        sAssetName = ""
        AXCase1.Text = ""
        Exit Sub
    End If
    
End Sub

Private Sub cmdCancel_Click()
    strTemplate = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    With MSFlexGrid
        If .row > 0 Then
            strTemplate = .TextMatrix(.row, 2)
            stypenum = .TextMatrix(.row, 0)
            sAssetName = .TextMatrix(.row, 1)
            Unload Me
        Else
            MsgBox "资产类别(或显示模板)不能为空！", vbOKOnly, "操作提示"
        End If
    End With
End Sub

Private Sub refurbish()
    Dim sqlstr As String
    Dim tempRs As New ADODB.Recordset
    Dim i As Long
On Error GoTo Err
    row = 1
    Me.Icon = frmMain.Icon
    MSFlexGrid.Clear
    With MSFlexGrid
        .ScrollBars = 2
        .rows = 1
        .cols = 4
        .AllowUserResizing = 3
        .FormatString = "^  国标分类码 |^  资产名称  "
        .colwidth(0) = 2000
        .colwidth(1) = 4800
        .colwidth(2) = 0
        .colwidth(3) = 0
        .ColAlignment(0) = 2
        .ColAlignment(1) = 2
        sqlstr = "select   snum,sname,showid,printid from fa_AssetTypes where lchilds=0   "
        If Trim(stypenum_Text.Text) <> "" Then
            sqlstr = sqlstr & " and snum like '" & Trim(stypenum_Text.Text) & "%'"
        End If
        If Trim(sassetname_Text.Text) <> "" Then
            sqlstr = sqlstr & " and  sname like '%" & Trim(sassetname_Text.Text) & "%'"
        End If
        sqlstr = sqlstr & "order by snum"
        tempRs.Open sqlstr, DBConn, adOpenKeyset, adLockReadOnly
        If tempRs.RecordCount > 0 Then
            .rows = tempRs.RecordCount + 1
        End If
        i = 1
        While Not tempRs.EOF
            .TextMatrix(i, 0) = tempRs.Fields(0)
            .TextMatrix(i, 1) = tempRs.Fields(1)
            .TextMatrix(i, 2) = tempRs.Fields(2)
            .TextMatrix(i, 3) = tempRs.Fields(3)
            If i Mod 2 = 1 Then
                .row = i
                .col = 0
                .CellBackColor = &HC0FFFF
                .col = 1
                .CellBackColor = &HC0FFFF
            Else
                .row = i
                .col = 0
                .CellBackColor = &H8000000E
                .col = 1
                .CellBackColor = &H8000000E
            End If
            i = i + 1
            tempRs.MoveNext
        Wend
        .row = 1
        .col = 0
        .CellBackColor = &HFF8080
        .col = 1
        .CellBackColor = &HFF8080
        row = .row
    End With
    tempRs.Close

    Set tempRs = Nothing
    Exit Sub
Err:
    Set tempRs = Nothing
End Sub

Private Sub Form_Load()
    refurbish
End Sub

Private Sub MSFlexGrid_Click()
Dim i As Long
Dim j As Long
    With MSFlexGrid
    j = row
    If j > 1 Then
        .row = j
        .col = 0
        .CellBackColor = &HFF8080
        .col = 1
        .CellBackColor = &HFF8080
        row = .row
        End If
    End With
End Sub

Private Sub MSFlexGrid_DblClick()
    With MSFlexGrid
        If .row > 0 Then
            strTemplate = .TextMatrix(.row, 2)
            stypenum = .TextMatrix(.row, 0)
            sAssetName = .TextMatrix(.row, 1)
            Unload Me
        Else
            MsgBox "资产类别(或显示模板)不能为空！", vbOKOnly, "操作提示"
        End If
    End With
End Sub

Private Sub MSFlexGrid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    With MSFlexGrid
        If .row > 0 Then
            strTemplate = .TextMatrix(.row, 2)
            stypenum = .TextMatrix(.row, 0)
            sAssetName = .TextMatrix(.row, 1)
            Unload Me
        Else
            MsgBox "资产类别(或显示模板)不能为空！", vbOKOnly, "操作提示"
        End If
    End With
End If
End Sub

Private Sub MSFlexGrid_SelChange()
Dim j As Long
    j = row
    row = MSFlexGrid.row
    If j > 0 Then
        If j Mod 2 = 1 Then
            MSFlexGrid.row = j
            MSFlexGrid.col = 0
            MSFlexGrid.CellBackColor = &HC0FFFF
            MSFlexGrid.col = 1
            MSFlexGrid.CellBackColor = &HC0FFFF
        Else
            MSFlexGrid.row = j
            MSFlexGrid.col = 0
            MSFlexGrid.CellBackColor = &H8000000E
            MSFlexGrid.col = 1
            MSFlexGrid.CellBackColor = &H8000000E
        End If
    End If
End Sub

Private Sub sassetname_Text_Change()
    refurbish
End Sub

Private Sub sassetname_Text_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        With MSFlexGrid
            If .row > 0 Then
                strTemplate = .TextMatrix(.row, 2)
                stypenum = .TextMatrix(.row, 0)
                sAssetName = .TextMatrix(.row, 1)
                Unload Me
            Else
                MsgBox "资产类别(或显示模板)不能为空！", vbOKOnly, "操作提示"
            End If
        End With
    End If
End Sub

Private Sub stypenum_Text_Change()
    refurbish
End Sub

Private Sub stypenum_Text_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        With MSFlexGrid
            If .row > 0 Then
                strTemplate = .TextMatrix(.row, 2)
                stypenum = .TextMatrix(.row, 0)
                sAssetName = .TextMatrix(.row, 1)
                Unload Me
            Else
                MsgBox "资产类别(或显示模板)不能为空！", vbOKOnly, "操作提示"
            End If
        End With
    End If
End Sub
