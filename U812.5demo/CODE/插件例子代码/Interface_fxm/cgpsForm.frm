VERSION 5.00
Object = "{9ADF72AD-DDA9-11D1-9D4B-000021006D51}#1.31#0"; "UFSpGrid.ocx"
Begin VB.Form cgpsForm 
   BackColor       =   &H80000009&
   Caption         =   "采购评审界面"
   ClientHeight    =   8475
   ClientLeft      =   3660
   ClientTop       =   2820
   ClientWidth     =   14925
   Icon            =   "cgpsForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   14925
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFE3C0&
      Height          =   550
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6015
      Begin VB.CommandButton btnTiHuan 
         Caption         =   "替换"
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   130
         Width           =   1095
      End
      Begin VB.CommandButton btnBack 
         Caption         =   "返回"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   130
         Width           =   1095
      End
   End
   Begin MsSuperGrid.SuperGrid grid 
      Height          =   2175
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   3836
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EditBorderStyle =   0
      Redraw          =   1
      GridColorFixed  =   -2147483632
      GridColor       =   -2147483633
      ForeColorSel    =   -2147483634
      ForeColorFixed  =   -2147483630
      Cols            =   3
      BackColorSel    =   -2147483635
      BackColorFixed  =   16769984
      BackColorBkg    =   -2147483636
   End
End
Attribute VB_Name = "cgpsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private gridformat As String
'Private vouch As Object
'Private doms As New DOMDocument


Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub btnTiHuan_Click()
    Dim i As Integer
    Dim bl As Boolean
    
    
    bl = False
    
    For i = 1 To grid.rows - 1
    
        If grid.TextMatrix(i, 0) = "Y" Then
        
            cInv = grid.TextMatrix(i, 1)
            bl = True
            Exit For
        
        End If
        
    
    Next i
    
    
    If bl Then
        cgthForm.Show 1
    Else
        MsgBox "请选择需要替换的材料！"
    End If
    
    
    
    
End Sub

Private Sub Form_Load()
    gridformat = "^  选择  |^材料编码|^材料名称|^材料规格|^计量单位|^总需求量|^结存数量|^可用量|^是否替代"
'    gridformat2 = "^  选择  |^订单号|^订单行号|^产品编码|^产品名称|^材料编码|^材料名称|^单位用量|^总需求量|^结存数量|^可用量|^替代料编码|^替代料名称|^计量单位|^原物料保留量|^替代数量|^结存数量|^可用数量"
    
    
           
    Frame1.Top = 0
    Frame1.Width = cgpsForm.Width
    Frame1.Height = 550
    Frame1.Left = 0
    
    grid.Top = Frame1.Height + 50
    grid.Left = 0
    grid.Width = cgpsForm.Width
    grid.Height = cgpsForm.Height - Frame1.Height - 50 - 500
    
    '初始grid
    gridLoad
    
    '加载数据到grid
    readData
    
End Sub

Public Function init(voucher As Object, cn As ADODB.Connection, code As String, id As Long)
    Set con = cn
    ccode = code
    SoKey = id
End Function

Private Function gridLoad()
    Dim oDisColor As Long
    oDisColor = &HC0C0C0
    
    grid.FormatString = gridformat
    DoForm Me, 2
    
    With grid
        .rows = 1: .Cols = 9
        .ColAlignment(0) = 4
        .ColAlignment(1) = 4
        .ColAlignment(2) = 4
        .ColAlignment(3) = 4
        .ColAlignment(4) = 4
        .ColAlignment(5) = 1
        .ColAlignment(6) = flexAlignRightCenter
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(8) = flexAlignRightCenter


'        .ColWidth(0) = 900
'        .ColWidth(1) = 1200
'        .ColWidth(2) = 1200
'        .ColWidth(3) = 1200
'        .ColWidth(4) = 1200
'        .ColWidth(5) = 1200
'        .ColWidth(6) = 1200
'        .ColWidth(7) = 1300
'        .ColWidth(8) = 1200
        
        .ColWidth(0) = 800
        .ColWidth(1) = (grid.Width - 800) * 0.125
        .ColWidth(2) = (grid.Width - 800) * 0.125
        .ColWidth(3) = (grid.Width - 800) * 0.125
        .ColWidth(4) = (grid.Width - 800) * 0.125
        .ColWidth(5) = (grid.Width - 800) * 0.125
        .ColWidth(6) = (grid.Width - 800) * 0.125
        .ColWidth(7) = (grid.Width - 800) * 0.125
        .ColWidth(8) = (grid.Width - 800) * 0.125
         
         
'         .SetColProperty
         
        .AddDisColor oDisColor
'        .ColButton(2) = BrowCom
'        .ColButton(3) = BrowNull
'        .ColButton(4) = DateBrowButton
'        .ColButton(5) = DblBrowButton
'        .ColButton(6) = UserBrowButton
                
    End With
    
End Function

Private Sub DoForm(Frm As Form, BorderStyle As Integer)
    On Error GoTo Next11
    Select Case BorderStyle
    Case 2
        Frm.Icon = Nothing ' LoadResPicture(103, 1)
    Case 3
        Frm.Icon = Nothing
        Frm.Left = (Screen.Width - Frm.Width) / 2
        Frm.Top = (Screen.Height - Frm.Height) / 2
    End Select
    Dim ctl As Control
    For Each ctl In Frm.Controls
        If TypeName(ctl) = "Edit" Then
            ctl.BadStr = "<>&*_%'|?;"""
        End If
    Next
Next11:
End Sub

Private Sub Form_Resize()
    
    Frame1.Top = 0
    Frame1.Width = cgpsForm.Width
    Frame1.Height = 550
    Frame1.Left = 0
    
    grid.Top = Frame1.Height + 50
    grid.Left = 0
    grid.Width = cgpsForm.Width
    grid.Height = IIf((cgpsForm.Height - Frame1.Height - 50 - 500) < 10, grid.Height, cgpsForm.Height - Frame1.Height - 50 - 500)
    
    With grid
        .ColWidth(0) = 800
        .ColWidth(1) = (grid.Width - 800) * 0.125
        .ColWidth(2) = (grid.Width - 800) * 0.125
        .ColWidth(3) = (grid.Width - 800) * 0.125
        .ColWidth(4) = (grid.Width - 800) * 0.125
        .ColWidth(5) = (grid.Width - 800) * 0.125
        .ColWidth(6) = (grid.Width - 800) * 0.125
        .ColWidth(7) = (grid.Width - 800) * 0.125
        .ColWidth(8) = (grid.Width - 800) * 0.125
    End With
End Sub

Private Sub readData()
    Dim strSql As String
    Dim i As Integer

    On Error GoTo ExitSub

    strSql = "select  材料编码,材料名称,材料规格,计量单位,sum(总需求量),结存数量,可用数量,是否替代,销售订单ID from EF_V_BOMChenage "
    strSql = strSql & " where 销售订单ID ='" & SoKey & "'"
    strSql = strSql & " group by 材料编码,材料名称,材料规格,计量单位,结存数量,可用数量,是否替代,销售订单ID "
    
    
    If nvRs.State = 1 Then nvRs.Close
    nvRs.CursorLocation = adUseClient
'    Set Rs = con.Execute(strsql)
    nvRs.Open strSql, con.ConnectionString, 1, 2
'    nvRs.Open strsql, con.ConnectionString
    
    
    If nvRs.EOF Then
        grid.rows = 2
    Else
        
        grid.rows = nvRs.RecordCount + 1
        For i = 1 To nvRs.RecordCount
            grid.TextMatrix(i, 1) = IIf(IsNull(nvRs.Fields(0)), "", nvRs.Fields(0))
            grid.TextMatrix(i, 2) = IIf(IsNull(nvRs.Fields(1)), "", nvRs.Fields(1))
            grid.TextMatrix(i, 3) = IIf(IsNull(nvRs.Fields(2)), "", nvRs.Fields(2))
            grid.TextMatrix(i, 4) = IIf(IsNull(nvRs.Fields(3)), "", nvRs.Fields(3))
            grid.TextMatrix(i, 5) = IIf(IsNull(nvRs.Fields(4)), "", nvRs.Fields(4))
            grid.TextMatrix(i, 6) = IIf(IsNull(nvRs.Fields(5)), "", nvRs.Fields(5))
            grid.TextMatrix(i, 7) = IIf(IsNull(nvRs.Fields(6)), "", nvRs.Fields(6))
            grid.TextMatrix(i, 8) = IIf(IsNull(nvRs.Fields(7)), "", nvRs.Fields(7))
            
            nvRs.MoveNext
        Next i
'        Set grid.Recordset = Rs
    
    End If
    
    If nvRs.State = 1 Then nvRs.Close

    Exit Sub
ExitSub:

    MsgBox "错误【" & Err.Description & "】"
    If nvRs.State = 1 Then nvRs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
        pubMoid = -1
End Sub

Private Sub grid_BeforeEdit(Cancel As Boolean, sReturnText As String)
    Cancel = True
End Sub

Private Sub grid_DblClick()
    Dim i As Integer
    If grid.TextMatrix(grid.rowsel, 0) = "Y" Then
        grid.TextMatrix(grid.rowsel, 0) = ""
    Else
        grid.TextMatrix(grid.rowsel, 0) = "Y"
        For i = 1 To grid.rows - 1
            grid.TextMatrix(i, 0) = ""
        Next i
        grid.TextMatrix(grid.rowsel, 0) = "Y"
    End If
End Sub
