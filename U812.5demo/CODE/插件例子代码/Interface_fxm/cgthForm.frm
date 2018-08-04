VERSION 5.00
Object = "{9ADF72AD-DDA9-11D1-9D4B-000021006D51}#1.31#0"; "UFSpGrid.ocx"
Begin VB.Form cgthForm 
   BackColor       =   &H80000009&
   Caption         =   "替代料明细"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15285
   Icon            =   "cgthForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   15285
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFE3C0&
      Height          =   550
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.CommandButton btnDelLine 
         Caption         =   "删除行"
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton btnCopyLine 
         Caption         =   "复制行"
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   130
         Width           =   1095
      End
      Begin VB.CommandButton btnSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "保存"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   130
         Width           =   1095
      End
      Begin VB.CommandButton btnBack 
         Caption         =   "返回"
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   130
         Width           =   1095
      End
   End
   Begin MsSuperGrid.SuperGrid grid 
      Height          =   2175
      Left            =   120
      TabIndex        =   5
      Top             =   1440
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
      BackColorSel    =   -2147483635
      BackColorFixed  =   16769984
      BackColorBkg    =   -2147483636
   End
End
Attribute VB_Name = "cgthForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnBack_Click()
    Unload Me
End Sub

Private Sub btnCopyLine_Click()
    Dim bl As Boolean
    
    bl = False
    For i = 1 To grid.rows - 1
        If LCase(grid.TextMatrix(i, 0)) = LCase("Y") Then
            grid.AddItem ""
                        
            grid.TextMatrix(grid.rows - 1, 1) = grid.TextMatrix(i, 1)
            grid.TextMatrix(grid.rows - 1, 2) = grid.TextMatrix(i, 2)
            grid.TextMatrix(grid.rows - 1, 3) = grid.TextMatrix(i, 3)
            grid.TextMatrix(grid.rows - 1, 4) = grid.TextMatrix(i, 4)
            grid.TextMatrix(grid.rows - 1, 5) = grid.TextMatrix(i, 5)
            grid.TextMatrix(grid.rows - 1, 6) = grid.TextMatrix(i, 6)
            grid.TextMatrix(grid.rows - 1, 7) = grid.TextMatrix(i, 7)
            grid.TextMatrix(grid.rows - 1, 8) = grid.TextMatrix(i, 8)
            grid.TextMatrix(grid.rows - 1, 9) = grid.TextMatrix(i, 9)
            grid.TextMatrix(grid.rows - 1, 10) = grid.TextMatrix(i, 10)
            grid.TextMatrix(grid.rows - 1, 11) = grid.TextMatrix(i, 11)
            grid.TextMatrix(grid.rows - 1, 12) = grid.TextMatrix(i, 12)
            grid.TextMatrix(grid.rows - 1, 13) = grid.TextMatrix(i, 13)
            grid.TextMatrix(grid.rows - 1, 14) = grid.TextMatrix(i, 14)
            grid.TextMatrix(grid.rows - 1, 15) = grid.TextMatrix(i, 15)
            grid.TextMatrix(grid.rows - 1, 16) = grid.TextMatrix(i, 16)
            grid.TextMatrix(grid.rows - 1, 17) = grid.TextMatrix(i, 17)
            grid.TextMatrix(grid.rows - 1, 18) = grid.TextMatrix(i, 18)
            bl = True
        End If
    Next i
    
    If Not bl Then
        MsgBox "未选择要复制的行，请双击需要复制的行！"
    End If
    
End Sub

Private Sub btnDelLine_Click()
    Dim bl As Boolean
    
    bl = False
    For i = 1 To grid.rows - 1
        If LCase(grid.TextMatrix(i, 0)) = LCase("Y") Then
            If grid.TextMatrix(i, 19) = "R" Then
                MsgBox "该行只读，不可删除！"
                Exit Sub
            Else
                grid.RemoveItem i
                bl = True
                Exit For
            End If
        End If
    Next i
    
    If Not bl Then
        MsgBox "未选择要删除的行，请双击需要删除的行！"
    End If
    
End Sub

Private Sub btnSave_Click()
    Save
'    RefProduct SoKey, ccode
    MsgBox "保存成功！"
    
End Sub

Private Sub Form_Load()
    Frame1.Top = 0
    Frame1.Width = cgthForm.Width
    Frame1.Height = 550
    Frame1.Left = 0
    
    grid.Top = Frame1.Height + 50
    grid.Left = 0
    grid.Width = cgthForm.Width
    grid.Height = cgthForm.Height - Frame1.Height - 50 - 500
    
    '初始grid
    gridLoad
    
    '加载数据到grid
    readData
End Sub

Private Function gridLoad()
    Dim oDisColor As Long
    
    oDisColor = &HC0C0C0
    
    gridformat = "^选择|^订单号|^订单行号|^产品编码|^产品名称|^材料编码|^材料名称|^单位用量|^总需求量|^结存数量|^可用量|^替代料编码|^替代料名称|^计量单位|^原物料保留量|^替代数量|^结存数量|^可用数量|^销售订单ID"
    grid.FormatString = gridformat
    DoForm Me, 2
    
    With grid
        .rows = 1: .Cols = 21
        .ColAlignment(0) = 4
        .ColAlignment(1) = 4
        .ColAlignment(2) = 4
        .ColAlignment(3) = 4
        .ColAlignment(4) = 4
        .ColAlignment(5) = 4
        .ColAlignment(6) = 4
        .ColAlignment(7) = 4
        .ColAlignment(8) = 4
        .ColAlignment(9) = 4
        .ColAlignment(10) = 4
        .ColAlignment(11) = 4
        .ColAlignment(12) = 4
        .ColAlignment(13) = 4
        .ColAlignment(14) = 4
        .ColAlignment(15) = 4
        .ColAlignment(16) = 4
        .ColAlignment(17) = 4
        .ColAlignment(18) = 4
        .ColAlignment(19) = 4
        .ColAlignment(20) = 4
                
        .ColWidth(0) = 700
        .ColWidth(1) = 1200
        .ColWidth(2) = 1000
        .ColWidth(3) = 1400
        .ColWidth(4) = 1600
        .ColWidth(5) = 1300
        .ColWidth(6) = 1600
        .ColWidth(7) = 1000
        .ColWidth(8) = 1000
        
        .ColWidth(9) = 1000
        .ColWidth(10) = 800
        .ColWidth(11) = 1600
        .ColWidth(12) = 1600
        .ColWidth(13) = 1300
        .ColWidth(14) = 1200
        .ColWidth(15) = 1000
        .ColWidth(16) = 1000
        .ColWidth(17) = 1000
        .ColWidth(18) = 1
        .ColWidth(19) = 1
        .ColWidth(20) = 1
         
        .AddDisColor oDisColor
        
        .ColButton(11) = UserBrowButton
        .ColButton(14) = DblBrowButton
        .ColButton(15) = DblBrowButton
                
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
    Frame1.Width = cgthForm.Width
    Frame1.Height = 550
    Frame1.Left = 0
    
    grid.Top = Frame1.Height + 50
    grid.Left = 0
    grid.Width = cgthForm.Width
    grid.Height = IIf((cgthForm.Height - Frame1.Height - 50 - 500) < 10, grid.Height, cgthForm.Height - Frame1.Height - 50 - 500)
End Sub

Private Sub grid_BeforeEdit(Cancel As Boolean, sReturnText As String)
    Select Case Me.grid.Col
        Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 13, 16, 17
            Cancel = True
    End Select
End Sub

Private Sub grid_BrowUser(RetValue As String, ByVal r As Long, ByVal c As Long)
    Dim rstClass As ADODB.Recordset
    Dim rstGrid As ADODB.Recordset
    Dim strError As String
    Dim clsCtlRefer As New U8RefService.IService
    Dim rRs As New ADODB.Recordset
    Dim sSql As String
    
    On Error GoTo Exit1
    
    Select Case c
        Case 11
            clsCtlRefer.RefID = "inventory_aa"
            clsCtlRefer.FillText = RetValue
            clsCtlRefer.FilterSQL = ""
            clsCtlRefer.MetaXML = ""
            If clsCtlRefer.ShowRef(nvLogin, rstClass, rstGrid, strError) Then
        '                clsCtlRefer.Show
                If Not rstGrid Is Nothing Then
                    grid.TextMatrix(r, 11) = rstGrid("cinvcode")
                    grid.TextMatrix(r, 12) = rstGrid("cinvname")
                    grid.TextMatrix(r, 13) = rstGrid("ccomunitname")
                    RetValue = grid.TextMatrix(r, 11)
'                    Set referPara.rstGrid = rstGrid
'                    ShowEnumReferCtrl = FillItemsAfterBrowse(clsVoucher, voucher, strCardSection, strFieldName, rstGrid, lngRow)
                    sSql = "select cinvcode,iQuantity ,fAvaQuantity  from CurrentStock "
                    If rRs.State = 1 Then rRs.Close
                    rRs.CursorLocation = adUseClient
                    rRs.Open sSql, con.ConnectionString, 1, 2
                   
                    If Not rRs.EOF And Not rRs.BOF Then
                        grid.TextMatrix(r, 16) = rRs.Fields("iQuantity")
                        grid.TextMatrix(r, 17) = rRs.Fields("fAvaQuantity")
                    End If
                    If rRs.State = 1 Then rRs.Close
                End If
            End If
    End Select
    
    Exit Sub
Exit1:
    MsgBox "错误信息： 【" & Err.Description & "】"
    If rRs.State = 1 Then rRs.Close
End Sub


Private Sub readData()
    Dim strSql As String
    Dim i As Integer
    Dim j As Integer

    On Error GoTo ExitSub

    strSql = "select * from EF_V_BOMChenage where 销售订单ID = '" & SoKey & "' and 材料编码 = '" & cInv & "' order by 销售订单号,订单行号"
    
    If nvRs.State = 1 Then nvRs.Close
    nvRs.CursorLocation = adUseClient
    nvRs.Open strSql, con.ConnectionString, 1, 2
    
    
    If nvRs.EOF Then
        grid.rows = 2
    Else
        grid.rows = nvRs.RecordCount + 1
        For i = 1 To nvRs.RecordCount
            grid.TextMatrix(i, 1) = IIf(IsNull(nvRs.Fields("销售订单号")), "", nvRs.Fields("销售订单号")) 'nvRs.Fields("销售订单号")
            grid.TextMatrix(i, 2) = IIf(IsNull(nvRs.Fields("订单行号")), "", nvRs.Fields("订单行号")) 'nvRs.Fields("订单行号")
            grid.TextMatrix(i, 3) = IIf(IsNull(nvRs.Fields("产品编码")), "", nvRs.Fields("产品编码")) 'nvRs.Fields("产品编码")
            grid.TextMatrix(i, 4) = IIf(IsNull(nvRs.Fields("产品名称")), "", nvRs.Fields("产品名称")) 'nvRs.Fields("产品名称")
            grid.TextMatrix(i, 5) = IIf(IsNull(nvRs.Fields("材料编码")), "", nvRs.Fields("材料编码")) 'nvRs.Fields("材料编码")
            grid.TextMatrix(i, 6) = IIf(IsNull(nvRs.Fields("材料名称")), "", nvRs.Fields("材料名称")) 'nvRs.Fields("材料名称")
            grid.TextMatrix(i, 7) = IIf(IsNull(nvRs.Fields("单位用量")), "", nvRs.Fields("单位用量")) 'nvRs.Fields("单位用量")
            grid.TextMatrix(i, 8) = IIf(IsNull(nvRs.Fields("总需求量")), "", nvRs.Fields("总需求量")) 'nvRs.Fields("总需求量")
            grid.TextMatrix(i, 9) = IIf(IsNull(nvRs.Fields("结存数量")), "", nvRs.Fields("结存数量")) 'nvRs.Fields("结存数量")
            grid.TextMatrix(i, 10) = IIf(IsNull(nvRs.Fields("可用数量")), "", nvRs.Fields("可用数量")) 'nvRs.Fields("可用数量")
            grid.TextMatrix(i, 11) = IIf(IsNull(nvRs.Fields("替代物料编码")), "", nvRs.Fields("替代物料编码"))
            grid.TextMatrix(i, 12) = IIf(IsNull(nvRs.Fields("替代物料名称")), "", nvRs.Fields("替代物料名称"))
            grid.TextMatrix(i, 13) = IIf(IsNull(nvRs.Fields("替代料计量单位")), "", nvRs.Fields("替代料计量单位"))
            grid.TextMatrix(i, 14) = IIf(IsNull(nvRs.Fields("原物料保留量")), "", nvRs.Fields("原物料保留量"))
            grid.TextMatrix(i, 15) = IIf(IsNull(nvRs.Fields("替代数量")), "", nvRs.Fields("替代数量"))
            grid.TextMatrix(i, 16) = IIf(IsNull(nvRs.Fields("替代料结存数量")), "", nvRs.Fields("替代料结存数量"))
            grid.TextMatrix(i, 17) = IIf(IsNull(nvRs.Fields("替代料可用数量")), "", nvRs.Fields("替代料可用数量"))
            grid.TextMatrix(i, 18) = IIf(IsNull(nvRs.Fields("autoid")), "", nvRs.Fields("autoid"))
            grid.TextMatrix(i, 19) = "R"
                        
            nvRs.MoveNext
        Next i
        
        For i = 1 To grid.rows - 1
            If grid.TextMatrix(i, 19) = "R" Then
                For j = 1 To grid.Cols - 1
                    grid.SetCellBackColor i, j, &H8080FF
                Next j
            End If
        Next i
    
    End If
    
    If nvRs.State = 1 Then nvRs.Close

    Exit Sub
ExitSub:

    MsgBox "错误 【" & Err.Description & "】"
    If nvRs.State = 1 Then nvRs.Close
End Sub


Private Static Sub Save()
    Dim i As Integer
    Dim j As Integer
    
    Dim BeginTrans As Boolean
    Dim tInvcode As String
    Dim yQty As Double
    Dim tQty As Double
    
    Dim TunitQty As Double
    
    Dim rds As New ADODB.Recordset
    Dim bomErrstr As String
    Dim iQuantity As Double
    
    Dim cInv As String
    Dim domtemp As New DOMDocument
    
    Dim ufu8mbom As New UFU8MBOMISrv.clsBom
    Dim trues As Boolean
    
    On Error GoTo ExitSub1
    
    BeginTrans = False
    
'   开始事务
    If BeginTrans = False Then
        con.BeginTrans
        BeginTrans = True
    End If
        
    
    For i = 1 To grid.rows - 1
        con.Execute "delete ef_bomchenage where cSOID = '" & grid.TextMatrix(i, 20) & "' and cinvcode='" & grid.TextMatrix(i, 5) & "'"
    Next i
    
    SetNVBom "select * from ef_bom where ID='" & grid.TextMatrix(1, 20) & "' and cInvCode='" & grid.TextMatrix(1, 5) & "'", grid
    
    For i = 1 To grid.rows - 1
        cInv = grid.TextMatrix(i, 11)
        If grid.TextMatrix(i, 15) = "" Then
            iQuantity = 0
        Else
            iQuantity = CDbl(grid.TextMatrix(i, 15))
        End If
        
  
        nvsql = "insert into ef_bomchenage ([id],[autoid],[cSOCode],[cinvcode],[unitQty],[tInvcode],[yQty],[tQty],[TunitQty]) values(newid()," & grid.TextMatrix(i, 18) & ",'"
        nvsql = nvsql & grid.TextMatrix(i, 1) & "','" & grid.TextMatrix(i, 5) & "'," & grid.TextMatrix(i, 7) & ",'" & IIf(grid.TextMatrix(i, 11) = "", "", grid.TextMatrix(i, 11)) & "',"
        nvsql = nvsql & IIf(grid.TextMatrix(i, 14) = "", "NULL", grid.TextMatrix(i, 14)) & "," & IIf(grid.TextMatrix(i, 15) = "", "NULL", grid.TextMatrix(i, 15)) & "," & TunitQty & ")"
                
        con.Execute nvsql
                
                
        If BeginTrans Then
            con.CommitTrans
            BeginTrans = False
        End If
    Next i
    
    Exit Sub
ExitSub1:

    MsgBox "保存失败，错误信息：" & Err.Description
    If BeginTrans Then
        con.RollbackTrans
        BeginTrans = False
    End If
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
