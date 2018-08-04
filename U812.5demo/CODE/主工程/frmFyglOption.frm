VERSION 5.00
Object = "{9ADF72AD-DDA9-11D1-9D4B-000021006D51}#1.31#0"; "UFSpGrid.ocx"
Object = "{5E4640D0-A415-404B-A457-72980C429D2F}#10.25#0"; "U8RefEdit.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFyglOption 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "费用选项"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdSure 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   300
      Left            =   3450
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4770
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   300
      Left            =   4710
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4770
      Width           =   1050
   End
   Begin VB.Timer TZcy 
      Enabled         =   0   'False
      Left            =   -510
      Top             =   5775
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4500
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7938
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "凭证设置"
      TabPicture(0)   =   "frmFyglOption.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame7 
         Caption         =   "凭证信息"
         Height          =   3660
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   5805
         Begin MsSuperGrid.SuperGrid grid 
            Height          =   2520
            Left            =   150
            TabIndex        =   7
            Top             =   1050
            Width           =   5550
            _ExtentX        =   9790
            _ExtentY        =   4445
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
            BackColorFixed  =   -2147483633
            BackColorBkg    =   -2147483636
         End
         Begin U8Ref.RefEdit txtccode 
            Height          =   315
            Left            =   1095
            TabIndex        =   6
            Top             =   315
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   556
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "贷方科目："
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   5
            Top             =   780
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "借方科目："
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   4
            Top             =   390
            Width           =   900
         End
      End
   End
End
Attribute VB_Name = "frmFyglOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsCtlRefer As New U8RefService.IService

Dim oDisColor As Long

'列字段Col
Dim colysbm As Integer
Dim colysmc As Integer
Dim colkmbm As Integer
Dim colkmmc As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSure_Click()
    
    Dim i As Integer
'    Dim datComb As Date
    On Error GoTo SetErr
    Dim strsql As String
    If Not LockItem("EFFYGL040102", True) Then Exit Sub
    
  
    
    UpdateAccinfo "EF", "EFFYGL_FyygdcCode", "费用预估单借方科目", txtccode.Text    '借方科目
    SaveData
    
    LockItem "EFFYGL040102", 0
    Unload Me
    Exit Sub
SetErr:
    LockItem "EFFYGL040102", 0
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.ScaleWidth) / 2
    Me.Top = (Screen.Height - Me.ScaleHeight) / 2
    
     
    Me.HelpContextID = 20090401
    Set Me.Icon = frmMain.Icon
    
    initme
    readData
End Sub

Private Sub initme()
    Me.txtccode.Init m_Login, "code_GL"
    
    oDisColor = &HC0C0C0
    
    colysbm = 1
    colysmc = 2
    colkmbm = 3
    colkmmc = 4
    
    With grid
        .Rows = 2
        .cols = 5
        .AddDisColor oDisColor
        .ColButton(colkmbm) = UserBrowButton
        .ColButton(colkmmc) = UserBrowButton
        
    End With
    initGridStyle
End Sub

Private Sub initGridStyle()
    Dim i As Long
    With Me.grid
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, colysbm) = "要素编码"
        .TextMatrix(0, colysmc) = "要素名称"
        .TextMatrix(0, colkmbm) = "科目编码"
        .TextMatrix(0, colkmmc) = "科目名称"
        
        .colwidth(0) = 50
        .colwidth(colysbm) = 1100
        .colwidth(colysmc) = 1500
        .colwidth(colkmbm) = 1100
        .colwidth(colkmmc) = 1500
        
        For i = 1 To .Rows - 1
            .SetCellBackColor i, colysbm, oDisColor
            .SetCellBackColor i, colysmc, oDisColor
        Next
    End With
End Sub

Private Sub readData()
    Dim strsql As String: Dim rs As Object
    
    strsql = "select ele.celementcode,ele.celementname,o.ccode,code.ccode_name as cname from PM_ItemElement as ele"
    strsql = strsql & vbCrLf & "left outer join EFFYGL_ElementCCodeOption as o On ele.celementcode=o.celementcode"
    strsql = strsql & vbCrLf & "left outer join code on code.ccode=o.ccode"
    Set rs = DBConn.Execute(strsql)
    
    If rs.EOF Then
        grid.Rows = 2
    Else
        Set grid.Recordset = rs
    End If
    initGridStyle
    
    txtccode.Text = getAccinformation("EF", "EFFYGL_FyygdcCode")
End Sub

Private Sub SaveData()
    Dim strsql As String: Dim rs As Object
    Dim row As Long
    
    For row = 1 To grid.Rows - 1
        strsql = "IF NOT EXISTS(SELECT * FROM EFFYGL_ElementCCodeOption WHERE celementcode=" & reNSql(grid.TextMatrix(row, colysbm)) & ")" & vbCr & _
            "INSERT INTO EFFYGL_ElementCCodeOption(celementcode,ccode) VALUES (" & reNSql(grid.TextMatrix(row, colysbm)) & "," & reNSql(grid.TextMatrix(row, colkmbm)) & ") " & vbCr & _
            "Else" & vbCr & _
            "UPDATE EFFYGL_ElementCCodeOption SET ccode=" & reNSql(grid.TextMatrix(row, colkmbm)) & "" & " WHERE celementcode=" & reNSql(grid.TextMatrix(row, colysbm)) & ""
        DBConn.Execute strsql
    Next
End Sub

Private Function reNSql(ByVal Str As String) As String
    reNSql = "N'" & VBA.Replace(Str, "'", "''") & "'"
End Function

Private Sub UpdateAccinfo(strSysID As String, strName As String, strCap As String, strValue As String)
    Dim strsql As String
    strsql = "IF NOT EXISTS(SELECT * FROM Accinformation WHERE cSysID='" & strSysID & "' AND cName='" & strName & "')" & vbCr & _
         "INSERT INTO AccInformation(cSysID,cID,cName,cCaption,cType,cValue,cDefault,bVisible,bEnable) VALUES ('" & strSysID & "','FYGL','" & strName & "','" & strCap & "','Text','" & strValue & "',null,1,1) " & vbCr & _
         "Else" & vbCr & _
         "UPDATE Accinformation SET cValue='" & strValue & "'" & "WHERE cSysID='" & strSysID & "' AND cName='" & strName & "'"
    DBConn.Execute strsql

End Sub

Private Sub SetcontrolEnable()
    Dim Enable As Boolean
    Dim LockItem As Boolean
    If Not m_Login.TaskExec("EFFYGL040102", False) Then
        txtccode.Enabled = False
        grid.ReadOnly = True
    End If
End Sub

Private Sub grid_BeforeEdit(Cancel As Boolean, sReturnText As String)
    If grid.TextMatrix(grid.row, colysbm) = "" Then Cancel = True
End Sub

Private Sub grid_BrowUser(RetValue As String, ByVal R As Long, ByVal C As Long)
    Dim rstClass As ADODB.Recordset
    Dim rstGrid As ADODB.Recordset
    Dim strError As String
    
    Select Case C
        Case colkmbm, colkmmc
            clsCtlRefer.RefID = "code_GL"
            clsCtlRefer.FillText = RetValue
            clsCtlRefer.FilterSQL = ""
            clsCtlRefer.MetaXML = "" ' "<Ref><RefSet bControlledInvalidData='1' bAuth='0' bMultiSel= '" + IIf(CBool(strMuti), "1", "0") + "' iShowMode='" + nod.Attributes.getNamedItem("referdispmode").Text + "' iShowStyle='" + strShowStyle + "' /></Ref>" 'iShowMode
            If clsCtlRefer.ShowRef(m_Login, rstClass, rstGrid, strError) Then
        '                clsCtlRefer.Show
                If Not rstGrid Is Nothing Then
                    grid.TextMatrix(R, colkmbm) = rstGrid("ccode") & ""
                    grid.TextMatrix(R, colkmmc) = rstGrid("ccode_name") & ""
                    RetValue = grid.TextMatrix(R, C)
'                    Set referPara.rstGrid = rstGrid
'                    ShowEnumReferCtrl = FillItemsAfterBrowse(clsVoucher, voucher, strCardSection, strFieldName, rstGrid, lngRow)
                End If
            End If
    End Select
End Sub

Private Sub Grid_CellDataCheck(RetValue As String, RetState As MsSuperGrid.OpType, ByVal R As Long, ByVal C As Long)
    Dim strsql As String: Dim rs As Object
    
    Select Case C
        Case colkmbm, colkmmc
            If RetValue = "" Then
                grid.TextMatrix(R, colkmbm) = ""
                grid.TextMatrix(R, colkmmc) = ""
                Exit Sub
            End If
            strsql = "select * from code where ccode='" & VBA.Replace(RetValue, "'", "''") & "' or ccode_name='" & VBA.Replace(RetValue, "'", "''") & "'"
            Set rs = DBConn.Execute(strsql)
            If Not rs.EOF Then
                grid.TextMatrix(R, colkmbm) = rs("ccode") & ""
                grid.TextMatrix(R, colkmmc) = rs("ccode_name") & ""
            Else
                MsgBox "科目信息非法，请输入或选择正确的科目", vbInformation
                RetState = dbCandel
            End If
            rs.Close: Set rs = Nothing
            Exit Sub
    End Select
End Sub
