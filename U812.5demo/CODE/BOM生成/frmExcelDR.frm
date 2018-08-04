VERSION 5.00
Object = "{9FD12F62-6922-47E1-B1AC-3615BBD3D7A5}#1.0#0"; "UFLabel.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Object = "{5E4640D0-A415-404B-A457-72980C429D2F}#10.37#0"; "U8RefEdit.ocx"
Begin VB.Form frmExcelDR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EXCEL导入"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   Icon            =   "frmExcelDR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   -840
      TabIndex        =   5
      Top             =   840
      Width           =   11655
      Begin UFLABELLib.UFLabel UFLabel1 
         Height          =   375
         Left            =   3240
         TabIndex        =   17
         Top             =   1920
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   661
         _StockProps     =   111
         Caption         =   "仓库："
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin U8Ref.RefEdit txtcWhcode 
         Height          =   375
         Left            =   3840
         TabIndex        =   16
         Top             =   1920
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         BadStr          =   "<>'""|&,"
         BadStrException =   """|&,"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Property        =   1
         RefType         =   1
         SkinStyle       =   ""
         StopSkin        =   0   'False
      End
      Begin VB.CommandButton Command2 
         Caption         =   "刷新"
         Height          =   375
         Index           =   3
         Left            =   6960
         TabIndex        =   14
         Top             =   1440
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   3240
         TabIndex        =   12
         Top             =   1440
         Width           =   3615
      End
      Begin VB.PictureBox Picture2 
         Height          =   3855
         Left            =   960
         Picture         =   "frmExcelDR.frx":030A
         ScaleHeight     =   3795
         ScaleWidth      =   1995
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3240
         TabIndex        =   7
         Top             =   600
         Width           =   5655
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   375
         Index           =   1
         Left            =   9000
         TabIndex        =   6
         Top             =   600
         Width           =   495
      End
      Begin ComCtl2.Animation Animation1 
         Height          =   735
         Left            =   3240
         TabIndex        =   9
         Top             =   3360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1296
         _Version        =   327681
         FullWidth       =   161
         FullHeight      =   49
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "选择工作表："
         Height          =   180
         Left            =   3240
         TabIndex        =   13
         Top             =   1200
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "选择EXCEL文档："
         Height          =   180
         Left            =   3240
         TabIndex        =   11
         Top             =   240
         Width           =   1350
      End
      Begin VB.Label lblMes 
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         Top             =   2880
         Width           =   5535
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   8790
      TabIndex        =   2
      Top             =   0
      Width           =   8790
      Begin VB.Image Image1 
         Height          =   825
         Left            =   7680
         Picture         =   "frmExcelDR.frx":5214
         Top             =   15
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "EXCEL导入向导"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "请先选择EXCEL文档，确定后开始导入。"
         Height          =   180
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   405
         Width           =   3150
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "确定(&O)"
      Height          =   375
      Index           =   0
      Left            =   6240
      TabIndex        =   1
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出(&C)"
      Height          =   375
      Index           =   2
      Left            =   7560
      TabIndex        =   0
      Top             =   5160
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7080
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "选择文件"
   End
   Begin VB.Label Label4 
      Caption         =   "注"
      ForeColor       =   &H00FF0000&
      Height          =   2415
      Left            =   120
      TabIndex        =   15
      Top             =   5160
      Visible         =   0   'False
      Width           =   5775
   End
End
Attribute VB_Name = "frmExcelDR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSqlCnn As ADODB.Connection
Private oConn As ADODB.Connection

Private m_GZ(1000) As String
Private m_FL(1000) As String
Public m_TempDB As String

Private Sub Aniomation(FileName As String)
Animation1.Open GetAppPath() & FileName
Animation1.AutoPlay = True
End Sub

Private Sub Command2_Click(Index As Integer)
Select Case Index
       Case 0
        If txtcWhcode.Text = "" Then
            MsgBox "请选择仓库！"
            Exit Sub
        End If
        Call mSave
       Case 1
       Call mReadExcel
       Case 2
       Unload Me
       Case 3
       Call OpenSheet

End Select
End Sub
Private Sub mSave()
On Error GoTo AErr

'Dim tCurSor As WaitCursor
Dim strSql As String
Dim strCostHub As String, strCostCode As String, strInvCode As String, strCppId As String
Dim strInvName As String, strInvStd As String, strBatch As String
Dim iQty As Double, iSum As Double, iGZ As Double
Dim strRep As String, cOthAmoID As String, cOthAmoID1 As String, cMemo As String
Dim DisRss As ADODB.Recordset, rst As ADODB.Recordset
Dim iPer As Long
Dim bJxFlag As Boolean
Dim strOpcode As String, strOP As String
Dim bCheckFlag As Boolean
Dim dBDate As Date, dEDate As Date
Dim strPartId As String
Dim dValue(20) As Double, sValue(20) As String
Dim bFind As Boolean
 Dim objAppExcel As Object, objWorkbook As Object, objWorkSheet As Object
 Dim i As Long
 Dim lngRow As Long, lngCount As Long
 Dim filtname As String
 Dim n As Long
 Dim strqtytmpcode As String
 
If Trim(Text1.Text) = "" Then
   MsgBox "请先选择EXCEL文档！", vbInformation + vbOKOnly, "系统提示"
    m_OK = False
   Exit Sub
End If
If Trim(Combo1.Text) = "" Then
   MsgBox "请先选择工作表！", vbInformation + vbOKOnly, "系统提示"
    m_OK = False
   Exit Sub
End If
  

 
    OpenExcelFile Trim(Text1.Text), objAppExcel, objWorkbook, objWorkSheet, Trim(Combo1.Text)
    
    
    Dim iRowCnt As Long
    Dim iColCnt As Long
    Dim sqlcolstr As String
    Dim cellValue As String
    sqlcolstr = "("
    
    iRowCnt = objWorkSheet.UsedRange.rows.Count
    iColCnt = objWorkSheet.UsedRange.Columns.Count
         
    If iRowCnt > 10000 Then iRowCnt = 10000
    If iColCnt > 500 Then iColCnt = 500
     
     If iRowCnt < 2 Then
     
     MsgBox "EXCEL中数据有误，请检查", vbInformation, "提示"
     m_OK = False
     
     Exit Sub
     End If
      
    
    
    Randomize
    tmpTableName = "tmplsdgpinputpuapp" & Int((100000 * Rnd) + 1)
    
    strSql = "if exists (select 1 from sysobjects where id = object_id('" & tmpTableName & "')) "
    strSql = strSql & "    drop table " & tmpTableName & " ; "
    strSql = strSql & "create table   " & tmpTableName & " (  "
    
    sqlcolstr = "("
    For i = 1 To iColCnt
    
     n = InStr(1, Trim(objWorkSheet.Cells(1, i)), "/")
       If n > 0 Then
            filtname = Mid(Trim(objWorkSheet.Cells(1, i)), 1, n - 1)
            
            
             If i < iColCnt Then
               strqtytmpcode = strqtytmpcode & filtname & ","
            Else
               strqtytmpcode = strqtytmpcode & filtname
            End If
      
      Else
      
        Select Case Trim(objWorkSheet.Cells(1, i))
        
'            Case "中分类"
'            filtname = "invtype"
            Case "商品编号"
            filtname = "cinvcode"
            Case "商品名称"
            filtname = "cinvname"
            Case "单位"
            filtname = "dw"
            Case "单价"
            filtname = "price"
            
        
        End Select
       
      End If
      
        If filtname <> "" Then
            If i < iColCnt Then
             
        
               strSql = strSql & filtname & " nvarchar(255) null ,"
               sqlcolstr = sqlcolstr & filtname & ","
             Else
             
               strSql = strSql & filtname & " nvarchar(255) null"
               strSql = strSql & ",cWhCode nvarchar(10) null)"
               sqlcolstr = sqlcolstr & filtname
               sqlcolstr = sqlcolstr & ",cWhCode)"
             End If
        End If
        
      
    
    Next
     
    
    
    
    'writeLog strSQL
    
    gConn.Execute (strSql)
      
    
    For lngRow = 2 To iRowCnt
        lngCount = lngCount + 1
         
        strSql = "insert into " & tmpTableName & "   " & sqlcolstr & "  values ("
        
        For i = 1 To iColCnt
         
         If i < iColCnt Then
            cellValue = Trim(objWorkSheet.Cells(lngRow, i) & "")
            If (Trim(cellValue) = "") Then
                strSql = strSql & " NULL ,"
            Else
                strSql = strSql & "N'" & cellValue & "',"
            End If
            
         Else
         
            cellValue = Trim(objWorkSheet.Cells(lngRow, i) & "")
            If (Trim(cellValue) = "") Then
                strSql = strSql & " NULL"
                strSql = strSql & ",N'" & txtcWhcode.Text & "')"
            Else
                strSql = strSql & "N'" & cellValue & "'"
                strSql = strSql & ",N'" & txtcWhcode.Text & "')"
            End If
         
         End If
            
        Next
        
         
          gConn.Execute (strSql)
          
         'writeLog strSQL
    Next
    
    '处理excel导入数据,
    
    strSql = "delete  HY_LSDG_InputpuAppdata where  cMaker='" & goLogin.cUserId & "'  " & _
    "  insert into HY_LSDG_InputpuAppdata(price,cinvcode,cinvname,dw,cdeptcode,iqty,ddate,cMaker,istats,cWhCode)" & _
     "SELECT price,cinvcode,cinvname,dw,deptcode,iqty,'" & goLogin.CurDate & "','" & goLogin.cUserId & "',0 ,cWhCode " & _
"From(SELECT price,cinvcode,cinvname,dw,cWhCode, " & _
            strqtytmpcode & _
     " From " & tmpTableName & " )T " & _
" UNPIVOT " & _
"( iqty FOR deptcode  IN " & _
  "  (" & strqtytmpcode & " )) P "

    
    gConn.Execute (strSql)
    
     
  strSql = "   update A SET A.cmamo=isnull(A.cmamo,'')+' 存货编码不存在',istats=3  FROM  HY_LSDG_InputpuAppdata   A " & _
" LEFT JOIN  inventory inv on a.cinvcode =inv.cInvCode " & _
" where  isnull(inv.cInvCode,'')=''"
 
 gConn.Execute (strSql)
 
strSql = "   update A SET A.cmamo=isnull(A.cmamo,'')+' 部门编码不存在',istats=3  FROM  HY_LSDG_InputpuAppdata   A " & _
" left join  Department  dep on dep.cDepCode =a.cdeptcode " & _
" where  isnull( dep.cDepCode,'')='' "

     gConn.Execute (strSql)
     
     
     strSql = "insert into  HY_LSDG_InputpuAppdatalist select * from HY_LSDG_InputpuAppdata with(nolock) where cMaker='" & goLogin.cUserId & "' and id not in (select isnull(id,0) from HY_LSDG_InputpuAppdatalist) "
     gConn.Execute (strSql)
     
     
     
     strSql = "if exists (select 1 from sysobjects where id = object_id('" & tmpTableName & "')) "
    strSql = strSql & "    drop table " & tmpTableName & " ; "
     
     gConn.Execute (strSql)
     
    CloseExcelFile objAppExcel, objWorkbook, objWorkSheet


'mSqlCnn.CommitTrans
bSwFlag = False


m_OK = True
'MsgBox "导入临时表完成！", vbInformation + vbOKOnly, "系统提示"
MSG1:
'Set tCurSor = Nothing
Animation1.Close
'lblMes.Caption = ""
 
Unload Me
Exit Sub
AErr:
 'Set tCurSor = Nothing
Animation1.Close
lblMes.Caption = ""
CloseExcelFile objAppExcel, objWorkbook, objWorkSheet
MsgBox "数据导入失败，请检查" & Err.Description, vbInformation + vbOKOnly, "系统提示"
End Sub
Private Sub mReadExcel()
On Error GoTo AErr


CommonDialog1.DefaultExt = "*.xlsx"
CommonDialog1.Filter = "导入单据支持格式(*.xls;*.xlsx)|*.xls;*.xlsx" ' "*.xls|*.xls|*.xlsx|*.xlsx"
CommonDialog1.DialogTitle = "选择EXCEL"
CommonDialog1.ShowOpen
Text1.Text = CommonDialog1.FileName
Call OpenSheet
Exit Sub
AErr:
MsgBox Err.Description, vbInformation + vbOKOnly, "系统提示"
End Sub
Private Sub Form_Load()
On Error GoTo AErr

lblMes.Caption = ""
m_TempDB = ""
txtcWhcode.RefType = RefArchive
txtcWhcode.Init g_oLogin, "Warehouse_AA"
'Set tCurSor = Nothing
Exit Sub
AErr:
'Set tCurSor = Nothing
MsgBox Err.Description, vbInformation + vbOKOnly, "系统提示"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set mSqlCnn = Nothing
Set oConn = Nothing

End Sub

Private Sub OpenSheet()
On Error GoTo AErr

If Trim(Text1.Text) = "" Then
   MsgBox "请先选择EXCEL文档！", vbInformation + vbOKOnly, "系统提示"
   Exit Sub
End If
    Dim xlApp As Object ' New Excel.Application
    Dim XLWorkBook As Object ' New Excel.Workbook
    Dim i As Long, j As Long
    Set xlApp = CreateObject("Excel.Application")
    
    Set XLWorkBook = xlApp.Workbooks.Open(Text1.Text)
    Combo1.Clear
    Combo1.Text = ""
    For i = 1 To XLWorkBook.Worksheets.Count
        Combo1.AddItem XLWorkBook.Sheets(i).Name
    Next i
    xlApp.Quit
    If Combo1.ListCount > 0 Then Combo1.Text = Combo1.List(0)
Exit Sub
AErr:
MsgBox Err.Description, vbInformation + vbOKOnly, "系统提示"
End Sub

Private Function GetABCode(strTxt As String, strS As String) As String

    On Error GoTo AErr

    Dim sValue() As String

    sValue = Split(strTxt, strS)
    GetABCode = sValue(0)

    Exit Function

AErr:
    GetABCode = ""
End Function


'打开EXCEL文件
Public Function OpenExcelFile(ByVal strSourceFile As String, _
                                ByRef objAppExcel As Object, _
                                ByRef objWorkbook As Object, _
                                ByRef objWorkSheet As Object, _
                                ByVal vntWorkSheet As Variant) As Boolean
    Dim strFileName As String
    
    If strSourceFile = "" Then
        Exit Function
    End If
    strFileName = Mid(strSourceFile, InStrRev(strSourceFile, "\") + 1)
    
    On Error GoTo OpenErr
    Set objAppExcel = CreateObject("Excel.Application")
    objAppExcel.Workbooks.Open strSourceFile
    Set objWorkbook = objAppExcel.Workbooks(strFileName)
    Set objWorkSheet = objWorkbook.Worksheets(vntWorkSheet)
    objWorkSheet.Activate
    OpenExcelFile = True
    On Error GoTo 0
    Exit Function
OpenErr:
    MsgBox "打开EXCEL文件出错!" & vbCrLf & Err.Description, vbCritical, "信息"
End Function

'退出EXCEL文件
Public Sub CloseExcelFile(ByRef objAppExcel As Object, _
                            ByRef objWorkbook As Object, _
                            ByRef objWorkSheet As Object)
    On Error Resume Next
    objWorkbook.Close
    objAppExcel.Quit
    Set objWorkSheet = Nothing
    Set objWorkbook = Nothing
    Set objAppExcel = Nothing
    On Error GoTo 0
End Sub


Public Function GetAppPath() As String
    Dim Path$
    
    Path$ = App.Path
    If Right$(Path$, 1) <> "\" Then Path$ = Path$ + "\"
    GetAppPath = Path$
    
End Function
