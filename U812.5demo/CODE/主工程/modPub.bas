Attribute VB_Name = "modpub_PZ"
Option Explicit
Public Const TotalColor = &HC7F3F7
Private Const WM_NEXTDLGCTL = &H28
Dim m_oDataSource As Object
'-----------------api函数Unicode
Public oapi As Object
Public Declare Function GetTempPathW Lib "unicows" (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long
Public Declare Function GetEnvironmentVariableW Lib "unicows" (ByVal lpName As Long, ByVal lpBuffer As Long, ByVal nSize As Long) As Long


Public agcZDStyle(10, 2) As String

Public Type Taccount
    wksUfM          As Workspace        '工作区变量名称
    dbMData         As ADODB.Connection '数据库连接
    UfDbData        As UfDbKit.UfDatabase
    dbMTmp          As Database         '临时数据库变量名称
    cTempPath       As String           '临时数据库路径
    iExchRateDecDgt As Byte
End Type

Public Account          As Taccount
Public Pubzz            As ZzPub.clsPub
Public TmpZD1           As String
Public TmpZD2           As String
Public g_oLogin           As U8Login.clsLogin
'全局常量
Global Const EditColor = 16777215 'RGB(255, 255, 255)
Global Const DisColor = &HE8E8E8 '14143693 'RGB(205, 208, 215)
Global Const RowColor = 15527415 'RGB(247, 237, 236)
Global Const EspColor = 14085870 'RGB(238, 238, 214)
Global Const MustColor = &HCC3333
Global Const SCROLLWIDTH = 235


'取计算机名
Public Function MyComputer() As String
    Dim i As Long, cTemp As String * 128
    
    i = 128
    Call oapi.GetComputerNameA(cTemp, i)
    MyComputer = LeftEx(cTemp, i)
End Function

'打开记录集
Public Function OpenRecordset(ByVal Source As String, Optional ByVal CursorType As CursorTypeEnum = adOpenDynamic, _
    Optional LockType As LockTypeEnum = adLockReadOnly) As ADODB.Recordset
    Dim rstmp As New ADODB.Recordset
    On Error GoTo Err:
    rstmp.CursorType = CursorType
    rstmp.LockType = LockType
    rstmp.Open Source, Account.dbMData
    Set OpenRecordset = rstmp
    Set rstmp = Nothing
Err:
    Set rstmp = Nothing

End Function


'解决键盘死锁
Public Sub SetNextTab(oCtrl As Object, Frm As Form)
    oCtrl.Enabled = False
    oapi.SendMessageA Frm.hwnd, WM_NEXTDLGCTL, 0, 0
    oCtrl.Enabled = True
End Sub


'窗体是否装入
Public Function FrmIsShow(frmName As String) As Boolean
    Dim Frm As Form
    
    FrmIsShow = False
    For Each Frm In Forms
        If UCase(Frm.Name) = UCase(frmName) Then
            FrmIsShow = True
            Frm.ZOrder
            Set Frm = Nothing
            Exit Function
        End If
    Next
    Set Frm = Nothing
End Function


'功能：Null校验
'参数：FldValue任何合法数据类型数据(包括Null)
'返回：若FldValue=Null,返回空字符串 ""；否则返回FldValue。
Public Function NullToStr(ByVal FldValue) As Variant
    NullToStr = IIf(IsNull(FldValue), "", FldValue)
End Function

'空串转换为Null值
Public Function StrToNull(ByVal Str As String) As Variant
    If Str = "" Then
       StrToNull = Null
    Else
       StrToNull = Str
    End If
End Function

Public Function NullToZero(vValue) As Variant
    NullToZero = vValue
    If IsNull(vValue) Then NullToZero = 0
End Function


Public Sub DoGrid(Grid As Object)
    Dim i   As Long
    Dim sName$

On Error Resume Next
    With Grid
        sName = LCase(TypeName(Grid))
        Select Case sName
        Case "msflexgrid", "mshflexgrid", "supergrid", "ufflexgridctl"
            .BackColorSel = &H9F6646   'RowColor
            '.ForeColorSel = vbBlack
            .BackColor = EditColor
            .BackColorFixed = EditColor ' &H8000000F
            .BackColorFixed = &HFFE3C6   '&H8000000F
            .GridColor = &HAEAEAE
            .GridColorFixed = &H888888
            For i = 0 To .cols - 1
                .FixedAlignment(i) = flexAlignCenterCenter
                If .colwidth(i) = 0 Then
                    .TextArray(i) = ""
                    .ColAlignment(i) = flexAlignRightCenter
                End If
            Next
        Case Else
        End Select
    End With
End Sub

Public Sub DoForm(Frm As Form, BorderStyle As Integer)
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





