VERSION 5.00
Begin VB.Form FrmZDKJ 
   Caption         =   "制单临时"
   ClientHeight    =   2040
   ClientLeft      =   2205
   ClientTop       =   2190
   ClientWidth     =   4950
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2040
   ScaleWidth      =   4950
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Label UFFrmCaptionMgr 
      Caption         =   "制单临时"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2040
      TabIndex        =   0
      Top             =   1440
      Width           =   1200
   End
End
Attribute VB_Name = "FrmZDKJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cPZid As String
Public strsql As String
Public indxfrm As Integer
Public bHC As Boolean, bModify As Boolean
Public bSave As Boolean
Private WithEvents ARPZ As ZzPz.clsPZ
Attribute ARPZ.VB_VarHelpID = -1
Private sPZID   As String   '凭证线索

Dim pzbill As String  '因缺少变量而临时添加。

'凭证保存后的回写事件，（和凭证保存在一个事务中处理）
'by 客户化开发中心 2006/03/01
'070314_U870修改入口参数为ADODB.Recordset类型
Private Sub ARPZ_Save(rstCurrentVouch As ADODB.Recordset, IsSuccess As Boolean)
    Dim bRet As Boolean, cPZid As String
    Dim i As Long, Rst As New ADODB.Recordset, oRs As New ADODB.Recordset
    Dim sqltemp As String
    Dim AffectedLine As Long
    Dim strsql As String
    Dim ino_id As Integer   '凭证编号
    Dim objConn As ADODB.Connection
    
    On Error GoTo Jmp0
    '外部凭证数据更新为总帐凭证
    If rstCurrentVouch.RecordCount > 0 Then
        Set objConn = rstCurrentVouch.ActiveConnection
        strsql = "update GL_accvouch set coutaccset='',ioutyear='',coutsysname='',ioutperiod=0,coutsign='',doutdate=null,coutbillsign=null,coutid=null,ccodecontrol=null where csign=N'" & rstCurrentVouch!cSign & "' And iperiod=" & rstCurrentVouch!iPeriod & " And ino_id =" & rstCurrentVouch!ino_id
        objConn.Execute strsql, AffectedLine
        
        '更新表头制单人,凭证号
'            MsgBox rstCurrentVouch.Fields("coutno_id")
        If rstCurrentVouch.Fields("coutno_id") <> "" Then
            strsql = "update EFFYGL_Pcostbudget set bbuild=1,coutid='" & rstCurrentVouch.Fields("coutno_id") & "' where ccode in ('" & VBA.Replace(rstCurrentVouch.Fields("coutno_id"), ",", "','") & "')"
'            If Left(rstCurrentVouch.Fields("coutno_id"), 4) = "KJ03" Then         '预算经费核销单
'               strSql = "update kj_expenvouch set strmakerid = '" & rstCurrentVouch.Fields("cbill") & "',pzguid = '" & rstCurrentVouch.Fields("coutno_id") & "' where left(strtypecode,1)= '2' and strvouchid in ('" & Right(rstCurrentVouch.Fields("coutno_id").value, Len(rstCurrentVouch.Fields("coutno_id").value) - 4) & "')"
'            ElseIf Left(rstCurrentVouch.Fields("coutno_id"), 4) = "KJ01" Then
'               strSql = "update kj_loanvouch set strmakerid = '" & rstCurrentVouch.Fields("cbill") & "',pzguid = '" & rstCurrentVouch.Fields("coutno_id") & "' where left(strtypecode,1)= '3' and strvouchid in ('" & Right(rstCurrentVouch.Fields("coutno_id").value, Len(rstCurrentVouch.Fields("coutno_id").value) - 4) & "')"
'            End If
'            MsgBox strsql
            objConn.Execute strsql, AffectedLine
        End If
    End If
    
    bRet = True
    IsSuccess = bRet
    bSave = True
    Exit Sub
Jmp0:
    IsSuccess = False
    bSave = False
End Sub
Public Sub ShowMsg(ByVal sMsg As String)
    ARPZ.ShowMsg sMsg
    DoEvents
End Sub
Public Sub LoadZZPz()
    Dim cSql As String
    Dim Rst As New ADODB.Recordset
    Dim cTmpTable2 As String
    Dim cTmpTable1 As String
    Dim iPeriod As Byte, coutsign As String
    
    Set ARPZ = New clsPZ
    Set ARPZ.zzLogin = m_Login
    Set ARPZ.zzSys = Pubzz
    ARPZ.NewFLCode = "AP,#"
        
    cTmpTable1 = TmpZD1
    cTmpTable2 = TmpZD2

'    Account.dbMData.Execute "update " & Pubzz.WbTableName & " set inid = ltrim(right(coutid,len(coutid)-charindex(' ',coutid)))"
    
    Call Pubzz.InitPzTempTbl
    bSave = False
    If Me.bModify Then
    Else
        ARPZ.StartUpPz "AP", "GL0201", Pz_ZD, "CN"     '显示凭证
    End If
    On Error Resume Next
    
    If Not bHC Then
    '魏光增2001-01-02 modify
        If Me.bModify Then
            cSql = "Update GL_accvouch set coutid=rtrim(left(coutid,CharIndex(N' ',coutid)-1)) From GL_accvouch " & _
                    "Where coutno_id=N'" & cPZid & "'"
        Else
            cSql = "Update GL_accvouch set coutid=rtrim(left(coutid,CharIndex(N' ',coutid)-1)) From GL_accvouch a," & cTmpTable2 & " b" & _
                    " Where ltrim(Substring(coutid,CharIndex(N' ',coutid),len(coutid)))=b.cmergeno and a.coutsysname=b.t_coutsysname" & _
                    " and CharIndex(N' ',coutid)>0 And a.coutno_id=b.t_coutno_id"
        End If
        If bSave Then Account.dbMData.Execute cSql
    Else
        cSql = "Update GL_accvouch set coutsysname=null From GL_accvouch Where coutno_id in (N'" & cPZid & "',N'" & sPZID & "')"
        If bSave Then Account.dbMData.Execute cSql
    End If
        
    Account.dbMData.Execute "DELETE " + Pubzz.WbTableName
    Account.dbMData.Execute "Drop Table " & TmpZD1
    Account.dbMData.Execute "Drop Table " & TmpZD2
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Unload(Cancel As Integer)
'    Call WriteResourceLog
End Sub
Private Sub m_objPz_Save(rstCurrentVouch As UfDbKit.UfRecordset, IsSuccess As Boolean)

End Sub
