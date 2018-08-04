VERSION 5.00
Begin VB.Form frmZPHX 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "支票核销"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7380
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin UFHBTVmain.ctlGrid ctlGrid1 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _ExtentX        =   10927
      _ExtentY        =   5953
   End
End
Attribute VB_Name = "frmZPHX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_CtlIF             As IPlugInExCtl     '组件操作接口

Public Function LoadData(ByVal sDepCode As String, _
                         ByVal sItemCode As String, _
                         ByRef conn As ADODB.Connection, _
                         ByRef oLogin As Object)

    Dim oGrid As clsGridZPHX
    Set oGrid = New clsGridZPHX
    m_CtlIF.Init oLogin, oGrid
    m_CtlIF.DoOtherOperation "SetForm", Me
    m_CtlIF.DoOtherOperation "SetConnection", conn
    m_CtlIF.SetData "<data cdepcode='" & sDepCode & "' citemcode='" & sItemCode & "'/>"

End Function

Private Sub Form_Load()
    Set m_CtlIF = ctlGrid1.Object
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_CtlIF = Nothing
End Sub

