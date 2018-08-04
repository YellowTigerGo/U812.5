VERSION 5.00
Object = "{005151D4-33F6-471B-B651-222D4E411832}#4.5#0"; "UFFormPartner.ocx"
Object = "{9FD12F62-6922-47E1-B1AC-3615BBD3D7A5}#1.0#0"; "UFLabel.ocx"
Object = "{5E4640D0-A415-404B-A457-72980C429D2F}#10.36#0"; "U8RefEdit.ocx"
Begin VB.Form frmVerify 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "删除确认"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7545
   Icon            =   "frmVerify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "删除并驳回"
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "直接删除"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin UFLABELLib.UFLabel UFLabel1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   111
      Caption         =   "驳回原因："
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
   Begin U8Ref.RefEdit RefEdit1 
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   2355
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
      SkinStyle       =   ""
      StopSkin        =   0   'False
   End
   Begin UFFormPartner.UFFrmCaption UFFrmCaption1 
      Left            =   6840
      Top             =   600
      _ExtentX        =   1085
      _ExtentY        =   450
      Caption         =   "审批处理"
      DebugFlag       =   0   'False
      SkinStyle       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
End
Attribute VB_Name = "frmVerify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public iResult As Integer  '0:驳回；1审批通过;2：取消
Public cReason As String



Private Sub btnOK_Click()
    If Option1(0).Value = True Then
        iResult = 1
    Else
        iResult = 2
        cReason = RefEdit1.Text
    End If
    
    Unload Me
End Sub


Private Sub btnCancel_Click()
    iResult = 0
    Unload Me
End Sub

Private Sub Form_Load()
    iResult = 0
End Sub
