VERSION 5.00
Object = "{005151D4-33F6-471B-B651-222D4E411832}#4.5#0"; "UFFormPartner.ocx"
Object = "{9FD12F62-6922-47E1-B1AC-3615BBD3D7A5}#1.0#0"; "UFLabel.ocx"
Object = "{5E4640D0-A415-404B-A457-72980C429D2F}#10.37#0"; "U8RefEdit.ocx"
Begin VB.Form frmDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "批次日期"
   ClientHeight    =   1800
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3570
   Icon            =   "frmDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin UFFormPartner.UFFrmCaption UFFrmCaption1 
      Left            =   240
      Top             =   1320
      _ExtentX        =   1085
      _ExtentY        =   450
      Caption         =   "批次日期"
      DebugFlag       =   0   'False
      SkinStyle       =   ""
   End
   Begin UFLABELLib.UFLabel UFLabel1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   111
      Caption         =   "批次日期："
   End
   Begin U8Ref.RefEdit RefEdit1 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   2055
      _ExtentX        =   3625
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
      SkinStyle       =   ""
      StopSkin        =   0   'False
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "frmDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public cmDate As String
Private Sub Form_Load()
    RefEdit1.RefType = RefDate
    RefEdit1.Text = g_oLogin.CurDate
End Sub

Private Sub OKButton_Click()
    If RefEdit1.Text = "" Then
        MsgBox "请选择日期", vbInformation, "条码系统"
        Exit Sub
    End If
    cmDate = RefEdit1.Text
    Unload Me
End Sub
