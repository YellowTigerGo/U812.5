VERSION 5.00
Object = "{005151D4-33F6-471B-B651-222D4E411832}#4.5#0"; "UFFormPartner.ocx"
Begin VB.Form FrmMsgBox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "提示"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   Icon            =   "FrmMsgBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin UFFormPartner.UFFrmCaption UFFrmCaptionMgr 
      Left            =   480
      Top             =   4320
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "提示"
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
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   390
      Left            =   4800
      TabIndex        =   1
      Top             =   4275
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Height          =   3975
      Left            =   75
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5865
   End
End
Attribute VB_Name = "FrmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    SkinSE_Init Me.hWnd, True
    Skinse_SetStopChildSkinFlag Me.hWnd
    SkinSE_SetFrameTitleText Me.hWnd, StrPtr(Me.Caption) ' UFFrmCaptionMgr.Caption为赋值窗体的标题控件 如果没有用me.caption
End Sub
