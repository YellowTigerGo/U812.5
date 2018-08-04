VERSION 5.00
Object = "{005151D4-33F6-471B-B651-222D4E411832}#4.5#0"; "UFFormPartner.ocx"
Object = "{AF8BBBB7-94C6-4772-B826-624478C37D6A}#1.5#0"; "UFKEYHOOK.ocx"
Object = "{9FD12F62-6922-47E1-B1AC-3615BBD3D7A5}#1.0#0"; "UFLabel.ocx"
Object = "{5E4640D0-A415-404B-A457-72980C429D2F}#10.37#0"; "U8RefEdit.ocx"
Begin VB.Form frmScan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "扫描"
   ClientHeight    =   8940
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin UFLABELLib.UFLabel UFLabel1 
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   120
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   450
      _StockProps     =   111
      Caption         =   "原条码："
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
   Begin U8Ref.RefEdit txtQRCodeOld 
      Height          =   375
      Left            =   2400
      TabIndex        =   27
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
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
   Begin UFLABELLib.UFLabel UFLabel2 
      Height          =   255
      Left            =   1800
      TabIndex        =   26
      Top             =   600
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   450
      _StockProps     =   111
      Caption         =   "条"
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
   Begin UFLABELLib.UFLabel lblCount 
      Height          =   255
      Left            =   840
      TabIndex        =   25
      Top             =   600
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   111
      Caption         =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin UFLABELLib.UFLabel UFLabel3 
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   600
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   450
      _StockProps     =   111
      Caption         =   "扫描："
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
   Begin UFKeyHook.UFKeyHookCtrl UFKeyHookCtrl1 
      Left            =   5280
      Top             =   6840
      _ExtentX        =   1905
      _ExtentY        =   529
   End
   Begin U8Ref.RefEdit txtMsg 
      Height          =   615
      Left            =   240
      TabIndex        =   23
      Top             =   8160
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1085
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
      ForeColor       =   255
      MultiLine       =   -1  'True
      LockedEdit      =   -1  'True
      SkinStyle       =   ""
      StopSkin        =   0   'False
   End
   Begin U8Ref.RefEdit txtQRCode 
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   600
      Width           =   3615
      _ExtentX        =   6376
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
      ForeColor       =   255
      SkinStyle       =   ""
      StopSkin        =   0   'False
   End
   Begin U8Ref.RefEdit txt 
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   1
      Top             =   1080
      Width           =   3615
      _ExtentX        =   6376
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
   Begin UFLABELLib.UFLabel lbl 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   111
      Caption         =   "识别码："
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
   Begin UFFormPartner.UFFrmCaption UFF 
      Left            =   6840
      Top             =   0
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "扫描"
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
   Begin UFLABELLib.UFLabel lbl 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   111
      Caption         =   "生产订单号："
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
   Begin U8Ref.RefEdit txt 
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   3
      Top             =   1680
      Width           =   3615
      _ExtentX        =   6376
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
   Begin UFLABELLib.UFLabel lbl 
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   111
      Caption         =   "行号："
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
   Begin U8Ref.RefEdit txt 
      Height          =   375
      Index           =   2
      Left            =   2400
      TabIndex        =   6
      Top             =   2280
      Width           =   3615
      _ExtentX        =   6376
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
   Begin U8Ref.RefEdit txt 
      Height          =   375
      Index           =   3
      Left            =   2400
      TabIndex        =   7
      Top             =   2880
      Width           =   3615
      _ExtentX        =   6376
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
   Begin UFLABELLib.UFLabel lbl 
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   111
      Caption         =   "材料批次号："
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
   Begin UFLABELLib.UFLabel lbl 
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   111
      Caption         =   "料号："
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
   Begin UFLABELLib.UFLabel lbl 
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   10
      Top             =   4080
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   111
      Caption         =   "流水号："
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
   Begin UFLABELLib.UFLabel lbl 
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   11
      Top             =   4680
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   111
      Caption         =   "容量："
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
   Begin UFLABELLib.UFLabel lbl 
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   12
      Top             =   5280
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   111
      Caption         =   "内阻："
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
   Begin UFLABELLib.UFLabel lbl 
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   13
      Top             =   5880
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   111
      Caption         =   "自放电："
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
   Begin U8Ref.RefEdit txt 
      Height          =   375
      Index           =   4
      Left            =   2400
      TabIndex        =   14
      Top             =   3480
      Width           =   3615
      _ExtentX        =   6376
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
   Begin U8Ref.RefEdit txt 
      Height          =   375
      Index           =   5
      Left            =   2400
      TabIndex        =   15
      Top             =   4080
      Width           =   3615
      _ExtentX        =   6376
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
   Begin U8Ref.RefEdit txt 
      Height          =   375
      Index           =   6
      Left            =   2400
      TabIndex        =   16
      Top             =   4680
      Width           =   3615
      _ExtentX        =   6376
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
   Begin U8Ref.RefEdit txt 
      Height          =   375
      Index           =   7
      Left            =   2400
      TabIndex        =   17
      Top             =   5280
      Width           =   3615
      _ExtentX        =   6376
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
   Begin U8Ref.RefEdit txt 
      Height          =   375
      Index           =   8
      Left            =   2400
      TabIndex        =   18
      Top             =   5880
      Width           =   3615
      _ExtentX        =   6376
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
   Begin UFLABELLib.UFLabel lbl 
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   19
      Top             =   6480
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   111
      Caption         =   "等级："
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
   Begin UFLABELLib.UFLabel lbl 
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   20
      Top             =   7080
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   111
      Caption         =   "允许入库："
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
   Begin U8Ref.RefEdit txt 
      Height          =   375
      Index           =   9
      Left            =   2400
      TabIndex        =   21
      Top             =   6480
      Width           =   3615
      _ExtentX        =   6376
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
   Begin U8Ref.RefEdit txt 
      Height          =   375
      Index           =   10
      Left            =   2400
      TabIndex        =   22
      Top             =   7080
      Width           =   3615
      _ExtentX        =   6376
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
   Begin UFLABELLib.UFLabel lbl 
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   29
      Top             =   7680
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   111
      Caption         =   "单体数量："
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
   Begin U8Ref.RefEdit txt 
      Height          =   375
      Index           =   11
      Left            =   2400
      TabIndex        =   30
      Top             =   7680
      Width           =   3615
      _ExtentX        =   6376
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
End
Attribute VB_Name = "frmScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim m_iCount As Integer

Private m_oldQRCode As String
'Public cmocode As String
Public gaiz As Boolean '是否改制品入库
Public Frm As FrmList
Private Sub Init()
    Dim i As Integer
    For i = 0 To txt.Count - 1
        txt(i).TabStop = False
        txt(i).LockedEdit = True
    Next
    txtMsg.TabStop = False
'    txt(6).Alignment = ccRight
'    txt(7).Alignment = ccRight
'    txt(8).Alignment = ccRight
End Sub

Private Sub Form_Load()
    If Not Me.gaiz Then
        UFLabel1.Visible = False
        txtQRCodeOld.Visible = False
    Else
        UFLabel1.Visible = True
        txtQRCodeOld.Visible = True
        txtQRCodeOld.TabIndex = 1
    End If
    Init
    m_iCount = 0
    m_oldQRCode = ""
End Sub

Private Sub txtQRCode_GotFocus()
    If txtQRCode.Text <> "" Then
        txtQRCode.SelStart = 0
        txtQRCode.SelLength = Len(txtQRCode.Text)
    End If
End Sub

Private Sub txtQRCode_LostFocus()
    If gaiz Then
        If txtQRCodeOld.Text = "" Then
            txtMsg.DisplayText = "请先扫入原条码"
            txtQRCodeOld.SetFocus
            Exit Sub
        End If
        DoExecute
    End If
End Sub

Private Sub txtQRCodeOld_GotFocus()
    If txtQRCodeOld.Text <> "" Then
        txtQRCodeOld.SelStart = 0
        txtQRCodeOld.SelLength = Len(txtQRCodeOld.Text)
    End If
End Sub

Private Sub txtQRCodeOld_LostFocus()
    Dim sError As String
    If txtQRCodeOld.Text = "" Then
        Exit Sub
    End If
    
    If txtQRCode.Text <> "" Then
        DoExecute
    End If
End Sub

Private Sub UFKeyHookCtrl1_ContainerKeyDown(KeyCode As Integer, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If gaiz Then
                SendKeys "{TAB}"
            Else
                DoExecute
            End If
            
        Case vbKeyEscape
            DoEscape
    End Select
End Sub

Private Sub DoExecute()
    Dim sError As String
    
    If gaiz Then
        If Len(txtQRCodeOld.Text) <> GetQRCodeRuleLen(Left(txtQRCodeOld.Text, 1)) Then
            txtMsg.DisplayText = "原条码不符合规则"
            txtQRCodeOld.SetFocus
            Exit Sub
        End If
        If cQRCodeIsNotExist(txtQRCodeOld.Text, sError) Then
            txtMsg.DisplayText = "原条码信息不存在。"
            txtQRCodeOld.SetFocus
            Exit Sub
        End If
    End If
    If Len(txtQRCode.Text) = 0 Then
        Exit Sub
    End If
    txtQRCode.SelStart = 0
    txtQRCode.SelLength = Len(txtQRCode.Text)
'    txtQRCode.SetFocus
    If Len(txtQRCode.Text) <> GetQRCodeRuleLen(Left(txtQRCode.Text, 1)) Then
        ClearCtlValue
        txtMsg.DisplayText = "条码不符合规则"
        Exit Sub
    End If
    If m_oldQRCode = txtQRCode.Text Then
        txtMsg.DisplayText = "条码信息已扫描。"
        Exit Sub
    End If
    m_oldQRCode = txtQRCode.Text
    Analysis txtQRCode.Text
    If DoCheck Then

        If Not DoSave(sError) Then
            txtMsg.DisplayText = sError
        Else
            If gMoCode = "" Then
                gMoCode = txt(1).Text
            End If
            txtQRCode.Text = ""
            If gaiz Then
                txtQRCodeOld.SetFocus
            Else
                txtQRCode.SetFocus
            End If
            txtMsg.DisplayText = "OK"
            Frm.ExecRefresh
        End If
    End If
End Sub

Private Sub ClearCtlValue()
    Dim i As Integer
    For i = 0 To txt.Count - 1
        txt(i).Text = ""
        txt(i).DisplayText = ""
    Next
    txtQRCode.Text = ""
'    txtQRCodeOld.Text = ""
'    txtMsg.Text = ""
    m_oldQRCode = ""
End Sub
'条码解析
Private Sub Analysis(cQRcode As String)
    Dim rs As New ADODB.Recordset
    Dim strSql As String
    Dim citemCode As String
    Dim iStart As Integer
    Dim iLength As Integer
    Dim fCapacity As Double
    Dim fInternal As Double
    Dim fDischarge As Double
    Dim iDTqty As Double
    Dim irl As Double
    Dim inz As Double
    Dim izfd As Double
    On Error GoTo hErr
    
    iStart = 1
    Dim cIDNum As String
    cIDNum = Left(cQRcode, 1)
    If cIDNum = "M" Then
        strSql = "select * from EF_QRCodeRule where cLule='模组二维码规则' and bSelected=1 order by iOrder"
    Else
        strSql = "select * from EF_QRCodeRule where cLule='单体二维码规则' and bSelected=1 order by iOrder"
    End If
    
    rs.Open strSql, gConn
    If Not rs.BOF And Not rs.EOF Then
        While Not rs.EOF
            citemCode = UCase$(rs!citemCode)
            iLength = rs!icodelength
            Select Case citemCode
            
                Case "CIDNUM"
                    txt(0).Text = Mid$(cQRcode, iStart, iLength)
                Case "CMOCODE"
                    txt(1).Text = Mid$(cQRcode, iStart, iLength)
                Case "IMOSEQ"
                    txt(2).Text = Mid$(cQRcode, iStart, iLength)
                    
                Case "CMATERIALBATCH"
                    txt(3).Text = Mid$(cQRcode, iStart, iLength)
                Case "CPN"
                    txt(4).Text = Mid$(cQRcode, iStart, iLength)
                    If cIDNum <> "M" Then
                        GetTXxsw txt(4).Text, irl, inz, izfd
                    Else
                        GetMZTX txt(1).Text, txt(2).Text, txt(4).Text, fCapacity, fInternal, fDischarge, iDTqty
                        txt(11).Text = iDTqty
                        txt(6).Text = fCapacity
                        txt(7).Text = fInternal
                        txt(8).Text = fDischarge
                    End If
                Case "CSERIALNUM"
                    txt(5).Text = Mid$(cQRcode, iStart, iLength)
                Case "FCAPACITY"
                    If cIDNum = "M" Then
                    Else
                        fCapacity = ConvertStrToDbl(Mid$(cQRcode, iStart, iLength))
                    End If
                    txt(6).Text = fCapacity * irl
                    
                Case "FINTERNAL"
                    If cIDNum = "M" Then
                    Else
                        fInternal = ConvertStrToDbl(Mid$(cQRcode, iStart, iLength))
                    End If
                    txt(7).Text = fInternal * inz
                Case "FDISCHARGE"
                    If cIDNum = "M" Then
                    Else
                        fDischarge = ConvertStrToDbl(Mid$(cQRcode, iStart, iLength))
                    End If
                    txt(8).Text = fDischarge * izfd
            End Select
            iStart = iStart + iLength
            rs.MoveNext
        Wend
    End If
    Set rs = Nothing
    Exit Sub
hErr:
    Set rs = Nothing
End Sub

'特性小数位
Private Sub GetTXxsw(cPN As String, ByRef irl As Double, ByRef inz As Double, ByRef izfd As Double)
    Dim rs As New ADODB.Recordset
    Dim val As Double
    On Error GoTo hErr
    rs.Open "select cInvDefine11,cInvDefine12,cInvDefine13 from inventory where cInvDefine1='" & cPN & "'", gConn
    If Not rs.BOF And Not rs.EOF Then
        val = GetRstVal(rs, "cInvDefine11")
        txt(6).NumPoint = val
        If val = 0 Then
            irl = 1
        Else
            irl = ConvertXSW(val)
        End If
        val = GetRstVal(rs, "cInvDefine12")
        txt(7).NumPoint = val
        If val = 0 Then
            inz = 1
        Else
            inz = ConvertXSW(val)
        End If
        val = GetRstVal(rs, "cInvDefine13")
        txt(8).NumPoint = val
        If val = 0 Then
            izfd = 1
        Else
            izfd = ConvertXSW(val)
        End If
    End If
    Set rs = Nothing
    Exit Sub
hErr:
    Set rs = Nothing
End Sub

Private Function ConvertXSW(val As Double) As Double
    Dim i As Integer
    Dim r As Double
    r = 1
    On Error GoTo hErr
    For i = 1 To val
        r = r * 0.1
    Next
    ConvertXSW = r
    Exit Function
hErr:
    ConvertXSW = 1
End Function

''取得rst中的字段值，将null转换为0
Private Function GetRstVal(rst As ADODB.Recordset, FieldName As String) As Variant
    If IsNull(rst(FieldName)) = True Then
        'If rst(FieldName).Type = adChar Or rst(FieldName).Type = adVarChar Or rst(FieldName).Type = adDate Or rst(FieldName).Type = adDBDate Or rst(FieldName).Type = adDBTime Or rst(FieldName).Type = adDBTimeStamp Then
        If rst(FieldName).Type = adChar Or rst(FieldName).Type = adVarChar Or rst(FieldName).Type = adDate _
                Or rst(FieldName).Type = adDBDate Or rst(FieldName).Type = adDBTime Or rst(FieldName).Type = adDBTimeStamp _
                Or rst(FieldName).Type = adVarWChar Or rst(FieldName).Type = adLongVarChar Or rst(FieldName).Type = adLongVarWChar _
                Or rst(FieldName).Type = adWChar Or rst(FieldName).Type = adBSTR Then
            GetRstVal = ""
        Else
            GetRstVal = 0
        End If
    Else
        If rst(FieldName).Type = adBoolean Then
            GetRstVal = IIf(rst(FieldName), "是", "否")
        ElseIf rst(FieldName).Type = adDate Or rst(FieldName).Type = adDBDate Or rst(FieldName).Type = adDBTime Then
            GetRstVal = Format(rst(FieldName), "yyyy-mm-dd")
        Else
            GetRstVal = rst(FieldName)
        End If
    End If

End Function

'获取条码定义的总长度
Private Function GetQRCodeRuleLen(cIDNum As String) As Integer
    Dim rs As New ADODB.Recordset
    Dim strSql As String
    On Error GoTo hErr
    If cIDNum = "M" Then
        strSql = "select SUM(icodelength) as ilen from EF_QRCodeRule where bSelected=1 and cLule='模组二维码规则'"
    Else
        strSql = "select SUM(icodelength) as ilen from EF_QRCodeRule where bSelected=1 AND cLule='单体二维码规则'"
    End If
    rs.Open strSql, gConn
    If Not rs.BOF And Not rs.EOF Then
        GetQRCodeRuleLen = rs!iLen
    End If
    Set rs = Nothing
    Exit Function
hErr:
    Set rs = Nothing
End Function

Private Function DoCheck() As Boolean
    Dim sError As String
    Dim sErrorShow As String
    Dim cgradecode As String
    Dim cgradename As String
    Dim bstock As Boolean
    
    If gMoCode <> "" And gMoCode <> txt(1).Text Then
        txtMsg.DisplayText = "此生产订单号为" & txt(1).Text & "，与当前生产订单号不一致。"
        DoCheck = False
        Exit Function
    End If
    If cQRCodeIsExist(txtQRCode.Text, sError) Then
        txtMsg.DisplayText = sError
        DoCheck = False
        Exit Function
    End If
    If Not cmocodeIsExist(txt(1).Text, sError) Then
        sErrorShow = sErrorShow & sError
    End If
    If Not imoseqIsExist(txt(1).Text, txt(2).Text, txt(4).Text, sError) Then
        sErrorShow = sErrorShow & sError
    End If
    If Not cPNIsExist(txt(4).Text, sError) Then
        sErrorShow = sErrorShow & sError
    End If
    If txt(0).Text <> "M" Then
        If GetInventoryGradeSet(txt(4).Text, CDbl(txt(6).Text), CDbl(txt(7).Text), CDbl(txt(8).Text), cgradecode, cgradename, bstock, sError) Then
            txt(9).Text = cgradecode
            txt(9).DisplayText = cgradename
            txt(10).Text = IIf(bstock, "1", "0")
            txt(10).DisplayText = IIf(bstock, "是", "否")
        Else
            sErrorShow = sErrorShow & sError
        End If
        If txt(10).Text = "0" Then
            sErrorShow = sErrorShow & "此产品等级不允许入库"
        End If
    End If
    If Len(sErrorShow) > 0 Then
        txtMsg.DisplayText = sErrorShow
        DoCheck = False
    Else
        DoCheck = True
    End If
    
End Function

'保存扫描记录
Private Function DoSave(ByRef sError As String) As Boolean
    Dim sql As String
    Dim iType As Integer
    On Error GoTo hErr
    iType = IIf(gaiz, 11, 10)
    sql = " insert into " & TblName & "(cqrcode,cidcode,cmocode,cmaterialbatch,imoseq,cpn,cseq,fcapacity,finternal,fdischarge,cgradecode,cuser,ddate,itype,bOut,cOldQRCode)" & _
        " values('" & txtQRCode.Text & "','" & txt(0).Text & "','" & txt(1).Text & "','" & txt(3).Text & "','" & txt(2).Text & "','" & txt(4).Text & "','" & txt(5).Text & "'," & _
        "'" & txt(6).Text & "','" & txt(7).Text & "','" & txt(8).Text & "','" & txt(9).Text & "','" & g_oLogin.cUserName & "','" & g_oLogin.CurDate & "'," & iType & ",0,'" & txtQRCodeOld.Text & "')"
    gConn.Execute sql
    m_iCount = m_iCount + 1
    lblCount = Format(m_iCount, "#,##0")
    DoSave = True
    Exit Function
hErr:
    sError = "条码保存失败：" & Err.Description
End Function

Private Sub DoEscape()
    Unload Me
End Sub

'校验生产订单号是否存在
Private Function cmocodeIsExist(cmocode As String, ByRef sError As String) As Boolean
    Dim rs As New ADODB.Recordset
    On Error GoTo hErr
    rs.Open "SELECT MOCODE FROM mom_order WHERE MOCODE='" & cmocode & "' ", gConn
    If Not rs.BOF And Not rs.EOF Then
        cmocodeIsExist = True
    Else
        sError = "生产订单号不存在！"
    End If
    Set rs = Nothing
    Exit Function
hErr:
    Set rs = Nothing
    sError = Err.Description
End Function


'校验生产订单行号是否存在
Private Function imoseqIsExist(cmocode As String, imoseq As String, cPN As String, ByRef sError As String) As Boolean
    Dim rs As New ADODB.Recordset
    On Error GoTo hErr
    rs.Open "select i.cInvDefine1 from mom_order t left outer join mom_orderdetail b on t.MoId =b.MoId left outer join inventory i on i.cInvCode=b.InvCode where MOCODE='" & cmocode & "' and SortSeq=" & imoseq, gConn
    If Not rs.BOF And Not rs.EOF Then
        If rs!cInvDefine1 & "" <> cPN Then
            sError = cPN & "不属于此生产订单！"
        Else
            imoseqIsExist = True
        End If
    Else
        sError = "生产订单行号不存在！"
    End If
    Set rs = Nothing
    Exit Function
hErr:
    Set rs = Nothing
    sError = Err.Description
End Function

'校验条码是否已扫描
Private Function cQRCodeIsExist(cQRcode As String, ByRef sError As String) As Boolean
    Dim rs As New ADODB.Recordset
    On Error GoTo hErr
    rs.Open "select cqrcode from " & TblName & " where cqrcode='" & cQRcode & "' union all select cQRCode from EF_InScanDetail where isnull(bOut,0)=0 and cQRCode='" & cQRcode & "'", gConn
    If Not rs.BOF And Not rs.EOF Then
        cQRCodeIsExist = True
        sError = "条码信息已扫描！"
    End If
    Set rs = Nothing
    Exit Function
hErr:
    Set rs = Nothing
    sError = Err.Description
End Function

'校验原条码是否已存在
Private Function cQRCodeIsNotExist(cQRcode As String, ByRef sError As String) As Boolean
    Dim rs As New ADODB.Recordset
    On Error GoTo hErr
    rs.Open "select cqrcode from EF_InScanDetail where cqrcode='" & cQRcode & "'", gConn
    If Not rs.BOF And Not rs.EOF Then
    Else
        cQRCodeIsNotExist = True
        sError = "条码信息不存在"
    End If
    Set rs = Nothing
    Exit Function
hErr:
    Set rs = Nothing
    sError = Err.Description
End Function

'校验料号是否存在
Private Function cPNIsExist(cPN As String, ByRef sError As String) As Boolean
    Dim rs As New ADODB.Recordset
    On Error GoTo hErr
    rs.Open "select cInvCode from inventory where cInvDefine1='" & cPN & "' ", gConn
    If Not rs.BOF And Not rs.EOF Then
        cPNIsExist = True
    Else
        sError = "料号不存在！"
    End If
    Set rs = Nothing
    Exit Function
hErr:
    Set rs = Nothing
    sError = Err.Description
End Function

Private Function GetInventoryGradeSet(cPN As String, fCapacity As Double, fInternal As Double, fDischarge As Double, ByRef cgradecode As String, ByRef cgradename As String, ByRef bstock As Boolean, ByRef sError As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim sql As String
    On Error GoTo hErr
    sql = " select cgradecode,cgradename,bstock from EF_v_InventoryGradeSet where cpn='" & cPN & "' " & _
          " and " & fCapacity & ">=fCapacityStart and " & fCapacity & "<fCapacityEnd" & _
          " and " & fInternal & ">=fInternalStart and " & fInternal & "<finternalend" & _
          " and " & fDischarge & ">=fDischargeStart and " & fDischarge & "<fDischargeend "
    rs.Open sql, gConn
    If Not rs.BOF And Not rs.EOF Then
        GetInventoryGradeSet = True
        cgradecode = rs!cgradecode
        cgradename = rs!cgradename
        bstock = rs!bstock
    Else
        sError = "未找到匹配的等级信息！"
    End If
    Set rs = Nothing
    Exit Function
hErr:
    Set rs = Nothing
    sError = Err.Description
End Function

'获取模组特性值：即单体特性值合计
Private Sub GetMZTX(cmocode As String, imoseq As String, cPN As String, ByRef fCapacity As Double, ByRef fInternal As Double, ByRef fDischarge As Double, ByRef iDTqty As Double)
    Dim rs As New ADODB.Recordset
    Dim strSql As String
    Dim val As Double
    On Error GoTo hErr
    strSql = "select sum(d.fCapacity) as fCapacity,SUM(d.finternal) as finternal,SUM(fdischarge) as fdischarge,count(cqrcode) as iDTqty from rdrecord11 t " & _
            " left outer join rdrecords11 r on t.ID =r.ID left outer join inventory i on r.cInvCode=i.cInvCode" & _
            " left outer join EF_OutScanDetail d on t.cCode=d.ccode and d.itype='20'  AND D.cPN=I.cInvDefine1" & _
            " where r.cmocode='" & cmocode & "' and r.imoseq=" & imoseq
    rs.Open strSql, gConn
    If Not rs.BOF And Not rs.EOF Then
        fCapacity = ConvertStrToDbl(rs!fCapacity)
        fInternal = ConvertStrToDbl(rs!fInternal)
        fDischarge = ConvertStrToDbl(rs!fDischarge)
        iDTqty = ConvertStrToDbl(rs!iDTqty)
    End If
    rs.Close
    Set rs = Nothing
    Exit Sub
hErr:
    Set rs = Nothing
End Sub


