VERSION 5.00
Object = "{9FD12F62-6922-47E1-B1AC-3615BBD3D7A5}#1.0#0"; "UFLabel.ocx"
Object = "{AF8BBBB7-94C6-4772-B826-624478C37D6A}#1.5#0"; "UFKEYHOOK.ocx"
Object = "{005151D4-33F6-471B-B651-222D4E411832}#4.5#0"; "UFFormPartner.ocx"
Object = "{5E4640D0-A415-404B-A457-72980C429D2F}#10.37#0"; "U8RefEdit.ocx"
Object = "{D5646CCD-3DEF-4356-8564-4C2AB79D21E9}#2.3#0"; "UFRadio.ocx"
Object = "{BF022F1C-E440-4790-987F-252926B9B602}#5.1#0"; "UFFrames.ocx"
Begin VB.Form frmRouting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "报工"
   ClientHeight    =   8430
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6795
   Icon            =   "frmRouting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin UFKeyHook.UFKeyHookCtrl UFKeyHookCtrl1 
      Left            =   1200
      Top             =   7800
      _ExtentX        =   1905
      _ExtentY        =   529
   End
   Begin UFFormPartner.UFFrmCaption UFFrmCaption1 
      Left            =   960
      Top             =   7680
      _ExtentX        =   450
      _ExtentY        =   238
      Caption         =   "报工"
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
   Begin VB.CommandButton btnDel 
      Caption         =   "删除"
      Height          =   375
      Left            =   4080
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1095
   End
   Begin UFFrames.UFFrame UFFrame1 
      Height          =   7575
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   13361
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin U8Ref.RefEdit txt 
         Height          =   375
         Index           =   0
         Left            =   1680
         TabIndex        =   12
         Top             =   1440
         Width           =   4695
         _ExtentX        =   8281
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
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   111
         Caption         =   "生产员："
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
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   4095
         Begin UFRadioLib.UFRadio UFRadio2 
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            Style           =   0
            Caption         =   "白"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   0
            Appearance      =   1
            BackColor       =   -2147483648
            DisabledPicture =   "frmRouting.frx":000C
            DownPicture     =   "frmRouting.frx":0028
            Enabled         =   -1  'True
            ForeColor       =   -2147483630
            MaskColor       =   12632256
            MouseIcon       =   "frmRouting.frx":0044
            MousePointer    =   0
            Picture         =   "frmRouting.frx":0060
            OLEDropMode     =   0
            RightToLeft     =   0   'False
            UseMaskColor    =   0   'False
            Value           =   -1  'True
            SkinStyle       =   ""
         End
         Begin UFRadioLib.UFRadio UFRadio2 
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            Style           =   0
            Caption         =   "晚"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   0
            Appearance      =   1
            BackColor       =   -2147483648
            DisabledPicture =   "frmRouting.frx":007C
            DownPicture     =   "frmRouting.frx":0098
            Enabled         =   -1  'True
            ForeColor       =   -2147483630
            MaskColor       =   12632256
            MouseIcon       =   "frmRouting.frx":00B4
            MousePointer    =   0
            Picture         =   "frmRouting.frx":00D0
            OLEDropMode     =   0
            RightToLeft     =   0   'False
            UseMaskColor    =   0   'False
            Value           =   0   'False
            SkinStyle       =   ""
         End
      End
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   4095
         Begin UFRadioLib.UFRadio UFRadio1 
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            Style           =   0
            Caption         =   "产"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   0
            Appearance      =   1
            BackColor       =   -2147483648
            DisabledPicture =   "frmRouting.frx":00EC
            DownPicture     =   "frmRouting.frx":0108
            Enabled         =   -1  'True
            ForeColor       =   -2147483630
            MaskColor       =   12632256
            MouseIcon       =   "frmRouting.frx":0124
            MousePointer    =   0
            Picture         =   "frmRouting.frx":0140
            OLEDropMode     =   0
            RightToLeft     =   0   'False
            UseMaskColor    =   0   'False
            Value           =   -1  'True
            SkinStyle       =   ""
         End
         Begin UFRadioLib.UFRadio UFRadio1 
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            Style           =   0
            Caption         =   "修"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   0
            Appearance      =   1
            BackColor       =   -2147483648
            DisabledPicture =   "frmRouting.frx":015C
            DownPicture     =   "frmRouting.frx":0178
            Enabled         =   -1  'True
            ForeColor       =   -2147483630
            MaskColor       =   12632256
            MouseIcon       =   "frmRouting.frx":0194
            MousePointer    =   0
            Picture         =   "frmRouting.frx":01B0
            OLEDropMode     =   0
            RightToLeft     =   0   'False
            UseMaskColor    =   0   'False
            Value           =   0   'False
            SkinStyle       =   ""
         End
         Begin UFRadioLib.UFRadio UFRadio1 
            Height          =   255
            Index           =   2
            Left            =   2760
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            Style           =   0
            Caption         =   "检"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   0
            Appearance      =   1
            BackColor       =   -2147483648
            DisabledPicture =   "frmRouting.frx":01CC
            DownPicture     =   "frmRouting.frx":01E8
            Enabled         =   -1  'True
            ForeColor       =   -2147483630
            MaskColor       =   12632256
            MouseIcon       =   "frmRouting.frx":0204
            MousePointer    =   0
            Picture         =   "frmRouting.frx":0220
            OLEDropMode     =   0
            RightToLeft     =   0   'False
            UseMaskColor    =   0   'False
            Value           =   0   'False
            SkinStyle       =   ""
         End
      End
      Begin UFLABELLib.UFLabel UFLabel1 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   111
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
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   111
         Caption         =   "产品条码："
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
         Left            =   1680
         TabIndex        =   14
         Top             =   1920
         Width           =   4695
         _ExtentX        =   8281
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
         Left            =   120
         TabIndex        =   15
         Top             =   2520
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   111
         Caption         =   "设备："
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
         Left            =   1680
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2400
         Width           =   4695
         _ExtentX        =   8281
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
         SkinStyle       =   ""
         StopSkin        =   0   'False
      End
      Begin UFLABELLib.UFLabel lbl 
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   3000
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   111
         Caption         =   "车型："
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin U8Ref.RefEdit txt 
         Height          =   375
         Index           =   3
         Left            =   1680
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2880
         Width           =   4695
         _ExtentX        =   8281
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
         Index           =   4
         Left            =   120
         TabIndex        =   19
         Top             =   3480
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   111
         Caption         =   "产品："
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin U8Ref.RefEdit txt 
         Height          =   375
         Index           =   4
         Left            =   1680
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   3360
         Width           =   4695
         _ExtentX        =   8281
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
         Index           =   5
         Left            =   120
         TabIndex        =   21
         Top             =   3960
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   111
         Caption         =   "工序编码："
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin U8Ref.RefEdit txt 
         Height          =   375
         Index           =   5
         Left            =   1680
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3840
         Width           =   4695
         _ExtentX        =   8281
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
         Index           =   6
         Left            =   120
         TabIndex        =   23
         Top             =   4920
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   111
         Caption         =   "数量："
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
         Index           =   6
         Left            =   1680
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   4800
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
         Alignment       =   1
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
         Property        =   3
         SkinStyle       =   ""
         StopSkin        =   0   'False
      End
      Begin UFLABELLib.UFLabel lbl 
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   25
         Top             =   5400
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   111
         Caption         =   "备注："
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
         Index           =   7
         Left            =   1680
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   5280
         Width           =   4695
         _ExtentX        =   8281
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
         Index           =   8
         Left            =   120
         TabIndex        =   27
         Top             =   5880
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Enabled         =   0   'False
      End
      Begin U8Ref.RefEdit txt 
         Height          =   375
         Index           =   8
         Left            =   1680
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   5760
         Width           =   4695
         _ExtentX        =   8281
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
         Height          =   1095
         Index           =   9
         Left            =   120
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   6240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   1931
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
         Locked          =   -1
         ScrollBars      =   1
         LockedEdit      =   -1  'True
         SkinStyle       =   ""
         StopSkin        =   0   'False
      End
      Begin UFLABELLib.UFLabel lbl 
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   31
         Top             =   4440
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   111
         Caption         =   "工序名称："
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin U8Ref.RefEdit txt 
         Height          =   375
         Index           =   10
         Left            =   1680
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   4320
         Width           =   4695
         _ExtentX        =   8281
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
   Begin VB.CommandButton btnCancel 
      Caption         =   "退出"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "保存"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   7920
      Width           =   1215
   End
End
Attribute VB_Name = "frmRouting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_FormStatus As String '0:查看；1：增加；2：修改
Private m_id As String

Public Property Get formStatus() As String
    formStatus = m_FormStatus
End Property

Public Property Let formStatus(ByVal vNewValue As String)
    m_FormStatus = vNewValue
End Property

Public Property Get id() As String
    id = m_id
End Property

Public Property Let id(ByVal vNewValue As String)
    m_id = vNewValue
End Property

Private Sub init()
    txt(0).refType = RefArchive
    txt(0).init g_oLogin, "HR_HI_Person_AAGH"
    txt(1).refType = RefNone
    txt(2).refType = RefArchive
    txt(2).init g_oLogin, "EQData_MM"
    txt(2).RetStyle = Code_CodeName
    txt(3).refType = RefNone
    txt(4).refType = RefNone
    txt(4).init g_oLogin, "Inventory_AA"
    txt(4).RetStyle = Code_CodeName
    txt(5).refType = RefNone
'    txt(5).Init g_oLogin, "StdOperation_MM"
    txt(6).refType = RefCalculator
    txt(7).refType = RefRichText
    txt(8).refType = RefNone
    txt(3).Locked = True
    txt(4).LockedEdit = True
    txt(5).Locked = True
    txt(10).Locked = True
    txt(8).Locked = True
    LoadData
    DoEnabled
End Sub

Private Sub LoadData()
    Dim rs As New ADODB.Recordset
    Dim strsql As String
    On Error GoTo hErr
    
    If formStatus = 1 Then
        
    ElseIf formStatus = 2 Then
        strsql = "select * from EF_V_Routing where wxid='" & id & "'"
        rs.Open strsql, DBconn
        If Not rs.EOF And Not rs.BOF Then
            txt(0).Text = rs!cpersoncode & ""
            txt(0).DisplayText = rs!cpersonname & ""
            txt(1).Text = rs!cqrcode & ""
            txt(2).Text = rs!ceqcode & ""
            txt(2).DisplayText = rs!ceqname & ""
            txt(3).Text = rs!cmodel & ""
            txt(4).Text = rs!cInvCode & ""
            txt(4).DisplayText = rs!cinvname & ""
            txt(5).Text = rs!cprocedureid & ""
            txt(10).Text = rs!cprocedurename & ""
            txt(6).Text = CDbl(rs!iQty & "")
            txt(7).Text = rs!cmemo & ""
            txt(8).Text = rs!cRejectReason & ""
            If rs!cSource & "" = "产" Then
                UFRadio1(0).Value = True
            ElseIf rs!cSource & "" = "修" Then
                UFRadio1(1).Value = True
            Else
                UFRadio1(2).Value = True
            End If
            If rs!cShift & "" = "白" Then
                UFRadio2(0).Value = True
            Else
                UFRadio2(1).Value = True
            End If
        End If
    End If
    GoTo hFinish
hErr:

hFinish:
    If rs.State <> 0 Then rs.Close
    Set rs = Nothing
End Sub

Private Function DoCheck() As Boolean
    If txt(0).Text = "" Then
        MsgBox "生产员不能为空。", vbInformation, "提示"
        DoCheck = False
        txt(0).SetFocus
        Exit Function
    End If
    If txt(1).Text = "" Then
        MsgBox "产品条码不能为空。", vbInformation, "提示"
        DoCheck = False
        txt(1).SetFocus
        Exit Function
    End If
    If str2Dbl(txt(6).Text) <= 0 Then
        MsgBox "数量不能为空或为零。", vbInformation, "提示"
        DoCheck = False
        txt(6).SetFocus
        Exit Function
    End If
    If formStatus = 2 Then
        If IsQualityCheck(id) Then
            MsgBox "报工已检验", vbInformation, "提示"
            DoCheck = False
            Exit Function
        End If
    End If
    DoCheck = True
End Function

Private Function DoSave() As Boolean
    Dim strsql As String
    Dim cSource As String
    Dim cShift As String
    Dim cGuid As String
    If UFRadio1(0).Value = True Then
        cSource = "产"
    ElseIf UFRadio1(1).Value = True Then
        cSource = "修"
    ElseIf UFRadio1(2).Value = True Then
        cSource = "检"
    End If
    If UFRadio2(0).Value = True Then
        cShift = "白"
    ElseIf UFRadio2(1).Value = True Then
        cShift = "晚"
    End If
    On Error GoTo hErr
    If formStatus = 1 Then
        cGuid = CreateGUID()
        strsql = "insert into EF_Routing(wxid,dDate,dDateTime,cSource,cShift,cEQCode,cQRcode,cinvCode,cModel,cProcedureID,cProcedureName" & _
                ",iQty,cCreater,cUser,cMemo,cPersonCode)values(" & _
                "'" & cGuid & "',CONVERT(varchar(100), GETDATE(), 23),CONVERT(varchar(100), GETDATE(), 25),'" & cSource & "','" & cShift & "'," & _
                "'" & txt(2).Text & "','" & txt(1).Text & "','" & txt(4).Text & "','" & txt(3).Text & "','" & txt(5).Text & "','" & txt(10).Text & "'," & _
                "'" & str2Dbl(txt(6).Text) & "','" & g_oLogin.cUserName & "','" & g_oLogin.cUserName & "','" & txt(7).Text & "','" & txt(0).Text & "')"
        DBconn.Execute strsql
    ElseIf formStatus = 2 Then
        strsql = "update EF_Routing set cSource='" & cSource & "',cShift='" & cShift & "',cEQCode='" & txt(2).Text & "',iQty='" & CDbl(txt(6).Text) & "'" & _
                ",cMemo='" & txt(7).Text & "',cPersonCode='" & txt(0).Text & "' where wxid='" & id & "'"
        DBconn.Execute strsql
    End If
    DoSave = True
    GoTo hFinish
hErr:
    DoSave = False
    MsgBox Err.Description, vbCritical, "提示"
hFinish:

End Function

'清空控件值
Private Sub ClearCtl()
    Dim i As Integer
    For i = 1 To txt.Count - 1
        txt(i).Text = ""
        txt(i).DisplayText = ""
    Next
'    UFRadio1(0).Value = True
'    UFRadio2(0).Value = True
    id = ""
End Sub
Private Sub DoEnabled()
    If formStatus = 1 Then
        txt(1).Locked = False
        btnDel.Enabled = False
    ElseIf formStatus = 2 Then
        txt(1).Locked = True
        btnDel.Enabled = True
    End If
    UFRadio1(0).TabStop = False
    UFRadio1(1).TabStop = False
    UFRadio1(2).TabStop = False
    UFRadio2(0).TabStop = False
    UFRadio2(1).TabStop = False
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnDel_Click()
    If MsgBox("确认是否删除？", vbOKCancel, "提示") = vbCancel Then
        Exit Sub
    End If
    If IsQualityCheck(id) Then
        MsgBox "报工已检验", vbInformation, "提示"
        Exit Sub
    End If
    If doDel Then
        DoExit
    End If
End Sub

Private Function doDel() As Boolean
    Dim strsql As String
    On Error GoTo hErr
    strsql = "delete EF_Routing where wxid='" & id & "'"
    DBconn.Execute strsql
    doDel = True
    Exit Function
hErr:
    MsgBox Err.Description, vbCritical, "提示"
End Function

Private Sub btnSave_Click()
    If Not DoCheck Then
        Exit Sub
    End If
    If DoSave Then
        If m_FormStatus = 1 Then
            ClearCtl
            formStatus = 1
            txt(0).TabStop = False
            txt(1).TabStop = True
            SendKeys "{TAB}"
'            txt(1).SetFocus
            voucherForm.UFToolbar1.FireSysCommand enumButton, "tlbRefresh"
        ElseIf m_FormStatus = 2 Then
            DoExit
        End If
        DoEnabled
    End If
End Sub

Private Sub Form_Activate()
'    txt(0).SetFocus
End Sub

Private Sub Form_Load()
    init
    LoadData
End Sub

'解析工序条码:A##车型##产品编码##产品名称##工序编码##备注##数量##其他
Private Function doAnalyseQR(QR As String) As Boolean
    Dim arrQR() As String
    Dim cInvCode As String
    Dim cinvname As String
    Dim cprocedureid As String
    Dim cprocedurename As String
    Dim ceqcode As String
    Dim ceqname As String
    
    On Error GoTo hErr
    arrQR = Split(QR, "##")
    If arrQR(0) <> "A" Then
        txt(9).DisplayText = "非工序条码"
        doAnalyseQR = False
        GoTo hFinish
    End If
    If UBound(arrQR) <> 7 Then
        txt(9).DisplayText = "不符合工序条码规则"
        doAnalyseQR = False
        GoTo hFinish
    End If
    cInvCode = arrQR(2)
    cinvname = GetcInvName(cInvCode)
    If cinvname = "" Then
        txt(9).DisplayText = "存货编码不存在"
        doAnalyseQR = False
        GoTo hFinish
    End If
    cprocedureid = arrQR(4)
    cprocedurename = arrQR(5)
    If Not Getcprocedurename(cInvCode, cprocedureid, ceqcode, ceqname) Then
        txt(9).DisplayText = cInvCode & cinvname & "产品的工艺路线中未找到工序" & cprocedureid & cprocedurename
        doAnalyseQR = False
        GoTo hFinish
    End If
    txt(2).Text = ceqcode
    txt(2).DisplayText = ceqname
    txt(3).Text = arrQR(1)
    txt(4).Text = cInvCode
    txt(4).DisplayText = cInvCode & cinvname
    txt(5).Text = cprocedureid
    txt(10).Text = cprocedurename
    txt(6).Text = str2Dbl(arrQR(6))
    txt(7).Text = arrQR(7)
    txt(9).DisplayText = ""
    doAnalyseQR = True
    GoTo hFinish
hErr:
hFinish:

End Function

'存货+工序查找设备
Private Function Getcprocedurename(cInvCode As String, cprocedureid As String, ByRef ceqcode As String, ByRef ceqname As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim strsql As String
    On Error GoTo hErr
    strsql = "select ceqcode,ceqname from EF_V_operation where cinvcode='" & cInvCode & "' and opcode='" & cprocedureid & "'"
    rs.Open strsql, DBconn
    If Not rs.BOF And Not rs.EOF Then
        ceqcode = rs!ceqcode & ""
        ceqname = rs!ceqname & ""
        Getcprocedurename = True
    End If
    GoTo hFinish
hErr:
hFinish:
    If rs.State <> 0 Then rs.Close
    Set rs = Nothing
End Function

Private Function IsQualityCheck(id As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim strsql As String
    On Error GoTo hErr
    strsql = "select * from EF_Routing where wxid='" & id & "' and isnull(cqcid,'')<>''"
    rs.Open strsql, DBconn
    If Not rs.BOF And Not rs.EOF Then
        IsQualityCheck = True
    End If
    GoTo hFinish
hErr:
    MsgBox Err.Description, vbCritical, "提示"
hFinish:
    If rs.State <> 0 Then rs.Close
    Set rs = Nothing
End Function

Private Function GetcInvName(cInvCode As String) As String
    Dim rs As New ADODB.Recordset
    Dim strsql As String
    On Error GoTo hErr
    strsql = "select * from inventory where cinvcode='" & cInvCode & "'"
    rs.Open strsql, DBconn
    If Not rs.BOF And Not rs.EOF Then
        GetcInvName = rs!cinvname & ""
    End If
    GoTo hFinish
hErr:
    MsgBox Err.Description, vbCritical, "提示"
hFinish:
    If rs.State <> 0 Then rs.Close
    Set rs = Nothing
End Function

Private Sub txt_GotFocus(index As Integer)
    txt(index).SelStart = 0
    txt(index).SelLength = Len(txt(index).Text)
End Sub

Private Sub txt_LostFocus(index As Integer)
    If index = 1 Then
        If txt(index).Text = "" Then
            Exit Sub
        End If
        If Not doAnalyseQR(txt(index).Text) Then
            txt(index).Text = ""
            btnSave.Enabled = False
        Else
            btnSave.Enabled = True
            btnSave.SetFocus
        End If
    ElseIf index = 0 Then
'        txt(1).SetFocus
    End If
End Sub

Private Sub txt_UFValidate(index As Integer, Cancle As Boolean)
    Dim cpersoncode As String
    Dim cpersonname As String
    btnSave.Enabled = True
    If index = 0 Then
        cpersoncode = txt(index).Text
        If Len(cpersoncode) > 0 Then
            If Not ExistPerson(cpersoncode, cpersonname) Then
                Cancle = True
                txt(9).DisplayText = "生产员不存在"
                txt(index).Text = ""
                txt(index).SetFocus
                btnSave.Enabled = False
            Else
'                txt(index).Text = cpersoncode
'                txt(index).DisplayText = cpersonname
                btnSave.Enabled = False
                txt(9).DisplayText = ""
            End If
        End If
    End If
End Sub

Private Function ExistPerson(ByRef cpersoncode As String, ByRef cpersonname As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim strsql As String
    On Error GoTo hErr
    strsql = "select * from hr_hi_person where cpsn_num='" & cpersoncode & "' or JobNumber ='" & cpersoncode & "'"
    rs.Open strsql, DBconn
    If Not rs.BOF And Not rs.EOF Then
        cpersoncode = rs!cpsn_num & ""
        cpersonname = rs!cpsn_name & ""
        ExistPerson = True
    End If
    GoTo hFinish
hErr:
hFinish:
    If rs.State <> 0 Then rs.Close
    Set rs = Nothing
End Function

Private Sub UFKeyHookCtrl1_ContainerKeyDown(KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    ElseIf KeyCode = vbKeyEscape Then
        DoExit
    End If
End Sub

Private Sub DoExit()
    Unload Me
End Sub
