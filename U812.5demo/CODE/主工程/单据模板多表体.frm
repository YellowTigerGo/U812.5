VERSION 5.00
Object = "{4BF46141-D335-11D2-A41B-B0AB2ED82D50}#1.0#0"; "MDIExtender.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{456334B9-D052-4643-8F5F-2326B24BE316}#6.31#0"; "UAPvouchercontrol85.ocx"
Object = "{201FB79B-5556-47A4-AD9C-A46BA0C45A44}#6.25#0"; "UFToolBarCtrl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmVouchNew1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "0"
   ClientHeight    =   7395
   ClientLeft      =   1395
   ClientTop       =   3090
   ClientWidth     =   11400
   FillColor       =   &H00004040&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab picVoucher 
      Height          =   6495
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "付印通知单"
      TabPicture(0)   =   "单据模板多表体.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Labeldjmb"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LabelVoucherName"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "U8VoucherSorter1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "vs"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "voucher"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "hs"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ComboDJMB"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ComboVTID"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Picture1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Picture2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ProgressBar1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "StBar"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "ImageList1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "封面"
      TabPicture(1)   =   "单据模板多表体.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture12"
      Tab(1).Control(1)=   "Picture11"
      Tab(1).Control(2)=   "ComboVTID1"
      Tab(1).Control(3)=   "ComboDJMB1"
      Tab(1).Control(4)=   "hs1"
      Tab(1).Control(5)=   "Voucher1"
      Tab(1).Control(6)=   "vs1"
      Tab(1).Control(7)=   "Labeldjmb1"
      Tab(1).Control(8)=   "U8VoucherSorter2"
      Tab(1).Control(9)=   "LabelVoucherName1"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "送书信息"
      TabPicture(2)   =   "单据模板多表体.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture22"
      Tab(2).Control(1)=   "Picture21"
      Tab(2).Control(2)=   "ComboVTID2"
      Tab(2).Control(3)=   "ComboDJMB2"
      Tab(2).Control(4)=   "hs2"
      Tab(2).Control(5)=   "vs2"
      Tab(2).Control(6)=   "Voucher2"
      Tab(2).Control(7)=   "Labeldjmb2"
      Tab(2).Control(8)=   "U8VoucherSorter3"
      Tab(2).Control(9)=   "LabelVoucherName2"
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "内容及印装方法"
      TabPicture(3)   =   "单据模板多表体.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture32"
      Tab(3).Control(1)=   "Picture31"
      Tab(3).Control(2)=   "ComboVTID3"
      Tab(3).Control(3)=   "ComboDJMB3"
      Tab(3).Control(4)=   "hs3"
      Tab(3).Control(5)=   "Voucher3"
      Tab(3).Control(6)=   "vs3"
      Tab(3).Control(7)=   "Labeldjmb3"
      Tab(3).Control(8)=   "U8VoucherSorter4"
      Tab(3).Control(9)=   "LabelVoucherName3"
      Tab(3).ControlCount=   10
      TabCaption(4)   =   "纸张材料"
      TabPicture(4)   =   "单据模板多表体.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "LabelVoucherName4"
      Tab(4).Control(1)=   "U8VoucherSorter5"
      Tab(4).Control(2)=   "Labeldjmb4"
      Tab(4).Control(3)=   "vs4"
      Tab(4).Control(4)=   "Voucher4"
      Tab(4).Control(5)=   "hs4"
      Tab(4).Control(6)=   "ComboDJMB4"
      Tab(4).Control(7)=   "ComboVTID4"
      Tab(4).Control(8)=   "Picture41"
      Tab(4).Control(9)=   "Picture42"
      Tab(4).ControlCount=   10
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1920
         Top             =   960
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -1
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin MSComctlLib.StatusBar StBar 
         Height          =   135
         Left            =   720
         TabIndex        =   59
         Top             =   6360
         Visible         =   0   'False
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   238
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   1
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   495
         Left            =   1800
         TabIndex        =   58
         Top             =   4800
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.PictureBox Picture42 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   -74880
         ScaleHeight     =   555
         ScaleWidth      =   10755
         TabIndex        =   54
         Top             =   1440
         Width           =   10755
      End
      Begin VB.PictureBox Picture41 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   -69480
         ScaleHeight     =   300
         ScaleWidth      =   3495
         TabIndex        =   50
         Top             =   480
         Width           =   3495
      End
      Begin VB.ComboBox ComboVTID4 
         Height          =   300
         Left            =   -66960
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   840
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.ComboBox ComboDJMB4 
         Height          =   300
         Left            =   -66960
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   840
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.PictureBox Picture32 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   -74880
         ScaleHeight     =   555
         ScaleWidth      =   10875
         TabIndex        =   45
         Top             =   1320
         Width           =   10875
      End
      Begin VB.PictureBox Picture31 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   -69480
         ScaleHeight     =   300
         ScaleWidth      =   3495
         TabIndex        =   41
         Top             =   600
         Width           =   3495
      End
      Begin VB.ComboBox ComboVTID3 
         Height          =   300
         Left            =   -66960
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1080
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.ComboBox ComboDJMB3 
         Height          =   300
         Left            =   -66960
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   720
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.PictureBox Picture22 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   -74880
         ScaleHeight     =   555
         ScaleWidth      =   10755
         TabIndex        =   36
         Top             =   1320
         Width           =   10755
      End
      Begin VB.PictureBox Picture21 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   -69360
         ScaleHeight     =   300
         ScaleWidth      =   3495
         TabIndex        =   33
         Top             =   600
         Width           =   3495
      End
      Begin VB.ComboBox ComboVTID2 
         Height          =   300
         Left            =   -66840
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1080
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.ComboBox ComboDJMB2 
         Height          =   300
         Left            =   -66840
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   720
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.PictureBox Picture12 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   -74880
         ScaleHeight     =   555
         ScaleWidth      =   10755
         TabIndex        =   28
         Top             =   1320
         Width           =   10755
      End
      Begin VB.PictureBox Picture11 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   -69360
         ScaleHeight     =   300
         ScaleWidth      =   3495
         TabIndex        =   24
         Top             =   600
         Width           =   3495
      End
      Begin VB.ComboBox ComboVTID1 
         Height          =   300
         Left            =   -66840
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1080
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.ComboBox ComboDJMB1 
         Height          =   300
         Left            =   -66840
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   120
         ScaleHeight     =   555
         ScaleWidth      =   10755
         TabIndex        =   18
         Top             =   1320
         Width           =   10755
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   5640
         ScaleHeight     =   300
         ScaleWidth      =   3495
         TabIndex        =   14
         Top             =   480
         Width           =   3495
      End
      Begin VB.ComboBox ComboVTID 
         Height          =   300
         Left            =   8160
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.ComboBox ComboDJMB 
         Height          =   300
         Left            =   8160
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.PictureBox hs 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3960
         ScaleHeight     =   270
         ScaleWidth      =   2250
         TabIndex        =   11
         Top             =   5400
         Visible         =   0   'False
         Width           =   2280
      End
      Begin UAPVoucherControl85.ctlVoucher voucher 
         Height          =   2535
         Left            =   1920
         TabIndex        =   16
         Top             =   2160
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4471
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10446406
         DisabledColor   =   16777215
         ColAlignment0   =   9
         Rows            =   20
         Cols            =   20
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ControlScrollBars=   0
         ControlAutoScales=   0
         BaseOfVScrollPoint=   0
         ShowSorter      =   0   'False
         ShowFixColer    =   0   'False
      End
      Begin VB.PictureBox vs 
         Appearance      =   0  'Flat
         Height          =   2550
         Left            =   960
         ScaleHeight     =   2520
         ScaleWidth      =   240
         TabIndex        =   17
         Top             =   2520
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox hs1 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -71040
         ScaleHeight     =   270
         ScaleWidth      =   2250
         TabIndex        =   22
         Top             =   5520
         Visible         =   0   'False
         Width           =   2280
      End
      Begin UAPVoucherControl85.ctlVoucher Voucher1 
         Height          =   1695
         Left            =   -73200
         TabIndex        =   26
         Top             =   2280
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2990
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10446406
         DisabledColor   =   16777215
         ColAlignment0   =   9
         Rows            =   20
         Cols            =   20
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ControlScrollBars=   0
         ControlAutoScales=   0
         BaseOfVScrollPoint=   0
         ShowSorter      =   0   'False
         ShowFixColer    =   0   'False
      End
      Begin VB.PictureBox vs1 
         Appearance      =   0  'Flat
         Height          =   2550
         Left            =   -74040
         ScaleHeight     =   2520
         ScaleWidth      =   240
         TabIndex        =   27
         Top             =   2640
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox hs2 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -71040
         ScaleHeight     =   270
         ScaleWidth      =   2250
         TabIndex        =   31
         Top             =   5520
         Visible         =   0   'False
         Width           =   2280
      End
      Begin VB.PictureBox vs2 
         Appearance      =   0  'Flat
         Height          =   2550
         Left            =   -74040
         ScaleHeight     =   2520
         ScaleWidth      =   240
         TabIndex        =   35
         Top             =   2640
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox hs3 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -71160
         ScaleHeight     =   270
         ScaleWidth      =   2250
         TabIndex        =   39
         Top             =   5520
         Visible         =   0   'False
         Width           =   2280
      End
      Begin UAPVoucherControl85.ctlVoucher Voucher3 
         Height          =   1695
         Left            =   -72120
         TabIndex        =   43
         Top             =   2280
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2990
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10446406
         DisabledColor   =   16777215
         ColAlignment0   =   9
         Rows            =   20
         Cols            =   20
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ControlScrollBars=   0
         ControlAutoScales=   0
         BaseOfVScrollPoint=   0
         ShowSorter      =   0   'False
         ShowFixColer    =   0   'False
      End
      Begin VB.PictureBox vs3 
         Appearance      =   0  'Flat
         Height          =   2550
         Left            =   -74160
         ScaleHeight     =   2520
         ScaleWidth      =   240
         TabIndex        =   44
         Top             =   2640
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox hs4 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -71160
         ScaleHeight     =   270
         ScaleWidth      =   2250
         TabIndex        =   48
         Top             =   5640
         Visible         =   0   'False
         Width           =   2280
      End
      Begin UAPVoucherControl85.ctlVoucher Voucher4 
         Height          =   1695
         Left            =   -72120
         TabIndex        =   52
         Top             =   2400
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2990
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10446406
         DisabledColor   =   16777215
         ColAlignment0   =   9
         Rows            =   20
         Cols            =   20
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ControlScrollBars=   0
         ControlAutoScales=   0
         BaseOfVScrollPoint=   0
         ShowSorter      =   0   'False
         ShowFixColer    =   0   'False
      End
      Begin VB.PictureBox vs4 
         Appearance      =   0  'Flat
         Height          =   2550
         Left            =   -74160
         ScaleHeight     =   2520
         ScaleWidth      =   240
         TabIndex        =   53
         Top             =   2760
         Visible         =   0   'False
         Width           =   270
      End
      Begin UAPVoucherControl85.ctlVoucher Voucher2 
         Height          =   1695
         Left            =   -72240
         TabIndex        =   57
         Top             =   2760
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2990
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10446406
         DisabledColor   =   16777215
         ColAlignment0   =   9
         Rows            =   20
         Cols            =   20
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ControlScrollBars=   0
         ControlAutoScales=   0
         BaseOfVScrollPoint=   0
         ShowSorter      =   0   'False
         ShowFixColer    =   0   'False
      End
      Begin VB.Label Labeldjmb4 
         Alignment       =   2  'Center
         BackColor       =   &H00E4C9AF&
         BackStyle       =   0  'Transparent
         Caption         =   "打印模版："
         Height          =   420
         Left            =   -68280
         TabIndex        =   5
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Labeldjmb3 
         Alignment       =   2  'Center
         BackColor       =   &H00E4C9AF&
         BackStyle       =   0  'Transparent
         Caption         =   "打印模版："
         Height          =   420
         Left            =   -68400
         TabIndex        =   6
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label Labeldjmb2 
         Alignment       =   2  'Center
         BackColor       =   &H00E4C9AF&
         BackStyle       =   0  'Transparent
         Caption         =   "打印模版："
         Height          =   420
         Left            =   -68160
         TabIndex        =   8
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label Labeldjmb1 
         Alignment       =   2  'Center
         BackColor       =   &H00E4C9AF&
         BackStyle       =   0  'Transparent
         Caption         =   "打印模版："
         Height          =   420
         Left            =   -68280
         TabIndex        =   55
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label U8VoucherSorter5 
         BackColor       =   &H80000007&
         Caption         =   "3333"
         Height          =   255
         Left            =   -71640
         TabIndex        =   51
         Top             =   840
         Width           =   600
      End
      Begin VB.Label LabelVoucherName4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "单据"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -70200
         TabIndex        =   49
         Top             =   840
         Width           =   630
      End
      Begin VB.Label U8VoucherSorter4 
         BackColor       =   &H80000007&
         Caption         =   "3333"
         Height          =   255
         Left            =   -71640
         TabIndex        =   42
         Top             =   720
         Width           =   600
      End
      Begin VB.Label LabelVoucherName3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "单据"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -70200
         TabIndex        =   40
         Top             =   720
         Width           =   630
      End
      Begin VB.Label U8VoucherSorter3 
         BackColor       =   &H80000007&
         Caption         =   "3333"
         Height          =   255
         Left            =   -71520
         TabIndex        =   34
         Top             =   720
         Width           =   600
      End
      Begin VB.Label LabelVoucherName2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "单据"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -70080
         TabIndex        =   32
         Top             =   720
         Width           =   630
      End
      Begin VB.Label U8VoucherSorter2 
         BackColor       =   &H80000007&
         Caption         =   "3333"
         Height          =   255
         Left            =   -71520
         TabIndex        =   25
         Top             =   720
         Width           =   600
      End
      Begin VB.Label LabelVoucherName1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "单据"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -70080
         TabIndex        =   23
         Top             =   720
         Width           =   630
      End
      Begin VB.Label U8VoucherSorter1 
         BackColor       =   &H80000007&
         Caption         =   "3333"
         Height          =   255
         Left            =   3480
         TabIndex        =   15
         Top             =   600
         Width           =   600
      End
      Begin VB.Label LabelVoucherName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "单据"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4920
         TabIndex        =   13
         Top             =   600
         Width           =   630
      End
      Begin VB.Label Labeldjmb 
         Alignment       =   2  'Center
         BackColor       =   &H00E4C9AF&
         BackStyle       =   0  'Transparent
         Caption         =   "打印模版："
         Height          =   420
         Left            =   6960
         TabIndex        =   12
         Top             =   600
         Width           =   1080
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6240
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tbrvoucher 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   953
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      _Version        =   393216
      Begin UFToolBarCtrl.UFToolbar UFToolbar1 
         Height          =   240
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   423
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox labXJ 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2205
      ScaleHeight     =   375
      ScaleWidth      =   675
      TabIndex        =   2
      Top             =   6480
      Visible         =   0   'False
      Width           =   675
      Begin VB.Line Line4 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         X1              =   -15
         X2              =   909
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         X1              =   15
         X2              =   15
         Y1              =   15
         Y2              =   399
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         X1              =   -90
         X2              =   855
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         X1              =   660
         X2              =   660
         Y1              =   372
         Y2              =   0
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "现结"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1095
         Left            =   240
         TabIndex        =   3
         Top             =   75
         Width           =   480
      End
   End
   Begin VB.PictureBox labZF 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1365
      ScaleHeight     =   375
      ScaleWidth      =   705
      TabIndex        =   0
      Top             =   6480
      Visible         =   0   'False
      Width           =   705
      Begin VB.Line Line9 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         X1              =   0
         X2              =   924
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line8 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         X1              =   690
         X2              =   690
         Y1              =   372
         Y2              =   0
      End
      Begin VB.Line Line6 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   384
      End
      Begin VB.Line Line5 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   939
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "作废"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   75
         TabIndex        =   1
         Top             =   75
         Width           =   480
      End
   End
   Begin UAPVoucherControl85.ctlVoucher Voucher5 
      Height          =   1695
      Left            =   2040
      TabIndex        =   56
      Top             =   1920
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2990
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10446406
      DisabledColor   =   16777215
      ColAlignment0   =   9
      Rows            =   20
      Cols            =   20
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ControlScrollBars=   0
      ControlAutoScales=   0
      BaseOfVScrollPoint=   0
      ShowSorter      =   0   'False
      ShowFixColer    =   0   'False
   End
   Begin MDIEXTENDERLibCtl.MDIExtend MDIExtend1 
      Left            =   3120
      Top             =   6720
      _cx             =   847
      _cy             =   847
      PassiveMode     =   0   'False
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   924
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "frmVouchNew1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents clsVoucherCO As EFFYVoucherCo.ClsVoucherCO_GDZC_M
Attribute clsVoucherCO.VB_VarHelpID = -1
'sl 添加
Private WithEvents clsVoucherCO1 As EFFYVoucherCo.ClsVoucherCO_GDZC_M
Attribute clsVoucherCO1.VB_VarHelpID = -1
Private WithEvents clsVoucherCO2 As EFFYVoucherCo.ClsVoucherCO_GDZC_M
Attribute clsVoucherCO2.VB_VarHelpID = -1
Private WithEvents clsVoucherCO3 As EFFYVoucherCo.ClsVoucherCO_GDZC_M
Attribute clsVoucherCO3.VB_VarHelpID = -1
Private WithEvents clsVoucherCO4 As EFFYVoucherCo.ClsVoucherCO_GDZC_M
Attribute clsVoucherCO4.VB_VarHelpID = -1
Private WithEvents clsVoucherCO5 As EFFYVoucherCo.ClsVoucherCO_GDZC_M
Attribute clsVoucherCO5.VB_VarHelpID = -1
'//
' by ahzzd 2005/06/01
'修改后的程序指定常数的值
Private Type POINTAPI
        X As Long
        Y As Long
End Type

Enum MD_EdPanelB
  Addp = 0
  Delp = 1
  EdtP = 2
End Enum

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private WithEvents ARPZ As ZzPz.clsPZ
Attribute ARPZ.VB_VarHelpID = -1
Private maxRefDate As Date ''   最晚参照日期
Private strCurVoucherNO As String
Private strVouchType As String, bReturnFlag As Boolean '记录单据类型
Private bCheckVouch As Boolean '单据的审核状态2
Public bFrmCancel As Boolean
Dim strCardNum As String        ''单据的CardNum
Dim sTemplateID As String       ''单据默认模板号码
Dim sCurTemplateID As String    ''单据当前的模板号
Dim sCurTemplateID2 As String    ''单据当前的模板号

Private m_strToolBarName As String    '单据工具栏标识
Private clsTbl As New clsAutoToolBar  '单据工具栏格式化变量

'sl 添加
Dim s1trVouchType As String     ''单据类型
Dim s1trCardNum As String        ''单据的CardNum
Dim s1TemplateID As String       ''单据默认模板号码
Dim s1CurTemplateID As String    ''单据当前的模板号
Dim s1CurTemplateID2 As String    ''单据当前的模板号

Dim s2trVouchType As String      ''单据类型
Dim s2trCardNum As String        ''单据的CardNum
Dim s2TemplateID As String       ''单据默认模板号码
Dim s2CurTemplateID As String    ''单据当前的模板号
Dim s2CurTemplateID2 As String    ''单据当前的模板号

Dim s3trVouchType As String      ''单据类型
Dim s3trCardNum As String        ''单据的CardNum
Dim s3TemplateID As String       ''单据默认模板号码
Dim s3CurTemplateID As String    ''单据当前的模板号
Dim s3CurTemplateID2 As String    ''单据当前的模板号

Dim s4trVouchType As String      ''单据类型
Dim s4trCardNum As String        ''单据的CardNum
Dim s4TemplateID As String       ''单据默认模板号码
Dim s4CurTemplateID As String    ''单据当前的模板号
Dim s4CurTemplateID2 As String    ''单据当前的模板号


Dim s5trVouchType As String      ''单据类型
Dim s5trCardNum As String        ''单据的CardNum
Dim s5TemplateID As String       ''单据默认模板号码
Dim s5CurTemplateID As String    ''单据当前的模板号
Dim s5CurTemplateID2 As String    ''单据当前的模板号

'//

Private vName As String
Private BrowFlag As Boolean '标识是否调用Voucher.browuser事件
Dim strRefFldName As String '发生参照的字段名
Private iVouchState As Integer
Private bClickCancel As Boolean
Private bClickSave As Boolean
'参照
Dim clsRefer As New UFReferC.UFReferClient
Dim clsAuth As New U8RowAuthsvr.clsRowAuth
Dim Domhead As New DOMDocument
Dim Dombody As New DOMDocument
'sl 添加
Dim Domhead1 As New DOMDocument
Dim Dombody1 As New DOMDocument
Dim Domhead2 As New DOMDocument
Dim Dombody2 As New DOMDocument
Dim Domhead3 As New DOMDocument
Dim Dombody3 As New DOMDocument
Dim Domhead4 As New DOMDocument
Dim Dombody4 As New DOMDocument
Dim Domhead5 As New DOMDocument
Dim Dombody5 As New DOMDocument

'//
Dim vNewID As Variant               '单据id
Dim iHeadIndex As Integer, iBodyIndex As Integer
Private m_UFTaskID As String
Private DomFormat As New DOMDocument
Private GetvouchNO As String
Private bFirst As Boolean
Dim strFreeName1 As String
Dim strFreeName2 As String
Dim strFreeName3 As String
Dim strFreeName4 As String
Dim strFreeName5 As String
Dim strFreeName6 As String
Dim strFreeName7 As String
Dim strFreeName8 As String
Dim strFreeName9 As String
Dim strFreeName10 As String
Private cSBVCode As String, SBVID As String, mDom As DOMDocument, oDomB As DOMDocument
Public iShowMode As Integer    ''窗体模式  0：正常 1：浏览
Private bCreditCheck  As Boolean   ''是否通过信用检查
Dim bOnceRefer As Boolean
Private ButtonTaskID As String  ''按钮任务id
Private RstTemplate As ADODB.Recordset, preVTID As String      ''保存临时的单据模版记录集
Private RstTemplate2 As New ADODB.Recordset
Dim vtidPrn() As Long ''打印模版数组
Private bfillDjmb As Boolean, vtidDJMB() As Long
Private bfillDjmb1 As Boolean, bfillDjmb2 As Boolean, bfillDjmb3 As Boolean, bfillDjmb4 As Boolean
Private vtidDJMB1() As Long, vtidDJMB2() As Long, vtidDJMB3() As Long, vtidDJMB4() As Long

Private bManBodyChecked As Boolean '' 是否手工cellcheckedPrivate
Private bRefContract As Boolean '是否参照合同，为cellcheck处理为cinvname2

Private bCloseFHSingle As Boolean
Private obj_EA As Object, DOMEA As DOMDocument, strEAXML As String ''审批流
Private bLostFocus As Boolean
Private domConfig As New DOMDocument
Private domTmp As DOMDocument
Private o_crm As Object
Private moAutoFill As Object
Private dOriVoucherWidth As Double, dOriVoucherHeight As Double
Private col(1 To 22) As Long  '用数组记录关键字所在的位置
Private bFromCurrentStock As Boolean     '存货参照是否来源于现存量

'by lg070314 增加U870支持
Private m_Cancel As Integer
Private m_UnloadMode As Integer
Dim sGuid As String
Private WithEvents m_mht As UFPortalProxyMessage.IMessageHandler
Attribute m_mht.VB_VarHelpID = -1
'///////////////zhupb
Dim strReferString As String
Public clsVoucher As New clsSaVoucher
Dim clsVoucherRefer As New clsSaRefer
Dim strsql As String


Public Property Get strToolBarName() As String
    strToolBarName = m_strToolBarName
End Property

Public Property Let strToolBarName(ByVal vNewValue As String)
    m_strToolBarName = vNewValue
End Property

Function createOrder(cdeFlg As Boolean) As Boolean
    Dim i As Integer
    Dim id1 As String
    Dim id2 As String
    Dim hasBom As Boolean
    Dim hasOrder As Boolean
    
    Dim MoDId As String
    
    Dim APIobj As Object
    Dim myLogin As Object
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    Dim rds As New ADODB.Recordset
    
    Dim clsSys As clsSystem
    Dim dataxml As New DOMDocument
    Dim nodelistL As IXMLDOMNodeList
    Dim ele As IXMLDOMElement
    Dim dataHead As New DOMDocument
    Dim dataBody As New DOMDocument
    
    Dim rsDom As DOMDocument
    Dim rdsDom As DOMDocument
    Dim rdsNdList As IXMLDOMNodeList
    Dim rdsEle As IXMLDOMElement
    Dim NdList As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim eleMent As IXMLDOMElement
    Dim errStr As String
    errStr = ""
    Set clsSys = New clsSystem
    clsSys.Init m_Login
    Set APIobj = CreateObject("MRPAPI.API_interface")
'    Set APIobj = New MRPAPI.API_interface
    
    createOrder = False
    '定义新的Login对象,初始化接口用
    Set myLogin = CreateObject("U8Login.clsLogin")
    If Not myLogin.login("DP", m_Login.DataSource, m_Login.cIYear, m_Login.cUserId, m_Login.cUserPassWord, CStr(m_Login.CurDate), m_Login.cServer) Then
'        GoTo dbError
    End If
    APIobj.Init myLogin
    
     DBConn.BeginTrans
     On Error GoTo dbError
        hasOrder = False
        hasBom = False
        
        If Not cdeFlg Then
            '不是超定额时（超定额时不对BOM进行操作）
            '*****************************生成BOM************************************
            '逐行判断有没有生成bom
            strsql = " select * from bas_part where partid in (select ParentId from bom_parent) " & " and  invcode='" & Me.Voucher.headerText("cinvcode") & "'" & _
                  "   And (IsNull(bas_part.Free1, N'') = IsNull('" & Me.Voucher.headerText("cfree1") & "', N'') Or IsNull(bas_part.Free1, N'') = N'')" & _
                  "   And (IsNull(bas_part.Free2, N'') = IsNull('" & Me.Voucher.headerText("cfree2") & "', N'') Or IsNull(bas_part.Free2, N'') = N'')" & _
                  "   And (IsNull(bas_part.Free3, N'') = IsNull('" & Me.Voucher.headerText("cfree3") & "', N'') Or IsNull(bas_part.Free3, N'') = N'')  " & _
                  "   And (IsNull(bas_part.Free4, N'') = IsNull('" & Me.Voucher.headerText("cfree4") & "', N'') Or IsNull(bas_part.Free4, N'') = N'') " & _
                  "   And (IsNull(bas_part.Free5, N'') = IsNull('" & Me.Voucher.headerText("cfree5") & "', N'') Or IsNull(bas_part.Free5, N'') = N'') " & _
                  "   And (IsNull(bas_part.Free6, N'') = IsNull('" & Me.Voucher.headerText("cfree6") & "', N'') Or IsNull(bas_part.Free6, N'') = N'') " & _
                  "   And (IsNull(bas_part.Free7, N'') = IsNull('" & Me.Voucher.headerText("cfree7") & "', N'') Or IsNull(bas_part.Free7, N'') = N'') " & _
                  "   And (IsNull(bas_part.Free8, N'') = IsNull('" & Me.Voucher.headerText("cfree8") & "', N'') Or IsNull(bas_part.Free8, N'') = N'') " & _
                  "   And (IsNull(bas_part.Free9, N'') = IsNull('" & Me.Voucher.headerText("cfree9") & "', N'') Or IsNull(bas_part.Free9, N'') = N'') " & _
                  "   And (IsNull(bas_part.Free10, N'') = IsNull('" & Me.Voucher.headerText("cfree10") & "', N'') Or IsNull(bas_part.Free10, N'') = N'') "

            Set rds = DBConn.Execute(strsql)
            If Not rds.EOF Then
                 hasBom = True
            End If
'            循环取得分段号的所有数据 , 并按物料合计
'            eidt by jiang 20080828
            strsql = "select b.cinvcode,D.cinvname,D.cInvAddCode,a.cfree1,a.cfree2,a.cfree3,a.cfree4,a.cfree5,a.cfree6,a.cfree7,a.cfree8,a.cfree9,a.cfree10 " & _
                    " from EFYZGL_pressinform as a inner join EFYZGL_Sheet as b on a.id=b.id " & _
                    " left join  inventory D on b.cinvcode = D.cInvCode   " & _
                    " where isnull(b.cinvcode,'')<>'' " & _
                    " and a.cinvcode='" & Me.Voucher.headerText("cinvcode") & "' and isnull(a.cfree1,'')='" & Me.Voucher.headerText("cfree1") & "' and isnull(a.cfree2,'')='" & Me.Voucher.headerText("cfree2") & "' and isnull(a.cfree3,'')='" & Me.Voucher.headerText("cfree3") & "' " & _
                    " and isnull(a.cfree4,'')='" & Me.Voucher.headerText("cfree4") & "' and isnull(a.cfree5,'')='" & Me.Voucher.headerText("cfree5") & "' and isnull(a.cfree6,'')='" & Me.Voucher.headerText("cfree6") & "' and isnull(a.cfree7,'')='" & Me.Voucher.headerText("cfree7") & "' " & _
                    " and isnull(a.cfree8,'')='" & Me.Voucher.headerText("cfree8") & "' and isnull(a.cfree9,'')='" & Me.Voucher.headerText("cfree9") & "' and isnull(a.cfree10,'')='" & Me.Voucher.headerText("cfree10") & "' " & _
                    " group by a.cinvcode,D.cinvname,D.cInvAddCode,b.cinvcode,a.cfree1,a.cfree2,a.cfree3,a.cfree4,a.cfree5,a.cfree6,a.cfree7,a.cfree8,a.cfree9,a.cfree10 "
            Set rds = DBConn.Execute(strsql)
            dataxml.Load App.Path & "\\bom_bom.xml"
            If Not hasBom Then
                '如果没有生成BOM
                
                '每个分段号生成一个BOM新增或更新到库里
                i = 0
                If Not rds.EOF Then
                    '版本,版本说明
                    Set nodelistL = dataxml.selectNodes("//Bom//Version//Version")
                    Set ele = nodelistL(0)
                    ele.nodeTypedValue = "1"
    '                Set nodelistL = dataxml.selectNodes("//Bom//Version//VersionDesc")
    '                Set ele = nodelistL(0)
    '                ele.nodeTypedValue = "自动生成"
                    '版本日期
                    Set nodelistL = dataxml.selectNodes("//Bom//Version//VersionEffDate")
                    Set ele = nodelistL(0)
                    ele.nodeTypedValue = Me.Voucher.headerText("ddate")
                    '工号,名称
                    Set nodelistL = dataxml.selectNodes("//Bom//Parent//InvCode")
                    Set ele = nodelistL(0)
                    ele.nodeTypedValue = Me.Voucher.headerText("cinvcode")
                    Set nodelistL = dataxml.selectNodes("//Bom//Parent//InvName")
                    Set ele = nodelistL(0)
                    ele.nodeTypedValue = Me.Voucher.headerText("cinvname")
                    '自由项1-10
                    Set nodelistL = dataxml.selectNodes("//Bom//Parent//Free1")
                    Set ele = nodelistL(0)
                    ele.nodeTypedValue = Me.Voucher.headerText("cfree1")
                    Set nodelistL = dataxml.selectNodes("//Bom//Parent//Free2")
                    Set ele = nodelistL(0)
                    ele.nodeTypedValue = Me.Voucher.headerText("cfree2")
                    Set nodelistL = dataxml.selectNodes("//Bom//Parent//Free3")
                    Set ele = nodelistL(0)
                    ele.nodeTypedValue = Me.Voucher.headerText("cfree3")
                    Set nodelistL = dataxml.selectNodes("//Bom//Parent//Free4")
                    Set ele = nodelistL(0)
                    ele.nodeTypedValue = Me.Voucher.headerText("cfree4")
                    Set nodelistL = dataxml.selectNodes("//Bom//Parent//Free5")
                    Set ele = nodelistL(0)
                    ele.nodeTypedValue = Me.Voucher.headerText("cfree5")
                    Set nodelistL = dataxml.selectNodes("//Bom//Parent//Free6")
                    Set ele = nodelistL(0)
                    ele.nodeTypedValue = Me.Voucher.headerText("cfree6")
                    Set nodelistL = dataxml.selectNodes("//Bom//Parent//Free7")
                    Set ele = nodelistL(0)
                    ele.nodeTypedValue = Me.Voucher.headerText("cfree7")
                    Set nodelistL = dataxml.selectNodes("//Bom//Parent//Free8")
                    Set ele = nodelistL(0)
                    ele.nodeTypedValue = Me.Voucher.headerText("cfree8")
                    Set nodelistL = dataxml.selectNodes("//Bom//Parent//Free9")
                    Set ele = nodelistL(0)
                    ele.nodeTypedValue = Me.Voucher.headerText("cfree9")
                    Set nodelistL = dataxml.selectNodes("//Bom//Parent//Free10")
                    Set ele = nodelistL(0)
                    ele.nodeTypedValue = Me.Voucher.headerText("cfree10")
                End If
                id1 = GetVouchID_1("bom_opcomponent", clsSys, "", 0, errStr)
                id2 = GetVouchID_1("bom_opcomponentopt", clsSys, "", 0, errStr)
                While Not rds.EOF
                    
                    'OpComponentId
                    Set nodelistL = dataxml.selectNodes("//Bom//Component//OpComponentId")
                    Set ele = nodelistL(i)
                    ele.nodeTypedValue = CStr(CInt(id1) + i)
                    Set nodelistL = dataxml.selectNodes("//Bom//Component//OptionsId")
                    Set ele = nodelistL(i)
                    ele.nodeTypedValue = CStr(CInt(id2) + i)
                    '序号
                    Set nodelistL = dataxml.selectNodes("//Bom//Component//SortSeq")
                    Set ele = nodelistL(i)
                    ele.nodeTypedValue = CStr(10 * (i + 1))
                    '物料
                    Set nodelistL = dataxml.selectNodes("//Bom//Component//InvCode")
                    Set ele = nodelistL(i)
                    ele.nodeTypedValue = Me.Voucher4.bodyText(Voucher.row, "cinvcode")
                    'add by jiang start 20080828
                    '辅助单位
                    Set nodelistL = dataxml.selectNodes("//Bom//Component//AuxUnitCode")
                    Set ele = nodelistL(i)
                    ele.nodeTypedValue = Me.Voucher4.bodyText(Voucher.row, "cunitid")
                    'add by jiang end 20080828
                    '分子数量
                    Set nodelistL = dataxml.selectNodes("//Bom//Component//BaseQtyN")
                    Set ele = nodelistL(i)
                    ele.nodeTypedValue = Me.Voucher4.bodyText(Voucher.row, "iquantityhj")
                    '分母数量
                    Set nodelistL = dataxml.selectNodes("//Bom//Component//BaseQtyD")
                    Set ele = nodelistL(i)
                    ele.nodeTypedValue = Me.Voucher4.bodyText(Voucher.row, "iquantityhj")
                    '子件生效日
                    Set nodelistL = dataxml.selectNodes("//Bom//Component//EffBegDate")
                    Set ele = nodelistL(i)
                    ele.nodeTypedValue = Me.Voucher.headerText("ddate")
                    'OptionsId
                    Set nodelistL = dataxml.selectNodes("//Bom//ComponentOpt//OptionsId")
                    Set ele = nodelistL(i)
                    ele.nodeTypedValue = CStr(CInt(id2) + i)
                    Voucher.row = Voucher.row + 1
                    rds.MoveNext
                    If Not rds.EOF Then
                        Set nodelistL = dataxml.selectNodes("//Bom//Component")
                        Set node = nodelistL(i).cloneNode(True)
                        On Error Resume Next
                        Set dataxml = dataxml.childNodes(1).appendChild(node)
                        Set nodelistL = dataxml.selectNodes("//Bom//ComponentOpt")
                        Set node = nodelistL(i).cloneNode(True)
                        On Error Resume Next
                        Set dataxml = dataxml.childNodes(1).appendChild(node)
                        i = i + 1
                    End If
                APIobj.Add 1, dataxml, errStr
                Wend
'                APIobj.Add 1, dataxml, errStr
                If Len(errStr) <> 0 Then
                    GoTo dbError
                End If
            Else
               '如果已生成BOM
            MsgBox "物料清单已转入！", vbInformation
            Exit Function
            End If
        End If
    '提交事务
    DBConn.CommitTrans
    'edit by jiang 20080828
    MsgBox "BOM生成成功！", vbInformation
    createOrder = True
    Exit Function
dbError:
    '出错,回滚事务
    DBConn.RollbackTrans
    LoadVoucher ""
    Call VoucherFreeTask
'    ButtonClick "unsure", "弃审", False
    'edit by jiang 20080828
    MsgBox "BOM生成失败：" & errStr
End Function
 Function GetVouchID_1(strTableName As String, clsSys As clsSystem, lngIDs As String, lngsTableCount As Long, errMsg As String) As String
    Dim AdoComm As ADODB.Command
    On Error GoTo DoERR
    Set AdoComm = New ADODB.Command
    If clsSys Is Nothing Then
        Set clsSys = New clsSystem
        clsSys.Init m_Login
    End If
    With AdoComm
        .ActiveConnection = clsSys.dbSales
        .CommandText = "sp_GetID"
        .CommandType = adCmdStoredProc
        .Prepared = False
        .Parameters.Append .CreateParameter("RemoteId", adVarChar, adParamInput, 2, "00")
        .Parameters.Append .CreateParameter("cAcc_Id", adVarChar, adParamInput, 3, clsSys.CurrentAccID)
        .Parameters.Append .CreateParameter("VouchType", adVarChar, adParamInput, 50, strTableName)
        .Parameters.Append .CreateParameter("iAmount", adInteger, adParamInput, 8, lngsTableCount)
        .Parameters.Append .CreateParameter("MaxID", adBigInt, adParamOutput)
        .Parameters.Append .CreateParameter("MaxIDs", adBigInt, adParamOutput)
        .Execute
        GetVouchID_1 = CStr(.Parameters("MaxID"))
        lngIDs = .Parameters("MaxIDs") + 1 - lngsTableCount
    End With
    Set AdoComm = Nothing
    Exit Function
DoERR:
    errMsg = "获取单据ID发生错误：" & Err.Description
    Set AdoComm = Nothing
End Function





'by lg070314 增加U870支持
'修改3 每个窗体都需要这个方法。Cancel与UnloadMode的参数的含义与QueryUnload的参数相同
'请在此方法中调用窗体Exit(退出)方法，并将设置窗体Unload事件参数(如Cancel)的值同时传给此方法的参数
Public Sub ExitForm(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    
    Cancel = m_Cancel
    UnloadMode = m_UnloadMode
End Sub


'设置帮助的系统id
Private Sub SetHelpID()
    Select Case strVouchType
'        Case "32"
'            Me.HelpContextID = 10060203
'        Case Else
'            Me.HelpContextID = 10060203
         Case "16"
             Me.HelpContextID = 2009000400
         Case "97"
             Me.HelpContextID = 2009000420
         Case "29"
             Me.HelpContextID = 2009000380
         Case "26"
             Me.HelpContextID = 2009000070
         Case "06"
             Me.HelpContextID = 2009000090
         Case "28"
             Me.HelpContextID = 2009000350
         Case "16"
             Me.HelpContextID = 2009000410
    End Select
    Me.HelpContextID = 20090308
End Sub
 
''sKey :操作的按钮名称
''
Private Function VoucherTask(skey As String) As Boolean
    Dim strID As String
    
    Select Case strVouchType
        Case "16"
            Select Case skey
                Case "增加", "复制", "删除", "修改"
                    strID = "FA03000102"  '
                Case "审核", "弃审"
                    strID = "FA03000103"  '
                Case "关闭", "打开"
                    strID = "FA03000104"  '
            End Select
        Case "97"
            Select Case skey
                Case "增加", "复制", "删除", "修改"
                    strID = "FA03010101"  '
                Case "关闭", "打开"
                    strID = "FA03010102"  '
                Case "审核", "弃审"
                    strID = "FA03010103"  '
                Case "变更"
                    strID = "FA03010105"  '
            End Select
    End Select
    strID = clsVoucherCO.GetVoucherTaskID(skey, strVouchType, bReturnFlag)
    If strID <> "" Then
        ButtonTaskID = strID
        VoucherTask = LockItem(ButtonTaskID, True, True)
    Else
        VoucherTask = True
    End If
End Function
''释放功能申请
Private Function VoucherFreeTask() As Boolean
    If ButtonTaskID <> "" Then
        VoucherFreeTask = LockItem(ButtonTaskID, False, True)
        ButtonTaskID = ""
    End If
End Function
 

'Dim strAuthId As String     '权限号/gyp/2002/07/24
Private Function ChangeTempaltes(sNewTemplateID As String, Optional bChangDefalt As Boolean, Optional bCheckAuth As Boolean = True, Optional bFormload As Boolean = False) As Boolean
    Dim strDJAuth As String
    Dim bChanged As Boolean
    Dim rstTmp As New ADODB.Recordset
    Dim tmpDomhead As New DOMDocument
    Dim i As Long
    
    On Error GoTo DoERR
    bChanged = False
    If sNewTemplateID = "" Or sNewTemplateID = "0" Then
        Exit Function
    End If
    If bCheckAuth = True Then
        If m_Login.IsAdmin = False Then
            If clsAuth.IsHoldAuth("djmb", Trim(sNewTemplateID), , "W") = False Then
                strDJAuth = clsAuth.getAuthString("DJMB", , "W")
                If strDJAuth = "1=2" Then
                    MsgBox "你没有使用单据模版的权限！"
                    'Me.Hide
                    Exit Function
                Else
                    If clsAuth.IsHoldAuth("DJMB", sTemplateID, , "W") = False Then
                        rstTmp.Open "select vt_id from vouchertemplates where vt_cardnumber='" & strCardNum & "' and vt_id in (" & strDJAuth & ") order by vt_id", DBConn, adOpenForwardOnly, adLockReadOnly
                        If Not rstTmp.EOF Then
                            fillComBol False
                            sNewTemplateID = rstTmp(0)      'left(strDJAuth, IIf(InStr(1, strDJAuth, ",") - 1 = -1, Len(strDJAuth), InStr(1, strDJAuth, ",")))
                        Else
                            MsgBox "你没有使用单据模版的权限！"
                            Me.Hide
                            rstTmp.Close
                            Set rstTmp = Nothing
                            Exit Function
                        End If
                        rstTmp.Close
                        sTemplateID = sNewTemplateID
                    Else
                        sNewTemplateID = sTemplateID
                    End If
                End If
            End If
        End If
    End If
    If bFirst = True Then Call getCardNumber(sNewTemplateID)
    If RstTemplate Is Nothing Then Set RstTemplate = New ADODB.Recordset
    If Trim(sNewTemplateID) = "" Or sNewTemplateID = "0" Then
        If bChangDefalt = True Then
            sNewTemplateID = sTemplateID
            bChanged = True
        End If
    Else
        If sCurTemplateID <> sNewTemplateID Then
            bChanged = True
        Else
            If bChangDefalt = True Then
                bChanged = True
            End If
        End If
    End If
    If bChanged = True Then
        If preVTID = sNewTemplateID And Not RstTemplate Is Nothing Then
            If Not RstTemplate.RecordCount = 0 Then
                GoTo UsePre  ''记录已经取回
            End If
        End If
        Set RstTemplate2 = clsVoucherCO.GetVoucherFormat(sNewTemplateID, strCardNum)
        If RstTemplate2 Is Nothing Then
            If bChangDefalt = True Then
                Set RstTemplate2 = clsVoucherCO.GetVoucherFormat(sTemplateID, strCardNum)
                bChanged = True
            Else
                bChanged = False
            End If
        Else
            If RstTemplate2.State = 1 Then
                If RstTemplate2.EOF And RstTemplate2.BOF Then
                    If bChangDefalt = True Then
                        Set RstTemplate2 = clsVoucherCO.GetVoucherFormat(sTemplateID, strCardNum)
                        sCurTemplateID = sTemplateID
                        sCurTemplateID2 = sTemplateID
                        bChanged = True
                    Else
                        bChanged = False
                    End If
                Else
                   bChanged = True
                End If
            Else
                    If bChangDefalt = True Then
                        Set RstTemplate2 = clsVoucherCO.GetVoucherFormat(sTemplateID, strCardNum)
                        If RstTemplate2.State = adStateClosed Then
                                MsgBox "模版设置有问题"
                                ChangeTempaltes = False
                                Exit Function
                        End If
                        If Not RstTemplate2 Is Nothing Then
                            If Not RstTemplate2.EOF Then
                                bChanged = True
                            Else
                                MsgBox "模版设置有问题"
                                ChangeTempaltes = False
                                Exit Function
                            End If
                        Else
                            MsgBox "模版设置有问题"
                            ChangeTempaltes = False
                            Exit Function
                        End If
                    Else
                        bChanged = False
                    End If
            End If
        End If
    End If
    If bChanged = True Then
        sCurTemplateID = sNewTemplateID
        sCurTemplateID2 = sNewTemplateID
        preVTID = sNewTemplateID
        If bFormload = False Then
            Voucher.Visible = False
        End If
        
        '如果是调试状态，不处理附件，以放置弹出‘加载附件失败’窗口
        If clsSAWeb.IsDebug Then RstTemplate2.Fields("vchtblprimarykeynames") = ""
        
        Voucher.setTemplateData RstTemplate2
        dOriVoucherHeight = Voucher.Height
        dOriVoucherWidth = Voucher.Width
        Call Form_Resize
        If Voucher.VoucherStatus <> VSNormalMode Then
            SetItemState "modify"
        End If
        Call SetVocuhNameLabel
        If Not DomFormat Is Nothing Then
            If DomFormat.xml <> "" Then
                Me.Voucher.SetBillNumberRule DomFormat.xml
                If Me.Voucher.VoucherStatus <> VSNormalMode Then
                    Call SetVouchNoWriteble
                End If
            End If
        End If
        RstTemplate2.Save tmpDomhead, adPersistXML
        If RstTemplate.State = 1 Then RstTemplate.Close
        RstTemplate.Open tmpDomhead
        If strVouchType = "07" Then
            Me.Voucher.BodyMaxRows = -1
            SetVouchItemState "cinvname", "b", False: SetVouchItemState "ccusinvname", "b", False
            SetVouchItemState "cinvcode", "b", False: SetVouchItemState "ccusinvcode", "b", False
            SetVouchItemState "cwhname", "b", False
            SetVouchItemState "ccusabbname", "t", False
            SetVouchItemState "cpersonname", "t", False
            SetVouchItemState "cdepname", "t", False
            SetVouchItemState "cbustype", "t", False
            SetVouchItemState "cexch_name", "t", False
            SetVouchItemState "iExchRate", "t", False
            SetVouchItemState "ccode", "b", False
            SetVouchItemState "dmdate", "b", False
            SetVouchItemState "dvdate", "b", False
            SetVouchItemState "cbatch", "b", False
            For i = 1 To 10
                SetVouchItemState "cfree" & i, "b", False
            Next
        End If
        
        If bFormload = False Then
            Me.Voucher.Visible = True
            Me.Refresh
        End If
    End If
    Set rstTmp = Nothing
    ChangeTempaltes = True
    Call ChangeCaptionCol
    Exit Function
UsePre:
    sCurTemplateID = sNewTemplateID
    sCurTemplateID2 = sNewTemplateID
    If bFormload = False Then
        Me.Voucher.Visible = False
    End If
    Voucher.setTemplateData RstTemplate
    dOriVoucherHeight = Voucher.Height
    dOriVoucherWidth = Voucher.Width
    Call Form_Resize
    If Voucher.VoucherStatus <> VSNormalMode Then
        SetItemState "modify"
    End If
    Call SetVocuhNameLabel
    If bFormload = False Then
        Me.Voucher.Visible = True
        Me.Refresh
    End If
    Set rstTmp = Nothing
    ChangeTempaltes = True
    Exit Function
DoERR:
    MsgBox Err.Description
    ChangeTempaltes = False
    Set rstTmp = Nothing
End Function
 
Private Sub SetVocuhNameLabel()
    '//单据名称标题，主要是为了解决单据的名称的特殊显示问题，例如 "期初" XXX单据
    Me.LabelVoucherName.Caption = Me.Voucher.TitleCaption

    '//单据的名称
    Me.Voucher.TitleCaption = Me.Voucher.TitleCaption
    Me.Voucher.TitleCaption = ""
End Sub

 
''函数load单据,更改按纽状态,更改模板
Private Sub LoadVoucher(sMove As String, Optional vid As Variant, Optional bRefreshClick As Boolean = False)
    Dim errMsg As String
    Dim errmsg1 As String, errmsg2 As String, errmsg3 As String, errmsg4 As String
    Dim i As Integer
    On Error Resume Next
    Select Case LCase(sMove)
        Case ""
            errMsg = clsVoucherCO.GetVoucherData(Domhead, Dombody, vid)
            errmsg1 = clsVoucherCO1.GetVoucherData(Domhead1, Dombody1, vid)
            errmsg2 = clsVoucherCO2.GetVoucherData(Domhead2, Dombody2, vid)
            errmsg3 = clsVoucherCO3.GetVoucherData(Domhead3, Dombody3, vid)
            errmsg4 = clsVoucherCO4.GetVoucherData(Domhead4, Dombody4, vid)
        Case "tonext"
ToNext:
            i = i + 1
            errMsg = clsVoucherCO.MoveNext(Domhead, Dombody)
            errmsg1 = clsVoucherCO1.MoveNext1(Domhead1, Dombody1)
            errmsg2 = clsVoucherCO2.MoveNext2(Domhead2, Dombody2)
            errmsg3 = clsVoucherCO3.MoveNext3(Domhead3, Dombody3)
            errmsg4 = clsVoucherCO4.MoveNext4(Domhead4, Dombody4)
        Case "toprevious"
            errMsg = clsVoucherCO.MovePrevious(Domhead, Dombody)
            errmsg1 = clsVoucherCO1.MovePrevious1(Domhead1, Dombody1)
            errmsg2 = clsVoucherCO2.MovePrevious2(Domhead2, Dombody2)
            errmsg3 = clsVoucherCO3.MovePrevious3(Domhead3, Dombody3)
            errmsg4 = clsVoucherCO4.MovePrevious4(Domhead4, Dombody4)
        Case "tolast"
            errMsg = clsVoucherCO.MoveLast(Domhead, Dombody)
            errmsg1 = clsVoucherCO1.MoveLast1(Domhead1, Dombody1)
            errmsg2 = clsVoucherCO2.MoveLast2(Domhead2, Dombody2)
            errmsg3 = clsVoucherCO3.MoveLast3(Domhead3, Dombody3)
            errmsg4 = clsVoucherCO4.MoveLast4(Domhead4, Dombody4)
        Case "tofirst"
            errMsg = clsVoucherCO.MoveFirst(Domhead, Dombody)
            errmsg1 = clsVoucherCO1.MoveFirst1(Domhead1, Dombody1)
            errmsg2 = clsVoucherCO2.MoveFirst2(Domhead2, Dombody2)
            errmsg3 = clsVoucherCO3.MoveFirst3(Domhead3, Dombody3)
            errmsg4 = clsVoucherCO4.MoveFirst4(Domhead4, Dombody4)
    End Select
        If errMsg <> "" Then
            If bRefreshClick = False And sMove = "" And vid = "" Then
                
            Else
                MsgBox errMsg
            End If
            If i <= 3 Then GoTo ToNext
            Exit Sub
        End If
    ChangeTempaltes IIf(val(GetHeadItemValue(Domhead, "ivtid")) = 0, sCurTemplateID2, GetHeadItemValue(Domhead, "ivtid")), , False
    Me.Voucher.Visible = False
    
    Voucher.setVoucherDataXML Domhead, Dombody
    Voucher1.setVoucherDataXML Domhead1, Dombody1
    Voucher2.setVoucherDataXML Domhead2, Dombody2
    Voucher3.setVoucherDataXML Domhead3, Dombody3
    Voucher4.setVoucherDataXML Domhead4, Dombody4
    '审批流文本
    'Me.voucher.ExamineFlowAuditInfo = GetEAStream(strVouchType, Domhead, Me.voucher, DBConn)
    Call SetSum
    If Me.Voucher.headerText("cexch_name") <> "" Then
        Me.Voucher.ItemState("iexchrate", siheader).nNumPoint = clsSAWeb_M.GetExchRateDec(GetHeadItemValue(Domhead, "cexch_name"))
        Me.Voucher.headerText("iexchrate") = GetHeadItemValue(Domhead, "iexchrate")
    End If
    ChangeButtonsState
    Call Form_Resize
    Me.Voucher.Visible = True
    EditPanel_1 EdtP, 3, ""
    Dim strXml As String
    strXml = "<?xml version='1.0' encoding='GB2312'?>" & Chr(13)
    domConfig.loadXML strXml & "<EAI>0</EAI>"
End Sub
 Private Sub SetSum()
    Exit Sub
    Dim ele As IXMLDOMElement, NdList As IXMLDOMNodeList
    Dim iSum As Double, strSumDX As Variant
    'Dim oNum2Chinese As Object
    
    
    If strVouchType = "92" Or strVouchType = "95" Then Exit Sub
    strSumDX = ""
    iSum = 0
    Set NdList = Dombody.selectNodes("//z:row")
    For Each ele In NdList
        iSum = iSum + CDbl(val(IIf(IsNull(ele.getAttribute("isum")), 0, ele.getAttribute("isum"))))
    Next
    'Set oNum2Chinese = CreateObject("FormulaParse.Calculator")
    Num2Chinese Format(iSum, "#.00"), strSumDX
    'If strSumDX = "圆整" Then strSumDX = "零圆零角零分"
    Me.Voucher.headerText("isumdx") = strSumDX
    Me.Voucher.headerText("isumx") = iSum
    Me.Voucher.headerText("zdsumdx") = strSumDX
    Me.Voucher.headerText("zdsum") = iSum
    Set NdList = Nothing
    Set ele = Nothing
End Sub
Private Sub SetItemState(Optional sOperate As String)
    Dim i As Long
    With Me.Voucher
        .BodyMaxRows = 0
        Select Case strVouchType
            Case "97", "16", "06"
                If strVouchType = "97" Then
                    If Dombody.selectNodes("//z:row[@ccontractid !='']").length > 0 Then
                        .EnableHead "ccusabbname", False
                        .EnableHead "cexch_name", False
                        .EnableHead "cbustype", False
                        If Voucher.headerText("cstname") = "" Then
                            SetOriItemState "T", "cstname"
                        Else
                            .EnableHead "cstname", False
                        End If
                    Else
                        SetOriItemState "T", "ccusabbname"
                        SetOriItemState "T", "cexch_name"
                        SetOriItemState "T", "cbustype"
                        SetOriItemState "T", "cstname"
                    End If
                    If iVouchState = 2 Then
                        sCurTemplateID = ""
                        
                        .Visible = False
                        For iHeadIndex = 1 To .HeadInfoCount
                            .EnableHead iHeadIndex, False
                        Next iHeadIndex
                        SetOriItemState "T", "cmemo"
                        SetOriItemState "T", "dpredatebt"
                        SetOriItemState "T", "dpremodatebt"
                        .Visible = True
                        .SetFocus
                        .UpdateCmdBtn
                    End If
                End If
                If LCase(sOperate) = "copy" Or LCase(sOperate) = "modify" Then
                    .EnableHead "cbustype", False
                Else
                    SetOriItemState "T", "cbustype"
                End If
        End Select
    End With
End Sub
 
Private Function GetScrollWidth() As Single
    GetScrollWidth = Screen.TwipsPerPixelX * GetSystemMetrics(SM_CXVSCROLL)
End Function
 

''需要改变单据模版
Private Sub ComboDJMB_Click()
    Dim tmpVoucherState As Variant
    ComboDJMB.ToolTipText = ComboDJMB.Text
    If Not bfillDjmb Then
        Me.Voucher.Visible = False
        Me.Voucher.getVoucherDataXML Domhead, Dombody
        tmpVoucherState = Me.Voucher.VoucherStatus
        Call ChangeTempaltes(Str(vtidDJMB(ComboDJMB.ListIndex)), , False)
        Me.Voucher.VoucherStatus = tmpVoucherState
        Me.Voucher.setVoucherDataXML Domhead, Dombody
        Me.Voucher.Visible = True
        Me.Voucher.headerText("ivtid") = Str(vtidDJMB(ComboDJMB.ListIndex))
        sCurTemplateID = Str(vtidDJMB(ComboDJMB.ListIndex))
        sCurTemplateID2 = Str(vtidDJMB(ComboDJMB.ListIndex))
    Else
        bfillDjmb = False
    End If
End Sub
 
Private Sub ComboVTID_Click()
    ComboVTID.ToolTipText = ComboVTID.Text
End Sub
 

Private Sub CTBCtrl1_OnCommand(ByVal enumType As prjTBCtrl.ENUM_MENU_OR_BUTTON, ByVal cButtonId As String, ByVal cmenuid As String)
    bCloseFHSingle = False
    ButtonClick cButtonId, tbrvoucher.buttons(cButtonId).ToolTipText
End Sub
 
Private Sub Form_Activate()
    On Error Resume Next
    Me.picVoucher.BackColor = Me.picVoucher.BackColor
End Sub
 
Private Sub Form_Deactivate()
    With Me.Voucher
        If .VoucherStatus <> VSNormalMode Then
            bLostFocus = True
            .ProtectUnload2
            bLostFocus = False
        End If
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 
    If iShowMode <> 1 Then  ''??
        setKey KeyCode, Shift
    ElseIf KeyCode = vbKeyF4 Then
        setKey KeyCode, Shift
    End If
    
End Sub
 
Private Sub Form_Load()
    Dim strErr As String
    Dim bLock As Boolean
    Dim recTmp As UfRecordset
    Dim dD As Date
    Dim s As String
    On Error Resume Next
    Me.KeyPreview = True
    '////////////////zhupb
    Dim strErrorResId As String
    clsVoucher.Init strCardNum, strErrorResId
    clsVoucherRefer.Init strCardNum, strErrorResId
    App.HelpFile = App.Path & "\" & getChmFile("EFYZGL")
'设置单据排序控件

'//////////////////////////////////////////////////
'  860sp升级到861修改处1 注释    2006/03/08 改控件在861版本中已经集成到单据控件中了   所以要删除
' voucher.SetSortCallBackObject U8VoucherSorter1
'    With U8VoucherSorter1
'        .BackColor = voucher.BackColor
'        .Left = Me.Left + 550
'        .Top = Me.Picture1.Top
'        .ZOrder
'    End With
'//////////////////////////////////////////////////

    Me.Voucher.ControlAutoScales = AutoBoth
    Me.Voucher.ControlScrollBars = ScrollBoth
    Me.Voucher1.ControlAutoScales = AutoBoth
    Me.Voucher1.ControlScrollBars = ScrollBoth
    Me.Voucher2.ControlAutoScales = AutoBoth
    Me.Voucher2.ControlScrollBars = ScrollBoth
    Me.Voucher3.ControlAutoScales = AutoBoth
    Me.Voucher3.ControlScrollBars = ScrollBoth
    Me.Voucher4.ControlAutoScales = AutoBoth
    Me.Voucher4.ControlScrollBars = ScrollBoth

'by lg070314增加U870菜单融合功能
    ''''''''''''''''''''''''''''''''''''''
    If Not g_business Is Nothing Then
        Set Me.UFToolbar1.Business = g_business
    End If
    Call RegisterMessage
    '''''''''''''''''''''''''''''''''''''''
    
    'Call SetButton  '设置菜单按钮
    clsTbl.initToolBar Me.tbrvoucher, Me.strToolBarName, strErr
    
    If lngClr1 <> 0 And lngClr2 <> 0 Then
        Call Voucher.SetRuleColor(lngClr1, lngClr2)
    End If
    ChangeOneFormTbr Me, Me.tbrvoucher, Me.UFToolbar1
    'SetButtonStatus "Cancel"
    clsTbl.ChangeButtonState Voucher, Me.tbrvoucher, Me.UFToolbar1, Voucher.VoucherStatus
    
    Labeldjmb.BackColor = Me.Picture2.BackColor
    Picture1.BackColor = Me.Picture2.BackColor
    Labeldjmb.ForeColor = vbBlack
    If iShowMode = 1 Then
        If frmMain.WindowState = 1 Then frmMain.WindowState = 2
    End If
    Me.picVoucher.Move 0, Me.tbrvoucher.Height, Me.ScaleWidth, Me.ScaleHeight - Me.StBar.Height    '-ME.tbrvoucher
     Me.picVoucher.Move 0, Me.tbrvoucher.Height, Me.ScaleWidth, Me.ScaleHeight - Me.StBar.Height - Me.UFToolbar1.Height     '-ME.tbrvoucher
    Me.Picture2.Width = Me.Width
    Me.Picture1.BackColor = Me.Picture2.BackColor
    Me.Picture1.Move Me.Picture2.Width - Me.Picture1.Width - 5
    Me.StBar.ZOrder
    strFreeName1 = clsSAWeb_M.getDefName(DBConn, "cfree1")
    strFreeName2 = clsSAWeb_M.getDefName(DBConn, "cfree2")
    strFreeName3 = clsSAWeb_M.getDefName(DBConn, "cfree3")
    strFreeName4 = clsSAWeb_M.getDefName(DBConn, "cfree4")
    strFreeName5 = clsSAWeb_M.getDefName(DBConn, "cfree5")
    strFreeName6 = clsSAWeb_M.getDefName(DBConn, "cfree6")
    strFreeName7 = clsSAWeb_M.getDefName(DBConn, "cfree7")
    strFreeName8 = clsSAWeb_M.getDefName(DBConn, "cfree8")
    strFreeName9 = clsSAWeb_M.getDefName(DBConn, "cfree9")
    strFreeName10 = clsSAWeb_M.getDefName(DBConn, "cfree10")
    With StBar
        .Panels.Clear
        .Panels.Add 1, , ""
        .Panels(1).Width = Me.Width * 1 / 3
        .Panels.Add 2, , ""
        .Panels(2).Width = Me.Width * 1 / 3
        .Panels.Add 3, , ""
        .Panels(3).Width = Me.Width * 1 / 3
    End With
    Me.BackColor = Me.Voucher.BackColor
    Me.ForeColor = Me.Voucher.BackColor
    Set moAutoFill = CreateObject("ScmPublicSrv.clsAutoFill")
    ProgressBar1.Top = Voucher.Top - 1000
    ProgressBar1.Width = Voucher.Width - 2000
    ProgressBar1.Left = Voucher.Left - 1000
End Sub
 

'Private Sub Form_Resize()
'    On Error Resume Next
'    Me.UFToolbar1.Top = 0
'    Me.UFToolbar1.Width = Me.ScaleWidth
'    Me.picVoucher.Move 0, Me.tbrvoucher.Height, Me.ScaleWidth, Me.ScaleHeight - Me.tbrvoucher.Height - Me.StBar.Height
'
'    '//Voucher 第1个页面Resize
'    If Me.picVoucher.Height - 1 * GetScrollWidth < dOriVoucherHeight Then
'        Me.Voucher1.Height = Me.picVoucher.Height - Picture12.Height '- 1 * GetScrollWidth
'    Else
'        Me.Voucher1.Height = Me.picVoucher.Height - Picture12.Height '- 1 * GetScrollWidth
'    End If
'    If Me.picVoucher.Width - 1 * GetScrollWidth < dOriVoucherWidth Then
'        Me.Voucher1.Width = Me.picVoucher.Width '- 1 * GetScrollWidth
'    Else
'        Me.Voucher1.Width = Me.picVoucher.Width  '- 1 * GetScrollWidth
'    End If
'
'    'zcy修改 作废或现结需要减小宽度，否则表体滚动条看不到
'
'    'SetScrollBarValue
'
'    Me.Picture12.Width = Me.Width
'    Me.Picture12.Move 0, 500
'    Me.Picture11.BackColor = Me.Picture12.BackColor
'    Me.Picture11.Move Me.Picture12.Width - Me.Picture11.Width - 400
'    Me.LabelVoucherName1.Move (Me.Width - Me.LabelVoucherName1.Width) / 2
'
'    If labZF.Visible = False And labXJ.Visible = False Then
'        Me.Voucher1.Move 0, Me.Picture12.Height
'    Else
'        Me.Voucher1.Move 0, Me.Picture12.Height + Me.labZF.Height
'        Me.Voucher1.Width = Me.Voucher1.Width - 350   'zcy修改 作废或现结需要减小宽度，否则表体滚动条看不到
'        Me.Voucher1.Height = Me.Voucher1.Height - Me.labZF.Height - Picture12.Height
'    End If
'
'    '//排序按钮
''    With U8VoucherSorter2
''        .BackColor = Me.Picture12.BackColor
''        .Left = 550
''        .Top = Me.picVoucher.Top + Me.Picture11.Top
'''        .ZOrder
''    End With
'
'        '//////////////////////////////////////////////////////////////////////
'      '//Voucher 第2个页面Resize
'    If Me.picVoucher.Height - 1 * GetScrollWidth < dOriVoucherHeight Then
'        Me.Voucher2.Height = Me.picVoucher.Height - Picture22.Height '- 1 * GetScrollWidth
'    Else
'        Me.Voucher2.Height = Me.picVoucher.Height - Picture22.Height '- 1 * GetScrollWidth
'    End If
'    If Me.picVoucher.Width - 1 * GetScrollWidth < dOriVoucherWidth Then
'        Me.Voucher2.Width = Me.picVoucher.Width '- 1 * GetScrollWidth
'    Else
'        Me.Voucher2.Width = Me.picVoucher.Width  '- 1 * GetScrollWidth
'    End If
'
'    'zcy修改 作废或现结需要减小宽度，否则表体滚动条看不到
'
'    'SetScrollBarValue
'
'    Me.Picture22.Width = Me.Width
'    Me.Picture22.Move 0, 500
'    Me.Picture21.BackColor = Me.Picture22.BackColor
'    Me.Picture21.Move Me.Picture22.Width - Me.Picture21.Width - 400
'    Me.LabelVoucherName2.Move (Me.Width - Me.LabelVoucherName2.Width) / 2
'
'    If labZF.Visible = False And labXJ.Visible = False Then
'        Me.Voucher2.Move 0, Me.Picture22.Height
'    Else
'        Me.Voucher2.Move 0, Me.Picture22.Height + Me.labZF.Height
'        Me.Voucher2.Width = Me.Voucher2.Width - 350   'zcy修改 作废或现结需要减小宽度，否则表体滚动条看不到
'        Me.Voucher2.Height = Me.Voucher2.Height - Me.labZF.Height - Picture22.Height
'    End If
'
''    '//排序按钮
''    With U8VoucherSorter3
''        .BackColor = Me.Picture22.BackColor
''        .Left = 550
''        .Top = Me.picVoucher.Top + Me.Picture21.Top
'''        .ZOrder
''    End With
'    '//////////////////////////////////////////////////////////////////////
'
' '//////////////////////////////////////////////////////////////////////
'      '//Voucher 第3个页面Resize
'    If Me.picVoucher.Height - 1 * GetScrollWidth < dOriVoucherHeight Then
'        Me.Voucher3.Height = Me.picVoucher.Height - Picture32.Height '- 1 * GetScrollWidth
'    Else
'        Me.Voucher3.Height = Me.picVoucher.Height - Picture32.Height '- 1 * GetScrollWidth
'    End If
'    If Me.picVoucher.Width - 1 * GetScrollWidth < dOriVoucherWidth Then
'        Me.Voucher3.Width = Me.picVoucher.Width '- 1 * GetScrollWidth
'    Else
'        Me.Voucher3.Width = Me.picVoucher.Width  '- 1 * GetScrollWidth
'    End If
'
'    'zcy修改 作废或现结需要减小宽度，否则表体滚动条看不到
'
'    'SetScrollBarValue
'
'    Me.Picture32.Width = Me.Width
'    Me.Picture32.Move 0, 500
'    Me.Picture31.BackColor = Me.Picture32.BackColor
'    Me.Picture31.Move Me.Picture32.Width - Me.Picture31.Width - 400
'    Me.LabelVoucherName3.Move (Me.Width - Me.LabelVoucherName3.Width) / 2
'
'    If labZF.Visible = False And labXJ.Visible = False Then
'        Me.Voucher3.Move 0, Me.Picture32.Height
'    Else
'        Me.Voucher3.Move 0, Me.Picture32.Height + Me.labZF.Height
'        Me.Voucher3.Width = Me.Voucher3.Width - 350   'zcy修改 作废或现结需要减小宽度，否则表体滚动条看不到
'        Me.Voucher3.Height = Me.Voucher3.Height - Me.labZF.Height - Picture32.Height
'    End If
'
'    '//排序按钮
''    With U8VoucherSorter4
''        .BackColor = Me.Picture32.BackColor
''        .Left = 550
''        .Top = Me.picVoucher.Top + Me.Picture31.Top
'''        .ZOrder
''    End With
'    '//////////////////////////////////////////////////////////////////////
'
' '//////////////////////////////////////////////////////////////////////
'      '//Voucher 第4个页面Resize
'    If Me.picVoucher.Height - 1 * GetScrollWidth < dOriVoucherHeight Then
'        Me.Voucher4.Height = Me.picVoucher.Height - Picture42.Height '- 1 * GetScrollWidth
'    Else
'        Me.Voucher4.Height = Me.picVoucher.Height - Picture42.Height '- 1 * GetScrollWidth
'    End If
'    If Me.picVoucher.Width - 1 * GetScrollWidth < dOriVoucherWidth Then
'        Me.Voucher4.Width = Me.picVoucher.Width '- 1 * GetScrollWidth
'    Else
'        Me.Voucher4.Width = Me.picVoucher.Width  '- 1 * GetScrollWidth
'    End If
'
'    'zcy修改 作废或现结需要减小宽度，否则表体滚动条看不到
'
'    'SetScrollBarValue
'
'    Me.Picture42.Width = Me.Width
'    Me.Picture42.Move 0, 500
'    Me.Picture41.BackColor = Me.Picture42.BackColor
'    Me.Picture41.Move Me.Picture42.Width - Me.Picture41.Width - 400
'    Me.LabelVoucherName4.Move (Me.Width - Me.LabelVoucherName4.Width) / 2
'
'    If labZF.Visible = False And labXJ.Visible = False Then
'        Me.Voucher4.Move 0, Me.Picture42.Height
'    Else
'        Me.Voucher4.Move 0, Me.Picture42.Height + Me.labZF.Height
'        Me.Voucher4.Width = Me.Voucher4.Width - 350   'zcy修改 作废或现结需要减小宽度，否则表体滚动条看不到
'        Me.Voucher4.Height = Me.Voucher4.Height - Me.labZF.Height - Picture42.Height
'    End If
'
'    '//排序按钮
''    With U8VoucherSorter5
''        .BackColor = Me.Picture42.BackColor
''        .Left = 550
''        .Top = Me.picVoucher.Top + Me.Picture41.Top
'''        .ZOrder
''    End With
'    '//////////////////////////////////////////////////////////////////////
'     '第0个页面
'    If Me.picVoucher.Height - 1 * GetScrollWidth < dOriVoucherHeight Then
'        Me.voucher.Height = Me.picVoucher.Height - Picture2.Height '- 1 * GetScrollWidth
'    Else
'        Me.voucher.Height = Me.picVoucher.Height - Picture2.Height '- 1 * GetScrollWidth
'    End If
'    If Me.picVoucher.Width - 1 * GetScrollWidth < dOriVoucherWidth Then
'        Me.voucher.Width = Me.picVoucher.Width '- 1 * GetScrollWidth
'    Else
'        Me.voucher.Width = Me.picVoucher.Width '- 1 * GetScrollWidth
'    End If
'    Me.Picture2.Width = Me.Width
'    Me.Picture2.Move 0, 0
'    Me.Picture1.BackColor = Me.Picture2.BackColor
'    Me.Picture1.Move Me.Picture2.Width - Me.Picture1.Width - 400
'    Me.LabelVoucherName.Move (Me.Width - Me.LabelVoucherName.Width) / 2
'
'    If labZF.Visible = False And labXJ.Visible = False Then
'        Me.voucher.Move 0, Me.Picture2.Height
'    Else
'        Me.voucher.Move 0, Me.Picture2.Height + Me.labZF.Height
'        Me.voucher.Width = Me.voucher.Width - 350
'        Me.voucher.Height = Me.voucher.Height - Me.labZF.Height - Picture2.Height
'    End If
'
'        '//排序按钮
''    With U8VoucherSorter1
''        .BackColor = Me.Picture2.BackColor
''        .Left = 550
''        .Top = Me.picVoucher.Top + Me.Picture1.Top
'''        .ZOrder
''    End With
'
'    '//
'
'    labZF.Top = picVoucher.Top + Me.Picture2.Height  'Me.top - Me.tbrvoucher.top
'    labZF.Left = Me.voucher.Left
'    labXJ.Top = picVoucher.Top + Me.Picture2.Height ' Me.top - Me.tbrvoucher.top    'Me.StBar.height
'    labXJ.Left = Me.voucher.Left '+ labZF.Width
'    With StBar
'        .Panels(1).Width = Me.Width * 1 / 3
'        .Panels(2).Width = Me.Width * 1 / 3
'        .Panels(3).Width = Me.Width * 1 / 3
'    End With
''//////////////////////////////////////////////////
''  860sp升级到861修改 注释    2006/03/08 改控件在861版本中已经集成到单据控件中了 所以要删除
''    With U8VoucherSorter1
''        .BackColor = Me.Picture2.BackColor
''        .Left = 550
''        .Top = Me.picVoucher.Top + Me.Picture1.Top
''        .ZOrder
''    End With
'    Me.BackColor = Me.voucher.BackColor
'    Me.ForeColor = Me.voucher.BackColor
'    Me.picVoucher.BackColor = Me.voucher.BackColor
'End Sub
'
Private Sub Form_Resize()

    On Error Resume Next
    Me.picVoucher.Tab = 0
    Me.UFToolbar1.Top = 0
    Me.UFToolbar1.Width = Me.ScaleWidth
    Me.picVoucher.Move 0, Me.tbrvoucher.Height, Me.ScaleWidth, Me.ScaleHeight - Me.tbrvoucher.Height - Me.StBar.Height
    
    
    '//Voucher 第1个页面Resize
    If Me.picVoucher.Height - 1 * GetScrollWidth < dOriVoucherHeight Then
        Me.Voucher1.Height = Me.picVoucher.Height - Picture12.Height '- 1 * GetScrollWidth
    Else
        Me.Voucher1.Height = Me.picVoucher.Height - Picture12.Height '- 1 * GetScrollWidth
    End If
    If Me.picVoucher.Width - 1 * GetScrollWidth < dOriVoucherWidth Then
        Me.Voucher1.Width = Me.picVoucher.Width '- 1 * GetScrollWidth
    Else
        Me.Voucher1.Width = Me.picVoucher.Width  '- 1 * GetScrollWidth
    End If
    
    'zcy修改 作废或现结需要减小宽度，否则表体滚动条看不到

    'SetScrollBarValue
    
    Me.Picture12.Width = Me.Width
    Me.Picture12.Move 0, 500
    Me.Picture11.BackColor = Me.Picture12.BackColor
    Me.Picture11.Move Me.Picture12.Width - Me.Picture11.Width - 400
    Me.LabelVoucherName1.Move (Me.Width - Me.LabelVoucherName1.Width) / 2
    
    If labZF.Visible = False And labXJ.Visible = False Then
        Me.Voucher1.Move 0, Me.Picture12.Height
    Else
        Me.Voucher1.Move 0, Me.Picture12.Height + Me.labZF.Height
        Me.Voucher1.Width = Me.Voucher1.Width - 350   'zcy修改 作废或现结需要减小宽度，否则表体滚动条看不到
        Me.Voucher1.Height = Me.Voucher1.Height - Me.labZF.Height - Picture12.Height
    End If
    
    '//排序按钮
'    With U8VoucherSorter2
'        .BackColor = Me.Picture12.BackColor
'        .Left = 550
'        .Top = Me.picVoucher.Top + Me.Picture11.Top
''        .ZOrder
'    End With
   
        '//////////////////////////////////////////////////////////////////////
      '//Voucher 第2个页面Resize
    If Me.picVoucher.Height - 1 * GetScrollWidth < dOriVoucherHeight Then
        Me.Voucher2.Height = Me.picVoucher.Height - Picture22.Height '- 1 * GetScrollWidth
    Else
        Me.Voucher2.Height = Me.picVoucher.Height - Picture22.Height '- 1 * GetScrollWidth
    End If
    If Me.picVoucher.Width - 1 * GetScrollWidth < dOriVoucherWidth Then
        Me.Voucher2.Width = Me.picVoucher.Width '- 1 * GetScrollWidth
    Else
        Me.Voucher2.Width = Me.picVoucher.Width  '- 1 * GetScrollWidth
    End If
    
    'zcy修改 作废或现结需要减小宽度，否则表体滚动条看不到

    'SetScrollBarValue
    
    Me.Picture22.Width = Me.Width
    Me.Picture22.Move 0, 500
    Me.Picture21.BackColor = Me.Picture22.BackColor
    Me.Picture21.Move Me.Picture22.Width - Me.Picture21.Width - 400
    Me.LabelVoucherName2.Move (Me.Width - Me.LabelVoucherName2.Width) / 2
    
    If labZF.Visible = False And labXJ.Visible = False Then
        Me.Voucher2.Move 0, Me.Picture22.Height
    Else
        Me.Voucher2.Move 0, Me.Picture22.Height + Me.labZF.Height
        Me.Voucher2.Width = Me.Voucher2.Width - 350   'zcy修改 作废或现结需要减小宽度，否则表体滚动条看不到
        Me.Voucher2.Height = Me.Voucher2.Height - Me.labZF.Height - Picture22.Height
    End If
    
    '//排序按钮
'    With U8VoucherSorter3
'        .BackColor = Me.Picture22.BackColor
'        .Left = 550
'        .Top = Me.picVoucher.Top + Me.Picture21.Top
''        .ZOrder
'    End With
    '//////////////////////////////////////////////////////////////////////

 '//////////////////////////////////////////////////////////////////////
      '//Voucher 第3个页面Resize
    If Me.picVoucher.Height - 1 * GetScrollWidth < dOriVoucherHeight Then
        Me.Voucher3.Height = Me.picVoucher.Height - Picture32.Height '- 1 * GetScrollWidth
    Else
        Me.Voucher3.Height = Me.picVoucher.Height - Picture32.Height '- 1 * GetScrollWidth
    End If
    If Me.picVoucher.Width - 1 * GetScrollWidth < dOriVoucherWidth Then
        Me.Voucher3.Width = Me.picVoucher.Width '- 1 * GetScrollWidth
    Else
        Me.Voucher3.Width = Me.picVoucher.Width  '- 1 * GetScrollWidth
    End If
    
    'zcy修改 作废或现结需要减小宽度，否则表体滚动条看不到

    'SetScrollBarValue
    
    Me.Picture32.Width = Me.Width
    Me.Picture32.Move 0, 500
    Me.Picture31.BackColor = Me.Picture32.BackColor
    Me.Picture31.Move Me.Picture32.Width - Me.Picture31.Width - 400
    Me.LabelVoucherName3.Move (Me.Width - Me.LabelVoucherName3.Width) / 2
    
    If labZF.Visible = False And labXJ.Visible = False Then
        Me.Voucher3.Move 0, Me.Picture32.Height
    Else
        Me.Voucher3.Move 0, Me.Picture32.Height + Me.labZF.Height
        Me.Voucher3.Width = Me.Voucher3.Width - 350   'zcy修改 作废或现结需要减小宽度，否则表体滚动条看不到
        Me.Voucher3.Height = Me.Voucher3.Height - Me.labZF.Height - Picture32.Height
    End If
    
    '//排序按钮
'    With U8VoucherSorter4
'        .BackColor = Me.Picture32.BackColor
'        .Left = 550
'        .Top = Me.picVoucher.Top + Me.Picture31.Top
''        .ZOrder
'    End With
    '//////////////////////////////////////////////////////////////////////
    
 '//////////////////////////////////////////////////////////////////////
      '//Voucher 第4个页面Resize
    If Me.picVoucher.Height - 1 * GetScrollWidth < dOriVoucherHeight Then
        Me.Voucher4.Height = Me.picVoucher.Height - Picture42.Height '- 1 * GetScrollWidth
    Else
        Me.Voucher4.Height = Me.picVoucher.Height - Picture42.Height '- 1 * GetScrollWidth
    End If
    If Me.picVoucher.Width - 1 * GetScrollWidth < dOriVoucherWidth Then
        Me.Voucher4.Width = Me.picVoucher.Width '- 1 * GetScrollWidth
    Else
        Me.Voucher4.Width = Me.picVoucher.Width  '- 1 * GetScrollWidth
    End If
    
    'zcy修改 作废或现结需要减小宽度，否则表体滚动条看不到

    'SetScrollBarValue
    
    Me.Picture42.Width = Me.Width
    Me.Picture42.Move 0, 500
    Me.Picture41.BackColor = Me.Picture42.BackColor
    Me.Picture41.Move Me.Picture42.Width - Me.Picture41.Width - 400
    Me.LabelVoucherName4.Move (Me.Width - Me.LabelVoucherName4.Width) / 2
    
    If labZF.Visible = False And labXJ.Visible = False Then
        Me.Voucher4.Move 0, Me.Picture42.Height
    Else
        Me.Voucher4.Move 0, Me.Picture42.Height + Me.labZF.Height
        Me.Voucher4.Width = Me.Voucher4.Width - 350   'zcy修改 作废或现结需要减小宽度，否则表体滚动条看不到
        Me.Voucher4.Height = Me.Voucher4.Height - Me.labZF.Height - Picture42.Height
    End If
    
    '//排序按钮
'    With U8VoucherSorter5
'        .BackColor = Me.Picture42.BackColor
'        .Left = 550
'        .Top = Me.picVoucher.Top + Me.Picture41.Top
''        .ZOrder
'    End With
    '//////////////////////////////////////////////////////////////////////
    
    '//Voucher 第0个页面Resize
    If Me.picVoucher.Height - 1 * GetScrollWidth < dOriVoucherHeight Then
        Me.Voucher.Height = Me.picVoucher.Height - Picture2.Height '- 1 * GetScrollWidth
    Else
        Me.Voucher.Height = Me.picVoucher.Height - Picture2.Height '- 1 * GetScrollWidth
    End If
    If Me.picVoucher.Width - 1 * GetScrollWidth < dOriVoucherWidth Then
        Me.Voucher.Width = Me.picVoucher.Width '- 1 * GetScrollWidth
    Else
        Me.Voucher.Width = Me.picVoucher.Width  '- 1 * GetScrollWidth
    End If
    
    'zcy修改 作废或现结需要减小宽度，否则表体滚动条看不到

    'SetScrollBarValue
    
    Me.Picture2.Width = Me.Width
    Me.Picture2.Move 0, 500
    Me.Picture1.BackColor = Me.Picture2.BackColor
    Me.Picture1.Move Me.Picture2.Width - Me.Picture1.Width - 400
    Me.LabelVoucherName.Move (Me.Width - Me.LabelVoucherName.Width) / 2
    Me.Picture2.ZOrder
    Me.Picture1.ZOrder
    
    If labZF.Visible = False And labXJ.Visible = False Then
        Me.Voucher.Move 0, Me.Picture2.Height
    Else
        Me.Voucher.Move 0, Me.Picture2.Height + Me.labZF.Height
        Me.Voucher.Width = Me.Voucher.Width - 350   'zcy修改 作废或现结需要减小宽度，否则表体滚动条看不到
        Me.Voucher.Height = Me.Voucher.Height - Me.labZF.Height - Picture2.Height
    End If
    
    '//排序按钮
'    With U8VoucherSorter1
'        .BackColor = Me.Picture2.BackColor
'        .Left = 550
'        .Top = Me.picVoucher.Top + Me.Picture1.Top
''        .ZOrder
'    End With
    '//////////////////////////////////////////////////////////////////////
    
    
    
    'Me.Voucher.Move 0, 0, 12000, 12000
    ''  置作废、现结的位置
    labZF.Top = picVoucher.Top + Me.Picture2.Height  'Me.top - Me.tbrvoucher.top
    labZF.Left = Me.Voucher.Left
    labXJ.Top = picVoucher.Top + Me.Picture2.Height ' Me.top - Me.tbrvoucher.top    'Me.StBar.height
    labXJ.Left = Me.Voucher.Left '+ labZF.Width
    With StBar
        .Panels(1).Width = Me.Width * 1 / 3
        .Panels(2).Width = Me.Width * 1 / 3
        .Panels(3).Width = Me.Width * 1 / 3
    End With
    
    Me.BackColor = Me.Voucher.BackColor
    Me.ForeColor = Me.Voucher.BackColor
    Me.picVoucher.BackColor = Me.Voucher.BackColor
    
End Sub
 

Private Function FillVoucher(Domhead As DOMDocument, Dombody As DOMDocument, Optional bClearBody As Boolean = False) As Boolean
    Dim lngCol As Long, lngRow As Long, rows As Long
    Dim i  As Long
    Dim ele As IXMLDOMElement
    Dim ns As IXMLDOMNode
    Dim nodS As IXMLDOMNode
    Dim NODs2 As IXMLDOMNode
    Dim elelist As IXMLDOMNodeList
    Dim elelist2 As IXMLDOMNodeList
    Dim eleTmp As IXMLDOMElement
    Dim linedom As DOMDocument
    Dim oDomH As New DOMDocument
    
    Dim ndRS    As IXMLDOMNode
    Dim nd      As IXMLDOMNode
    With Me.Voucher
        Set linedom = New DOMDocument
        .getVoucherDataXML oDomH, linedom
        Set ns = linedom.selectSingleNode("//rs:data")
        Set elelist = linedom.selectNodes("//z:row[@cinvcode = '']")
        If (Not ns Is Nothing) And elelist.length <> 0 Then
            For Each nodS In elelist
                ns.removeChild nodS
            Next
        End If
        If bClearBody = True Then
            Call ClearAllLineByDom(linedom)
        End If
        .setVoucherDataXML oDomH, linedom
        .BodyMaxRows = 0
        rows = .BodyRows
        If Not Domhead Is Nothing Then
             For Each ele In Domhead.selectNodes("//R")
                
                If LCase(ele.getAttribute("K")) = "cexch_name" Then
                    .ItemState("iexchrate", siheader).nNumPoint = clsSAWeb_M.GetExchRateDec(.headerText("cexch_name"))
                End If
                If LCase(ele.getAttribute("K")) = "minddate" Then
                    maxRefDate = CDate(ele.getAttribute("V"))
                End If
            Next
        End If
        If Not Dombody Is Nothing Then
            lngCol = .BodyRows
            Set elelist2 = Dombody.selectNodes("//z:row")
            If elelist2.length > 5 Then
                lngCol = lngCol + elelist2.length
                .getVoucherDataXML Domhead, linedom
                Set ns = linedom.selectSingleNode("//rs:data")
                If ns Is Nothing Then
                    Set ns = linedom.createElement("rs:data")
                    linedom.selectSingleNode("xml").appendChild ns
                End If
                Set ns = linedom.selectSingleNode("//rs:data")
                For Each ele In elelist2
                    '.AddLine
                    ele.setAttribute "editprop", "A"
                    ns.appendChild ele
                Next
                .setVoucherDataXML Domhead, linedom
                .row = lngCol
            Else
                For Each NODs2 In elelist2
                    .AddLine
                    lngCol = lngCol + 1
                    .row = lngCol
                    Set ns = linedom.selectSingleNode("//rs:data")
                    If ns Is Nothing Then
                        Set eleTmp = linedom.createElement("rs:data")
                        linedom.documentElement.appendChild eleTmp
                    End If
                    Set elelist = linedom.selectNodes("//z:row")
                    For Each nodS In elelist
                        ns.removeChild nodS
                    Next
                    Set ns = linedom.selectSingleNode("//rs:data")
                    ns.appendChild NODs2
                    .UpdateLineData linedom ', lngCol
                Next
            End If
        End If
    End With
End Function

Public Sub ButtonClick(s As String, sTaskKey As String, Optional bCloseSingle As Boolean = False)
    Dim i As Long
    Dim j As Long
    Dim row As Long
    Dim objGoldTax As Object
    Dim strError As String
    Dim strXMLHead As String
    Dim strXMLBody As String
    Dim lngRow As Integer
    Dim lngCol As Integer
    Dim strID As Variant
    Dim ele As IXMLDOMElement
    Dim strAuthID As String
    Dim elelist As IXMLDOMNodeList
    Dim ndRS    As IXMLDOMNode
    Dim nd      As IXMLDOMNode
    Dim bEAlast As Boolean
    Dim sPrnTmplate As Long
    Dim VoucherGrid As Object
    Dim strsql As String
        Dim AffectedLine As Long
    Dim Frm As New frmVouchNew
    On Error GoTo Err
    
    If clsTbl.ButtonKeyDown(m_Login, s) Then
    
    bCloseFHSingle = bCloseSingle
    strErrMsg = ""
    i = 0
    Set domPrint = Nothing
    With Voucher
        Select Case LCase(s)
           Case "openorder"  '打开
              Screen.MousePointer = vbHourglass
                If VoucherTask(sTaskKey) = False Then Screen.MousePointer = vbDefault: Exit Sub
                strError = ""
                Set Domhead = Me.Voucher.GetHeadDom
                         Select Case strVouchType
                         Case "26"
                            strsql = "update EFYZGL_pressinform set ccloser='' where id =" & Me.Voucher.headerText("id") & ""
                         End Select
                             AffectedLine = 0
                             DBConn.Execute strsql, AffectedLine
                If strError <> "" Then
                        Call ShowErrDom(strError, Domhead)
                End If
                LoadVoucher ""
                Call SetButtonStatus(s)
                Call VoucherFreeTask
           Case "closeorder"  '关闭
              Screen.MousePointer = vbHourglass
                If VoucherTask(sTaskKey) = False Then Screen.MousePointer = vbDefault: Exit Sub
                strError = ""
                Set Domhead = Me.Voucher.GetHeadDom
                        Select Case strVouchType
                         Case "26"
                            strsql = "update EFYZGL_pressinform set ccloser='" & m_Login.cUserName & "' where id =" & Me.Voucher.headerText("id") & ""
                        End Select
                            AffectedLine = 0
                            DBConn.Execute strsql, AffectedLine
                If strError <> "" Then
                        Call ShowErrDom(strError, Domhead)
                End If
                LoadVoucher ""
                Call SetButtonStatus(s)
                Call VoucherFreeTask
                    
                
           Case "shiftto"  '转入
                Call createOrder(False)
                LoadVoucher ""
                Call SetButtonStatus(s)
                Call VoucherFreeTask
                
           Case "filter"  '过滤
            'by lg070315　增加u870单据新的定位过滤
''                voucher.ShowFindDlg
'                Dim Frmlist As New frmVoucherList
'                With Frmlist
'                Select Case strVouchType
'                 Case "00"
'                        .Sysid = "FA"
'                        .VouchKey = "FA110"
'                        .strTaskId = strAuthId
'                        .VouchType = strVouchType
'                  Case "26"   '付印通知单
'                        .Sysid = "EFYZGL"
'                        .VouchKey = "EFYZGL03"
'                        .strTaskId = strAuthId
'                        .VouchType = strVouchType
'                        .Caption = "付印通知单列表"
'                        .HelpContextID = 2009000080
'                End Select
'                If .Filter Then
'                        If g_business Is Nothing Then
'                            .Show
'                        Else
'                            Call g_business.ShowForm(Frmlist, "EFYZGL", .strsguid, False, True, .Object_vfd)
'                        End If
'                End If
'                End With
            Case "add"            '//增加
                picVoucher.Tab = 0
                If ChangeDJMBForEdit = False Then Screen.MousePointer = vbDefault: Exit Sub
                If VoucherTask(sTaskKey) = False Then Screen.MousePointer = vbDefault: Exit Sub
                Screen.MousePointer = vbHourglass
                EditPanel_1 EdtP, 3, ""
                labZF.Visible = False
                labXJ.Visible = False
                Me.Voucher.AddNew ANMNormalAdd, Domhead, Dombody '
               '//其他几个模板的增加
                Me.Voucher1.AddNew ANMNormalAdd, Domhead1, Dombody1
                Me.Voucher2.AddNew ANMNormalAdd, Domhead2, Dombody2
                Me.Voucher3.AddNew ANMNormalAdd, Domhead3, Dombody3
                Me.Voucher4.AddNew ANMNormalAdd, Domhead4, Dombody4
                Call SetVouchNoWriteble      '设置单据号是否可以编辑
                Call AddNewVouch(, Voucher)             '设置新增单据的初始值
                Call AddNewVouch(, Voucher1)
                Call AddNewVouch(, Voucher2)
                Call AddNewVouch(, Voucher3)
                Call AddNewVouch(, Voucher4)
                Me.Voucher.AddNew ANMCopyALL, Domhead, Dombody
                
                Me.Voucher1.AddNew ANMCopyALL, Domhead1, Dombody1
                Me.Voucher2.AddNew ANMCopyALL, Domhead2, Dombody2
                Me.Voucher3.AddNew ANMCopyALL, Domhead3, Dombody3
                Me.Voucher4.AddNew ANMCopyALL, Domhead4, Dombody4

                Set Domhead = Me.Voucher.GetHeadDom
                If iShowMode = 2 Then

                    Me.Voucher.setVoucherDataXML mDom, oDomB 'DomBody
                End If
                iVouchState = 0
                Call SetButtonStatus(s)
                Call SetItemState(s)
                If iShowMode = 2 Then
                    With Me.Voucher
                        .Visible = False
                        .EnableHead "ccusabbname", False
                        .EnableHead "cdepname", False
                        .EnableHead "cpersonname", False
                        .Visible = True
                    End With
                End If
                
            Case "chenged" '变更
'                If strVouchType = "97" Or strVouchType = "96" Then
'                    Call Frm.ShowVoucher(gdzckpxg, Me.voucher.headerText("id"))
'               End If
            '/////////////////////////////////////////////////////////////////////////////////////////
            '  860sp升级到861修改处1 注释    2006/03/09 861版本中单据控件增加单据附件功能（附件的可以是文件，图片附件的上限大小为1M）
                Me.Voucher.SelectFile
               
            Case "modify"              '//修改
                picVoucher.Tab = 0
                If CheckDJMBAuth(Me.Voucher.headerText("ivtid"), "W") = False Then
                    MsgBox "当前操作员没有当前单据模版的使用权限！"
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                If VoucherTask(sTaskKey) = False Then Screen.MousePointer = vbDefault: Exit Sub
                Me.Voucher.row = 0
                Me.Voucher.col = 0
                Screen.MousePointer = vbHourglass
                .getVoucherDataXML Domhead, Dombody
                Voucher1.getVoucherDataXML Domhead1, Dombody1
                Voucher2.getVoucherDataXML Domhead2, Dombody2
                Voucher3.getVoucherDataXML Domhead3, Dombody3
                Voucher4.getVoucherDataXML Domhead4, Dombody4
                Me.Voucher.VoucherStatus = VSeEditMode
                Me.Voucher1.VoucherStatus = VSeEditMode
                Me.Voucher2.VoucherStatus = VSeEditMode
                Me.Voucher3.VoucherStatus = VSeEditMode
                Me.Voucher4.VoucherStatus = VSeEditMode
                Call SetItemState(s)
                Call AddNewVouch("modify", Voucher)
                Call SetVouchNoWriteble
                Call SetButtonStatus(s)
                iVouchState = 1
                .SetFocus
                .UpdateCmdBtn
                
            Case "erase"                 '//删除
                If CheckDJMBAuth(Me.Voucher.headerText("ivtid"), "W") = False Then
                    MsgBox "当前操作员没有当前单据模版的使用权限！"
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                If MsgBox("确实要删除本张单据吗？", vbYesNo + vbQuestion) = vbNo Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                If VoucherTask(sTaskKey) = False Then Screen.MousePointer = vbDefault: Exit Sub
                Set Domhead = Me.Voucher.GetHeadDom
                bCreditCheck = False
                Screen.MousePointer = vbHourglass
                strError = clsVoucherCO.Delete(Domhead)
                If strError <> "" Then
                    ShowErrDom strError, Domhead
                    LoadVoucher ""
                Else
                    LoadVoucher "tonext"
                End If
                Call VoucherFreeTask

            Case "copy"                         '//复制
                If ChangeDJMBForEdit = False Then Screen.MousePointer = vbDefault: Exit Sub
                If VoucherTask(sTaskKey) = False Then Exit Sub
                labZF.Visible = False
                labXJ.Visible = False
                Screen.MousePointer = vbHourglass
                AddNewVouch "copy", Voucher
                Call AddNewVouch("copy", Voucher1)
                Call AddNewVouch("copy", Voucher2)
                Call AddNewVouch("copy", Voucher3)
                Call AddNewVouch("copy", Voucher4)
                Me.Voucher.AddNew ANMCopyALL, Domhead, Dombody
                Me.Voucher1.AddNew ANMCopyALL, Domhead1, Dombody1
                Me.Voucher2.AddNew ANMCopyALL, Domhead2, Dombody2
                Me.Voucher3.AddNew ANMCopyALL, Domhead3, Dombody3
                Me.Voucher4.AddNew ANMCopyALL, Domhead4, Dombody4
                Call SetVouchNoWriteble
                Call SetButtonStatus(s)
                iVouchState = 0
                Call SetItemState(s)
                .SetFocus
                .UpdateCmdBtn
            Case "addrow", "addline"                      '//增加一行
                'if picVoucher.Tab=
                Select Case picVoucher.Tab
                    Case 0
                        Voucher.AddLine
                    Case 1
                        Voucher1.AddLine
                    Case 2
                        Voucher2.AddLine
                    Case 3
                        Voucher3.AddLine
                    Case 4
                        Voucher4.AddLine
                End Select
'                With Me.voucher
'                    .AddLine
'                End With
'                  Me.Voucher1.AddLine
'                  Me.Voucher2.AddLine
'                  Me.Voucher3.AddLine
'                  Me.Voucher4.AddLine
            Case "delrow", "delline"                    '//删除一行
                    Dim tmpRow As Variant
                        Select Case picVoucher.Tab
                         Case 0
                            tmpRow = Me.Voucher.row - 1
                            Me.Voucher.DelLine Me.Voucher.row
                            If tmpRow <> 0 Then
                                Me.Voucher.row = tmpRow
                            Else
                                Me.Voucher.row = 0
                                Me.Voucher.col = 0
                            End If
                         Case 1
                            tmpRow = Me.Voucher1.row - 1
                            Me.Voucher1.DelLine Me.Voucher1.row
                            If tmpRow <> 0 Then
                                Me.Voucher1.row = tmpRow
                            Else
                                Me.Voucher1.row = 0
                                Me.Voucher1.col = 0
                            End If
                         Case 2
                            tmpRow = Me.Voucher2.row - 1
                            Me.Voucher2.DelLine Me.Voucher2.row
                            If tmpRow <> 0 Then
                                Me.Voucher2.row = tmpRow
                            Else
                                Me.Voucher2.row = 0
                                Me.Voucher2.col = 0
                            End If
                         Case 3
                            tmpRow = Me.Voucher3.row - 1
                            Me.Voucher3.DelLine Me.Voucher3.row
                            If tmpRow <> 0 Then
                                Me.Voucher3.row = tmpRow
                            Else
                                Me.Voucher3.row = 0
                                Me.Voucher3.col = 0
                            End If
                        Case 4
                            tmpRow = Me.Voucher4.row - 1
                            Me.Voucher4.DelLine Me.Voucher4.row
                            If tmpRow <> 0 Then
                                Me.Voucher4.row = tmpRow
                            Else
                                Me.Voucher4.row = 0
                                Me.Voucher4.col = 0
                            End If
                        End Select
                
            Case "sure"           '//审核
                Screen.MousePointer = vbHourglass
                If VoucherTask(sTaskKey) = False Then Screen.MousePointer = vbDefault: Exit Sub
                Call AddNewVouch(s, Voucher)
                bCreditCheck = True
                Set Domhead = Me.Voucher.GetHeadDom
                strError = clsVoucherCO.VerifyVouch(Domhead, bCreditCheck)
                Call ShowErrDom(strError, Domhead)
                ''刷新当前单据
                LoadVoucher ""
                Call VoucherFreeTask
                
            Case "unsure"            '//弃审
                Screen.MousePointer = vbHourglass
                If VoucherTask(sTaskKey) = False Then Screen.MousePointer = vbDefault: Exit Sub
                Call AddNewVouch(s, Voucher)
                bCreditCheck = False
                Set Domhead = Me.Voucher.GetHeadDom
                strError = clsVoucherCO.VerifyVouch(Domhead, bCreditCheck)
                Call ShowErrDom(strError, Domhead)
                ''刷新当前单据
                LoadVoucher ""
                Call VoucherFreeTask
            Case "cancel"                  '//取消
                picVoucher.Tab = 0
                bClickCancel = True
                Voucher.VoucherStatus = VSNormalMode
                Voucher1.VoucherStatus = VSNormalMode
                Voucher2.VoucherStatus = VSNormalMode
                Voucher3.VoucherStatus = VSNormalMode
                Voucher4.VoucherStatus = VSNormalMode
                LoadVoucher ""
                bOnceRefer = False
                Call SetButtonStatus(s)
                ChangeButtonsState
                bClickCancel = False
                Call VoucherFreeTask
            Case "save"                    '//保存
                picVoucher.Tab = 0
                Screen.MousePointer = vbHourglass
                Voucher.ProtectUnload2
                Voucher1.ProtectUnload2
                Voucher2.ProtectUnload2
                Voucher3.ProtectUnload2
                Voucher4.ProtectUnload2
                bClickCancel = False
                bClickSave = True
                strError = ""
                
                If s1trVouchType = "99" Then    '封面
                    For i = Me.Voucher1.BodyRows To 1 Step -1
                        If Me.Voucher1.bodyText(i, "citemcode") = "" Then
                            Me.Voucher1.DelLine i
                        End If
                    Next i
                End If
                If s2trVouchType = "07" Then   '送书地址和数量
                    For i = Me.Voucher2.BodyRows To 1 Step -1
                        If Me.Voucher2.bodyText(i, "address") = "" Then
                            Me.Voucher2.DelLine i
                        End If
                    Next i
                End If
                If s3trVouchType = "95" Then   '内容和印装方法
                    For i = Me.Voucher3.BodyRows To 1 Step -1
                        If Me.Voucher3.bodyText(i, "content") = "" Then
                            Me.Voucher3.DelLine i
                        End If
                    Next i
                End If
                If s4trVouchType = "27" Then  '纸张材料
                    For i = Me.Voucher4.BodyRows To 1 Step -1
                        If Me.Voucher4.bodyText(i, "cinvcode") = "" Then
                            Me.Voucher4.DelLine i
                        End If
                    Next i
                End If
                
                If Me.Voucher.BodyRows = 0 And strVouchType <> "26" Then
                    MsgBox "表体没有记录，请录入！"
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                
                If Me.Voucher1.BodyRows = 0 Then
                    MsgBox "封面信息表体没有记录，请录入！"
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                
                If Me.Voucher2.BodyRows = 0 Then
                    MsgBox "送书信息表体没有记录，请录入！"
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                
                If Me.Voucher3.BodyRows = 0 Then
                    MsgBox "内容信息表体没有记录，请录入！"
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                
                If Me.Voucher4.BodyRows = 0 Then
                    MsgBox "纸张材料信息表体没有记录，请录入！"
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                If .headVaildIsNull2(strError) = False Then
                    MsgBox "表头项目" + strError
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                
                strError = ""
                If .bodyVaildIsNull2(strError) = False Then
                    MsgBox "表体项目" + strError
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                strError = ""
                Call AddNewVouch("Save", Voucher)
                Voucher.getVoucherDataXML Domhead, Dombody
                Voucher1.getVoucherDataXML Domhead1, Dombody1
                Voucher2.getVoucherDataXML Domhead2, Dombody2
                Voucher3.getVoucherDataXML Domhead3, Dombody3
                Voucher4.getVoucherDataXML Domhead4, Dombody4
                
                 If Voucher4.BodyRows >= 10 Then
                    Set ndRS = Dombody4.selectSingleNode("//rs:data")

                    Set elelist = Dombody4.selectNodes("//z:row[@cinvcode = '']")

                    If (Not ndRS Is Nothing) And elelist.length <> 0 Then
                        For Each nd In elelist
                            ndRS.removeChild nd
                        Next
                    End If
                    .setVoucherDataXML Domhead4, Dombody4
                    If Dombody4.selectNodes("//z:row").length = 0 Then
                        MsgBox "纸张材料表体记录为0！"
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
                
                If Voucher1.BodyRows >= 10 Then
                    Set ndRS = Dombody1.selectSingleNode("//rs:data")

                    Set elelist = Dombody1.selectNodes("//z:row[@citemcode = '']")

                    If (Not ndRS Is Nothing) And elelist.length <> 0 Then
                        For Each nd In elelist
                            ndRS.removeChild nd
                        Next
                    End If
                    .setVoucherDataXML Domhead1, Dombody1
                    If Dombody1.selectNodes("//z:row").length = 0 Then
                        MsgBox "封面信息表体记录为0！"
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
                
                
                If Voucher2.BodyRows >= 10 Then
                    Set ndRS = Dombody2.selectSingleNode("//rs:data")

                    Set elelist = Dombody2.selectNodes("//z:row[@address = '']")

                    If (Not ndRS Is Nothing) And elelist.length <> 0 Then
                        For Each nd In elelist
                            ndRS.removeChild nd
                        Next
                    End If
                    .setVoucherDataXML Domhead2, Dombody2
                    If Dombody2.selectNodes("//z:row").length = 0 Then
                        MsgBox "送书信息表体记录为0！"
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
                
                If Voucher3.BodyRows >= 10 Then
                    Set ndRS = Dombody3.selectSingleNode("//rs:data")

                    Set elelist = Dombody3.selectNodes("//z:row[@content = '']")

                    If (Not ndRS Is Nothing) And elelist.length <> 0 Then
                        For Each nd In elelist
                            ndRS.removeChild nd
                        Next
                    End If
                    .setVoucherDataXML Domhead3, Dombody3
                    If Dombody3.selectNodes("//z:row").length = 0 Then
                        MsgBox "内容信息表体记录为0！"
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                End If

                
                '//给单据表体赋模板号
                Set elelist = Dombody1.selectNodes("//z:row")
                For i = 0 To elelist.length - 1
                    Set ele = Dombody1.selectNodes("//z:row").Item(i)
                    ele.setAttribute "ivtid", s1CurTemplateID2
                Next
                
                Set elelist = Dombody2.selectNodes("//z:row")
                For i = 0 To elelist.length - 1
                    Set ele = Dombody2.selectNodes("//z:row").Item(i)
                    ele.setAttribute "ivtid", s2CurTemplateID2
                Next
                 
                Set elelist = Dombody3.selectNodes("//z:row")
                For i = 0 To elelist.length - 1
                    Set ele = Dombody3.selectNodes("//z:row").Item(i)
                    ele.setAttribute "ivtid", s3CurTemplateID2
                Next
                
                Set elelist = Dombody4.selectNodes("//z:row")
                For i = 0 To elelist.length - 1
                    Set ele = Dombody4.selectNodes("//z:row").Item(i)
                    ele.setAttribute "ivtid", s4CurTemplateID2
                Next
                
                Set elelist = Dombody.selectNodes("//z:row")
                For i = 0 To elelist.length - 1
                    Set ele = Dombody.selectNodes("//z:row").Item(i)
                    ele.setAttribute "ivtid", sCurTemplateID2
                Next
                '////////////////////////////////////////////////////////////////////////////////////////////////
                '860sp升级到861修改处   2006/03/08   增加单据附件功能
                If SetAttachXML(Domhead) = False Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                '////////////////////////////////////////////////////////////////////////////////////////////////
                If bFirst = True Then
                    Domhead.selectSingleNode("//z:row").Attributes.getNamedItem("bfirst").nodeValue = "1"
                End If
                bCreditCheck = False
ToSave:
                strError = clsVoucherCO.SaveMultVouch(Domhead, Dombody, iVouchState, vNewID, domConfig, Dombody1, Dombody2, Dombody3, Dombody4)
                If strError <> "" Then
                    If InStr(1, strError, "<", vbTextCompare) <> 0 Then
                        ShowErrDom strError, Domhead
                    Else
                        MsgBox IIf(Trim(strError) = "当前操作不成功，请重新再试!", "", strError)
                        If Domhead.selectNodes("//z:row").length = 1 Then
                            If .headerText(getVoucherCodeName) <> GetHeadItemValue(Domhead, getVoucherCodeName) And strVouchType <> "92" Then
                                .headerText(getVoucherCodeName) = GetHeadItemValue(Domhead, getVoucherCodeName)
                            End If
                        End If
                    End If
                Else
                    Voucher.VoucherStatus = VSNormalMode
                    Voucher1.VoucherStatus = VSNormalMode
                    Voucher2.VoucherStatus = VSNormalMode
                    Voucher3.VoucherStatus = VSNormalMode
                    Voucher4.VoucherStatus = VSNormalMode
                    LoadVoucher "", IIf(vNewID <> "", vNewID, 0)
                    bOnceRefer = False
                    Call SetButtonStatus(s)
                    ChangeButtonsState
                    Call VoucherFreeTask
                End If
                bClickSave = False
                If strVouchType = "98" Then
                    Unload Me
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            Case "print"            '打印
                LoadData
                
                If Me.ComboVTID.ListCount = 0 Then
                    MsgBox "当前操作员没有可以使用的打印模版，请检查！"
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                
                sPrnTmplate = CLng(vtidPrn(Me.ComboVTID.ListIndex))
                VoucherPrn strVouchType, Voucher5, "EFYZGL030301", clsSAWeb_M.GetVTID("pbprint", DBConn, s5trCardNum), , True
                .VoucherStatus = VSNormalMode
                LoadVoucher ""
           
            Case "preview"
                LoadData
                
                If Me.ComboVTID.ListCount = 0 Then
                    MsgBox "当前操作员没有可以使用的打印模版，请检查！"
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                    
                    sPrnTmplate = CLng(vtidPrn(Me.ComboVTID.ListIndex))
                    VoucherPrn strVouchType, Voucher5, "EFYZGL030301", clsSAWeb_M.GetVTID("pbprint", DBConn, s5trCardNum), "Preview", True
                    .VoucherStatus = VSNormalMode
                    LoadVoucher ""
                    
            Case "output"
                    VouchOutPut Voucher, CLng(sTemplateID), strCardNum
            Case "exit"
                Unload Me
                Screen.MousePointer = vbDefault
                Exit Sub
            Case "seek"
                '-----------------------------------------------------
                '由单据联查凭证
                 Find_GL_accvouch
            
            Case "paint"
                Screen.MousePointer = vbHourglass
                LoadVoucher "", , True
                
            Case LCase("ToPrevious")   '上一张
                picVoucher.Tab = 0
                Screen.MousePointer = vbHourglass
                LoadVoucher s, , True
                  
                Voucher.VoucherStatus = VSNormalMode
            
            Case LCase("ToNext")   '下一张
                picVoucher.Tab = 0
                Screen.MousePointer = vbHourglass
                LoadVoucher s, , True
        
                Voucher.VoucherStatus = VSNormalMode
                
            Case LCase("ToLast")   '末张
                picVoucher.Tab = 0
                Screen.MousePointer = vbHourglass
                LoadVoucher s, , True
 
                Voucher.VoucherStatus = VSNormalMode
                
            Case LCase("ToFirst")   '首张
                picVoucher.Tab = 0
                Screen.MousePointer = vbHourglass
                LoadVoucher s, , True
 
                Voucher.VoucherStatus = VSNormalMode
                
            Case "attached"            '增加附件
                Me.Voucher.SelectFile
            Case "refresh"               '刷新
                Screen.MousePointer = vbHourglass
'                If val(GetHeadItemValue(Domhead, "vt_id")) <> 0 Then
                    LoadVoucher "", , True
'                End If
            Case LCase("CopyRow")
                Set Dombody = .GetLineDom
                Set Domhead = New DOMDocument
                clsVoucherCO.CopyRow Dombody
                i = Voucher.BodyMaxRows
                Voucher.BodyMaxRows = 0
                .AddLine
                Voucher.BodyMaxRows = i
                .UpdateLineData Dombody, .BodyRows
            Case LCase("LookVeri")  ''查询审批流
'                If .VoucherStatus = VSeEditMode Then .ProtectUnload2
'                Set Domhead = .GetHeadDom
'                If obj_EA.NeedEAFControl(clsSAWeb_M.GetEAsCode(strVouchType, Domhead), GetHeadItemValue(Domhead, clsSAWeb_M.getVouchMainIDName(strVouchType))) Then
'                    If (obj_EA.ResearchEAStream(clsSAWeb_M.GetEAsCode(strVouchType, Domhead), .headerText(clsSAWeb_M.getVouchMainIDName(strVouchType)))) = False Then
'                        MsgBox obj_EA.ErrDescript
'                    End If
'                Else
'                    MsgBox "该单据未进入审批流!"
'                End If
           Case "help"
'            SendKeys "{F1}"
            On Error Resume Next
            ShowContextHelp Me.hwnd, App.HelpFile, Me.HelpContextID
        End Select
        
        clsTbl.ChangeButtonState Voucher, Me.tbrvoucher, Me.UFToolbar1, Voucher.VoucherStatus
   End With
   Set ele = Nothing
   Screen.MousePointer = vbDefault
   ProgressBar1.Visible = False
FreeTask:
   clsTbl.ButtonKeyUp m_Login, s
   End If
   
   Exit Sub
Err:
    ProgressBar1.Visible = False
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
End Sub

Private Function IsInFormsByTag(strTag As String) As Boolean
    Dim frmTmp As Form
    
    IsInFormsByTag = False
    For Each frmTmp In Forms
        If frmTmp.Tag = strTag Then
            frmTmp.ZOrder 0
            IsInFormsByTag = True
            Exit For
        End If
    Next
End Function


Private Sub Form_Unload(Cancel As Integer)
doNext:
    If Me.Voucher.VoucherStatus <> VSNormalMode Then
        Select Case MsgBox("是否保存对当前单据的编辑？", vbYesNoCancel + vbQuestion)
            Case vbYes
                ButtonClick "Save", "保存"
                If Me.Voucher.VoucherStatus = VSNormalMode Then
                    GoTo DoQuit
                End If
            Case vbNo
                VoucherFreeTask
                GoTo DoQuit
            Case vbCancel
                
        End Select

        bFrmCancel = True
        Me.ZOrder
        Cancel = 3
    Else
DoQuit:
        On Error Resume Next
        bFrmCancel = False
'by lg070314增加U870菜单融合，关闭时处理Business
        Set UFToolbar1.Business = Nothing
        
        Set clsVoucherCO = Nothing
        Set clsAuth = Nothing
        Set clsRefer = Nothing
        Set RstTemplate = Nothing
        Set Domhead = Nothing
        Set Dombody = Nothing
        Set DomFormat = Nothing
        Set RstTemplate = Nothing
        Set RstTemplate2 = Nothing
        Set obj_EA = Nothing
        Set DOMEA = Nothing
        If m_UFTaskID <> "" Then
            m_Login.TaskExec m_UFTaskID, 0
        End If
        'minForm
   End If
End Sub
 
Private Sub mdiAddRow_Click()
    ButtonClick "AddRow", ""
End Sub
 
Private Sub mdiDelRow_Click()
    ''右键菜单
    ButtonClick "DelRow", ""
End Sub
  
 

Private Sub Label14_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub Label8_Click()

End Sub
'///////zhupb
Private Sub picVoucher_Click(PreviousTab As Integer)
    Dim strErrorResId As String
    Select Case picVoucher.Caption
        Case "付印通知单"
              strCardNum = "EFYZGL030301"
              clsVoucher.Init strCardNum, strErrorResId
              clsVoucherRefer.Init strCardNum, strErrorResId
        Case "封面"
              strCardNum = "EFYZGL030301"
              clsVoucher.Init strCardNum, strErrorResId
              clsVoucherRefer.Init strCardNum, strErrorResId
        Case "送书信息"
              strCardNum = "EFYZGL030301"
              clsVoucher.Init strCardNum, strErrorResId
              clsVoucherRefer.Init strCardNum, strErrorResId
        Case "内容及印装方法"
              strCardNum = "EFYZGL030301"
              clsVoucher.Init strCardNum, strErrorResId
              clsVoucherRefer.Init strCardNum, strErrorResId
        Case "纸张材料"
              strCardNum = "EFYZGL030301"
              clsVoucher.Init strCardNum, strErrorResId
              clsVoucherRefer.Init strCardNum, strErrorResId
    End Select
End Sub

Private Sub tbrvoucher_ButtonClick(ByVal Button As MSComctlLib.Button)
    bCloseFHSingle = False
    ButtonClick Button.key, Button.ToolTipText
End Sub
Private Function getVoucherCodeName() As String
    Dim KeyCode As String
    Select Case strVouchType

        Case "05", "00"
            KeyCode = "cdlcode"
        Case "06"
           KeyCode = "ccode" 'sl
        Case "27", "28", "29", "26"
            KeyCode = "ccode"
        Case "97"
            KeyCode = "ccode"
        Case "98"
            KeyCode = "cevcode"
        Case "99"
            KeyCode = "cspvcode"
        Case "07", "16"
            KeyCode = "ccode"
        Case "95", "92"
            KeyCode = "cwlcode"
        Case Else
            KeyCode = "ccode"
    End Select
    getVoucherCodeName = KeyCode
End Function

Private Sub UFToolbar1_OnCommand(ByVal enumType As UFToolBarCtrl.ENUM_MENU_OR_BUTTON, ByVal cButtonId As String, ByVal cmenuid As String)
    ButtonClick IIf(enumType = enumButton, cButtonId, cmenuid), ""
End Sub

'--------表头自定义项自动带入 added by jzq--------------------------------------
Private Sub Voucher_AutoFillBackEvent(vtIndex As Variant, ByVal vtCurrentValue As Variant, ByVal vtCurrentFieldObject As Variant, ByVal vtAutoFieldInfo As Variant)
 On Error GoTo ErrHandle
    Dim sErrMsg As String
    If moAutoFill.AutoFillRelations(DBConn, Voucher, _
            vtCurrentFieldObject, vtAutoFieldInfo, sErrMsg) = False Then
        MsgBox "填写表头自定义项错误。" & vbCrLf & sErrMsg, vbInformation + vbOKOnly, "提示信息"
    End If
    Exit Sub
ErrHandle:
    MsgBox "填写表头自定义项错误。" & vbCrLf & Err.Description, vbInformation + vbOKOnly, "提示信息"
End Sub




'--------表头自定义项自动带入 added by jzq--------------------------------------
Private Sub Voucher1_AutoFillBackEvent(vtIndex As Variant, ByVal vtCurrentValue As Variant, ByVal vtCurrentFieldObject As Variant, ByVal vtAutoFieldInfo As Variant)
 On Error GoTo ErrHandle
    Dim sErrMsg As String
    If moAutoFill.AutoFillRelations(DBConn, Voucher, _
            vtCurrentFieldObject, vtAutoFieldInfo, sErrMsg) = False Then
        MsgBox "填写表头自定义项错误。" & vbCrLf & sErrMsg, vbInformation + vbOKOnly, "提示信息"
    End If
    Exit Sub
ErrHandle:
    MsgBox "填写表头自定义项错误。" & vbCrLf & Err.Description, vbInformation + vbOKOnly, "提示信息"
End Sub



Private Sub Voucher1_bodyBrowUser(ByVal row As Long, ByVal col As Long, sRet As Variant, referPara As UapVoucherControl85.ReferParameter)
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strErrorResId As String
'    clsVoucher.Init strCardNum, strErrorResId
'    clsVoucherRefer.Init strCardNum, strErrorResId
    lngRow = row
    lngCol = col
    strReferString = clsVoucherRefer.ShowReferCtl(clsVoucher, Voucher1, sibody, col, referPara, row)
    If referPara.Cancel Then
        sRet = Voucher1.bodyText(lngRow, lngCol)
    End If
End Sub

'--------表头自定义项自动带入 added by jzq--------------------------------------
Private Sub Voucher2_AutoFillBackEvent(vtIndex As Variant, ByVal vtCurrentValue As Variant, ByVal vtCurrentFieldObject As Variant, ByVal vtAutoFieldInfo As Variant)
 On Error GoTo ErrHandle
    Dim sErrMsg As String
    If moAutoFill.AutoFillRelations(DBConn, Voucher, _
            vtCurrentFieldObject, vtAutoFieldInfo, sErrMsg) = False Then
        MsgBox "填写表头自定义项错误。" & vbCrLf & sErrMsg, vbInformation + vbOKOnly, "提示信息"
    End If
    Exit Sub
ErrHandle:
    MsgBox "填写表头自定义项错误。" & vbCrLf & Err.Description, vbInformation + vbOKOnly, "提示信息"
End Sub



'--------表头自定义项自动带入 added by jzq--------------------------------------
Private Sub Voucher3_AutoFillBackEvent(vtIndex As Variant, ByVal vtCurrentValue As Variant, ByVal vtCurrentFieldObject As Variant, ByVal vtAutoFieldInfo As Variant)
 On Error GoTo ErrHandle
    Dim sErrMsg As String
    If moAutoFill.AutoFillRelations(DBConn, Voucher, _
            vtCurrentFieldObject, vtAutoFieldInfo, sErrMsg) = False Then
        MsgBox "填写表头自定义项错误。" & vbCrLf & sErrMsg, vbInformation + vbOKOnly, "提示信息"
    End If
    Exit Sub
ErrHandle:
    MsgBox "填写表头自定义项错误。" & vbCrLf & Err.Description, vbInformation + vbOKOnly, "提示信息"
End Sub


'--------表头自定义项自动带入 added by jzq--------------------------------------
Private Sub Voucher4_AutoFillBackEvent(vtIndex As Variant, ByVal vtCurrentValue As Variant, ByVal vtCurrentFieldObject As Variant, ByVal vtAutoFieldInfo As Variant)
 On Error GoTo ErrHandle
    Dim sErrMsg As String
    If moAutoFill.AutoFillRelations(DBConn, Voucher, _
            vtCurrentFieldObject, vtAutoFieldInfo, sErrMsg) = False Then
        MsgBox "填写表头自定义项错误。" & vbCrLf & sErrMsg, vbInformation + vbOKOnly, "提示信息"
    End If
    Exit Sub
ErrHandle:
    MsgBox "填写表头自定义项错误。" & vbCrLf & Err.Description, vbInformation + vbOKOnly, "提示信息"
End Sub


Private Sub Voucher_BillNumberChecksucceed()
    Dim errMsg As String, strVouchNo As String, KeyCode As String
    Dim tmpDOM As New DOMDocument
    KeyCode = getVoucherCodeName()
    If strVouchType = "92" Then Exit Sub
    With Me.Voucher
        If Not (LCase(DomFormat.selectSingleNode("//单据编号").Attributes.getNamedItem("重号自动重取").nodeValue) = "false" And _
        LCase(DomFormat.selectSingleNode("//单据编号").Attributes.getNamedItem("允许手工修改").nodeValue) = "true") Then
            Set tmpDOM = .GetHeadDom

            If clsVoucherCO.GetVoucherNO(tmpDOM, strVouchNo, errMsg) = False Then
                MsgBox errMsg
                strCurVoucherNO = ""
            Else
                .headerText(KeyCode) = strVouchNo
                If strVouchType = "97" Then
                    .headerText("sassetnum") = strVouchNo
                End If
                strCurVoucherNO = strVouchNo
            End If
        End If
    End With
End Sub

Private Sub Voucher1_BillNumberChecksucceed()
    Dim errMsg As String, strVouchNo As String, KeyCode As String
    Dim tmpDOM As New DOMDocument
    
    KeyCode = getVoucherCodeName()
    If strVouchType = "92" Then Exit Sub
    With Me.Voucher1
        'If .headerText(KeyCode) = "" And KeyCode <> "" Then
        If Not (LCase(DomFormat.selectSingleNode("//单据编号").Attributes.getNamedItem("重号自动重取").nodeValue) = "false" And _
        LCase(DomFormat.selectSingleNode("//单据编号").Attributes.getNamedItem("允许手工修改").nodeValue) = "true") Then
            Set tmpDOM = .GetHeadDom
        
            If clsVoucherCO.GetVoucherNO(tmpDOM, strVouchNo, errMsg) = False Then
                MsgBox errMsg
                strCurVoucherNO = ""
            Else
                .headerText(KeyCode) = strVouchNo
                strCurVoucherNO = strVouchNo
            End If
        End If
        'End If
    End With
End Sub

Private Sub Voucher2_BillNumberChecksucceed()
    Dim errMsg As String, strVouchNo As String, KeyCode As String
    Dim tmpDOM As New DOMDocument
    
    KeyCode = getVoucherCodeName()
    If strVouchType = "92" Then Exit Sub
    With Me.Voucher2
        'If .headerText(KeyCode) = "" And KeyCode <> "" Then
        If Not (LCase(DomFormat.selectSingleNode("//单据编号").Attributes.getNamedItem("重号自动重取").nodeValue) = "false" And _
        LCase(DomFormat.selectSingleNode("//单据编号").Attributes.getNamedItem("允许手工修改").nodeValue) = "true") Then
            Set tmpDOM = .GetHeadDom
        
            If clsVoucherCO.GetVoucherNO(tmpDOM, strVouchNo, errMsg) = False Then
                MsgBox errMsg
                strCurVoucherNO = ""
            Else
                .headerText(KeyCode) = strVouchNo
                strCurVoucherNO = strVouchNo
            End If
        End If
        'End If
    End With
End Sub


Private Sub Voucher3_BillNumberChecksucceed()
    Dim errMsg As String, strVouchNo As String, KeyCode As String
    Dim tmpDOM As New DOMDocument
    
    KeyCode = getVoucherCodeName()
    If strVouchType = "92" Then Exit Sub
    With Me.Voucher3
        'If .headerText(KeyCode) = "" And KeyCode <> "" Then
        If Not (LCase(DomFormat.selectSingleNode("//单据编号").Attributes.getNamedItem("重号自动重取").nodeValue) = "false" And _
        LCase(DomFormat.selectSingleNode("//单据编号").Attributes.getNamedItem("允许手工修改").nodeValue) = "true") Then
            Set tmpDOM = .GetHeadDom
        
            If clsVoucherCO.GetVoucherNO(tmpDOM, strVouchNo, errMsg) = False Then
                MsgBox errMsg
                strCurVoucherNO = ""
            Else
                .headerText(KeyCode) = strVouchNo
                strCurVoucherNO = strVouchNo
            End If
        End If
        'End If
    End With
End Sub

Private Sub Voucher4_BillNumberChecksucceed()
    Dim errMsg As String, strVouchNo As String, KeyCode As String
    Dim tmpDOM As New DOMDocument
    
    KeyCode = getVoucherCodeName()
    If strVouchType = "92" Then Exit Sub
    With Me.Voucher4
        'If .headerText(KeyCode) = "" And KeyCode <> "" Then
        If Not (LCase(DomFormat.selectSingleNode("//单据编号").Attributes.getNamedItem("重号自动重取").nodeValue) = "false" And _
        LCase(DomFormat.selectSingleNode("//单据编号").Attributes.getNamedItem("允许手工修改").nodeValue) = "true") Then
            Set tmpDOM = .GetHeadDom
        
            If clsVoucherCO.GetVoucherNO(tmpDOM, strVouchNo, errMsg) = False Then
                MsgBox errMsg
                strCurVoucherNO = ""
            Else
                .headerText(KeyCode) = strVouchNo
                strCurVoucherNO = strVouchNo
            End If
        End If
        'End If
    End With
End Sub
Public Function IsInForms(strFormName As String, Optional iIndex As Long) As Boolean
Dim frmIndex As Long
For frmIndex = Forms.Count To 1 Step -1
    If Forms(frmIndex - 1).Caption = strFormName Then
       IsInForms = True
       If Forms(frmIndex - 1).WindowState = 1 Then Forms(frmIndex - 1).WindowState = 2
       iIndex = frmIndex - 1
       Exit Function
    End If
Next
IsInForms = 0
End Function

 
 
Private Sub Voucher_bodyBrowUser(ByVal row As Long, ByVal col As Long, sRet As Variant, referPara As UapVoucherControl85.ReferParameter)
    Dim lngRow As Long
    Dim lngCol As Long
    lngRow = row
    lngCol = col
    strReferString = clsVoucherRefer.ShowReferCtl(clsVoucher, Voucher, sibody, col, referPara, row)
    If referPara.Cancel Then
        sRet = Voucher.bodyText(lngRow, lngCol)
    End If
'    Dim strSql As String, cInvCode As String, cCusCode As String
'    Dim Domhead As DOMDocument, Dombody As DOMDocument
'    Dim Dombodys_str1 As String
'    Dim Dombodys_str2 As String
'    Dim lngRow As Integer
'    Dim I As Integer, lRecord As Long
'    Dim j As Long
'    Dim sKey As String
'    Dim sKeyValue As String, strAuth As String
'    Dim tmpRow As Long, tmpCol As Long
'    Dim tmpCol2 As Long
'    Dim strClass As String
'    Dim strDate As String
'    Dim strGrid As String
'    Dim bMulitSelect As Boolean
'    Dim tmpbFromCurrentStock As Boolean
'    Dim bInvRefSuce As Boolean
'    Dim adoField As ADODB.Field
'    Dim sFormat As String, strExchName As String
'    Dim referParatmp As UAPVoucherControl85.ReferParameter
'        Dim tmpDOMBody As DOMDocument
'    Dim ifalg As Boolean
'    tmpRow = row
'    tmpCol = voucher.col
'    tmpCol2 = col
'    On Error Resume Next
'    With voucher
'        .MultiLineSelect = False ''设置多选默认
'        clsRefer.SetReferSQLString ""
'        clsRefer.SetRWAuth "INVENTORY", "R", False
'        clsRefer.SetReferDisplayMode enuGrid
'        sKey = .ItemState(col, sibody).sFieldName
'        sKeyValue = .bodyText(row, col)
'        Select Case LCase(sKey)
'         Case "htid"
'                clsRefer.StrRefInit m_login, False, "", "select htid,htcontent,htmemo from EFBWGL_dbcbht ", "合同编码,合同内容,合同备注", "", False, 1, 1, 1
'                clsRefer.Show
'                If Not clsRefer.recmx Is Nothing Then
'                  sRet = clsRefer.recmx.Fields("htid")
'                  Me.voucher.bodyText(row, "htcontent") = clsRefer.recmx.Fields("htcontent")
'                  Me.voucher.bodyText(row, "htmeno") = clsRefer.recmx.Fields("htmeno")
'                End If
'
'            Case "cmemo"
'                clsRefer.SetReferDataType 17
'                clsRefer.Show
'
'                If Not clsRefer.recmx Is Nothing Then
'                    sRet = clsRefer.recmx.Fields("ctext")
'
'                End If
'                Exit Sub
'            Case "citemcode"
'                If Me.voucher.bodyText(row, "citem_class") = "" Then
'                    clsRefer.EnumRefInit m_login, enuTreeViewAndGrid, False, enuItem
''                    MsgBox "请先选择项目大类！"
''                    Exit Sub
'                Else
'                    clsRefer.ItemRefInit m_login, False, Me.voucher.bodyText(row, "citem_class")
'                End If
'                'strSQL = "select * from userdefine where cid=" & GetDefineID(LCase(Me.Voucher.ItemState(Col, Sibody).sFieldName))
'                'Call clsRefer.StrRefInit(m_login, False, "", strSQL, "", "UserDefine")
'                'clsRefer.SetReferSQLString strSQL
'                clsRefer.SetReferFilterString getReferString(sKey, sKeyValue)
'                clsRefer.Show
'                If Me.voucher.bodyText(row, "citem_class") = "" Then
'                    If Not clsRefer.RstSelClass.EOF Then
'                        .bodyText(.row, "citem_class") = clsRefer.RstSelClass("citem_class")
'                        .bodyText(.row, "citem_cname") = clsRefer.RstSelClass("citem_name")
'                    End If
'                End If
'            Case "citem_class"
'
'                strSql = "select citem_class ,citem_name from fitem"
'                If clsRefer.StrRefInit(m_login, False, "", strSql, "项目大类编码,项目大类名称", "") = False Then Exit Sub
'                'clsRefer.SetReferSQLString strsql
'                'clsRefer.SetReferDataType enuItem
'                'clsRefer.SetReferFilterString getReferString(sKey, sKeyValue)
'                clsRefer.Show
'        End Select
'End With
''by lg070315　增加U870 UAP单据控件新的参照处理
'referPara.Cancel = True
'
'End Sub
'Private Sub Voucher1_bodyBrowUser(ByVal row As Long, ByVal col As Long, sRet As Variant, referPara As ReferParameter)
'    Dim strSql As String, cInvCode As String, cCusCode As String
'    Dim Domhead As DOMDocument, Dombody As DOMDocument
'    Dim lngRow As Integer
'    Dim I As Integer, lRecord As Long
'    Dim tmpDOMBody As DOMDocument
'    Dim sKey As String, sKeyValue As String, strAuth As String
'    Dim tmprst As ADODB.Recordset
'    Dim tmpRow As Long, tmpCol As Long
'    Dim sFormat As String, strExchName As String
'    Dim adoField As ADODB.Field
'    Dim tmpbFromCurrentStock As Boolean
'    Dim tmpCol2 As Long
'    Dim bInvRefSuce As Boolean
'    Dim strDate As String
'    Dim bMulitSelect As Boolean
'    Dim strCusInv As String
'
'
'    tmpRow = row
'    tmpCol = Voucher1.col
'    tmpCol2 = col
'
'
'    On Error Resume Next
'
'    With Voucher1
'        .MultiLineSelect = True ''设置多选默认
'        clsRefer.SetReferSQLString ""
'        clsRefer.SetRWAuth "INVENTORY", "R", False
'        clsRefer.SetReferDisplayMode enuGrid
'        sKey = .ItemState(col, sibody).sFieldName
'        sKeyValue = .bodyText(row, col)
'
'        '.ItemState(C, Sibody).sFieldName
'        Select Case LCase(sKey)
'
'            Case "cmemo"
'                clsRefer.SetReferDataType 17
'                clsRefer.Show
'
'                If Not clsRefer.recmx Is Nothing Then
'                    sRet = clsRefer.recmx.Fields("ctext")
'
'                End If
'                Exit Sub
'            Case "citemcode"
'                If Me.Voucher1.bodyText(row, "citem_class") = "" Then
'                    clsRefer.EnumRefInit m_login, enuTreeViewAndGrid, False, enuItem
'                Else
'                    clsRefer.ItemRefInit m_login, True, Me.Voucher1.bodyText(row, "citem_class")
'                End If
'                clsRefer.SetReferFilterString getReferString(sKey, sKeyValue)
'                clsRefer.Show
'                If Me.Voucher1.bodyText(row, "citem_class") = "" Then
'                    If Not clsRefer.RstSelClass.EOF Then
'                         sRet = clsRefer.recmx.Fields("citemcode").value
'                        .bodyText(.row, "citem_class") = clsRefer.RstSelClass("citem_class")
'                        .bodyText(.row, "citem_name") = clsRefer.RstSelClass("citem_name") 'sl add 项目大类名称
'                        .bodyText(.row, "citemname") = clsRefer.recmx.Fields("citemname").value
'                    End If
'                Else
'                  If Not clsRefer.recmx.EOF Then
'                   sRet = clsRefer.recmx.Fields("citemcode").value
'                   .bodyText(.row, "citemcode") = clsRefer.recmx.Fields("citemcode").value
'                   .bodyText(.row, "citemname") = clsRefer.recmx.Fields("citemname").value
'                   .bodyText(.row, "citem_class") = clsRefer.RstSelClass("citem_class")
'                   .bodyText(.row, "citem_name") = clsRefer.RstSelClass("citem_name")
'                   For I = 1 To clsRefer.RstSelCount - 1
'                     clsRefer.recmx.MoveNext
'                     .AddLine 1
'                     .bodyText(.BodyRows, "citemcode") = clsRefer.recmx.Fields("citemcode").value
'                     .bodyText(.BodyRows, "citemname") = clsRefer.recmx.Fields("citemname").value
'                     .bodyText(.BodyRows, "citem_class") = .bodyText(.row - 1, "citem_class")
'                     .bodyText(.BodyRows, "citem_name") = .bodyText(.row - 1, "citem_name")
'                   Next
'
'                  End If
'                End If
'
'            Case "citem_class"
'
'                strSql = "select citem_class ,citem_name from fitem"
'                clsRefer.StrRefInit m_login, False, "", strSql, "项目大类编码,项目大类名称", "", False, 1, 1, 1
'                clsRefer.Show
'                sRet = clsRefer.recmx.Fields("citem_class")
'                .bodyText(.row, "citem_class") = clsRefer.recmx.Fields("citem_class")
'                .bodyText(.row, "citem_name") = clsRefer.recmx.Fields("citem_name")
'        End Select
'
'    End With
'    referPara.Cancel = True
End Sub
 Private Sub Voucher2_bodyBrowUser(ByVal row As Long, ByVal col As Long, sRet As Variant, referPara As UapVoucherControl85.ReferParameter)
    
    Dim lngRow As Long
    Dim lngCol As Long
    lngRow = row
    lngCol = col
    referPara.Cancel = True
    strReferString = clsVoucherRefer.ShowReferCtl(clsVoucher, Voucher2, sibody, col, referPara, row)
    If referPara.Cancel Then
        sRet = Voucher2.bodyText(lngRow, lngCol)
    End If
'    Dim strSql As String, cInvCode As String, cCusCode As String
'    Dim Domhead As DOMDocument, Dombody As DOMDocument
'    Dim lngRow As Integer
'    Dim I As Integer, lRecord As Long
'    Dim tmpDOMBody As DOMDocument
'    Dim sKey As String, sKeyValue As String, strAuth As String
'    Dim tmprst As ADODB.Recordset
'    Dim tmpRow As Long, tmpCol As Long
'    Dim sFormat As String, strExchName As String
'    Dim adoField As ADODB.Field
'    Dim tmpbFromCurrentStock As Boolean
'    Dim tmpCol2 As Long
'    Dim bInvRefSuce As Boolean
'    Dim strDate As String
'    Dim bMulitSelect As Boolean
'    Dim strCusInv As String
'
'
'    tmpRow = row
'    tmpCol = Voucher2.col
'    tmpCol2 = col
'
'
'    On Error Resume Next
'
'    With Voucher2
'        .MultiLineSelect = False ''设置多选默认
'        clsRefer.SetReferSQLString ""
'        clsRefer.SetRWAuth "INVENTORY", "R", False
'        clsRefer.SetReferDisplayMode enuGrid
'        sKey = .ItemState(col, sibody).sFieldName
'        sKeyValue = .bodyText(row, col)
'
'        '.ItemState(C, Sibody).sFieldName
'        Select Case LCase(sKey)
'          Case "ccusabbname"
'                clsRefer.SetReferDataType enuCustomer
'                If myinfo.bAuth_Cus Then
'                    'strAuth = clsAuth.GetAuthString("CUSTOMER")
'                    'clsRefer.SetRWAuth strAuth
'                    clsRefer.SetRWAuth "CUSTOMER", "W", True
'                Else
'                    clsRefer.SetRWAuth "CUSTOMER", "R", False
'                End If
'
'                clsRefer.SetReferFilterString " isnull(dEndDate,'9999-12-31')>'" + strDate + "' " & IIf(getReferString(sKey, sKeyValue) <> "", " and " & getReferString(sKey, sKeyValue), "")
'
'                clsRefer.Show
'                If Not clsRefer.recmx Is Nothing Then
'                    Me.Voucher2.bodyText(row, "ccuscode") = clsRefer.recmx("ccuscode")  'sl 修改 保存子表的客户编码
'                    sRet = clsRefer.recmx.Fields("ccusabbname")
'                    Me.Voucher2.bodyText(row, "ccusaddress") = clsRefer.recmx("ccusaddress")
'                End If
'            Case "cmemo"
'                clsRefer.SetReferDataType 17
'                clsRefer.Show
'
'                If Not clsRefer.recmx Is Nothing Then
'                    sRet = clsRefer.recmx.Fields("ctext")
'
'                End If
'                Exit Sub
'            Case "citemcode"
'                If Me.Voucher2.bodyText(row, "citem_class") = "" Then
'                    clsRefer.EnumRefInit m_login, enuTreeViewAndGrid, False, enuItem
'                Else
'                    clsRefer.ItemRefInit m_login, False, Me.Voucher2.bodyText(row, "citem_class")
'                End If
'                clsRefer.SetReferFilterString getReferString(sKey, sKeyValue)
'                clsRefer.Show
'                If Me.Voucher2.bodyText(row, "citem_class") = "" Then
'                    If Not clsRefer.RstSelClass.EOF Then
'                        sRet = clsRefer.RstSelClass("citemcode")
'                        .bodyText(.row, "citem_class") = clsRefer.RstSelClass("citem_class")
'                        .bodyText(.row, "citem_cname") = clsRefer.RstSelClass("citem_name")
'                        .bodyText(.row, "citemname") = clsRefer.RstSelClass("citemname")
'                    End If
'                End If
'            Case "citem_class"
'
'                strSql = "select citem_class ,citem_name from fitem"
'                If clsRefer.StrRefInit(m_login, False, "", strSql, "项目大类编码,项目大类名称", "") = False Then Exit Sub
'                clsRefer.Show
'
'        End Select
'
'    End With
'        referPara.Cancel = True
End Sub

Private Sub Voucher3_bodyBrowUser(ByVal row As Long, ByVal col As Long, sRet As Variant, referPara As UapVoucherControl85.ReferParameter)
    Dim lngRow As Long
    Dim lngCol As Long
    lngRow = row
    lngCol = col
    referPara.Cancel = True
    strReferString = clsVoucherRefer.ShowReferCtl(clsVoucher, Voucher3, sibody, col, referPara, row)
    If referPara.Cancel Then
        sRet = Voucher3.bodyText(lngRow, lngCol)
    End If
'    Dim strSql As String, cInvCode As String, cCusCode As String
'    Dim Domhead As DOMDocument, Dombody As DOMDocument
'    Dim lngRow As Integer
'    Dim I As Integer, lRecord As Long
'    Dim tmpDOMBody As DOMDocument
'    Dim sKey As String, sKeyValue As String, strAuth As String
'    Dim tmprst As ADODB.Recordset
'    Dim tmpRow As Long, tmpCol As Long
'    Dim sFormat As String, strExchName As String
'    Dim adoField As ADODB.Field
'    Dim tmpbFromCurrentStock As Boolean
'    Dim tmpCol2 As Long
'    Dim bInvRefSuce As Boolean
'    Dim strDate As String
'    Dim bMulitSelect As Boolean
'    Dim strCusInv As String
'
'
'    tmpRow = row
'    tmpCol = Voucher3.col
'    tmpCol2 = col
'
'
'    On Error Resume Next
'
'    With Voucher3
'        .MultiLineSelect = False ''设置多选默认
'        clsRefer.SetReferSQLString ""
'        clsRefer.SetRWAuth "INVENTORY", "R", False
'        clsRefer.SetReferDisplayMode enuGrid
'        sKey = .ItemState(col, sibody).sFieldName
'        sKeyValue = .bodyText(row, col)
'
'        Select Case LCase(sKey)
'
'            Case "cmemo", "content"
'                clsRefer.SetReferDataType 17
'                clsRefer.Show
'
'                If Not clsRefer.recmx Is Nothing Then
'                    sRet = clsRefer.recmx.Fields("ctext")
'
'                End If
'            Case "citemcode"
'                If Me.Voucher3.bodyText(row, "citem_class") = "" Then
'                    clsRefer.EnumRefInit m_login, enuTreeViewAndGrid, False, enuItem
'                Else
'                    clsRefer.ItemRefInit m_login, False, Me.Voucher3.bodyText(row, "citem_class")
'                End If
'                clsRefer.SetReferFilterString getReferString(sKey, sKeyValue)
'                clsRefer.Show
'                If Me.Voucher3.bodyText(row, "citem_class") = "" Then
'                    If Not clsRefer.RstSelClass.EOF Then
'                        .bodyText(.row, "citem_class") = clsRefer.RstSelClass("citem_class")
'                        .bodyText(.row, "citem_cname") = clsRefer.RstSelClass("citem_name")
'                    End If
'                End If
'            Case "citem_class"
'
'                strSql = "select citem_class ,citem_name from fitem"
'                If clsRefer.StrRefInit(m_login, False, "", strSql, "项目大类编码,项目大类名称", "") = False Then Exit Sub
'
'                clsRefer.Show
'
'        End Select
'    End With
'        referPara.Cancel = True
End Sub
Private Sub Voucher4_bodyBrowUser(ByVal row As Long, ByVal col As Long, sRet As Variant, referPara As UapVoucherControl85.ReferParameter)
    Dim lngRow As Long
    Dim lngCol As Long
    lngRow = row
    lngCol = col
    referPara.Cancel = True
    strReferString = clsVoucherRefer.ShowReferCtl(clsVoucher, Voucher4, sibody, col, referPara, row)
    If referPara.Cancel Then
        sRet = Voucher4.bodyText(lngRow, lngCol)
    End If
'    Dim strSql As String, cInvCode As String, cCusCode As String
'    Dim Domhead As DOMDocument, Dombody As DOMDocument
'    Dim lngRow As Integer
'    Dim I As Integer, lRecord As Long
'    Dim tmpDOMBody As DOMDocument
'    Dim sKey As String, sKeyValue As String, strAuth As String
'    Dim tmprst As ADODB.Recordset
'    Dim tmpRow As Long, tmpCol As Long
'    Dim sFormat As String, strExchName As String
'    Dim adoField As ADODB.Field
'    Dim tmpbFromCurrentStock As Boolean
'    Dim tmpCol2 As Long
'    Dim bInvRefSuce As Boolean
'    Dim strDate As String
'    Dim bMulitSelect As Boolean
'    Dim strCusInv As String
'    Dim ppcom As Object
'
'    tmpRow = Row
'    tmpCol = Voucher4.Col
'    tmpCol2 = Col
'
'
'    On Error Resume Next
'
'    With Voucher4
'        .MultiLineSelect = False ''设置多选默认
'        clsRefer.SetReferSQLString ""
'        clsRefer.SetRWAuth "INVENTORY", "R", False
'        clsRefer.SetReferDisplayMode enuGrid
'        sKey = .ItemState(Col, sibody).sFieldName
'        sKeyValue = .bodyText(Row, Col)
'
'
'        Select Case LCase(sKey)
'           Case "iinvexchrate"
'                .bodyText(Row, "inum") = .bodyText(Row, "iquantity") / .bodyText(Row, "iinvexchrate")
'            Case "iinvexchratejf"   'sl 添加  付印通知单—纸张 加放率
'                .bodyText(Row, "iquantityjf") = .bodyText(Row, "iquantity") * .bodyText(Row, "iinvexchratejf")
'                .bodyText(Row, "inumjf") = .bodyText(Row, "iquantity") * .bodyText(Row, "iinvexchratejf") / .bodyText(Row, "iinvexchrate")
'                .bodyText(Row, "iquantityhj") = .bodyText(Row, "iquantity") + .bodyText(Row, "iquantity") * .bodyText(Row, "iinvexchratejf")
'                .bodyText(Row, "inumhj") = .bodyText(Row, "inum") + .bodyText(Row, "iquantity") * .bodyText(Row, "iinvexchratejf") / .bodyText(Row, "iinvexchrate")
'
'            Case "iquantityjf"
'                  If val(.bodyText(Row, "iinvexchratejf")) <> 0 Then
'                    .bodyText(Row, "iquantity") = .bodyText(Row, "iquantityjf") / .bodyText(Row, "iinvexchratejf")
'                     .bodyText(Row, "inum") = .bodyText(Row, "iquantityjf") / .bodyText(Row, "iinvexchratejf") / .bodyText(Row, "iinvexchrate")
'                    .bodyText(Row, "inumjf") = .bodyText(Row, "iquantityjf") / .bodyText(Row, "iinvexchratejf") * .bodyText(Row, "iinvexchratejf") / .bodyText(Row, "iinvexchrate")
'                    .bodyText(Row, "iquantityhj") = .bodyText(Row, "iquantity") + .bodyText(Row, "iquantity") * .bodyText(Row, "iinvexchratejf")
'                    .bodyText(Row, "inumhj") = .bodyText(Row, "inum") + .bodyText(Row, "iquantity") * .bodyText(Row, "iinvexchratejf") / .bodyText(Row, "iinvexchrate")
'                  End If
'
'            Case "inumjf"
'
'            Case "iquantityhj"
'
'            Case "inumhj"
'
'
'            Case "cinvname", "cinvcode"
'
'                strDate = Me.voucher.headerText("ddate")
'
'                If strDate <> "" Then
'                    strDate = Left(strDate, 10)
'                Else
'                    strDate = "1900-01-01"
'                End If
'
'                bMulitSelect = False
'                clsRefer.SetReferDisplayMode enuTreeViewAndGrid, bMulitSelect
'                If myinfo.bAuth_Inv Then
'                    clsRefer.SetRWAuth "INVENTORY", "W", True
'                Else
'                    clsRefer.SetRWAuth "INVENTORY", "R", False
'                End If
'
'                    bFromCurrentStock = False
'                    tmpbFromCurrentStock = False
'                If .bodyText(.Row, "sheetsouce") = "非带料" Then 'sl 制版单表体明细存货属性为应税劳务＝1
'                  strSql = "bpurchase=1 and bcomsume=1"  'sl add 存货属性 应税劳务＝1
'                Else
'                  strSql = "bpurchase=1 and bcomsume=1"
'                End If
'
'                If .bodyText(.Row, "cinvcode") = "" Then
'                Else
'                  strSql = strSql & "and cinvname like '%" & .bodyText(.Row, "cinvcode") & "%' "
'                End If
'                    If clsRefer.EnumRefInit(m_login, 2, bMulitSelect, "inventory", strSql) = False Then Exit Sub
'                clsRefer.Show
'                bInvRefSuce = False
'                If Not clsRefer.recmx Is Nothing Then
'                    If Not clsRefer.recmx.EOF Then
'                             bInvRefSuce = True
'                            .bodyText(Row, "cinvcode") = clsRefer.recmx("cinvcode")
'                            .bodyText(Row, "cinvname") = clsRefer.recmx("cinvname")
'                            .bodyText(Row, "cinvstd") = clsRefer.recmx("cinvstd")
'                        If bMulitSelect = True Then
'                            bFromCurrentStock = tmpbFromCurrentStock
'                            strRefFldName = ""
'                            Call Voucher4_bodyCellCheck("", 1, tmpRow, tmpCol2, referPara)
'                            bFromCurrentStock = tmpbFromCurrentStock
'                            For lRecord = 1 To clsRefer.RstSelCount - 1
'                                bFromCurrentStock = tmpbFromCurrentStock
'                                clsRefer.recmx.MoveNext
'                                .AddLine 1
'                                If bFromCurrentStock Then
'                                    For Each adoField In clsRefer.recmx.Fields
'                                        .bodyText(.BodyRows, adoField.Name) = adoField.value
'                                    Next adoField
'                                Else
'                                    .bodyText(.BodyRows, "cinvname") = clsRefer.recmx("cinvname")
'                                    .bodyText(.BodyRows, "cinvcode") = clsRefer.recmx("cinvcode")
'                                    .bodyText(Row, "cinvstd") = clsRefer.recmx("cinvstd")
'                                End If
'                                bFromCurrentStock = tmpbFromCurrentStock
'                                Call Voucher4_bodyCellCheck("", 1, CLng(.BodyRows), Col, referPara)
'
'                            Next lRecord
'                            bFromCurrentStock = False
'                            clsRefer.recmx.MoveFirst
'                            strRefFldName = sKey '"cinvname"
'                            BrowFlag = True
'                            Call .BodyProtectUnload
'                            BrowFlag = False
'                        End If
'                    Else
'                        bFromCurrentStock = False
'                        tmpbFromCurrentStock = False
'                    End If
'                Else
'                    bFromCurrentStock = False
'                    tmpbFromCurrentStock = False
'                End If
'
'            Case "cmemo"
'                clsRefer.SetReferDataType 17
'                clsRefer.Show
'
'                If Not clsRefer.recmx Is Nothing Then
'                    sRet = clsRefer.recmx.Fields("ctext")
'
'                End If
'                Exit Sub
'
'            Case "citemcode"
'                If Me.Voucher4.bodyText(Row, "citem_class") = "" Then
'                    clsRefer.EnumRefInit m_login, enuTreeViewAndGrid, False, enuItem
'
'                Else
'                    clsRefer.ItemRefInit m_login, False, Me.Voucher4.bodyText(Row, "citem_class")
'                End If
'
'                clsRefer.SetReferFilterString getReferString(sKey, sKeyValue)
'                clsRefer.Show
'                If Me.Voucher4.bodyText(Row, "citem_class") = "" Then
'                    If Not clsRefer.RstSelClass.EOF Then
'                        .bodyText(.Row, "citem_class") = clsRefer.RstSelClass("citem_class")
'                        .bodyText(.Row, "citem_cname") = clsRefer.RstSelClass("citem_name")
'                    End If
'                End If
'            Case "citem_class"
'
'                strSql = "select citem_class ,citem_name from fitem"
'                If clsRefer.StrRefInit(m_login, False, "", strSql, "项目大类编码,项目大类名称", "") = False Then Exit Sub
'
'                clsRefer.Show
'
'            Case "cinva_unit"
'                If val(.bodyText(Row, "igrouptype")) <> 1 Then
'                    Exit Sub
'                End If
'
'                If clsRefer.EnumRefInit(m_login, 1, False, "ComputationUnit", "cGroupCode='" & .bodyText(Row, "cGroupCode") & "'") = False Then Exit Sub
'                clsRefer.Show
'                If Not clsRefer.recmx Is Nothing And Not clsRefer.recmx.EOF Then
'                      sRet = clsRefer.recmx.Fields!cComUnitName
'                End If
'                Exit Sub
'        End Select
'
'        ''设置存货多选
'        If (LCase(sKey) = "cinvname" Or LCase(sKey) = "cinvcode") And bMulitSelect = True And bInvRefSuce = True Then
'            .MultiLineSelect = True
'            strRefFldName = ""
'            If bMulitSelect = True Then
'                .Row = tmpRow
'                .Col = tmpCol
'            Else
'                Call Voucher4_bodyCellCheck("", 1, tmpRow, tmpCol2, referPara)
'                .Row = tmpRow
'                .Col = tmpCol
'            End If
'        End If
'
'    End With
'        referPara.Cancel = True
End Sub






Private Sub Voucher_bodyCellCheck(RetValue As Variant, bChanged As Long, ByVal r As Long, ByVal c As Long, referPara As ReferParameter)
    Dim intNumPoint As Integer
    Dim strFieldName As String
    Dim lngOldRow As Long
    Dim lngOldRows As Long
    Dim i As Long
    
    referPara.bValid = True
    strFieldName = Voucher.ItemState(c, sibody).sFieldName
    lngOldRow = r
    lngOldRows = Voucher.BodyRows
    If Not referPara.rstGrid Is Nothing Then
        strReferString = clsVoucherRefer.FillItemsAfterBrowse(clsVoucher, Voucher, "B", strFieldName, referPara.rstGrid, r)
        If Voucher.BodyRows > lngOldRows Then
            For i = lngOldRows + 1 To Voucher.BodyRows
                clsVoucher.CellCheck Voucher, "B", Voucher.ItemState(c, sibody).sFieldName, bChanged, clsVoucherRefer, i
            Next
        End If
        r = lngOldRow
    Else
        strFieldName = clsVoucherRefer.CellCheck("", Voucher, "B", strFieldName, r)
        If strFieldName <> "" Then
            If strFieldName = "cancel" Then
                bChanged = Cancel
            Else
                RetValue = ""
'                bChanged = retry
            End If
            Exit Sub
        End If
    End If
'    clsVoucherRefer.CellCheck strReferString, Voucher, "B", Voucher.ItemState(C, sibody).sFieldName, R
    clsVoucher.CellCheck Voucher, "B", Voucher.ItemState(c, sibody).sFieldName, bChanged, clsVoucherRefer, r
    RetValue = Voucher.bodyText(lngOldRow, c)
    If Voucher.ItemState(c, sibody).nFieldType = 4 Then
        intNumPoint = Voucher.ItemState(c, sibody).nNumPoint
        RetValue = Format(RetValue, IIf(intNumPoint = 0, "###0", "###0." & String(intNumPoint, "0")))
    End If
    If Voucher.ItemState(c, sibody).nReferType = 3 Then
        RetValue = Format(RetValue, "YYYY-MM-DD")
    End If
'    Dim lngRow As Long
'    Dim strError As String
'    Dim tmpDomhead As DOMDocument, tmpDOMBody As DOMDocument
'    Dim sKey As String
'    Dim ele As IXMLDOMElement
'    Dim sKeyValue As String
'
'
'    On Error GoTo DoERR
'    ''是否放弃
'
'    With Me.voucher
'        sKey = LCase(.ItemState(c, sibody).sFieldName)
'        '存货参照是否来源于现存量
'        If bFromCurrentStock = True Then
'            If sKey = "cinvname" Or sKey = "cinvcode" Then sKey = "currentstock"
'            bFromCurrentStock = False
'        End If
'
'
'        If Left(sKey, 7) = "cdefine" Or Left(sKey, 5) = "cfree" Then
'        ''使用新的自定义项目
'            If .bodyText(r, sKey) <> RetValue Then
'                sKeyValue = RetValue
'            Else
'                sKeyValue = .bodyText(r, sKey)
'            End If
'            RetValue = CellCheckDefine(.LookUpArray(sKey, sibody), sibody, sKeyValue)
'            If Left(sKey, 5) = "cfree" Then
'                If RetValue <> "" Then
'                    .bodyText(r, sKey) = RetValue
'                Else
'                    Exit Sub
'                End If
'            ElseIf Left(sKey, 7) = "cdefine" Then
'                Exit Sub
'            End If
'
'        End If
'
'            Set tmpDomhead = .GetHeadDom
'            Set tmpDOMBody = .GetLineDom(r)
'            Set Domhead = tmpDomhead
'            Set Dombody = tmpDOMBody
'
'        If BrowFlag = True And sKey = LCase(strRefFldName) Then
'            BrowFlag = False
'            strRefFldName = ""
'            Exit Sub
'        End If
'
'
'        strError = clsVoucherCO.BodyCheck(sKey, Dombody, Domhead, r)
'        If strError <> "" Then
'            MsgBox strError, vbInformation, "印制管理"
'            sKey = LCase(sKey)
'            If sKey = "cinvname" Or sKey = "cinvcode" Then
'                bChanged = success 'Cancel  'retry
'                .bodyText(r, "cinvcode") = ""
'                .bodyText(r, "cinvname") = ""
'                RetValue = ""
'            'zzg-860-Develop-Bug19615
'            ElseIf sKey = "cinva_unit" Then
'                bChanged = Cancel
'                RetValue = ""
'            Else
'                bChanged = IIf(bLostFocus, Cancel, retry) 'Cancel
'            End If
'            sKey = LCase(.ItemState(c, sibody).sFieldName)
'
'            If sKey = "cinvname" Or sKey = "cinvcode" Then
'                If Not Dombody Is Nothing Then
'                    If Dombody.selectNodes("//R").length > 0 Then
'                            For Each ele In Dombody.selectNodes("//R")
'                                .bodyText(r, ele.getAttribute("K")) = ele.getAttribute("V")
'                                If LCase(ele.getAttribute("K")) = sKey Then
'                                    RetValue = ""
'                                End If
'                            Next
'                    End If
'                End If
'            End If
'            Exit Sub
'        End If
'        sKey = LCase(.ItemState(c, sibody).sFieldName)
'        If Not Dombody Is Nothing Then
'            If Dombody.selectNodes("//R").length > 0 Then
'                For Each ele In Dombody.selectNodes("//R")
'                    .bodyText(r, ele.getAttribute("K")) = ele.getAttribute("V")
'                    If LCase(ele.getAttribute("K")) = sKey Then
'                        RetValue = ele.getAttribute("V")
'                    End If
'                    If sKey = "cinvname" Or sKey = "cinvcode" Then
'                        If strVouchType = "05" Or strVouchType = "06" Or strVouchType = "26" Or strVouchType = "27" Or strVouchType = "28" Or strVouchType = "29" Then
'                            If (LCase(ele.getAttribute("K")) = "binvtype" Or LCase(ele.getAttribute("K")) = "bservice") And ele.getAttribute("V") <> "" Then
'                                If CBool(ele.getAttribute("V")) = True Then
'                                    .bodyText(r, "cwhname") = "": .bodyText(r, "cwhcode") = ""
'                                End If
'                            End If
'                        End If
'                    End If
'                Next
'            End If
'        End If
'
'            Dim strInvCode As String
'            Dim dblQuan As Double
'            Dim dblNum As Double
'            strInvCode = .bodyText(r, "cinvcode")
'            If strInvCode <> "" Or sKey = "cinvcode" Then
'                If LCase(sKey) = "cinvname" Or LCase(sKey) = "cwhname" Or LCase(sKey) = "cbatch" Or LCase(sKey) = "cfree1" _
'                    Or LCase(sKey) = "cfree2" Or LCase(sKey) = "cfree3" Or LCase(sKey) = "cfree4" Or LCase(sKey) = "cfree5" _
'                    Or LCase(sKey) = "cfree6" Or LCase(sKey) = "cfree7" Or LCase(sKey) = "cfree8" Or LCase(sKey) = "cfree9" Or _
'                    LCase(sKey) = "cfree10" Then
'                    If strInvCode = "" Then strInvCode = RetValue
'                    If clsSAWeb_M.GetSumQuantity(DBConn, strInvCode, dblQuan, dblNum, .bodyText(r, "cfree1"), _
'                        .bodyText(r, "cfree2"), .bodyText(r, "cfree3"), .bodyText(r, "cfree4"), .bodyText(r, "cfree5"), _
'                        .bodyText(r, "cfree6"), .bodyText(r, "cfree7"), .bodyText(r, "cfree8"), .bodyText(r, "cfree9"), _
'                        .bodyText(r, "cfree10"), .bodyText(r, "cbatch"), .bodyText(r, "cwhcode")) Then
'                        .headerText("fstockquan") = dblQuan
'                        .headerText("fcanusequan") = dblNum
'                    Else
'                        MsgBox "取可用量失败"
'                    End If
'                End If
'            End If
'    If .headerText("itaxrate") = "" And iVouchState <> 2 Then
'       .headerText("itaxrate") = .bodyText(r, "itaxrate")
'    End If
'    End With
'    Exit Sub
'DoERR:
'    MsgBox Err.Description
End Sub
 
  Private Sub Voucher1_bodyCellCheck(RetValue As Variant, bChanged As Long, ByVal r As Long, ByVal c As Long, referPara As ReferParameter)
    Dim intNumPoint As Integer
    Dim strFieldName As String
    Dim lngOldRow As Long
    Dim lngOldRows As Long
    Dim i As Long
    
    referPara.bValid = True
    strFieldName = Voucher1.ItemState(c, sibody).sFieldName
    lngOldRow = r
    lngOldRows = Voucher1.BodyRows
    If Not referPara.rstGrid Is Nothing Then
        strReferString = clsVoucherRefer.FillItemsAfterBrowse(clsVoucher, Voucher1, "B", strFieldName, referPara.rstGrid, r)
        If Voucher1.BodyRows > lngOldRows Then
            For i = lngOldRows + 1 To Voucher1.BodyRows
                clsVoucher.CellCheck Voucher1, "B", Voucher1.ItemState(c, sibody).sFieldName, bChanged, clsVoucherRefer, i
            Next
        End If
        r = lngOldRow
    Else
        strFieldName = clsVoucherRefer.CellCheck("", Voucher1, "B", strFieldName, r)
        If strFieldName <> "" Then
            If strFieldName = "cancel" Then
                bChanged = Cancel
            Else
                RetValue = ""
'                bChanged = retry
            End If
            Exit Sub
        End If
    End If
'    clsVoucherRefer.CellCheck strReferString, Voucher, "B", Voucher.ItemState(C, sibody).sFieldName, R
    clsVoucher.CellCheck Voucher1, "B", Voucher1.ItemState(c, sibody).sFieldName, bChanged, clsVoucherRefer, r
    RetValue = Voucher1.bodyText(lngOldRow, c)
    If Voucher1.ItemState(c, sibody).nFieldType = 4 Then
        intNumPoint = Voucher1.ItemState(c, sibody).nNumPoint
        RetValue = Format(RetValue, IIf(intNumPoint = 0, "###0", "###0." & String(intNumPoint, "0")))
    End If
    If Voucher1.ItemState(c, sibody).nReferType = 3 Then
        RetValue = Format(RetValue, "YYYY-MM-DD")
    End If
'    Dim lngRow As Long
'    Dim strError As String
'    Dim tmpDomhead As DOMDocument, tmpDOMBody As DOMDocument
'    Dim sKey As String
'    Dim ele As IXMLDOMElement
'    Dim sKeyValue As String
'
'    'Dim DOMHead As New DOMDocument
'   ' Dim DomBody As New DOMDocument
'
'   ' On Error Resume Next
'    On Error GoTo DoERR
'    ''是否放弃
'    'If bClickCancel Then Exit Sub
'
'    With Me.Voucher1
'        sKey = LCase(.ItemState(c, sibody).sFieldName)
'
'        If Left(sKey, 7) = "cdefine" Or Left(sKey, 5) = "cfree" Then
'        ''使用新的自定义项目
'            'sRet = RefDefine((.LookUpArray(sKey, Sibody)), Sibody)
'            'RetValue = CellCheckDefine(.LookUpArray(sKey, sibody), sibody, .bodyText(R, sKey))
'            If .bodyText(r, sKey) <> RetValue Then
'                sKeyValue = RetValue
'            Else
'                sKeyValue = .bodyText(r, sKey)
'            End If
'            RetValue = CellCheckDefine(.LookUpArray(sKey, sibody), sibody, sKeyValue)
'            If Left(sKey, 5) = "cfree" Then
'                If RetValue <> "" Then
'                    .bodyText(r, sKey) = RetValue
'                Else
'                    Exit Sub
'                End If
'            ElseIf Left(sKey, 7) = "cdefine" Then
'                Exit Sub
'            End If
'
'        End If
'
'
'        Set tmpDomhead = .GetHeadDom
'        Set tmpDOMBody = .GetLineDom(r)
'        Set Domhead1 = tmpDomhead
'        Set Dombody1 = tmpDOMBody
'        If BrowFlag = True And sKey = LCase(strRefFldName) Then
'            BrowFlag = False
'            strRefFldName = ""
'            Exit Sub
'        End If
'
'      strError = clsVoucherCO.BodyCheck(sKey, Dombody1, Domhead1, r)
'        If strError <> "" Then
'            MsgBox strError, vbInformation, "印制管理"
'            sKey = LCase(sKey)
'            If sKey = "cinvname" Or sKey = "cinvcode" Then
'                bChanged = success 'Cancel  'retry
'                .bodyText(r, "cinvcode") = ""
'                .bodyText(r, "cinvname") = ""
'                RetValue = ""
'            'zzg-860-Develop-Bug19615
'            ElseIf sKey = "cinva_unit" Then
'                bChanged = Cancel
'                RetValue = ""
'            Else
'                bChanged = IIf(bLostFocus, Cancel, retry) 'Cancel
'            End If
'            sKey = LCase(.ItemState(c, sibody).sFieldName)
'
'            Exit Sub
'        End If
'        sKey = LCase(.ItemState(c, sibody).sFieldName)
'        If Not Dombody1 Is Nothing Then
'            If Dombody1.selectNodes("//R").length > 0 Then
'                For Each ele In Dombody1.selectNodes("//R")
'                    .bodyText(r, ele.getAttribute("K")) = ele.getAttribute("V")
'                    If LCase(ele.getAttribute("K")) = sKey Then
'                        RetValue = ele.getAttribute("V")
'                    End If
'                Next
'            End If
'        End If
'
'    End With
'    Exit Sub
'DoERR:
'    MsgBox Err.Description

End Sub
 
  Private Sub Voucher2_bodyCellCheck(RetValue As Variant, bChanged As Long, ByVal r As Long, ByVal c As Long, referPara As ReferParameter)
    Dim intNumPoint As Integer
    Dim strFieldName As String
    Dim lngOldRow As Long
    Dim lngOldRows As Long
    Dim i As Long
    
    referPara.bValid = True
    strFieldName = Voucher2.ItemState(c, sibody).sFieldName
    lngOldRow = r
    lngOldRows = Voucher2.BodyRows
    If Not referPara.rstGrid Is Nothing Then
        strReferString = clsVoucherRefer.FillItemsAfterBrowse(clsVoucher, Voucher2, "B", strFieldName, referPara.rstGrid, r)
        If Voucher2.BodyRows > lngOldRows Then
            For i = lngOldRows + 1 To Voucher2.BodyRows
                clsVoucher.CellCheck Voucher2, "B", Voucher2.ItemState(c, sibody).sFieldName, bChanged, clsVoucherRefer, i
            Next
        End If
        r = lngOldRow
    Else
        strFieldName = clsVoucherRefer.CellCheck("", Voucher2, "B", strFieldName, r)
        If strFieldName <> "" Then
            If strFieldName = "cancel" Then
                bChanged = Cancel
            Else
                RetValue = ""
'                bChanged = retry
            End If
            Exit Sub
        End If
    End If
'    clsVoucherRefer.CellCheck strReferString, Voucher, "B", Voucher.ItemState(C, sibody).sFieldName, R
    clsVoucher.CellCheck Voucher2, "B", Voucher2.ItemState(c, sibody).sFieldName, bChanged, clsVoucherRefer, r
    RetValue = Voucher2.bodyText(lngOldRow, c)
    If Voucher2.ItemState(c, sibody).nFieldType = 4 Then
        intNumPoint = Voucher2.ItemState(c, sibody).nNumPoint
        RetValue = Format(RetValue, IIf(intNumPoint = 0, "###0", "###0." & String(intNumPoint, "0")))
    End If
    If Voucher2.ItemState(c, sibody).nReferType = 3 Then
        RetValue = Format(RetValue, "YYYY-MM-DD")
    End If
    
    'RetValue = voucher.bodyText(R, C)
    strFieldName = Voucher2.ItemState(c, sibody).sFieldName
    Select Case LCase(strFieldName)
        Case "iquantity" ''送书册数
            If val(RetValue) < 0 Then
                MsgBox "送书册数不能小于0", vbInformation
                bChanged = Cancel: Exit Sub
            End If
    End Select
    
'    Dim lngRow As Long
'    Dim strError As String
'    Dim tmpDomhead As DOMDocument, tmpDOMBody As DOMDocument
'    Dim sKey As String
'    Dim ele As IXMLDOMElement
'    Dim sKeyValue As String
'
'    'Dim DOMHead As New DOMDocument
'   ' Dim DomBody As New DOMDocument
'
'   ' On Error Resume Next
'    On Error GoTo DoERR
'    ''是否放弃
'    'If bClickCancel Then Exit Sub
'
'    With Me.Voucher2
'        sKey = LCase(.ItemState(c, sibody).sFieldName)
'        '存货参照是否来源于现存量
'        If bFromCurrentStock = True Then
'            If sKey = "cinvname" Or sKey = "cinvcode" Then sKey = "currentstock"
'            bFromCurrentStock = False
'        End If
'        ''添客户默认仓库、订单预发货日期
'        If sKey = "cinvname" Or sKey = "cinvcode" Or sKey = "ccusinvcode" Or sKey = "ccusinvname" Then
'            If strVouchType = "05" Or strVouchType = "06" Or strVouchType = "26" Or strVouchType = "27" _
'                Or strVouchType = "28" Or strVouchType = "29" Then
'                If .bodyText(r, "cwhname") = "" And .headerText("ccuscode") <> "" Then
'                End If
'            End If
'            If strVouchType = "97" Then
'                '从表头带入预完工日期，预发货日期  edit by jzq 2004.09.02
'                .bodyText(r, "dpredate") = .headerText("dpredatebt")
'                .bodyText(r, "dpremodate") = .headerText("dpremodatebt")
''                If .bodyText(R, "dpremodate") = "" Then
''                    .bodyText(R, "dpremodate") = IIf(.bodyText(R, "dpredate") <> "", .bodyText(R, "dpredate"), .headerText("ddate"))
''                End If
'                '--------------end edit by jzq
'            End If
'        End If
'        '合同处理
'        If sKey = "cinvcode" And bRefContract Then sKey = "cinvname2"
'
'        If Left(sKey, 7) = "cdefine" Or Left(sKey, 5) = "cfree" Then
'        ''使用新的自定义项目
'            'sRet = RefDefine((.LookUpArray(sKey, Sibody)), Sibody)
'            'RetValue = CellCheckDefine(.LookUpArray(sKey, sibody), sibody, .bodyText(R, sKey))
'            If .bodyText(r, sKey) <> RetValue Then
'                sKeyValue = RetValue
'            Else
'                sKeyValue = .bodyText(r, sKey)
'            End If
'            RetValue = CellCheckDefine(.LookUpArray(sKey, sibody), sibody, sKeyValue)
'            If Left(sKey, 5) = "cfree" Then
'                If RetValue <> "" Then
'                    .bodyText(r, sKey) = RetValue
'                Else
'                    Exit Sub
'                End If
'            ElseIf Left(sKey, 7) = "cdefine" Then
'                Exit Sub
'            End If
'
'        End If
'        'If LCase(sKey) <> "itax" Then
'            Set tmpDomhead = .GetHeadDom
'            Set tmpDOMBody = .GetLineDom(r)
'            Set Domhead = tmpDomhead
'            Set Dombody = tmpDOMBody
'        'Else
'        '    .getVoucherDataXML domHead, DomBody
'        'End If
'        If BrowFlag = True And sKey = LCase(strRefFldName) Then
'            BrowFlag = False
'            strRefFldName = ""
'            Exit Sub
'        End If
'
'        If sKey = "ccode" Then
'            If strVouchType = "05" Or strVouchType = "06" Then
'                If val(.bodyText(r, "icorid")) <> 0 And val(.bodyText(r, "iquantity")) <= 0 Then
'                    bChanged = Cancel
'                    MsgBox "数量小于等于零时，入库单号不可输入或者修改!"
'                    Exit Sub
'                End If
'            Else
'                If val(.bodyText(r, "iquantity")) <= 0 Then
'                   MsgBox "数量小于等于零时，入库单号不可输入或者修改!"
'                   RetValue = ""
'                   .bodyText(r, "ibatch") = ""
'                   Exit Sub
'                End If
'            End If
'        End If
'        ''期初单据不校验
'        If sKey = "cbatch" Or sKey = "ccode" Then
'            If bFirst Then Exit Sub
'        End If
'
'        strError = clsVoucherCO.BodyCheck(sKey, Dombody, Domhead, r)
'        If strError <> "" Then
'            MsgBox strError, vbInformation, "销售管理"
'            sKey = LCase(sKey)
'            If sKey = "cinvname" Or sKey = "cinvcode" Then
'                bChanged = success 'Cancel  'retry
'                .bodyText(r, "cinvcode") = ""
'                .bodyText(r, "cinvname") = ""
'                RetValue = ""
'            'zzg-860-Develop-Bug19615
'            ElseIf sKey = "cinva_unit" Then
'                bChanged = Cancel
'                RetValue = ""
'            Else
'                bChanged = IIf(bLostFocus, Cancel, retry) 'Cancel
'            End If
'            sKey = LCase(.ItemState(c, sibody).sFieldName)
'
'            If sKey = "cinvname" Or sKey = "cinvcode" Then
'                If Not Dombody Is Nothing Then
'                    If Dombody.selectNodes("//R").length > 0 Then
'                            For Each ele In Dombody.selectNodes("//R")
'                                .bodyText(r, ele.getAttribute("K")) = ele.getAttribute("V")
'                                If LCase(ele.getAttribute("K")) = sKey Then
'                                    RetValue = ""
'                                End If
'                            Next
'                    End If
'                End If
'            End If
'            Exit Sub
'        End If
'        sKey = LCase(.ItemState(c, sibody).sFieldName)
'        If Not Dombody Is Nothing Then
'            If Dombody.selectNodes("//R").length > 0 Then
'                For Each ele In Dombody.selectNodes("//R")
'                    .bodyText(r, ele.getAttribute("K")) = ele.getAttribute("V")
'                    If LCase(ele.getAttribute("K")) = sKey Then
'                        RetValue = ele.getAttribute("V")
'                    End If
'                    If sKey = "cinvname" Or sKey = "cinvcode" Then
'                        If strVouchType = "05" Or strVouchType = "06" Or strVouchType = "26" Or strVouchType = "27" Or strVouchType = "28" Or strVouchType = "29" Then
'                            If (LCase(ele.getAttribute("K")) = "binvtype" Or LCase(ele.getAttribute("K")) = "bservice") And ele.getAttribute("V") <> "" Then
'                                If CBool(ele.getAttribute("V")) = True Then
'                                    .bodyText(r, "cwhname") = "": .bodyText(r, "cwhcode") = ""
'                                End If
'                            End If
'                        End If
'                    End If
'                Next
'            End If
'        End If
'
'            Dim strInvCode As String
'            Dim dblQuan As Double
'            Dim dblNum As Double
'            strInvCode = .bodyText(r, "cinvcode")
'            If strInvCode <> "" Or sKey = "cinvcode" Then
'                If LCase(sKey) = "cinvname" Or LCase(sKey) = "cwhname" Or LCase(sKey) = "cbatch" Or LCase(sKey) = "cfree1" _
'                    Or LCase(sKey) = "cfree2" Or LCase(sKey) = "cfree3" Or LCase(sKey) = "cfree4" Or LCase(sKey) = "cfree5" _
'                    Or LCase(sKey) = "cfree6" Or LCase(sKey) = "cfree7" Or LCase(sKey) = "cfree8" Or LCase(sKey) = "cfree9" Or _
'                    LCase(sKey) = "cfree10" Then
'                    If strInvCode = "" Then strInvCode = RetValue
'                    If clsSAWeb_M.GetSumQuantity(DBConn, strInvCode, dblQuan, dblNum, .bodyText(r, "cfree1"), _
'                        .bodyText(r, "cfree2"), .bodyText(r, "cfree3"), .bodyText(r, "cfree4"), .bodyText(r, "cfree5"), _
'                        .bodyText(r, "cfree6"), .bodyText(r, "cfree7"), .bodyText(r, "cfree8"), .bodyText(r, "cfree9"), _
'                        .bodyText(r, "cfree10"), .bodyText(r, "cbatch"), .bodyText(r, "cwhcode")) Then
'                        .headerText("fstockquan") = dblQuan
'                        .headerText("fcanusequan") = dblNum
'                    Else
'                        MsgBox "取可用量失败"
'                    End If
'                End If
'            End If
'    If .headerText("itaxrate") = "" And iVouchState <> 2 Then
'       .headerText("itaxrate") = .bodyText(r, "itaxrate")
'    End If
'    End With
'    Exit Sub
'DoERR:
'    MsgBox Err.Description
End Sub
 
  Private Sub Voucher3_bodyCellCheck(RetValue As Variant, bChanged As Long, ByVal r As Long, ByVal c As Long, referPara As ReferParameter)
    Dim intNumPoint As Integer
    Dim strFieldName As String
    Dim lngOldRow As Long
    Dim lngOldRows As Long
    Dim i As Long
    
    referPara.bValid = True
    strFieldName = Voucher3.ItemState(c, sibody).sFieldName
    lngOldRow = r
    lngOldRows = Voucher3.BodyRows
    If Not referPara.rstGrid Is Nothing Then
        strReferString = clsVoucherRefer.FillItemsAfterBrowse(clsVoucher, Voucher3, "B", strFieldName, referPara.rstGrid, r)
        If Voucher3.BodyRows > lngOldRows Then
            For i = lngOldRows + 1 To Voucher3.BodyRows
                clsVoucher.CellCheck Voucher3, "B", Voucher3.ItemState(c, sibody).sFieldName, bChanged, clsVoucherRefer, i
            Next
        End If
        r = lngOldRow
    Else
        strFieldName = clsVoucherRefer.CellCheck("", Voucher3, "B", strFieldName, r)
        If strFieldName <> "" Then
            If strFieldName = "cancel" Then
                bChanged = Cancel
            Else
                RetValue = ""
'                bChanged = retry
            End If
            Exit Sub
        End If
    End If
'    clsVoucherRefer.CellCheck strReferString, Voucher, "B", Voucher.ItemState(C, sibody).sFieldName, R
    clsVoucher.CellCheck Voucher3, "B", Voucher3.ItemState(c, sibody).sFieldName, bChanged, clsVoucherRefer, r
    RetValue = Voucher3.bodyText(lngOldRow, c)
    If Voucher3.ItemState(c, sibody).nFieldType = 4 Then
        intNumPoint = Voucher3.ItemState(c, sibody).nNumPoint
        RetValue = Format(RetValue, IIf(intNumPoint = 0, "###0", "###0." & String(intNumPoint, "0")))
    End If
    If Voucher3.ItemState(c, sibody).nReferType = 3 Then
        RetValue = Format(RetValue, "YYYY-MM-DD")
    End If
    
    'RetValue = voucher.bodyText(R, C)
    strFieldName = Voucher3.ItemState(c, sibody).sFieldName
    Select Case LCase(strFieldName)
        Case "mianquantity" ''面数
            If val(RetValue) < 0 Then
                MsgBox "面数不能小于0", vbInformation
                bChanged = Cancel: Exit Sub
            End If
    End Select
'    Dim lngRow As Long
'    Dim strError As String
'    Dim tmpDomhead As DOMDocument, tmpDOMBody As DOMDocument
'    Dim sKey As String
'    Dim ele As IXMLDOMElement
'    Dim sKeyValue As String
'
'    'Dim DOMHead As New DOMDocument
'   ' Dim DomBody As New DOMDocument
'
'   ' On Error Resume Next
'    On Error GoTo DoERR
'    ''是否放弃
'    'If bClickCancel Then Exit Sub
'
'    With Me.Voucher3
'        sKey = LCase(.ItemState(c, sibody).sFieldName)
'        '存货参照是否来源于现存量
'        If bFromCurrentStock = True Then
'            If sKey = "cinvname" Or sKey = "cinvcode" Then sKey = "currentstock"
'            bFromCurrentStock = False
'        End If
'        ''添客户默认仓库、订单预发货日期
'        If sKey = "cinvname" Or sKey = "cinvcode" Or sKey = "ccusinvcode" Or sKey = "ccusinvname" Then
'            If strVouchType = "05" Or strVouchType = "06" Or strVouchType = "26" Or strVouchType = "27" _
'                Or strVouchType = "28" Or strVouchType = "29" Then
'                If .bodyText(r, "cwhname") = "" And .headerText("ccuscode") <> "" Then
'
'                End If
'            End If
'            If strVouchType = "97" Then
'                '从表头带入预完工日期，预发货日期  edit by jzq 2004.09.02
'                .bodyText(r, "dpredate") = .headerText("dpredatebt")
'                .bodyText(r, "dpremodate") = .headerText("dpremodatebt")
''                If .bodyText(R, "dpremodate") = "" Then
''                    .bodyText(R, "dpremodate") = IIf(.bodyText(R, "dpredate") <> "", .bodyText(R, "dpredate"), .headerText("ddate"))
''                End If
'                '--------------end edit by jzq
'            End If
'        End If
'        '合同处理
'        If sKey = "cinvcode" And bRefContract Then sKey = "cinvname2"
'
'        If Left(sKey, 7) = "cdefine" Or Left(sKey, 5) = "cfree" Then
'        ''使用新的自定义项目
'            'sRet = RefDefine((.LookUpArray(sKey, Sibody)), Sibody)
'            'RetValue = CellCheckDefine(.LookUpArray(sKey, sibody), sibody, .bodyText(R, sKey))
'            If .bodyText(r, sKey) <> RetValue Then
'                sKeyValue = RetValue
'            Else
'                sKeyValue = .bodyText(r, sKey)
'            End If
'            RetValue = CellCheckDefine(.LookUpArray(sKey, sibody), sibody, sKeyValue)
'            If Left(sKey, 5) = "cfree" Then
'                If RetValue <> "" Then
'                    .bodyText(r, sKey) = RetValue
'                Else
'                    Exit Sub
'                End If
'            ElseIf Left(sKey, 7) = "cdefine" Then
'                Exit Sub
'            End If
'
'        End If
'        'If LCase(sKey) <> "itax" Then
'            Set tmpDomhead = .GetHeadDom
'            Set tmpDOMBody = .GetLineDom(r)
'            Set Domhead = tmpDomhead
'            Set Dombody = tmpDOMBody
'        'Else
'        '    .getVoucherDataXML domHead, DomBody
'        'End If
'        If BrowFlag = True And sKey = LCase(strRefFldName) Then
'            BrowFlag = False
'            strRefFldName = ""
'            Exit Sub
'        End If
'
'        If sKey = "ccode" Then
'            If strVouchType = "05" Or strVouchType = "06" Then
'                If val(.bodyText(r, "icorid")) <> 0 And val(.bodyText(r, "iquantity")) <= 0 Then
'                    bChanged = Cancel
'                    MsgBox "数量小于等于零时，入库单号不可输入或者修改!"
'                    Exit Sub
'                End If
'            Else
'                If val(.bodyText(r, "iquantity")) <= 0 Then
'                   MsgBox "数量小于等于零时，入库单号不可输入或者修改!"
'                   RetValue = ""
'                   .bodyText(r, "ibatch") = ""
'                   Exit Sub
'                End If
'            End If
'        End If
'        ''期初单据不校验
'        If sKey = "cbatch" Or sKey = "ccode" Then
'            If bFirst Then Exit Sub
'        End If
'
'        strError = clsVoucherCO.BodyCheck(sKey, Dombody, Domhead, r)
'        If strError <> "" Then
'            MsgBox strError, vbInformation, "销售管理"
'            sKey = LCase(sKey)
'            If sKey = "cinvname" Or sKey = "cinvcode" Then
'                bChanged = success 'Cancel  'retry
'                .bodyText(r, "cinvcode") = ""
'                .bodyText(r, "cinvname") = ""
'                RetValue = ""
'            'zzg-860-Develop-Bug19615
'            ElseIf sKey = "cinva_unit" Then
'                bChanged = Cancel
'                RetValue = ""
'            Else
'                bChanged = IIf(bLostFocus, Cancel, retry) 'Cancel
'            End If
'            sKey = LCase(.ItemState(c, sibody).sFieldName)
'
'            If sKey = "cinvname" Or sKey = "cinvcode" Then
'                If Not Dombody Is Nothing Then
'                    If Dombody.selectNodes("//R").length > 0 Then
'                            For Each ele In Dombody.selectNodes("//R")
'                                .bodyText(r, ele.getAttribute("K")) = ele.getAttribute("V")
'                                If LCase(ele.getAttribute("K")) = sKey Then
'                                    RetValue = ""
'                                End If
'                            Next
'                    End If
'                End If
'            End If
'            Exit Sub
'        End If
'        sKey = LCase(.ItemState(c, sibody).sFieldName)
'        If Not Dombody Is Nothing Then
'            If Dombody.selectNodes("//R").length > 0 Then
'                For Each ele In Dombody.selectNodes("//R")
'                    .bodyText(r, ele.getAttribute("K")) = ele.getAttribute("V")
'                    If LCase(ele.getAttribute("K")) = sKey Then
'                        RetValue = ele.getAttribute("V")
'                    End If
'                    If sKey = "cinvname" Or sKey = "cinvcode" Then
'                        If strVouchType = "05" Or strVouchType = "06" Or strVouchType = "26" Or strVouchType = "27" Or strVouchType = "28" Or strVouchType = "29" Then
'                            If (LCase(ele.getAttribute("K")) = "binvtype" Or LCase(ele.getAttribute("K")) = "bservice") And ele.getAttribute("V") <> "" Then
'                                If CBool(ele.getAttribute("V")) = True Then
'                                    .bodyText(r, "cwhname") = "": .bodyText(r, "cwhcode") = ""
'                                End If
'                            End If
'                        End If
'                    End If
'                Next
'            End If
'        End If
'
'            Dim strInvCode As String
'            Dim dblQuan As Double
'            Dim dblNum As Double
'            strInvCode = .bodyText(r, "cinvcode")
'            If strInvCode <> "" Or sKey = "cinvcode" Then
'                If LCase(sKey) = "cinvname" Or LCase(sKey) = "cwhname" Or LCase(sKey) = "cbatch" Or LCase(sKey) = "cfree1" _
'                    Or LCase(sKey) = "cfree2" Or LCase(sKey) = "cfree3" Or LCase(sKey) = "cfree4" Or LCase(sKey) = "cfree5" _
'                    Or LCase(sKey) = "cfree6" Or LCase(sKey) = "cfree7" Or LCase(sKey) = "cfree8" Or LCase(sKey) = "cfree9" Or _
'                    LCase(sKey) = "cfree10" Then
'                    If strInvCode = "" Then strInvCode = RetValue
'                    If clsSAWeb_M.GetSumQuantity(DBConn, strInvCode, dblQuan, dblNum, .bodyText(r, "cfree1"), _
'                        .bodyText(r, "cfree2"), .bodyText(r, "cfree3"), .bodyText(r, "cfree4"), .bodyText(r, "cfree5"), _
'                        .bodyText(r, "cfree6"), .bodyText(r, "cfree7"), .bodyText(r, "cfree8"), .bodyText(r, "cfree9"), _
'                        .bodyText(r, "cfree10"), .bodyText(r, "cbatch"), .bodyText(r, "cwhcode")) Then
'                        .headerText("fstockquan") = dblQuan
'                        .headerText("fcanusequan") = dblNum
'                    Else
'                        MsgBox "取可用量失败"
'                    End If
'                End If
'            End If
'    If .headerText("itaxrate") = "" And iVouchState <> 2 Then
'       .headerText("itaxrate") = .bodyText(r, "itaxrate")
'    End If
'    End With
'    Exit Sub
'DoERR:
'    MsgBox Err.Description
End Sub


  Private Sub Voucher4_bodyCellCheck(RetValue As Variant, bChanged As Long, ByVal r As Long, ByVal c As Long, referPara As ReferParameter)
    Dim intNumPoint As Integer
    Dim strFieldName As String
    Dim lngOldRow As Long
    Dim lngOldRows As Long
    Dim i As Long
    
    referPara.bValid = True
    strFieldName = Voucher4.ItemState(c, sibody).sFieldName
    lngOldRow = r
    lngOldRows = Voucher4.BodyRows
    If Not referPara.rstGrid Is Nothing Then
        strReferString = clsVoucherRefer.FillItemsAfterBrowse(clsVoucher, Voucher4, "B", strFieldName, referPara.rstGrid, r)
        If Voucher4.BodyRows > lngOldRows Then
            For i = lngOldRows + 1 To Voucher4.BodyRows
                clsVoucher.CellCheck Voucher4, "B", Voucher4.ItemState(c, sibody).sFieldName, bChanged, clsVoucherRefer, i
            Next
        End If
        r = lngOldRow
    Else
        strFieldName = clsVoucherRefer.CellCheck("", Voucher4, "B", strFieldName, r)
        If strFieldName <> "" Then
            If strFieldName = "cancel" Then
                bChanged = Cancel
            Else
                RetValue = ""
'                bChanged = retry
            End If
            Exit Sub
        End If
    End If
'    clsVoucherRefer.CellCheck strReferString, Voucher, "B", Voucher.ItemState(C, sibody).sFieldName, R
    clsVoucher.CellCheck Voucher4, "B", Voucher4.ItemState(c, sibody).sFieldName, bChanged, clsVoucherRefer, r
    RetValue = Voucher4.bodyText(lngOldRow, c)
    If Voucher4.ItemState(c, sibody).nFieldType = 4 Then
        intNumPoint = Voucher4.ItemState(c, sibody).nNumPoint
        RetValue = Format(RetValue, IIf(intNumPoint = 0, "###0", "###0." & String(intNumPoint, "0")))
    End If
    If Voucher4.ItemState(c, sibody).nReferType = 3 Then
        RetValue = Format(RetValue, "YYYY-MM-DD")
    End If
    
    RetValue = Voucher4.bodyText(r, c)
    strFieldName = Voucher4.ItemState(c, sibody).sFieldName
    Select Case LCase(strFieldName)
        Case "iquantity" ''令数
            If val(RetValue) < 0 Then
                MsgBox "令数不能小于0", vbInformation
                bChanged = Cancel: Exit Sub
            End If
        Case "iinvexchratejf" ''加放率
            If val(RetValue) < 0 Then
                MsgBox "加放率不能小于0", vbInformation
                bChanged = Cancel: Exit Sub
            End If
        Case "iquantityjf" ''加放数量
            If val(RetValue) < 0 Then
                MsgBox "加放数量不能小于0", vbInformation
                bChanged = Cancel: Exit Sub
            End If
        Case "iquantityhj" ''合计数量
            If val(RetValue) < 0 Then
                MsgBox "合计数量不能小于0", vbInformation
                bChanged = Cancel: Exit Sub
            End If
    End Select
'zhupeibi  根据令数、加放率求加放数量和合计数量
     strFieldName = Voucher4.ItemState(c, sibody).sFieldName
     
     Select Case LCase(strFieldName)
           Case "iinvexchrate"
                If val(Voucher4.bodyText(Voucher4.row, "iinvexchrate")) <> 0 Then
                    Voucher4.bodyText(Voucher4.row, "inum") = val(Voucher4.bodyText(Voucher4.row, "iquantity")) / val(Voucher4.bodyText(Voucher4.row, "iinvexchrate"))
                End If
           Case "iquantity"
                If val(Voucher4.bodyText(Voucher4.row, "iinvexchratejf")) <> 0 Then
                Voucher4.bodyText(Voucher4.row, "iquantityjf") = val(Voucher4.bodyText(Voucher4.row, "iquantity")) * val(Voucher4.bodyText(Voucher4.row, "iinvexchratejf")) / 1000
                Voucher4.bodyText(Voucher4.row, "iquantityhj") = val(Voucher4.bodyText(Voucher4.row, "iquantity")) + val(Voucher4.bodyText(Voucher4.row, "iquantity")) * val(Voucher4.bodyText(Voucher4.row, "iinvexchratejf")) / 1000
                End If
           Case "iinvexchratejf"
                If val(Voucher4.bodyText(Voucher4.row, "iinvexchratejf")) <> 0 Then
                    Voucher4.bodyText(Voucher4.row, "iquantityjf") = val(Voucher4.bodyText(Voucher4.row, "iquantity")) * val(Voucher4.bodyText(Voucher4.row, "iinvexchratejf")) / 1000
                    Voucher4.bodyText(Voucher4.row, "iquantityhj") = val(Voucher4.bodyText(Voucher4.row, "iquantity")) + val(Voucher4.bodyText(Voucher4.row, "iquantity")) * val(Voucher4.bodyText(Voucher4.row, "iinvexchratejf")) / 1000
                End If
           Case "iquantityjf"
                  If val(Voucher4.bodyText(Voucher4.row, "iinvexchratejf")) <> 0 Then
                    Voucher4.bodyText(Voucher4.row, "iquantity") = val(Voucher4.bodyText(Voucher4.row, "iquantityjf")) / (val(Voucher4.bodyText(Voucher4.row, "iinvexchratejf")) / 1000)
                    Voucher4.bodyText(Voucher4.row, "iquantityhj") = val(Voucher4.bodyText(Voucher4.row, "iquantity")) + val(Voucher4.bodyText(Voucher4.row, "iquantity")) * val(Voucher4.bodyText(Voucher4.row, "iinvexchratejf")) / 1000
                   
                  End If
     End Select
     
'    Dim lngRow As Long
'    Dim strError As String
'    Dim tmpDomhead As DOMDocument, tmpDOMBody As DOMDocument
'    Dim sKey As String
'    Dim ele As IXMLDOMElement
'    Dim sKeyValue As String
'
'    'Dim DOMHead As New DOMDocument
'   ' Dim DomBody As New DOMDocument
'
'   ' On Error Resume Next
'    On Error GoTo DoERR
'    ''是否放弃
'    'If bClickCancel Then Exit Sub
'
'    With Me.Voucher4
'        sKey = LCase(.ItemState(c, sibody).sFieldName)
'        If Left(sKey, 7) = "cdefine" Or Left(sKey, 5) = "cfree" Then
'        ''使用新的自定义项目
'            'sRet = RefDefine((.LookUpArray(sKey, Sibody)), Sibody)
'            'RetValue = CellCheckDefine(.LookUpArray(sKey, sibody), sibody, .bodyText(R, sKey))
'            If .bodyText(r, sKey) <> RetValue Then
'                sKeyValue = RetValue
'            Else
'                sKeyValue = .bodyText(r, sKey)
'            End If
'            RetValue = CellCheckDefine(.LookUpArray(sKey, sibody), sibody, sKeyValue)
'            If Left(sKey, 5) = "cfree" Then
'                If RetValue <> "" Then
'                    .bodyText(r, sKey) = RetValue
'                Else
'                    Exit Sub
'                End If
'            ElseIf Left(sKey, 7) = "cdefine" Then
'                Exit Sub
'            End If
'
'        End If
'        Set tmpDomhead = .GetHeadDom
'        Set tmpDOMBody = .GetLineDom(r)
'        Set Domhead4 = tmpDomhead
'        Set Dombody4 = tmpDOMBody
'        If BrowFlag = True And sKey = LCase(strRefFldName) Then
'            BrowFlag = False
'            strRefFldName = ""
'            Exit Sub
'        End If
'
'        If strVouchType <> "26" And sKey <> "iquantity" Then
'        strError = clsVoucherCO.BodyCheck(sKey, Dombody4, Domhead4, r)
'        End If
'        If strError <> "" Then
'            MsgBox strError, vbInformation, "销售管理"
'            sKey = LCase(sKey)
'            If sKey = "cinvname" Or sKey = "cinvcode" Then
'                bChanged = success 'Cancel  'retry
'                .bodyText(r, "cinvcode") = ""
'                .bodyText(r, "cinvname") = ""
'                RetValue = ""
'            'zzg-860-Develop-Bug19615
'            ElseIf sKey = "cinva_unit" Then
'                bChanged = Cancel
'                RetValue = ""
'            Else
'                bChanged = IIf(bLostFocus, Cancel, retry) 'Cancel
'            End If
'            sKey = LCase(.ItemState(c, sibody).sFieldName)
'
'            If sKey = "cinvname" Or sKey = "cinvcode" Then
'                If Not Dombody4 Is Nothing Then
'                    If Dombody4.selectNodes("//R").length > 0 Then
'                            For Each ele In Dombody4.selectNodes("//R")
'                                .bodyText(r, ele.getAttribute("K")) = ele.getAttribute("V")
'                                If LCase(ele.getAttribute("K")) = sKey Then
'                                    RetValue = ""
'                                End If
'                            Next
'                    End If
'                End If
'            End If
'            Exit Sub
'        End If
'        sKey = LCase(.ItemState(c, sibody).sFieldName)
'        If Not Dombody4 Is Nothing Then
'            If Dombody4.selectNodes("//R").length > 0 Then
'                For Each ele In Dombody4.selectNodes("//R")
'                    .bodyText(r, ele.getAttribute("K")) = ele.getAttribute("V")
'                    If LCase(ele.getAttribute("K")) = sKey Then
'                        RetValue = ele.getAttribute("V")
'                    End If
'                    If sKey = "cinvname" Or sKey = "cinvcode" Then
'                        If strVouchType = "05" Or strVouchType = "06" Or strVouchType = "26" Or strVouchType = "27" Or strVouchType = "28" Or strVouchType = "29" Then
'                            If (LCase(ele.getAttribute("K")) = "binvtype" Or LCase(ele.getAttribute("K")) = "bservice") And ele.getAttribute("V") <> "" Then
'                                If CBool(ele.getAttribute("V")) = True Then
'                                    .bodyText(r, "cwhname") = "": .bodyText(r, "cwhcode") = ""
'                                End If
'                            End If
'                        End If
'                    End If
'                Next
'            End If
'        End If
'
'            Dim strInvCode As String
'            Dim dblQuan As Double
'            Dim dblNum As Double
'            strInvCode = Voucher4.bodyText(r, "cinvcode")
'            If strInvCode <> "" Or skey = "cinvcode" Then
'                If LCase(skey) = "cinvname" Or LCase(skey) = "cwhname" Or LCase(skey) = "cbatch" Or LCase(skey) = "cfree1" _
'                    Or LCase(skey) = "cfree2" Or LCase(skey) = "cfree3" Or LCase(skey) = "cfree4" Or LCase(skey) = "cfree5" _
'                    Or LCase(skey) = "cfree6" Or LCase(skey) = "cfree7" Or LCase(skey) = "cfree8" Or LCase(skey) = "cfree9" Or _
'                    LCase(skey) = "cfree10" Then
'                    If strInvCode = "" Then strInvCode = RetValue
'                    If clsSAWeb_M.GetSumQuantity(DBConn, strInvCode, dblQuan, dblNum, Voucher4.bodyText(r, "cfree1"), _
'                        Voucher4.bodyText(r, "cfree2"), Voucher4.bodyText(r, "cfree3"), Voucher4.bodyText(r, "cfree4"), Voucher4.bodyText(r, "cfree5"), _
'                        Voucher4.bodyText(r, "cfree6"), Voucher4.bodyText(r, "cfree7"), Voucher4.bodyText(r, "cfree8"), Voucher4.bodyText(r, "cfree9"), _
'                        Voucher4.bodyText(r, "cfree10"), Voucher4.bodyText(r, "cbatch"), Voucher4.bodyText(r, "cwhcode")) Then
'                        .headerText("fstockquan") = dblQuan
'                        .headerText("fcanusequan") = dblNum
'                    Else
'                        MsgBox "取可用量失败"
'                    End If
'                End If
'            End If
'        If skey = "iinvexchratejf" Then
'            If val(Me.voucher.headerText("numprint")) >= 5000 Then
'                If Voucher4.bodyText(r, "iinvexchratejf") <> "" Then          'sl 添加 校验 加放率*数量＝加放数量
'                   Voucher4.bodyText(r, "iquantityjf") = Voucher4.bodyText(r, "iquantity") * Voucher4.bodyText(r, "iinvexchratejf")
'                   Voucher4.bodyText(r, "iquantityhj") = Voucher4.bodyText(r, "iquantity") * Voucher4.bodyText(r, "iinvexchratejf") + Voucher4.bodyText(r, "iquantity")
'                   Voucher4.bodyText(r, "inumjf") = Voucher4.bodyText(r, "inum") * Voucher4.bodyText(r, "iinvexchratejf")
'                   Voucher4.bodyText(r, "inumhj") = Voucher4.bodyText(r, "inum") * Voucher4.bodyText(r, "iinvexchratejf") + Voucher4.bodyText(r, "inum")
'                End If
'            Else
'                If Voucher4.bodyText(r, "iinvexchratejf") <> "" Then          'sl 添加 校验 加放率*数量＝加放数量
'                   Voucher4.bodyText(r, "iquantityjf") = 50 * Voucher4.bodyText(r, "iinvexchratejf")
'                   Voucher4.bodyText(r, "iquantityhj") = 50 * Voucher4.bodyText(r, "iinvexchratejf") + Voucher4.bodyText(r, "iquantity")
'                   Voucher4.bodyText(r, "inumjf") = Voucher4.bodyText(r, "inum") * Voucher4.bodyText(r, "iinvexchratejf")
'                   Voucher4.bodyText(r, "inumhj") = Voucher4.bodyText(r, "inum") * Voucher4.bodyText(r, "iinvexchratejf") + Voucher4.bodyText(r, "inum")
'                End If
'            End If
'        End If
'        If skey = "iquantityjf" Then
''           If val(Me.voucher.headerText("numprint")) >= 5000 Then
'               If val(Voucher4.bodyText(r, "iinvexchratejf")) <> 0 Then
'                Voucher4.bodyText(r, "iinvexchratejf") = Voucher4.bodyText(r, "iquantityjf") / Voucher4.bodyText(r, "iquantity")
'                Voucher4.bodyText(r, "inumjf") = Voucher4.bodyText(r, "inum") * (Voucher4.bodyText(r, "iquantityjf") / Voucher4.bodyText(r, "iquantity"))
'                Voucher4.bodyText(r, "iquantityhj") = val(Voucher4.bodyText(r, "iquantity")) + val(Voucher4.bodyText(r, "iquantityjf"))
'                Voucher4.bodyText(r, "inumhj") = Voucher4.bodyText(r, "inum") + Voucher4.bodyText(r, "inum") * (Voucher4.bodyText(r, "iquantityjf") / Voucher4.bodyText(r, "iquantity"))
'               End If
''            Else
''                If val(voucher4.bodyText(R, "iinvexchratejf")) <> 0 Then
''                  voucher4.bodyText(R, "iinvexchratejf") = voucher4.bodyText(R, "iquantityjf") / voucher4.bodyText(R, "iquantity")
''                  voucher4.bodyText(R, "inumjf") = voucher4.bodyText(R, "inum") * (voucher4.bodyText(R, "iquantityjf") / voucher4.bodyText(R, "iquantity"))
''                  voucher4.bodyText(R, "iquantityhj") = val(voucher4.bodyText(R, "iquantity")) + val(voucher4.bodyText(R, "iquantityjf"))
''                  voucher4.bodyText(R, "inumhj") = voucher4.bodyText(R, "inum") + voucher4.bodyText(R, "inum") * (voucher4.bodyText(R, "iquantityjf") / voucher4.bodyText(R, "iquantity"))
''                End If
''            End If
'
'        End If
'       If skey = "inumjf" Then
''            If val(Me.voucher.headerText("numprint")) >= 5000 Then
''                If val(voucher4.bodyText(R, "iinvexchratejf")) <> 0 Then
''                  voucher4.bodyText(R, "iinvexchratejf") = voucher4.bodyText(R, "inumjf") / voucher4.bodyText(R, "inum")
''                   voucher4.bodyText(R, "iquantityjf") = voucher4.bodyText(R, "iquantity") * voucher4.bodyText(R, "iinvexchratejf")
''                  voucher4.bodyText(R, "iquantityhj") = val(voucher4.bodyText(R, "iquantity")) + val(voucher4.bodyText(R, "iquantityjf"))
''                  voucher4.bodyText(R, "inumhj") = val(voucher4.bodyText(R, "inum")) + val(voucher4.bodyText(R, "inumjf"))
''                End If
''            Else
'                 If val(Voucher4.bodyText(r, "iinvexchratejf")) <> 0 Then
'                  Voucher4.bodyText(r, "iinvexchratejf") = Voucher4.bodyText(r, "inumjf") / Voucher4.bodyText(r, "inum")
'                   Voucher4.bodyText(r, "iquantityjf") = Voucher4.bodyText(r, "iquantity") * Voucher4.bodyText(r, "iinvexchratejf")
'                  Voucher4.bodyText(r, "iquantityhj") = val(Voucher4.bodyText(r, "iquantity")) + val(Voucher4.bodyText(r, "iquantityjf"))
'                  Voucher4.bodyText(r, "inumhj") = val(Voucher4.bodyText(r, "inum")) + val(Voucher4.bodyText(r, "inumjf"))
'                End If
''            End If
'
'       End If
'    End With
'    Exit Sub
'DoERR:
'    MsgBox Err.Description
End Sub

 
 
 
 

 
 

 
 

 
 

 

 



 

  ''自定义项目检查，调用u8defpro
Private Function CellCheckDefine(Index As Variant, iVoucherSec As Integer, KeyValue As String) As String
    Dim clsDef As U8DefPro.clsDefPro
    Dim nDataSource As Long         '数据来源
    Dim nEnterType As Long         '输入方式
    Dim sDataRule As String       '数据公式
    Dim bValidityCheck As Boolean      '是否合法性检测
    Dim bBuildArchives As Boolean      '是否建档
    Dim sVouchType As String
    Dim sTableName As String, sFieldName As String, sCardNumber As String
    Dim sDefWhere As String
    Dim intReturn As Integer
    Dim bFixlenth As Boolean
    Dim lngFixLenth As Long
    Dim strTmp As String
    Dim strDataType As String
    Dim sKeyName As String      ''cellcheck的项目名称
    Dim bFree As Boolean        ''是否自由项
    
    If KeyValue = "" Then
        CellCheckDefine = ""
        Exit Function
    End If
    Set clsDef = New U8DefPro.clsDefPro
    If Not clsDef.Init(False, DBConn.ConnectionString, m_Login.cUserId) Then
        CellCheckDefine = ""
        MsgBox "初始化自定义项组件失败！"
        Set clsDef = Nothing
        Exit Function
    End If
    With Me.Voucher
        sKeyName = LCase(.ItemState(Index, iVoucherSec).sFieldName)
        If Left(sKeyName, 5) = "cfree" Then
            bFree = True
        Else
            bFree = False
        End If
        If bFree = False Then
            nDataSource = .ItemState(Index, iVoucherSec).nDataSource
            nEnterType = .ItemState(Index, iVoucherSec).nEnterType
            sDataRule = .ItemState(Index, iVoucherSec).sDataRule
            bValidityCheck = .ItemState(Index, iVoucherSec).bValidityCheck
            strTmp = .ItemState(Index, iVoucherSec).sDefaultValue
            bFixlenth = IIf(Left(strTmp, 1) = "1", True, False)
            lngFixLenth = val(Mid(strTmp, 3))
            strDataType = .ItemState(Index, iVoucherSec).nFieldType
            ''是否建立档
            bBuildArchives = .ItemState(Index, iVoucherSec).bBuildArchives
            If bValidityCheck = False Then
                CellCheckDefine = KeyValue
                If strDataType = "3" Then         '
                    If Abs(val(CellCheckDefine)) > 2147483647 Then
                        MsgBox "当前项目设置为整型，输入的值" & CellCheckDefine & "的输入超出可取值范围（-2,147,483,647到2,147,483,647）!"
                        CellCheckDefine = ""
                    End If
                End If
                If CellCheckDefine <> "" And bFixlenth Then
                    If GetStrTrueLenth(CellCheckDefine) <> lngFixLenth Then
                        CellCheckDefine = ""
                        MsgBox "项目为定长" & lngFixLenth & ",长度不符合！"
                    End If
                End If
                If LCase(sTableName) = "" Or LCase(sTableName) = "userdefine" Then
                    'Call clsDef.UsrDefCodeToValue(iVoucherSec, .ItemState(Index, iVoucherSec).sFieldName, KeyValue)
                End If
                Set clsDef = Nothing
                Exit Function
            End If
            ''定长时如果长度不等于定长不建档
            If bFixlenth Then
                If GetStrTrueLenth(KeyValue) <> lngFixLenth Then
                    bBuildArchives = False
                End If
            End If
            
            Select Case nDataSource  '0表示手工输入；1表示档案；2表示单据
                Case 0
                    sTableName = "UserDefine"
                    sFieldName = "cValue"
                    sVouchType = ""
                Case 1
                    sTableName = Left(sDataRule, InStr(1, sDataRule, ",") - 1)
                    sFieldName = Mid(sDataRule, InStr(1, sDataRule, ",") + 1)
                    sVouchType = ""
                Case 2
                    sCardNumber = Left(sDataRule, InStr(1, sDataRule, ",") - 1)
                    sFieldName = Mid(sDataRule, InStr(1, sDataRule, ",") + 1)
            End Select
        Else
            ''自由项
            bBuildArchives = .ItemState(Index, iVoucherSec).bBuildArchives
        End If
        
        
        If bFree = False Then
            intReturn = clsDef.ValidateAr(nDataSource, iVoucherSec, .ItemState(Index, iVoucherSec).sFieldName, sTableName, sFieldName, KeyValue, sCardNumber, "", bBuildArchives)
        Else
            intReturn = clsDef.ValidateFreeAr(.ItemState(Index, iVoucherSec).sFieldName, KeyValue, bBuildArchives)
        End If
        Select Case intReturn
            Case 0  '合法性检查成功
                CellCheckDefine = KeyValue
            Case 1  '合法性检查失败，建档成功；
                CellCheckDefine = KeyValue
                'MsgBox "合法性检查失败，建档成功"
            Case -2 '-2表示合法性检查失败，建档失败
                If bValidityCheck = True Then
                    CellCheckDefine = ""
                    MsgBox "合法性检查失败，建档失败"
                Else
                    If KeyValue <> "" Then CellCheckDefine = KeyValue
                End If
            Case -1 '-1表示合法性检查失败；
                If bValidityCheck = True Then
                    CellCheckDefine = ""
                    'MsgBox "合法性检查失败"
                    MsgBox "录入不合法，请检查"
                Else
                    If KeyValue <> "" Then CellCheckDefine = KeyValue
                End If
        End Select
        If CellCheckDefine <> "" And bFixlenth Then
            If GetStrTrueLenth(CellCheckDefine) <> lngFixLenth Then
                CellCheckDefine = ""
                MsgBox "项目为定长" & lngFixLenth & ",长度不符合！"
            End If
        End If
        If strDataType = "3" Then
                If Abs(val(CellCheckDefine)) > 2147483647 Then
                    MsgBox "当前项目设置为整型，输入的值" & CellCheckDefine & "的输入超出可取值范围（-2,147,483,647到2,147,483,647）!"
                    CellCheckDefine = ""
                End If
            End If
    End With
    Set clsDef = Nothing
        
End Function

Private Sub Voucher_FillHeadComboBox(ByVal Index As Long, pCom As Object)
    Select Case LCase(Me.Voucher.ItemState(Index, siheader).sFieldName)

        Case "usestate"  '状态
            pCom.Clear
            pCom.AddItem "在用"
            pCom.AddItem "在建"
            
        Case "sdevsource"   '设备来源
            pCom.Clear
            pCom.AddItem "购置"
            pCom.AddItem "基建"
            pCom.AddItem "无偿调入"
            pCom.AddItem "接受捐赠"
            pCom.AddItem "盘盈"
            pCom.AddItem "自制"
            pCom.AddItem "其它"
            
        Case "buildform"   '建筑结构 1.钢混2.砖混3.砖木4.其它
            pCom.Clear
            pCom.AddItem "钢混"
            pCom.AddItem "砖混"
            pCom.AddItem "砖木"
            pCom.AddItem "其它"
            
        Case "buildmod"   '购建形式      1.自建2.购买3.伪产4.其它
            pCom.Clear
            pCom.AddItem "自建"
            pCom.AddItem "购买"
            pCom.AddItem "伪产"
            pCom.AddItem "其它"
            
        Case "motfuel"   '燃料      1.汽油2.柴油3.天然气4.双燃料5.太阳能6.电7.其它
            pCom.Clear
            pCom.AddItem "汽油"
            pCom.AddItem "柴油"
            pCom.AddItem "天然气"
            pCom.AddItem "双燃料"
            pCom.AddItem "太阳能"
            pCom.AddItem "电"
            pCom.AddItem "其它"
            
       Case "direction"  ' 舵向
            pCom.Clear
            pCom.AddItem "左"
            pCom.AddItem "右"
            
       Case Else
            pCom.Clear
    End Select
    
End Sub
 
Private Sub Voucher_FillList(ByVal r As Long, ByVal c As Long, pCom As Object)
    Dim sFieldName As String
    sFieldName = LCase(Me.Voucher.ItemState(c, sibody).sFieldName)
    Select Case sFieldName
        Case "usestate_after"
            If Trim(Voucher.bodyText(r, "usestate_before")) = "在建" Then
                pCom.Clear
                pCom.AddItem ""
                pCom.AddItem "在用"
'                pCom.AddItem "在建"
            End If
        Case "ending"
            pCom.Clear
            pCom.AddItem "盘亏"      '"盘亏"，”盘实“，”变更“，盘盈
            pCom.AddItem "盘实"
            pCom.AddItem "变更"
            pCom.AddItem "盘盈"
            
        Case "ifalg"
            pCom.Clear
            pCom.AddItem "是"
            pCom.AddItem "否"
    
        Case "sdevsource"   '设备来源
            pCom.Clear
            pCom.AddItem "购置"
            pCom.AddItem "基建"
            pCom.AddItem "无偿调入"
            pCom.AddItem "接受捐赠"
            pCom.AddItem "盘盈"
            pCom.AddItem "自制"
            pCom.AddItem "其它"
            
        Case "buildform"   '建筑结构 1.钢混2.砖混3.砖木4.其它
            pCom.Clear
            pCom.AddItem "钢混"
            pCom.AddItem "砖混"
            pCom.AddItem "砖木"
            pCom.AddItem "其它"
            
        Case "buildmod"   '购建形式      1.自建2.购买3.伪产4.其它
            pCom.Clear
            pCom.AddItem "自建"
            pCom.AddItem "购买"
            pCom.AddItem "伪产"
            pCom.AddItem "其它"
            
        Case "motfuel"   '燃料      1.汽油2.柴油3.天然气4.双燃料5.太阳能6.电7.其它
            pCom.Clear
            pCom.AddItem "汽油"
            pCom.AddItem "柴油"
            pCom.AddItem "天然气"
            pCom.AddItem "双燃料"
            pCom.AddItem "太阳能"
            pCom.AddItem "电"
            pCom.AddItem "其它"
            
       Case Else
            pCom.Clear
    
    End Select
End Sub
 

''Private Sub Voucher_controlError(ByVal nErr As Long, ByVal Description As String)
''    MsgBox Description
''End Sub
 
Private Sub Voucher_headBrowUser(ByVal Index As Variant, sRet As Variant, referPara As ReferParameter)    '付印通知单主体单据
    strReferString = clsVoucherRefer.ShowReferCtl(clsVoucher, Voucher, siheader, CLng(Index), referPara)
    sRet = Voucher.headerText(Index)
    
'    Dim iElement As IXMLDOMElement
'    Dim sKey As String, sKeyValue As String
'    Dim strSql As String
'    Dim sFormat As String
'    Dim strAuth As String
'    Dim strDate As String
'    Dim strCusInv As String
'    Dim oCRMServer As Object ' New UFCRMSRVSALE.clsOpportuntity
'    Dim oCRMServerOPP As Object 'New UFCRMSRVSALE.clsOppPro
'    Dim oCRMServerAct As Object 'New UFCRMSRVSALE.clsActivity
'    Dim rst As New ADODB.Recordset
'
'
'    On Error Resume Next
'
'    clsRefer.referMulti = False
'    clsRefer.SetReferDisplayMode enuGrid
'    clsRefer.SetReferSQLString ""
'    clsRefer.SetRWAuth "INVENTORY", "R", False
'    strAuth = ""
'    sKey = Me.voucher.ItemState(Index, siheader).sFieldName
'    sKeyValue = Me.voucher.headerText(Index)
'    Select Case LCase(sKey)
'
'        Case "cvenabbname", "cvenname"   'sl add 供应商参照
'         If Me.voucher.headerText("cvenabbname") = "" Then
'            clsRefer.SetReferDataType enuVendor
'            clsRefer.SetRWAuth "VENDOR", "W", True
'            clsRefer.Show
'            If Not clsRefer.recmx Is Nothing Or clsRefer.recmx.EOF = True Then
'                  sRet = clsRefer.recmx.Fields(Me.voucher.ItemState(Index, siheader).sFieldName)
'                   Me.voucher.headerText("cvencode") = clsRefer.recmx.Fields("cvencode")
'            End If
'          Else
'            strSql = "select cvencode,cvenabbname from vendor where cvenabbname like '%" & Me.voucher.headerText("cvenabbname") & "%' order by cvencode"
'            clsRefer.StrRefInit m_login, False, "", strSql, "供应商编码,供应商简称", "", False, 1, 1, 1
'            clsRefer.Show
'             If Not clsRefer.recmx Is Nothing Then
'               sRet = clsRefer.recmx.Fields("cvenabbname")
'               Me.voucher.headerText("cvencode") = clsRefer.recmx.Fields("cvencode")
'             End If
'          End If
'        Case "cdepname"
'            clsRefer.SetReferDataType enuDepartment
'            If myinfo.bAuth_Dep Then
'                clsRefer.SetRWAuth "DEPARTMENT", "W", True
'            Else
'                clsRefer.SetRWAuth "DEPARTMENT", "R", False
'            End If
'                clsRefer.SetReferFilterString "(bdepend=1) " & IIf(getReferString(sKey, sKeyValue) = "", "", " and (" & getReferString(sKey, sKeyValue) & ")")
'                clsRefer.Show
'            If Not clsRefer.recmx Is Nothing Or clsRefer.recmx.EOF = True Then
'                 sRet = clsRefer.recmx.Fields(Me.voucher.ItemState(Index, siheader).sFieldName)
'            End If
'        Case "cpersonname"
'            clsRefer.SetReferDataType enuPerson
'            If myinfo.bAuth_Per Then
'                'strAuth = clsAuth.GetAuthString("PERSON")
'                clsRefer.SetRWAuth "PERSON", "W", True
'            Else
'                clsRefer.SetRWAuth "PERSON", "R", False
'            End If
'
'            clsRefer.SetReferFilterString getReferString(sKey, sKeyValue) & IIf(voucher.headerText("cdepcode") <> "", IIf(getReferString(sKey, sKeyValue) = "", "", " and ") & " person.cdepcode='" & voucher.headerText("cdepcode") & "'", "")
'            clsRefer.Show
'            If Not clsRefer.recmx Is Nothing Then
'                sRet = clsRefer.recmx.Fields(Me.voucher.ItemState(Index, siheader).sFieldName)
'                Me.voucher.headerText("cPersonCode") = clsRefer.recmx.Fields("cPersonCode")
'                Me.voucher.headerText("cdepcode") = clsRefer.recmx.Fields("cdepcode")
'                Me.voucher.headerText("cdepname") = clsRefer.recmx.Fields("cdepname")
'                If strVouchType = "92" Or strVouchType = "95" Then
'                    Me.voucher.headerText("cHandler") = clsRefer.recmx.Fields("cPersonCode")
'                End If
'            End If
'        Case "cexch_name"
'            clsRefer.SetReferDataType enuforeigncurrency
'            clsRefer.SetReferFilterString getReferString(sKey, sKeyValue)
'            clsRefer.Show
'            If Not clsRefer.recmx Is Nothing Then
'                sRet = clsRefer.recmx.Fields(Me.voucher.ItemState(Index, siheader).sFieldName)
'            End If
'        Case "cinvname", "cinvcode"    'sl 修改为完全根据存货档案和发稿记录关联的字段参照
'           'strSQL = "select a.cinvcode,a.cinvname,a.cinvstd,b.isbncode,b.cpubdegree,b.cprintdegree from inventory a inner join EFBWGL_distrecord b on a.cinvcode=b.cbookcode"
'           If Me.voucher.headerText("cinvcode") = "" Then
'                strSql = "select a.cinvcode,a.cinvname,a.cinvstd,b.isbncode,b.cpubdegree,b.cprintdegree,a.iinvrcost,c.cformatcode,c.cbookbindingcode,EFBWGL_dbformat.cformatname,EFBWGL_dbbindway.cbindwayname,EFBWGL_seldeclare.cbookbindingcode as cbookbindingcode1,EFBWGL_dbbookbinding.cbookbindingname,EFBWGL_seldeclare.numprint" & _
'                       " from inventory a left join EFBWGL_distrecord b on a.cinvcode=b.cbookcode left join EFYZGL_pmaking c on a.cinvcode=c.cinvcode left join EFBWGL_dbformat on c.cformatcode=EFBWGL_dbformat.cformatcode  left join EFBWGL_dbbindway on c.cbookbindingcode=EFBWGL_dbbindway.cbindway  left join" & _
'                       " EFBWGL_selregister on b.selid=EFBWGL_selregister.id left join EFBWGL_seldeclare on EFBWGL_selregister.sdid=EFBWGL_seldeclare.id left join EFBWGL_dbbookbinding on EFBWGL_seldeclare.cbookbindingcode=EFBWGL_dbbookbinding.cbookbindingcode WHERE (a.cInvCode NOT IN (SELECT cinvcode FROM dbo.EFYZGL_V_pressinformT))"
'                clsRefer.StrRefInit m_login, False, "", strSql, "图书编码,图书名称,图书规格,标准书号,版次,印次,定价,开本编码,装订方法编码,开本名称,装订方法名称,装订规格编码,装订规格名称,印数", "", False, 1, 1, 1
'                clsRefer.Show
'                 If Not clsRefer.recmx Is Nothing Then
'                     sRet = clsRefer.recmx.Fields(Me.voucher.ItemState(Index, siheader).sFieldName)
'                     Me.voucher.headerText("cInvCode") = clsRefer.recmx.Fields("cInvCode")
'                     Me.voucher.headerText("cInvName") = clsRefer.recmx.Fields("cInvName")
'                     Me.voucher.headerText("cinvstd") = clsRefer.recmx.Fields("cinvstd")  '规格型号
'                     Me.voucher.headerText("isbn") = clsRefer.recmx.Fields("isbncode")
'                     Me.voucher.headerText("editionnum") = clsRefer.recmx.Fields("cpubdegree")
'                     Me.voucher.headerText("printnum") = clsRefer.recmx.Fields("cprintdegree")
'                     Me.voucher.headerText("iprice") = clsRefer.recmx.Fields("iinvrcost")
'                     Me.voucher.headerText("cformatcode") = clsRefer.recmx.Fields("cformatcode")
'                     Me.voucher.headerText("cformatname") = clsRefer.recmx.Fields("cformatname")
'                     Me.voucher.headerText("bindmanner") = clsRefer.recmx.Fields("cbookbindingcode")
'                     Me.voucher.headerText("cbindwayname") = clsRefer.recmx.Fields("cbindwayname")
'                     Me.voucher.headerText("cbookbindingcode") = clsRefer.recmx.Fields("cbookbindingcode1")
'                     Me.voucher.headerText("cbookbindingname") = clsRefer.recmx.Fields("cbookbindingname")
'                     Me.voucher.headerText("numprint") = clsRefer.recmx.Fields("numprint")
'                 End If
'            Else
'             strSql = "select a.cinvcode,a.cinvname,a.cinvstd,b.isbncode,b.cpubdegree,b.cprintdegree,a.iinvrcost,c.cformatcode,c.cbookbindingcode,EFBWGL_dbformat.cformatname,EFBWGL_dbbindway.cbindwayname,EFBWGL_seldeclare.cbookbindingcode as cbookbindingcode1,EFBWGL_dbbookbinding.cbookbindingname,EFBWGL_seldeclare.numprint" & _
'                       " from inventory a left join EFBWGL_distrecord b on a.cinvcode=b.cbookcode left join EFYZGL_pmaking c on a.cinvcode=c.cinvcode left join EFBWGL_dbformat on c.cformatcode=EFBWGL_dbformat.cformatcode  left join EFBWGL_dbbindway on c.cbookbindingcode=EFBWGL_dbbindway.cbindway  left join" & _
'                       " EFBWGL_selregister on b.selid=EFBWGL_selregister.id left join EFBWGL_seldeclare on EFBWGL_selregister.sdid=EFBWGL_seldeclare.id left join EFBWGL_dbbookbinding on EFBWGL_seldeclare.cbookbindingcode=EFBWGL_dbbookbinding.cbookbindingcode where (a.cinvname like '%" & Me.voucher.headerText("cinvcode") & "%' or a.cInvCode LIKE '%" & Me.voucher.headerText("cinvcode") & "%') and (a.cInvCode NOT IN (SELECT cinvcode FROM dbo.EFYZGL_V_pressinformT))"
'                clsRefer.StrRefInit m_login, False, "", strSql, "图书编码,图书名称,图书规格,标准书号,版次,印次,定价,开本编码,装订方法编码,开本名称,装订方法名称,装订规格编码,装订规格名称,印数", "", False, 1, 1, 1
'                clsRefer.Show
'                 If Not clsRefer.recmx Is Nothing Then
'                     sRet = clsRefer.recmx.Fields(Me.voucher.ItemState(Index, siheader).sFieldName)
'                     Me.voucher.headerText("cInvCode") = clsRefer.recmx.Fields("cInvCode")
'                     Me.voucher.headerText("cInvName") = clsRefer.recmx.Fields("cInvName")
'                     Me.voucher.headerText("cinvstd") = clsRefer.recmx.Fields("cinvstd")  '规格型号
'                     Me.voucher.headerText("isbn") = clsRefer.recmx.Fields("isbncode")
'                     Me.voucher.headerText("editionnum") = clsRefer.recmx.Fields("cpubdegree")
'                     Me.voucher.headerText("printnum") = clsRefer.recmx.Fields("cprintdegree")
'                     Me.voucher.headerText("iprice") = clsRefer.recmx.Fields("iinvrcost")
'                     Me.voucher.headerText("cformatcode") = clsRefer.recmx.Fields("cformatcode")
'                     Me.voucher.headerText("cformatname") = clsRefer.recmx.Fields("cformatname")
'                     Me.voucher.headerText("bindmanner") = clsRefer.recmx.Fields("cbookbindingcode")
'                     Me.voucher.headerText("cbindwayname") = clsRefer.recmx.Fields("cbindwayname")
'                     Me.voucher.headerText("cbookbindingcode") = clsRefer.recmx.Fields("cbookbindingcode1")
'                     Me.voucher.headerText("cbookbindingname") = clsRefer.recmx.Fields("cbookbindingname")
'                     Me.voucher.headerText("numprint") = clsRefer.recmx.Fields("numprint")
'                 End If
'            End If
'        Case "cmemo"
'            clsRefer.SetReferDataType 17
'            clsRefer.Show
'
'            If Not clsRefer.recmx Is Nothing Then
'                sRet = clsRefer.recmx.Fields("ctext")
'
'            End If
'        Case "ivtid"
'            strAuth = clsAuth.getAuthString("DJMB")
'            'clsRefer.SetRWAuth strAuth
'            clsRefer.SetReferSQLString "SELECT VT_Name as 单据模版名称, VT_ID as 单据模版编号 From VoucherTemplates WHERE (VT_CardNumber = '" & strCardNum & "') AND (VT_TemplateMode = 0) " & IIf(strAuth <> "", "and vt_id in (" & IIf(strAuth = "1=2", "0", strAuth) & ")", "")
'            clsRefer.Show
'            If Not clsRefer.recmx Is Nothing Then
'                sRet = clsRefer.recmx.Fields("单据模版编号")
'
'            End If
'        Case "cformatcode"                  'sl 开本参照基础档案
'             clsRefer.StrRefInit m_login, False, "", "select cformatcode,cformatname from EFBWGL_dbformat ", "开本编码,开本名称", "", False, 1, 1, 1
'             clsRefer.Show
'             If Not clsRefer.recmx Is Nothing Then
'               sRet = clsRefer.recmx.Fields("cformatcode")
'               Me.voucher.headerText("cformatname") = clsRefer.recmx.Fields("cformatname")
'             End If
'        Case "cformatname"                  'sl 开本名称参照基础档案
'             clsRefer.StrRefInit m_login, False, "", "select cformatcode,cformatname from EFBWGL_dbformat ", "开本编码,开本名称", "", False, 1, 1, 1
'             clsRefer.Show
'             If Not clsRefer.recmx Is Nothing Then
'               sRet = clsRefer.recmx.Fields("cformatname")
'               Me.voucher.headerText("cformatcode") = clsRefer.recmx.Fields("cformatcode")
'             End If
'
'        Case "cbookbindingcode"             'sl 正文装订规格参照基础档案EFBWGL_dbbookbinding
'             clsRefer.StrRefInit m_login, False, "", "select cbookbindingcode,cbookbindingname from EFBWGL_dbbookbinding", "装订规格编码,装订规格名称", "", False, 1, 1, 1
'             clsRefer.Show
'             If Not clsRefer.recmx Is Nothing Then
'               sRet = clsRefer.recmx.Fields("cbookbindingcode")
'               Me.voucher.headerText("cbookbindingname") = clsRefer.recmx.Fields("cbookbindingname")
'             End If
'        Case "cbookbindingname"             'sl 正文装订规格参照基础档案EFBWGL_dbbookbinding
'             clsRefer.StrRefInit m_login, False, "", "select cbookbindingcode,cbookbindingname from EFBWGL_dbbookbinding ", "装订规格编码,装订规格名称", "", False, 1, 1, 1
'             clsRefer.Show
'             If Not clsRefer.recmx Is Nothing Then
'               sRet = clsRefer.recmx.Fields("cbookbindingname")
'               Me.voucher.headerText("cbindway") = clsRefer.recmx.Fields("cbindway")
'             End If
'
'        Case "cbookbindingcode1"             'sl 装订订式参照基础档案EFBWGL_Dbbindmanner
'             clsRefer.StrRefInit m_login, False, "", "select cbindmanner,cbindmannername from EFBWGL_dbbindmanner ", "订式编码,订式名称", "", False, 1, 1, 1
'             clsRefer.Show
'             If Not clsRefer.recmx Is Nothing Then
'               sRet = clsRefer.recmx.Fields("cbindmanner")
'               Me.voucher.headerText("cbindmannername") = clsRefer.recmx.Fields("cbindmannername")
'             End If
'        Case "cbindmannername"             'sl 装订订式参照基础档案EFBWGL_dbbindmanner
'             clsRefer.StrRefInit m_login, False, "", "select cbindmanner,cbindmannername from EFBWGL_dbbindmanner", "订式编码,订式名称", "", False, 1, 1, 1
'             clsRefer.Show
'             If Not clsRefer.recmx Is Nothing Then
'               sRet = clsRefer.recmx.Fields("cbindmannername")
'               Me.voucher.headerText("cbookbindingcode1") = clsRefer.recmx.Fields("cbindmanner")
'             End If
'
'        Case "bindmanner"             'sl 装订订法参照基础档案EFBWGL_Dbbindway
'             clsRefer.StrRefInit m_login, False, "", "select cbindway,cbindwayname from EFBWGL_dbbindway ", "订法编码,订法名称", "", False, 1, 1, 1
'             clsRefer.Show
'             If Not clsRefer.recmx Is Nothing Then
'               sRet = clsRefer.recmx.Fields("cbindway")
'               Me.voucher.headerText("cbindwayname") = clsRefer.recmx.Fields("cbindwayname")
'             End If
'        Case "cbindwayname"             'sl 装订订法参照基础档案EFBWGL_dbbindway
'             clsRefer.StrRefInit m_login, False, "", "select cbindway,cbindwayname from EFBWGL_dbbindway ", "订法编码,订法名称", "", False, 1, 1, 1
'             clsRefer.Show
'             If Not clsRefer.recmx Is Nothing Then
'               sRet = clsRefer.recmx.Fields("cbindwayname")
'               Me.voucher.headerText("bindmanner") = clsRefer.recmx.Fields("cbindway")
'             End If
'    End Select
'    If Left(sKey, 7) = "cdefine" Then
'        With Me.voucher
'            sRet = RefDefine(Index, siheader)
'        End With
'    End If
''by lg070315　增加U870 UAP单据控件新的参照处理
'    referPara.Cancel = True
'
'    If rst.state = 1 Then rst.Close
'    Set rst = Nothing
'    Exit Sub
End Sub

Private Function RefDefine(Index As Variant, iVoucherSec As Integer) As String
    Dim clsDef As U8DefPro.clsDefPro
    Dim nDataSource As Long         '数据来源
    Dim nEnterType As Long         '输入方式
    Dim sDataRule As String       '数据公式
    Dim bValidityCheck As Boolean      '是否合法性检测
    Dim bBuildArchives As Boolean      '是否建档
    Dim sVouchType As String
    Dim sTableName As String, sFieldName As String, sCardNumber As String
    Dim sDefWhere As String
    Dim strKeyValue As String
    Set clsDef = New U8DefPro.clsDefPro
        With Me.Voucher
            If iVoucherSec = siheader Then
                strKeyValue = .headerText(Index)
            Else
                strKeyValue = .bodyText(.row, Index)
            End If
            nDataSource = .ItemState(Index, iVoucherSec).nDataSource
            nEnterType = .ItemState(Index, iVoucherSec).nEnterType
            sDataRule = .ItemState(Index, iVoucherSec).sDataRule
            bValidityCheck = .ItemState(Index, iVoucherSec).bValidityCheck
            bBuildArchives = .ItemState(Index, iVoucherSec).bBuildArchives
            Select Case nDataSource  '0表示手工输入；1表示档案；2表示单据
                Case 0
                    sTableName = "UserDefine"
                    sFieldName = "cValue"
                    sVouchType = ""
                Case 1
                    sTableName = Left(sDataRule, InStr(1, sDataRule, ",") - 1)
                    sFieldName = Mid(sDataRule, InStr(1, sDataRule, ",") + 1)
                    sVouchType = ""
                Case 2
                    sCardNumber = Left(sDataRule, InStr(1, sDataRule, ",") - 1)
                    sFieldName = Mid(sDataRule, InStr(1, sDataRule, ",") + 1)
            End Select
            If Not clsDef.Init(False, DBConn.ConnectionString, m_Login.cUserId) Then
                RefDefine = ""
                MsgBox "初始化自定义项组件失败！"
                Exit Function
            End If
            RefDefine = clsDef.GetRefVal(nDataSource, iVoucherSec, .ItemState(Index, iVoucherSec).sFieldName, sTableName, sFieldName, sCardNumber, strKeyValue, False, 40, 1)
        End With
        Set clsDef = Nothing
End Function

Private Sub Voucher_headCellCheck(Index As Variant, RetValue As String, bChanged As UapVoucherControl85.CheckRet, referPara As ReferParameter)
    Dim pt          As POINTAPI
    Dim hwnd        As Long
    Dim sClsName    As String * 100
    Dim ele As IXMLDOMElement
    Dim sSkeyCode As String
    Dim strsql As String
    Dim rds As New ADODB.Recordset
    Dim strCellCheckType As String
    Dim blnTrue As Boolean
    Dim domTmp As New DOMDocument
    Dim strFieldName As String
    Dim strCellCheck As String
    Call GetCursorPos(pt)
    hwnd = WindowFromPoint(pt.X, pt.Y)
    referPara.bValid = True
    strFieldName = Voucher.ItemState(Index, siheader).sFieldName
    If Not referPara.rstGrid Is Nothing Then
        strReferString = clsVoucherRefer.FillItemsAfterBrowse(clsVoucher, Voucher, "T", strFieldName, referPara.rstGrid)
    Else
        strCellCheck = clsVoucherRefer.CellCheck("", Voucher, "T", strFieldName)
        If strCellCheck <> "" Then
            RetValue = ""
            'bChanged = retry
            Exit Sub
        End If
    End If
    RetValue = Voucher.headerText(Index)
    Select Case LCase(strFieldName)
        Case "numprint" ''印数
            If val(RetValue) < 0 Then
                MsgBox "印数不能小于0", vbInformation
                bChanged = Cancel: Exit Sub
            End If
        Case "iyzprint" ''印张
            If val(RetValue) < 0 Then
                MsgBox "印张不能小于0", vbInformation
                bChanged = Cancel: Exit Sub
            End If
        Case "iprice" ''定价
            If val(RetValue) < 0 Then
                MsgBox "定价不能小于0", vbInformation
                bChanged = Cancel: Exit Sub
            End If
        Case "idiscount" ''折扣额
            If val(RetValue) < 0 Then
                MsgBox "折扣额不能小于0", vbInformation
                bChanged = Cancel: Exit Sub
            End If
        Case "inatdiscount" ''本币折扣额
            If val(RetValue) < 0 Then
                MsgBox "本币折扣额不能小于0", vbInformation
                bChanged = Cancel: Exit Sub
            End If
        Case "inatdiscount" ''本币折扣额
            If val(RetValue) < 0 Then
                MsgBox "本币折扣额不能小于0", vbInformation
                bChanged = Cancel: Exit Sub
            End If
    End Select
'    strCellCheckType = clsVoucher.CellCheck(ctlVoucher1, "T", strFieldName, bChanged, clsVoucherRefer)
'    If strCellCheckType <> "" Then
'        blnTrue = ReferVoucherByInput(strCellCheckType)
'        If Not blnTrue Then
'            bChanged = Cancel
'            domTmp.loadXML strCellCheckType
'            If Not domTmp.documentElement.Attributes.getNamedItem("errresid") Is Nothing Then
'                MsgBox GetString(domTmp.documentElement.Attributes.getNamedItem("errresid").Text), vbExclamation
'            End If
'            Set domTmp = Nothing
'        End If
'    End If
    
    
'///////871旧参照
'    Dim pt          As POINTAPI
'    Dim hwnd        As Long
'    Dim sClsName    As String * 100
'    Dim ele As IXMLDOMElement
'    Dim sSkeyCode As String
'    'If bClickCancel = True Then
'        Call GetCursorPos(pt)
'        hwnd = WindowFromPoint(pt.x, pt.y)
'        If hwnd <> 0 Then
'           GetClassName hwnd, sClsName, 100
'           sClsName = LCase(Trim(sClsName))
'           If sClsName = "msvb_lib_toolbar" Or sClsName = "toolbar20wndclass" Or Trim(sClsName) = "msocommandbar" Then
'                If Not bClickSave Then Exit Sub
'           End If
'        End If
'    'End If
'    Dim lngRow As Long
'    Dim lngCol As Long
'    Dim strAuth As String
''    Dim domHead As New DOMDocument
''    Dim DomBody As New DOMDocument
'    Dim strError As String
'    Dim rstTmp As New ADODB.Recordset
'    Dim strKey As String
'    Dim strRefersql As String
'    Dim tmprst As ADODB.Recordset
'
'    ''是否放弃
'    'If bClickCancel Then Exit Sub
'
'    strKey = LCase(Me.voucher.ItemState(Index, siheader).sFieldName)
'
''    Select Case LCase(strKey)
''        Case "cpersonname"
''            strKey = "cpersoncode"
''        Case "ccusabbname"
''            strKey = "ccuscode"
''        Case "cdepname"
''            strKey = "cdepcode"
''    End Select
'    With Me.voucher
'        If Left(strKey, 7) = "cdefine" Then
'        ''使用新的自定义项目
'            'sRet = RefDefine((.LookUpArray(sKey, Sibody)), Sibody)
'            RetValue = CellCheckDefine(Index, siheader, .headerText(strKey))
'            Exit Sub
'        End If
'
'
'        Select Case LCase(strKey)
'
'            Case "cwlcode"
'                If strVouchType = "95" Or .headerText("cwlcode") = "" Then
'                    Exit Sub
'                End If
'                strRefersql = "select dDate, cWLcode,cDepCode,cDepName,cCusCode,cCusAbbName,cInvCode,cInvName, (iWLAmount-isnull(iHandBackAmount,0)) as iWLAmount, (iWLMoney-isnull(iHandBackMoney,0)) as iWLMoney,cCusInvCode,cCusInvName from SA_WrapLeaseT left join (select cWLCode as THcWLCode,sum(iHandBackAmount) as iHandBackAmount, "
'                strRefersql = strRefersql + "sum(iHandBackMoney) as iHandBackMoney from  SA_WrapLease where bIWLType=0 group by cWLCode ) "
'                strRefersql = strRefersql + " as   SA_WrapLeaseTH on SA_WrapLeaseT.cWLCode=SA_WrapLeaseTH.THcWLCode where bIWLType=1  and ((iWLAmount-isnull(iHandBackAmount,0))<>0 or (iWLMoney-isnull(iHandBackMoney,0))<>0) and cwlcode='" & .headerText("cwlcode") & "'"
'                rstTmp.Open strRefersql, DBConn, adOpenForwardOnly, adLockReadOnly
'                If rstTmp.EOF Then
'                    rstTmp.Close
'                    MsgBox "错误的包装物退回单据号码或者已经全部退回，请检查!"
'                    bChanged = retry
'                    RetValue = ""
'                    Exit Sub
'                Else
'                    RetValue = rstTmp.Fields(Me.voucher.ItemState(Index, siheader).sFieldName)
'                    Me.voucher.headerText("cWLcode") = rstTmp.Fields("cWLcode")
'                    Me.voucher.headerText("cDepCode") = rstTmp.Fields("cDepCode")
'                    Me.voucher.headerText("cDepName") = rstTmp.Fields("cDepName")
'                    Me.voucher.headerText("cCusCode") = rstTmp.Fields("cCusCode")
'                    Me.voucher.headerText("cCusAbbName") = rstTmp.Fields("cCusAbbName")
'                    Me.voucher.headerText("iHandBackAmount") = rstTmp.Fields("iWLAmount")
'                    Me.voucher.headerText("iHandBackMoney") = rstTmp.Fields("iWLMoney")
'                    Me.voucher.headerText("cInvCode") = rstTmp.Fields("cInvCode")
'                    Me.voucher.headerText("cInvName") = rstTmp.Fields("cInvName")
'                    Me.voucher.headerText("cCusInvCode") = rstTmp.Fields("cCusInvCode")
'                    Me.voucher.headerText("cCusInvName") = rstTmp.Fields("cCusInvName")
'                    Me.voucher.headerText("bIWLType") = "0"
'                    Call Voucher_headCellCheck(Me.voucher.LookUpArray("ccusabbname", siheader), "", Cancel, referPara)           ' .ItemState("kl", Sibody))
'                    Me.voucher.headerText("cDepCode") = rstTmp.Fields("cDepCode")
'                    Me.voucher.headerText("cDepName") = rstTmp.Fields("cDepName")
'                    Index = Me.voucher.LookUpArray("cwlcode", siheader)
'                    rstTmp.Close
'                End If
'            Case "cinvcode"
'                Set tmprst = New ADODB.Recordset
'                tmprst.Open "select cinvcode,cInvName,cinvstd  from Inventory where cinvcode='" & Me.voucher.headerText("cinvcode") & "' ", DBConn, adOpenForwardOnly, adLockReadOnly
'                If Not tmprst.EOF Then
'                    Me.voucher.headerText("cinvcode") = ""
'                    Me.voucher.headerText("cinvname") = ""
'                    Me.voucher.headerText("cinvstd") = ""
'                    RetValue = tmprst.Fields(0)
'                    Me.voucher.headerText("cinvcode") = tmprst.Fields(0)
'                    Me.voucher.headerText("cinvname") = tmprst.Fields(1)
'                    Me.voucher.headerText("cinvstd") = tmprst.Fields(2)
'                Else
'                    Me.voucher.headerText("cinvcode") = ""
'                    Me.voucher.headerText("cinvname") = ""
'                    Me.voucher.headerText("cinvstd") = ""
'                    RetValue = ""
'                    MsgBox "图书编码错误，请检查后重试！"
'                    Exit Sub
'                End If
'                tmprst.Close
'                Set tmprst = Nothing
'
'            Case "ivtid"
'                strAuth = clsAuth.getAuthString("DJMB", , "W")
'                'clsRefer.SetRWAuth strAuth
'                If strAuth = "1=2" Then
'                    MsgBox "你没有权限使用任何单据模版！"
'                End If
'                If strAuth <> "" Then
'                    rstTmp.Open "SELECT VT_Name , VT_ID  From VoucherTemplates WHERE (VT_CardNumber = '" & strCardNum & "') AND (VT_TemplateMode = 0) and vt_id =" & .headerText(strKey) & " and vt_id in (" & strAuth & ")", _
'                            DBConn, adOpenForwardOnly, adLockReadOnly
'                    If rstTmp.EOF Then
'                        rstTmp.Close
'                        MsgBox "错误的模版号码或者没有权限，请检查!"
'                        bChanged = retry
'                        RetValue = ""
'                        Exit Sub
'                    End If
'                    rstTmp.Close
'                End If
'
'                Set rstTmp = clsVoucherCO.GetVoucherFormat(Me.voucher.headerText("ivtid"), strCardNum)
'                If rstTmp Is Nothing Or rstTmp.state = 0 Then
'
'                        MsgBox "错误的模版号码，请检查!"
'                Else
'                    If rstTmp.EOF Then
'                        MsgBox "错误的模版号码，请检查!"
'                    Else
'                        Dim tmpVoucherState As Integer
'                        Me.voucher.Visible = False
'                        Me.voucher.getVoucherDataXML Domhead, Dombody
'                        tmpVoucherState = Me.voucher.VoucherStatus
'                        Me.voucher.setTemplateData rstTmp
'                        Me.voucher.VoucherStatus = tmpVoucherState
'                        Me.voucher.setVoucherDataXML Domhead, Dombody
'                        Me.voucher.Visible = True
'                        sCurTemplateID = Me.voucher.headerText("ivtid")
'                        sCurTemplateID2 = Me.voucher.headerText("ivtid")
'                    End If
'                End If
'
'            Case Else
'                If strKey = "coppcode" And strVouchType <> "99" Then
'                    Call DelFreeLine
'                End If
'                If LCase(strKey) = "iexchrate" Or LCase(strKey) = "itaxrate" Or LCase(strKey) = "ccuscode" Or _
'                    LCase(strKey) = "ccusabbname" Or LCase(strKey) = "coppcode" Then
'                    .getVoucherDataXML Domhead, Dombody
'                    If strKey = "coppcode" Then
'                        If .headerText(strKey) <> "" Then
'                            If .BodyRows > 0 And strVouchType <> "99" Then
'                                Set ele = Domhead.selectSingleNode("//z:row")
'
'                                If VBA.MsgBox("如果商机数据包含存货信息，是否删除表体数据？", vbYesNo, "销售管理") = vbYes Then
'                                    '单据控件不支持两次msgbox
'                                    ele.setAttribute "bclearbody", "1"
'                                Else
'                                    ele.setAttribute "bclearbody", "0"
'                                End If
'                            End If
'                        End If
'                    End If
'                Else
'                    Set Domhead = .GetHeadDom
'                End If
'                strError = ""
'                    If strError <> "" Then
'                        MsgBox strError, vbInformation, "销售管理"
'                        If strKey = "coppcode" Then
'                            RetValue = ""
'                            'bChanged = IIf(bLostFocus, Cancel, retry)
'                            If .BodyRows > 0 Then
'                                .setVoucherDataXML Domhead, Dombody
'                                RetValue = ""
'                                bChanged = success
'                            Else
'                                bChanged = retry  'success
'                            End If
'                            Exit Sub
'                        Else
'                            RetValue = ""
'                            bChanged = IIf(bLostFocus, Cancel, retry)
'                        End If
'                        If bChanged = success Then
'                            If Not Domhead Is Nothing Then
'                                If Domhead.selectNodes("//R").length > 0 Then
'                                    For Each ele In Domhead.selectNodes("//R")
'                                        .headerText(ele.getAttribute("K")) = "" 'domHead.documentElement.childNodes.item(lngCol).Attributes.item(1).nodeValue
'                                    Next
'                                End If
'
'                            End If
'                        End If
'                        Exit Sub
'                    End If
'                    On Error Resume Next
'                    If Not Domhead Is Nothing Then
'                        For Each ele In Domhead.selectNodes("//R")
'                            If LCase(strKey) = "cpersonname" Then
'                                If ele.getAttribute("K") = "oricdepcode" And .headerText("cdepcode") <> "" Then
'                                    If ele.getAttribute("V") <> "" Then
'                                        If LCase(ele.getAttribute("V")) <> LCase(.headerText("cdepcode")) Then
'                                            If MsgBox("录入的业务员不属于录入的部门，是否继续？", vbYesNo + vbQuestion) = vbNo Then
'                                                RetValue = ""
'                                                .headerText("cpersonname") = ""
'                                                Exit Sub
'                                            End If
'                                        End If
'                                    End If
'                                End If
'                            End If
'                            .headerText(ele.getAttribute("K")) = ele.getAttribute("V")
'                            If LCase(.ItemState(Index, siheader).sFieldName) = LCase(ele.getAttribute("K")) Then
'                                RetValue = ele.getAttribute("V")
'                                If LCase(strKey) = "cexch_name" Then
'                                    .ItemState("iexchrate", siheader).nNumPoint = clsSAWeb_M.GetExchRateDec(.headerText("cexch_name"))
'                                End If
'                            End If
'                        Next
'                    End If
'                    '如果表头币种改变，需要更改表体汇率
'                    If LCase(strKey) = "coppcode" And (.headerText(strKey) <> "" Or strVouchType = "99") Then
'                        RetValue = GetHeadItemValue(Domhead, "coppcode")
''                        If strVouchType = "99" Then
''                            bChanged = success
''                            Exit Sub
''                        End If
'                        'RetValue = GetHeadItemValue(Domhead, strKey)
'                        .setVoucherDataXML Domhead, Dombody
'                        If GetHeadItemValue(Domhead, "cExch_name") <> myinfo.cCurrencyName Then
'                            SetOriItemState "T", "iExchRate"
'                        Else
'                            .EnableHead "iExchRate", False
'                        End If
'                         .ItemState("iexchrate", siheader).nNumPoint = clsSAWeb_M.GetExchRateDec(.headerText("cexch_name"))
'                         .headerText("iexchrate") = GetHeadItemValue(Domhead, "iexchrate")
'                        Exit Sub
'                    End If
'                    If LCase(strKey) = "cexch_name" Or LCase(strKey) = "iexchrate" Or LCase(strKey) = "ccuscode" Or LCase(strKey) = "ccusabbname" Then
'                        If LCase(.ItemState(Index, siheader).sFieldName) = "cexch_name" Then
'                            Me.voucher.getVoucherDataXML Domhead, Dombody
'                        End If
'                        If Not Dombody Is Nothing Then
'                            If LCase(strKey) = "ccuscode" Or LCase(strKey) = "ccusabbname" Then
'                                Set Domhead = .GetHeadDom
'                            End If
'                            .setVoucherDataXML Domhead, Dombody
''
'                        End If
'                        If RetValue <> myinfo.cCurrencyName Then
'                            SetOriItemState "T", "iExchRate"
'                        Else
'                            .EnableHead "iExchRate", False
'                        End If
'                    End If
'                    On Err GoTo DoERR
'        End Select
'
'    End With
'    Set rstTmp = Nothing
'    Exit Sub
'DoERR:
'    MsgBox Err.Description
End Sub
 


Private Sub fillComBol(bPrint As Boolean, Optional ComboVT As ComboBox, Optional ComboDJ As ComboBox, Optional strCardNum As String, Optional strVouchType As String)
    Dim tmprst As New ADODB.Recordset
    Dim strAuth As String
    Dim strsql As String
    Dim i As Long
    Dim sWhere As String
    
    If bPrint = True Then
        ComboVT.Clear
        
    Else
        ComboDJ.Clear
    End If
    strAuth = clsAuth.getAuthString("DJMB")
    If strAuth = "1=2" Then Exit Sub
    If bFirst = False Then
        If bPrint = True Then
            ''打印
            strsql = "SELECT VT_Name,VT_ID  From VoucherTemplates WHERE (VT_CardNumber = '" & strCardNum & "') AND (VT_TemplateMode = 1) " & IIf(strAuth <> "", "and vt_id in (" & strAuth & ")", "")
        Else
            ''显示
            strsql = "SELECT VT_Name,VT_ID  From VoucherTemplates WHERE (VT_CardNumber = '" & strCardNum & "') AND (VT_TemplateMode = 0) " & IIf(strAuth <> "", "and vt_id in (" & strAuth & ")", "")
        End If
    Else
        Select Case strVouchType
            Case "05"
                sWhere = " VT_CardNumber = '01' or VT_CardNumber = '03' "
            Case "06"
                sWhere = " VT_CardNumber = '05' or VT_CardNumber = '06' "
        End Select
        If bPrint = True Then
            ''打印
            strsql = "SELECT VT_Name,VT_ID  From VoucherTemplates WHERE (" & sWhere & ") AND (VT_TemplateMode = 1) " & IIf(strAuth <> "", "and vt_id in (" & strAuth & ")", "")
        Else
            ''显示
            strsql = "SELECT VT_Name,VT_ID  From VoucherTemplates WHERE ( " & sWhere & ") AND (VT_TemplateMode = 0) " & IIf(strAuth <> "", "and vt_id in (" & strAuth & ")", "")
        End If
    
    End If
    tmprst.CursorLocation = adUseClient
    tmprst.Open strsql, DBConn, adOpenForwardOnly, adLockReadOnly
    If tmprst.EOF Then
        i = 0
        If bPrint = True Then
            ComboVT.Clear
        Else
            ComboDJ.Clear
        End If
    Else
        i = tmprst.RecordCount - 1
        If bPrint Then
            Select Case ComboVT.Name
                Case "ComboVTID"
                    ReDim vtidPrn(i)
                Case "ComboVTID1"
                    ReDim vtidPrn1(i)
                Case "ComboVTID2"
                    ReDim vtidPrn2(i)
                Case "ComboVTID3"
                    ReDim vtidPrn3(i)
                Case "ComboVTID4"
                    ReDim vtidPrn4(i)
            End Select
        
        Else
            Select Case ComboVT.Name
                Case "ComboVTID"
                    ReDim vtidDJMB(i)
                
                Case "ComboVTID1"
                    ReDim vtidDJMB1(i)
                Case "ComboVTID2"
                    ReDim vtidDJMB2(i)
                Case "ComboVTID3"
                    ReDim vtidDJMB3(i)
                Case "ComboVTID4"
                    ReDim vtidDJMB4(i)
            End Select
        End If
    End If
    If Not tmprst.EOF Then
        If bPrint = True Then
            ComboVT.Clear
            i = 0
            Do While Not tmprst.EOF
                ComboVT.AddItem tmprst(0)
'                vtidPrn(i) = CLng(tmprst(1))
                Select Case ComboVT.Name
                    Case "ComboVTID"
                        vtidPrn(i) = CLng(tmprst(1))
                    Case "ComboVTID1"
                        vtidPrn1(i) = CLng(tmprst(1))
                    Case "ComboVTID2"
                        vtidPrn2(i) = CLng(tmprst(1))
                    Case "ComboVTID3"
                        vtidPrn3(i) = CLng(tmprst(1))
                    Case "ComboVTID4"
                        vtidPrn4(i) = CLng(tmprst(1))
                End Select
                i = i + 1
                tmprst.MoveNext
            Loop
            ComboVT.ListIndex = 0
            ComboVT.ToolTipText = ComboVT.Text
        Else
            ComboDJ.Clear
            i = 0
            Do While Not tmprst.EOF
                ComboDJ.AddItem tmprst(0)
                Select Case ComboVT.Name
                    Case "ComboVTID"
                        vtidDJMB(i) = CLng(tmprst(1))
                    Case "ComboVTID1"
                        vtidDJMB1(i) = CLng(tmprst(1))
                    Case "ComboVTID2"
                        vtidDJMB2(i) = CLng(tmprst(1))
                    Case "ComboVTID3"
                        vtidDJMB3(i) = CLng(tmprst(1))
                    Case "ComboVTID4"
                        vtidDJMB4(i) = CLng(tmprst(1))
                End Select
                i = i + 1
                tmprst.MoveNext
            Loop
            ComboDJ.ListIndex = 0
            ComboDJ.ToolTipText = ComboDJ.Text
        End If
    End If
    tmprst.Close
    bfillDjmb = True
    bfillDjmb1 = True
    bfillDjmb2 = True
    bfillDjmb3 = True
    bfillDjmb4 = True
    Set tmprst = Nothing
End Sub

'单据初始化
Public Function ShowVoucher(VoucherType As VoucherType, Optional vVoucherId As Variant, Optional imode As Integer)
    Dim tmpTemplateID As String
    Dim errMsg As String
'by lg070314 增加U870门户融合
    Dim vfd As Object
    sGuid = CreateGUID()
    If Not (g_business Is Nothing) Then
        Set vfd = g_business.CreateFormEnv(sGuid, Me)
    End If
    
    g_FormbillShow = False
    Screen.MousePointer = vbHourglass
    'frmFloat.m_oTimer.Enabled = False
    On Error GoTo DoERR
    If IsMissing(imode) = True Then
        iShowMode = 0
    Else
        iShowMode = imode
    End If
    Set clsVoucherCO = New EFFYVoucherCo.ClsVoucherCO_GDZC_M
    'by ahzzd 2005/05/09 单据初始化
    clsVoucherCO.Init VoucherType, m_Login, DBConn, "CS", clsSAWeb_M
    clsAuth.Init m_Login.UfDbName, m_Login.cUserId
    'sl 2008/02/20 871审批流改变 原有的u8ExamineAndApprove不能用
'    Set obj_EA = CreateObject("u8ExamineAndApprove.clsU8Examine")
'    Call obj_EA.Init(m_login)
    Select Case VoucherType
        Case pbpressinform
            strVouchType = "26"
            strCardNum = "EFYZGL030301"
        Case pbwrappage
            strVouchType = "99"
            strCardNum = "EFYZGL030301"
        Case pbgiveaddress
            strVouchType = "07"
            strCardNum = "EFYZGL030301"
        Case pbcontent
            strVouchType = "95"
            strCardNum = "EFYZGL030301"
        Case pbsheet
            strVouchType = "27"
            strCardNum = "EFYZGL030301"
    End Select
    ''设置按钮
 
   U8VoucherSorter1.Visible = False

 
    sTemplateID = clsSAWeb_M.GetVTID("pbpressinform", DBConn, strCardNum)
    s1TemplateID = clsSAWeb_M.GetVTID("pbwrappage", DBConn, "EFYZGL04")
    s2TemplateID = clsSAWeb_M.GetVTID("pbgiveaddress", DBConn, "EFYZGL05")
    s3TemplateID = clsSAWeb_M.GetVTID("pbcontent", DBConn, "EFYZGL06")
    s4TemplateID = clsSAWeb_M.GetVTID("pbsheet", DBConn, "EFYZGL07")
 
'
'    ''设置按钮
'    sTemplateID = clsSAWeb_M.GetVTID(DBConn, strCardNum)
    If iShowMode <> 2 Then
 
        errMsg = clsVoucherCO.GetVoucherData(Domhead, Dombody, vVoucherId)
 
        If errMsg <> "" Then
            MsgBox errMsg
            If iShowMode = 1 Then
                On Error Resume Next
                bFrmCancel = False
                Set clsVoucherCO = Nothing
                Set clsAuth = Nothing
                Set clsRefer = Nothing
                Set RstTemplate = Nothing
                Set Domhead = Nothing
                Set Dombody = Nothing
                Set DomFormat = Nothing
                Set RstTemplate = Nothing
                Set RstTemplate2 = Nothing
                If m_UFTaskID <> "" Then
                    m_Login.TaskExec m_UFTaskID, 0
                End If
                Screen.MousePointer = vbDefault
                Exit Function
            End If
        End If
 
        If iShowMode = 1 Then
            Call reInit(VoucherType, Domhead)
        End If
        If Not Domhead.selectSingleNode("//z:row") Is Nothing Then
            If Not Domhead.selectSingleNode("//z:row").Attributes.getNamedItem("ivtid") Is Nothing Then
                tmpTemplateID = Domhead.selectSingleNode("//z:row").Attributes.getNamedItem("ivtid").nodeValue
            Else
                tmpTemplateID = "0"
            End If
        Else
            tmpTemplateID = "0"
        End If
 
    Else
        errMsg = clsVoucherCO.GetVoucherData(Domhead, Dombody, 0)
        If errMsg <> "" Then
            MsgBox errMsg
        End If
        Set oDomB = New DOMDocument
        oDomB.loadXML Dombody.xml
    End If
    If tmpTemplateID = "" Or tmpTemplateID = "0" Then
        sCurTemplateID = sTemplateID    ''取默认模板
    Else
        sCurTemplateID = tmpTemplateID  ''新的模板
    End If
    sCurTemplateID2 = sCurTemplateID
    
    If sCurTemplateID = 0 Then
        Me.Hide
        MsgBox "您没有模版使用权限"
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '860sp升级到861修改处   2006/03/12  861 增加附件
    Call SetVoucherDataSource
    

    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    If ChangeTempaltes(sCurTemplateID, True, False, True) = False Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    On Error Resume Next
    If GetHeadItemValue(Domhead, "cexch_name") <> "" Then
        Me.Voucher.ItemState("iexchrate", siheader).nNumPoint = clsSAWeb_M.GetExchRateDec(GetHeadItemValue(Domhead, "cexch_name"))
    End If
    On Error GoTo DoERR
 
    
    Voucher.setVoucherDataXML Domhead, Dombody
    
    If VoucherType = pbpressinform Then
        setVouchDate1 pbwrappage, vVoucherId
        setVouchDate2 pbgiveaddress, vVoucherId
        setVouchDate3 pbcontent, vVoucherId
        setVouchDate4 pbsheet, vVoucherId
    End If
 
    'by ahzzd 2006/05/09   数据准备到单据上完成
    ' sl 在871 上注释 审批流文本
   ' Me.voucher.ExamineFlowAuditInfo = GetEAStream(strVouchType, Domhead, Me.voucher, DBConn)
    
 
    Call SetSum
    Call ChangeCaptionCol
    If Me.Caption = "0" Then
        Me.Caption = Voucher.TitleCaption
    End If

    If iShowMode <> 2 Then
        Voucher.VoucherStatus = VSNormalMode
        Voucher1.VoucherStatus = VSNormalMode
        Voucher2.VoucherStatus = VSNormalMode
        Voucher3.VoucherStatus = VSNormalMode
        Voucher4.VoucherStatus = VSNormalMode
        'ChangeButtonsState
        clsTbl.ChangeButtonState Voucher, Me.tbrvoucher, Me.UFToolbar1, Voucher.VoucherStatus
    Else
    End If
    If iShowMode <> 1 Then
 
        If clsVoucherCO.GetVoucherNO(Domhead, GetvouchNO, errMsg, DomFormat, True) = False Then
            MsgBox errMsg
        Else
 
            Me.Voucher.SetBillNumberRule DomFormat.xml
        End If
        clsRefer.SetLogin m_Login   ''初始化参照控件
    End If
    Me.Voucher.Visible = True
    Me.Voucher1.Visible = True
    Me.Voucher2.Visible = True
    Me.Voucher3.Visible = True
    Me.Voucher4.Visible = True
    Call fillComBol(True, ComboVTID4, ComboDJMB4, s4trCardNum, s4trVouchType)
    Call fillComBol(True, ComboVTID1, ComboDJMB1, s1trCardNum, s1trVouchType)
    Call fillComBol(True, ComboVTID2, ComboDJMB2, s2trCardNum, s2trVouchType)
    Call fillComBol(True, ComboVTID3, ComboDJMB3, s3trCardNum, s3trVouchType)
    Call fillComBol(True, ComboVTID, ComboDJMB, strCardNum, strVouchType)   ''填充模版选择
    If iShowMode <> 1 Then
        Call fillComBol(False, ComboVTID1, ComboDJMB1, s1trCardNum, s1trVouchType)
        Call fillComBol(False, ComboVTID2, ComboDJMB2, s2trCardNum, s2trVouchType)
        Call fillComBol(False, ComboVTID3, ComboDJMB3, s3trCardNum, s3trVouchType)
        Call fillComBol(False, ComboVTID4, ComboDJMB4, s4trCardNum, s4trVouchType)
        Call fillComBol(False, ComboVTID, ComboDJMB, strCardNum, strVouchType)
        bfillDjmb = False
        bfillDjmb1 = False
        bfillDjmb2 = False
        bfillDjmb3 = False
        bfillDjmb4 = False
    End If
    Call SetHelpID
    If Me.Caption = "" Then
        Me.Caption = Me.LabelVoucherName.Caption
    End If
    Dim strXml As String
    strXml = "<?xml version='1.0' encoding='GB2312'?>" & Chr(13)
    domConfig.loadXML strXml & "<EAI>0</EAI>"
    
'by lg070314 增加U870支持，窗体融合
    If g_business Is Nothing Then
        Me.Show
    Else
'        InitToolbarTag Me.tbrvoucher
        Call g_business.ShowForm(Me, "EF", sGuid, False, True, vfd)
        Set Me.Voucher.PortalBusinessObject = g_business
        Me.Voucher.PortalBizGUID = sGuid
    End If
    
    Me.BackColor = Me.Voucher.BackColor
    Me.picVoucher.BackColor = Me.Voucher.BackColor
    Me.Refresh
    Me.Voucher.SetFocus
    Screen.MousePointer = vbDefault
    g_FormbillShow = True
    U8VoucherSorter1.Visible = False
 
 
    Exit Function
DoERR:
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
End Function
Private Sub setVouchDate1(VoucherType As VoucherType, Optional vVoucherId As Variant, Optional imode As Integer)
    '//
    Dim errMsg As String
    Dim tmpTemplateID As String
    Set clsVoucherCO1 = Nothing
    Set clsVoucherCO1 = New EFFYVoucherCo.ClsVoucherCO_GDZC_M
    clsVoucherCO1.Init VoucherType, m_Login, DBConn, "CS", clsSAWeb_M              ', m_Conn
    
    Select Case VoucherType
        Case pbwrappage     'sl 付印通知单—封皮
            s1trVouchType = "99"
            s1trCardNum = "EFYZGL030301"
    End Select
    ''设置按钮
    
    If s1trVouchType = "95" Or s1trVouchType = "92" Then
        U8VoucherSorter2.Visible = False
    End If
    s1TemplateID = clsSAWeb_M.GetVTID("pbwrappage", DBConn, s1trCardNum)
    
    If iShowMode <> 2 Then
        errMsg = clsVoucherCO1.GetVoucherData(Domhead1, Dombody1, vVoucherId)
        ''读取单据上的模板
        If errMsg <> "" Then
            MsgBox errMsg
            If iShowMode = 1 Then
                'Call Form_Unload(0)
                On Error Resume Next
                
                bFrmCancel = False
                Set clsVoucherCO1 = Nothing
                Set clsAuth = Nothing
                Set clsRefer = Nothing
                Set RstTemplate = Nothing
                Set Domhead1 = Nothing
                Set Dombody1 = Nothing
                Set DomFormat = Nothing
                Set RstTemplate = Nothing
                Set RstTemplate2 = Nothing
                If m_UFTaskID <> "" Then
                    m_Login.TaskExec m_UFTaskID, 0
                End If
                Screen.MousePointer = vbDefault
                
                Exit Sub
            End If
        End If
        
        If iShowMode = 1 Then
            Call reInit(VoucherType, Domhead1)
        End If
        
'        sKey = LCase(sKey)
'        If Not Dombody.selectNodes("//z:row").item(r).Attributes.getNamedItem(sKey) Is Nothing Then
'            GetBodyItemValue = Dombody.selectNodes("//z:row").item(r).Attributes.getNamedItem(sKey).nodeValue
'        Else
'            GetBodyItemValue = ""
'        End If
        If Dombody1.selectNodes("//z:row").length > 0 Then
            If Not Dombody1.selectNodes("//z:row").Item(0).Attributes.getNamedItem("ivtid") Is Nothing Then
                tmpTemplateID = Dombody1.selectNodes("//z:row").Item(0).Attributes.getNamedItem("ivtid").nodeValue
            Else
                tmpTemplateID = "0"
            End If
        Else
            tmpTemplateID = "0"
        End If
    Else
        errMsg = clsVoucherCO1.GetVoucherData(Domhead1, Dombody1, 0)
        If errMsg <> "" Then
            MsgBox errMsg
        End If
        Set oDomB = New DOMDocument
        oDomB.loadXML Dombody1.xml
    End If
    If tmpTemplateID = "" Or tmpTemplateID = "0" Then
        s1CurTemplateID = s1TemplateID    ''取默认模板
    Else
        s1CurTemplateID = tmpTemplateID  ''新的模板
    End If
    s1CurTemplateID2 = s1CurTemplateID
    
    If s1CurTemplateID = 0 Then
        Me.Hide
        MsgBox "您没有模版使用权限"
        Screen.MousePointer = vbDefault
        
        Exit Sub
    End If
'    If ChangeTempaltes(sCurTemplateID, True, False, True) = False Then
'        Screen.MousePointer = vbDefault
'
'        Exit Sub
'    End If
    Dim rstmp As Object
    Set rstmp = clsVoucherCO1.GetVoucherFormat(s1CurTemplateID, s1trCardNum)
        
    '如果是调试状态，不处理附件，以放置弹出‘加载附件失败’窗口
    If clsSAWeb.IsDebug Then rstmp.Fields("vchtblprimarykeynames") = ""
    
    Voucher1.setTemplateData rstmp
    On Error Resume Next
    If GetHeadItemValue(Domhead1, "cexch_name") <> "" Then

            'Me.Voucher1.ItemState("iexchrate", siheader).nNumPoint = clsSAWeb_M.GetExchRateDec(GetHeadItemValue(Domhead1, "cexch_name"))

    End If
    On Error GoTo DoERR
    Voucher1.setVoucherDataXML Domhead1, Dombody1
DoERR:
End Sub

Private Sub setVouchDate2(VoucherType As VoucherType, Optional vVoucherId As Variant, Optional imode As Integer)
    Dim errMsg As String
    Dim tmpTemplateID As String
    Set clsVoucherCO2 = Nothing
    Set clsVoucherCO2 = New EFFYVoucherCo.ClsVoucherCO_GDZC_M
    clsVoucherCO2.Init VoucherType, m_Login, DBConn, "CS", clsSAWeb_M              ', m_Conn
    Select Case VoucherType
        Case pbgiveaddress
            s2trVouchType = "07"
            s2trCardNum = "EFYZGL030301"
    End Select
    ''设置按钮
    
    If s2trVouchType = "95" Or s2trVouchType = "92" Then
        U8VoucherSorter3.Visible = False
    End If
    s2TemplateID = clsSAWeb_M.GetVTID("pbgiveaddress", DBConn, s2trCardNum)
    
    If iShowMode <> 2 Then
        errMsg = clsVoucherCO2.GetVoucherData(Domhead2, Dombody2, vVoucherId)
        ''读取单据上的模板
        If errMsg <> "" Then
            MsgBox errMsg
            If iShowMode = 1 Then
                'Call Form_Unload(0)
                On Error Resume Next
                
                bFrmCancel = False
                Set clsVoucherCO2 = Nothing
                Set clsAuth = Nothing
                Set clsRefer = Nothing
                Set RstTemplate = Nothing
                Set Domhead2 = Nothing
                Set Dombody2 = Nothing
                Set DomFormat = Nothing
                Set RstTemplate = Nothing
                Set RstTemplate2 = Nothing
                If m_UFTaskID <> "" Then
                    m_Login.TaskExec m_UFTaskID, 0
                End If
                Screen.MousePointer = vbDefault
                
                Exit Sub
            End If
        End If
        
        If iShowMode = 1 Then
            Call reInit(VoucherType, Domhead1)
        End If
        
'        sKey = LCase(sKey)
'        If Not Dombody.selectNodes("//z:row").item(r).Attributes.getNamedItem(sKey) Is Nothing Then
'            GetBodyItemValue = Dombody.selectNodes("//z:row").item(r).Attributes.getNamedItem(sKey).nodeValue
'        Else
'            GetBodyItemValue = ""
'        End If
        If Dombody2.selectNodes("//z:row").length > 0 Then
            If Not Dombody2.selectNodes("//z:row").Item(0).Attributes.getNamedItem("ivtid") Is Nothing Then
                tmpTemplateID = Dombody2.selectNodes("//z:row").Item(0).Attributes.getNamedItem("ivtid").nodeValue
            Else
                tmpTemplateID = "0"
            End If
        Else
            tmpTemplateID = "0"
        End If
    Else
        errMsg = clsVoucherCO2.GetVoucherData(Domhead2, Dombody2, 0)
        If errMsg <> "" Then
            MsgBox errMsg
        End If
        Set oDomB = New DOMDocument
        oDomB.loadXML Dombody2.xml
    End If
    If tmpTemplateID = "" Or tmpTemplateID = "0" Then
        s2CurTemplateID = s2TemplateID    ''取默认模板
    Else
        s2CurTemplateID = tmpTemplateID  ''新的模板
    End If
    s2CurTemplateID2 = s2CurTemplateID
    
    If s2CurTemplateID = 0 Then
        Me.Hide
        MsgBox "您没有模版使用权限"
        Screen.MousePointer = vbDefault
        
        Exit Sub
    End If
'    If ChangeTempaltes(sCurTemplateID, True, False, True) = False Then
'        Screen.MousePointer = vbDefault
'
'        Exit Sub
'    End If
    Dim rstmp As Object
    Set rstmp = clsVoucherCO2.GetVoucherFormat(s2CurTemplateID, s2trCardNum)
        
    '如果是调试状态，不处理附件，以放置弹出‘加载附件失败’窗口
    If clsSAWeb.IsDebug Then rstmp.Fields("vchtblprimarykeynames") = ""
    
    Voucher2.setTemplateData rstmp
    On Error Resume Next
    If GetHeadItemValue(Domhead2, "cexch_name") <> "" Then

'            Me.Voucher2.ItemState("cexchrate", siheader).nNumPoint = clsSAWeb_M.GetExchRateDec(GetHeadItemValue(Domhead2, "cexch_name"))

    End If
    On Error GoTo DoERR
    Voucher2.setVoucherDataXML Domhead2, Dombody2
DoERR:
End Sub
Private Sub setVouchDate3(VoucherType As VoucherType, Optional vVoucherId As Variant, Optional imode As Integer)
    Dim errMsg As String
    Dim tmpTemplateID As String
    Set clsVoucherCO3 = Nothing
    Set clsVoucherCO3 = New EFFYVoucherCo.ClsVoucherCO_GDZC_M
    clsVoucherCO3.Init VoucherType, m_Login, DBConn, "CS", clsSAWeb_M              ', m_Conn
    Select Case VoucherType
  
        Case pbcontent           'sl 付印通知单－内容及印装方法
            s3trVouchType = "95"
            s3trCardNum = "EFYZGL030301"
        
    End Select
    ''设置按钮
    
    If s3trVouchType = "95" Or s3trVouchType = "92" Then
        U8VoucherSorter3.Visible = False
    End If
    s3TemplateID = clsSAWeb_M.GetVTID("pbcontent", DBConn, s3trCardNum)
    
    If iShowMode <> 2 Then
        errMsg = clsVoucherCO3.GetVoucherData(Domhead3, Dombody3, vVoucherId)
        ''读取单据上的模板
        If errMsg <> "" Then
            MsgBox errMsg
            If iShowMode = 1 Then
                'Call Form_Unload(0)
                On Error Resume Next
                
                bFrmCancel = False
                Set clsVoucherCO1 = Nothing
                Set clsAuth = Nothing
                Set clsRefer = Nothing
                Set RstTemplate = Nothing
                Set Domhead3 = Nothing
                Set Dombody3 = Nothing
                Set DomFormat = Nothing
                Set RstTemplate = Nothing
                Set RstTemplate2 = Nothing
                If m_UFTaskID <> "" Then
                    m_Login.TaskExec m_UFTaskID, 0
                End If
                Screen.MousePointer = vbDefault
                
                Exit Sub
            End If
        End If
        
        If iShowMode = 1 Then
            Call reInit(VoucherType, Domhead3)
        End If
        
'        sKey = LCase(sKey)
'        If Not Dombody.selectNodes("//z:row").item(r).Attributes.getNamedItem(sKey) Is Nothing Then
'            GetBodyItemValue = Dombody.selectNodes("//z:row").item(r).Attributes.getNamedItem(sKey).nodeValue
'        Else
'            GetBodyItemValue = ""
'        End If
        If Dombody3.selectNodes("//z:row").length > 0 Then
            If Not Dombody3.selectNodes("//z:row").Item(0).Attributes.getNamedItem("ivtid") Is Nothing Then
                tmpTemplateID = Dombody3.selectNodes("//z:row").Item(0).Attributes.getNamedItem("ivtid").nodeValue
            Else
                tmpTemplateID = "0"
            End If
        Else
            tmpTemplateID = "0"
        End If
    Else
        errMsg = clsVoucherCO3.GetVoucherData(Domhead3, Dombody3, 0)
        If errMsg <> "" Then
            MsgBox errMsg
        End If
        Set oDomB = New DOMDocument
        oDomB.loadXML Dombody1.xml
    End If
    If tmpTemplateID = "" Or tmpTemplateID = "0" Then
        s3CurTemplateID = s3TemplateID    ''取默认模板
    Else
        s3CurTemplateID = tmpTemplateID  ''新的模板
    End If
    s3CurTemplateID2 = s3CurTemplateID
    
    If s3CurTemplateID = 0 Then
        Me.Hide
        MsgBox "您没有模版使用权限"
        Screen.MousePointer = vbDefault
        
        Exit Sub
    End If
    Dim rstmp As Object
    Set rstmp = clsVoucherCO3.GetVoucherFormat(s3CurTemplateID, s3trCardNum)
        
    '如果是调试状态，不处理附件，以放置弹出‘加载附件失败’窗口
    If clsSAWeb.IsDebug Then rstmp.Fields("vchtblprimarykeynames") = ""
    
    Voucher3.setTemplateData rstmp
    On Error Resume Next
    If GetHeadItemValue(Domhead3, "cexch_name") <> "" Then

'            Me.Voucher3.ItemState("cexchrate", siheader).nNumPoint = clsSAWeb_M.GetExchRateDec(GetHeadItemValue(Domhead3, "cexch_name"))

    End If
    On Error GoTo DoERR
    Voucher3.setVoucherDataXML Domhead3, Dombody3
DoERR:
End Sub
Private Sub setVouchDate4(VoucherType As VoucherType, Optional vVoucherId As Variant, Optional imode As Integer)
    Dim errMsg As String
    Dim tmpTemplateID As String
    Set clsVoucherCO4 = Nothing
    Set clsVoucherCO4 = New EFFYVoucherCo.ClsVoucherCO_GDZC_M
    clsVoucherCO4.Init VoucherType, m_Login, DBConn, "CS", clsSAWeb_M              ', m_Conn
    Select Case VoucherType
            
        Case pbsheet ' sl 付印通知单－纸张
            s4trVouchType = "27"
            s4trCardNum = "EFYZGL030301"
        
    End Select
    ''设置按钮
    
    If s4trVouchType = "95" Or s4trVouchType = "92" Then
        U8VoucherSorter5.Visible = False
    End If
    s4TemplateID = clsSAWeb_M.GetVTID("pbsheet", DBConn, s4trCardNum)
    
    If iShowMode <> 2 Then
        errMsg = clsVoucherCO4.GetVoucherData(Domhead4, Dombody4, vVoucherId)
        ''读取单据上的模板
        If errMsg <> "" Then
            MsgBox errMsg
            If iShowMode = 1 Then
                'Call Form_Unload(0)
                On Error Resume Next
                
                bFrmCancel = False
                Set clsVoucherCO4 = Nothing
                Set clsAuth = Nothing
                Set clsRefer = Nothing
                Set RstTemplate = Nothing
                Set Domhead4 = Nothing
                Set Dombody4 = Nothing
                Set DomFormat = Nothing
                Set RstTemplate = Nothing
                Set RstTemplate2 = Nothing
                If m_UFTaskID <> "" Then
                    m_Login.TaskExec m_UFTaskID, 0
                End If
                Screen.MousePointer = vbDefault
                
                Exit Sub
            End If
        End If
        
        If iShowMode = 1 Then
            Call reInit(VoucherType, Domhead1)
        End If
        
'        sKey = LCase(sKey)
'        If Not Dombody.selectNodes("//z:row").item(r).Attributes.getNamedItem(sKey) Is Nothing Then
'            GetBodyItemValue = Dombody.selectNodes("//z:row").item(r).Attributes.getNamedItem(sKey).nodeValue
'        Else
'            GetBodyItemValue = ""
'        End If
        If Dombody4.selectNodes("//z:row").length > 0 Then
            If Not Dombody4.selectNodes("//z:row").Item(0).Attributes.getNamedItem("ivtid") Is Nothing Then
                tmpTemplateID = Dombody4.selectNodes("//z:row").Item(0).Attributes.getNamedItem("ivtid").nodeValue
            Else
                tmpTemplateID = "0"
            End If
        Else
            tmpTemplateID = "0"
        End If
    Else
        errMsg = clsVoucherCO4.GetVoucherData(Domhead4, Dombody4, 0)
        If errMsg <> "" Then
            MsgBox errMsg
        End If
        Set oDomB = New DOMDocument
        oDomB.loadXML Dombody1.xml
    End If
    If tmpTemplateID = "" Or tmpTemplateID = "0" Then
        s4CurTemplateID = s4TemplateID    ''取默认模板
    Else
        s4CurTemplateID = tmpTemplateID  ''新的模板
    End If
    s4CurTemplateID2 = s4CurTemplateID
    
    If s4CurTemplateID = 0 Then
        Me.Hide
        MsgBox "您没有模版使用权限"
        Screen.MousePointer = vbDefault
        
        Exit Sub
    End If
'    If ChangeTempaltes(sCurTemplateID, True, False, True) = False Then
'        Screen.MousePointer = vbDefault
'
'        Exit Sub
'    End If
    Dim rstmp As Object
    Set rstmp = clsVoucherCO4.GetVoucherFormat(s4CurTemplateID, s4trCardNum)
        
    '如果是调试状态，不处理附件，以放置弹出‘加载附件失败’窗口
    If clsSAWeb.IsDebug Then rstmp.Fields("vchtblprimarykeynames") = ""
    
    Voucher4.setTemplateData rstmp
    On Error Resume Next
    If GetHeadItemValue(Domhead4, "cexch_name") <> "" Then

'            Me.Voucher4.ItemState("cexchrate", siheader).nNumPoint = clsSAWeb_M.GetExchRateDec(GetHeadItemValue(Domhead4, "cexch_name"))

    End If
    On Error GoTo DoERR
    Voucher4.setVoucherDataXML Domhead4, Dombody4
DoERR:
End Sub
Private Sub setVouchDate5(Optional VoucherType As VoucherType, Optional vVoucherId As Variant, Optional imode As Integer)
    '//
    Dim errMsg As String
    Dim tmpTemplateID As String
    Set clsVoucherCO5 = Nothing
    Set clsVoucherCO5 = New EFFYVoucherCo.ClsVoucherCO_GDZC_M
    clsVoucherCO5.Init VoucherType, m_Login, DBConn, "CS", clsSAWeb_M              ', m_Conn
    s5trCardNum = "EFYZGL030301"

    s5TemplateID = clsSAWeb_M.GetVTID("pbprint", DBConn, s5trCardNum)
    
    If iShowMode <> 2 Then
        errMsg = clsVoucherCO5.GetVoucherData(Domhead5, Dombody5, vVoucherId)
        ''读取单据上的模板
        If errMsg <> "" Then
            MsgBox errMsg
            If iShowMode = 1 Then
                'Call Form_Unload(0)
                On Error Resume Next
                
                bFrmCancel = False
                Set clsVoucherCO5 = Nothing
                Set clsAuth = Nothing
                Set clsRefer = Nothing
                Set RstTemplate = Nothing
                Set Domhead5 = Nothing
                Set Dombody5 = Nothing
                Set DomFormat = Nothing
                Set RstTemplate = Nothing
                Set RstTemplate2 = Nothing
                If m_UFTaskID <> "" Then
                    m_Login.TaskExec m_UFTaskID, 0
                End If
                Screen.MousePointer = vbDefault
                
                Exit Sub
            End If
        End If
        
        If iShowMode = 1 Then
            Call reInit(VoucherType, Domhead5)
        End If
        
'        sKey = LCase(sKey)
'        If Not Dombody.selectNodes("//z:row").item(r).Attributes.getNamedItem(sKey) Is Nothing Then
'            GetBodyItemValue = Dombody.selectNodes("//z:row").item(r).Attributes.getNamedItem(sKey).nodeValue
'        Else
'            GetBodyItemValue = ""
'        End If
'        If Dombody5.selectNodes("//z:row").length > 0 Then
'            If Not Dombody5.selectNodes("//z:row").Item(0).Attributes.getNamedItem("ivtid") Is Nothing Then
'                tmpTemplateID = Dombody5.selectNodes("//z:row").Item(0).Attributes.getNamedItem("ivtid").nodeValue
'            Else
'                tmpTemplateID = "0"
'            End If
'        Else
'            tmpTemplateID = "0"
'        End If
        tmpTemplateID = "" ' "60014"
    Else
        errMsg = clsVoucherCO5.GetVoucherData(Domhead5, Dombody5, 0)
        If errMsg <> "" Then
            MsgBox errMsg
        End If
        Set oDomB = New DOMDocument
        oDomB.loadXML Dombody5.xml
    End If
    
    If tmpTemplateID = "" Or tmpTemplateID = "0" Then
        s5CurTemplateID = s5TemplateID    ''取默认模板
    Else
        s5CurTemplateID = tmpTemplateID  ''新的模板
    End If
    s5CurTemplateID2 = s5CurTemplateID
    
    If s5CurTemplateID = 0 Then
        Me.Hide
        MsgBox "您没有模版使用权限"
        Screen.MousePointer = vbDefault
        
        Exit Sub
    End If
'    If ChangeTempaltes(sCurTemplateID, True, False, True) = False Then
'        Screen.MousePointer = vbDefault
'
'        Exit Sub
'    End If
    Dim rstmp As Object
    Set rstmp = clsVoucherCO5.GetVoucherFormat(s5CurTemplateID, s5trCardNum)
        
    '如果是调试状态，不处理附件，以放置弹出‘加载附件失败’窗口
    If clsSAWeb.IsDebug Then rstmp.Fields("vchtblprimarykeynames") = ""
    
    Voucher5.setTemplateData rstmp
    On Error Resume Next
    If GetHeadItemValue(Domhead5, "cexch_name") <> "" Then

            Me.Voucher5.ItemState("cexchrate", siheader).nNumPoint = clsSAWeb_M.GetExchRateDec(GetHeadItemValue(Domhead1, "cexch_name"))
    End If
    On Error GoTo DoERR
    Voucher5.setVoucherDataXML Domhead5, Dombody5
DoERR:
End Sub






''设置单据号是否可以编辑
Private Sub SetVouchNoWriteble()
    Dim KeyCode As String
    
    On Error Resume Next
    If strVouchType = "92" Then Exit Sub
    KeyCode = getVoucherCodeName()
    If Not DomFormat Is Nothing Then
        If DomFormat.xml <> "" Then
            If LCase(DomFormat.selectSingleNode("//单据编号").Attributes.getNamedItem("允许手工修改").nodeValue) = "false" And LCase(DomFormat.selectSingleNode("//单据编号").Attributes.getNamedItem("重号自动重取").nodeValue) = "false" Then
                Me.Voucher.EnableHead KeyCode, False
            Else
                Me.Voucher.EnableHead KeyCode, True
            End If
        End If
    End If
End Sub
 
Private Sub SetButton()
    Dim Index As Integer
    Dim btnX As MSComctlLib.Button
    On Error Resume Next
    Set Me.Icon = frmMain.Icon
    With tbrvoucher
        Set .ImageList = frmMain.imgBmp
        
            .buttons.Clear
         ''增加按钮
'by lg070314 修改U870门户菜单融合，所有Toolbar的Button增加Tag值
'Tag值 表示菜单上的图标文件名称   图标文件在 ..\U8SOFT\icons
        
            ''打印
            Set btnX = .buttons.Add(, "Print", strPrint, tbrDefault)
            btnX.ToolTipText = strPrint
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Print"
        
             ''预览
            Set btnX = .buttons.Add(, "Preview", strPreview, tbrDefault)
            btnX.ToolTipText = strPreview
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "print preview"
 
            ''增加
            Set btnX = .buttons.Add(, "Add", strAdd, tbrDefault)
            btnX.ToolTipText = strAdd
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Add"
            
            ''修改
            Set btnX = .buttons.Add(, "Modify", strModify, tbrDefault)
            btnX.ToolTipText = strModify
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "modify"
            
            
            ''删除
            Set btnX = .buttons.Add(, "Erase", strDelete, tbrDefault)
            btnX.ToolTipText = strDelete
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "delete"
             
            ''复制
            Set btnX = .buttons.Add(, "Copy", "复制", tbrDefault)
            btnX.ToolTipText = "复制"
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Copy"
                        
            ''保存
            Set btnX = .buttons.Add(, "Save", strSave, tbrDefault)
            btnX.ToolTipText = strSave
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "save"
            
            ''放弃
            Set btnX = .buttons.Add(, "Cancel", strDiscard, tbrDefault)
            btnX.ToolTipText = strDiscard
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Cancel"
        
        
            ''增行
            Set btnX = .buttons.Add(, "AddRow", strAddrecord, tbrDefault)
            btnX.ToolTipText = strAddrecord
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "add a row"
      
            ''删行
            Set btnX = .buttons.Add(, "DelRow", strDeleterecord, tbrDefault)
            btnX.ToolTipText = strDeleterecord
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Delete a row"
            
            '审核
            Set btnX = .buttons.Add(, "Sure", "审核", tbrDefault)
            btnX.ToolTipText = "审核"
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Approve"
            
            '弃审
            Set btnX = .buttons.Add(, "UnSure", "弃审", tbrDefault)
            btnX.ToolTipText = "弃审"
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Unapprove"
            

            '关闭
            Set btnX = .buttons.Add(, "CloseOrder", strClose, tbrDefault)
            btnX.ToolTipText = strClose
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "close"
            
            '打开
            Set btnX = .buttons.Add(, "OpenOrder", strOpen, tbrDefault)
            btnX.ToolTipText = strOpen
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "open"
            
            '转入
            Set btnX = .buttons.Add(, "ShiftTo", "转入", tbrDefault)
            btnX.ToolTipText = "转入"
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "importdir_wiz"

            '过滤
            Set btnX = .buttons.Add(, "Filter", strFilter, tbrDefault)
            btnX.ToolTipText = strFilter
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Filter"
                     
           
            
            ''首张
            Set btnX = .buttons.Add(, "ToFirst", strFirst, tbrDefault)
            btnX.ToolTipText = strFirst
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "first page"
            ''上张
            Set btnX = .buttons.Add(, "ToPrevious", strPrevious, tbrDefault)
            btnX.ToolTipText = strPrevious
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "previous page"
            ''下张
            Set btnX = .buttons.Add(, "ToNext", strNext, tbrDefault)
            btnX.ToolTipText = strNext
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "next page"
           ''末张
            Set btnX = .buttons.Add(, "ToLast", strLast, tbrDefault)
            btnX.ToolTipText = strLast
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Last page"

           ''刷新
            Set btnX = .buttons.Add(, "Paint", strRefresh, tbrDefault)
            btnX.ToolTipText = strRefresh
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "refresh"
            
            '' 帮助
            Set btnX = .buttons.Add(, "Help", "帮助", tbrDefault)
            btnX.ToolTipText = "帮助"
            btnX.Description = btnX.ToolTipText
            btnX.Tag = "Help"
                 
           
'           '退出
'            Set btnX = .buttons.Add(, "Exit", strExit, tbrDefault)
'            btnX.image = 1118
'            btnX.ToolTipText = strExit
'            btnX.Description = btnX.ToolTipText
'            btnX.Tag = "Exit"
            
    End With
    ''  置作废、现结的位置
    labZF.Top = picVoucher.Top 'Me.top - Me.tbrvoucher.top
    labZF.Left = Me.Voucher.Left
    labXJ.Top = picVoucher.Top ' Me.top - Me.tbrvoucher.top    'Me.StBar.height
    labXJ.Left = Me.Voucher.Left + labZF.Width
    'by lg070316增加初始化U870菜单
    Call InitToolbarTag(Me.tbrvoucher)
    
End Sub
''改变button的状态
Private Sub ChangeButtonsState()
    Dim i As Integer
    Dim strsql As String
    Dim rds As New ADODB.Recordset
    Exit Sub
    On Error Resume Next
    Me.labXJ.Visible = False
    Me.labZF.Visible = False
    With Me.Voucher
        If .headerText("ddate") <> "" Then
            Me.tbrvoucher.buttons("Copy").Enabled = True
            Me.tbrvoucher.buttons("Sure").Enabled = True
            Me.tbrvoucher.buttons("UnSure").Enabled = True
            Select Case strVouchType
               Case "26"  '付印通知单单
                    If .headerText("cCloser") = "" Then
                            Me.tbrvoucher.buttons("OpenOrder").Enabled = False
                            Me.tbrvoucher.buttons("CloseOrder").Enabled = True
                            Me.tbrvoucher.buttons("Modify").Enabled = True
                            Me.tbrvoucher.buttons("Erase").Enabled = True
                            Me.tbrvoucher.buttons("Sure").Enabled = True
                            Me.tbrvoucher.buttons("UnSure").Enabled = True
                    Else
                            Me.tbrvoucher.buttons("OpenOrder").Enabled = True
                            Me.tbrvoucher.buttons("CloseOrder").Enabled = False
                            Me.tbrvoucher.buttons("Modify").Enabled = False
                            Me.tbrvoucher.buttons("Erase").Enabled = False
                            Me.tbrvoucher.buttons("Sure").Enabled = False
                            Me.tbrvoucher.buttons("UnSure").Enabled = False
                    End If
                    If .headerText("cVerifier") <> "" Then
                        Me.tbrvoucher.buttons("OpenOrder").Visible = True
                        Me.tbrvoucher.buttons("CloseOrder").Visible = True
                        Me.tbrvoucher.buttons("Sure").Enabled = False
                        bCheckVouch = False
                        Me.tbrvoucher.buttons("Modify").Enabled = False
                        Me.tbrvoucher.buttons("Erase").Enabled = False
                        Me.tbrvoucher.buttons("lock").Enabled = False
                        Me.tbrvoucher.buttons("ShiftTo").Enabled = True
                       '判断有没有生成bom
                        strsql = " select * from bas_part where partid in (select ParentId from bom_parent) " & " and  invcode='" & Me.Voucher.headerText("cinvcode") & "'" & _
                        "   And (IsNull(bas_part.Free1, N'') = IsNull('" & Me.Voucher.headerText("cfree1") & "', N'') Or IsNull(bas_part.Free1, N'') = N'')" & _
                        "   And (IsNull(bas_part.Free2, N'') = IsNull('" & Me.Voucher.headerText("cfree2") & "', N'') Or IsNull(bas_part.Free2, N'') = N'')" & _
                        "   And (IsNull(bas_part.Free3, N'') = IsNull('" & Me.Voucher.headerText("cfree3") & "', N'') Or IsNull(bas_part.Free3, N'') = N'')  " & _
                        "   And (IsNull(bas_part.Free4, N'') = IsNull('" & Me.Voucher.headerText("cfree4") & "', N'') Or IsNull(bas_part.Free4, N'') = N'') " & _
                        "   And (IsNull(bas_part.Free5, N'') = IsNull('" & Me.Voucher.headerText("cfree5") & "', N'') Or IsNull(bas_part.Free5, N'') = N'') " & _
                        "   And (IsNull(bas_part.Free6, N'') = IsNull('" & Me.Voucher.headerText("cfree6") & "', N'') Or IsNull(bas_part.Free6, N'') = N'') " & _
                        "   And (IsNull(bas_part.Free7, N'') = IsNull('" & Me.Voucher.headerText("cfree7") & "', N'') Or IsNull(bas_part.Free7, N'') = N'') " & _
                        "   And (IsNull(bas_part.Free8, N'') = IsNull('" & Me.Voucher.headerText("cfree8") & "', N'') Or IsNull(bas_part.Free8, N'') = N'') " & _
                        "   And (IsNull(bas_part.Free9, N'') = IsNull('" & Me.Voucher.headerText("cfree9") & "', N'') Or IsNull(bas_part.Free9, N'') = N'') " & _
                        "   And (IsNull(bas_part.Free10, N'') = IsNull('" & Me.Voucher.headerText("cfree10") & "', N'') Or IsNull(bas_part.Free10, N'') = N'') "

                        Set rds = DBConn.Execute(strsql)
                          If Not rds.EOF Then
                                  Me.tbrvoucher.buttons("ShiftTo").Enabled = False
                                  Me.tbrvoucher.buttons("UnSure").Enabled = False
                          End If
                    Else
                        bCheckVouch = True
                        Me.tbrvoucher.buttons("OpenOrder").Visible = False
                        Me.tbrvoucher.buttons("CloseOrder").Visible = False
                        Me.tbrvoucher.buttons("ShiftTo").Enabled = False
                        If .headerText("cCloser") = "" Then
                            Me.tbrvoucher.buttons("Sure").Enabled = True
                            Me.tbrvoucher.buttons("Modify").Enabled = True
                            Me.tbrvoucher.buttons("Erase").Enabled = True
                        Else
                            Me.tbrvoucher.buttons("Sure").Enabled = False
                            Me.tbrvoucher.buttons("Modify").Enabled = False
                            Me.tbrvoucher.buttons("Erase").Enabled = False
                        End If
                    End If
                
                End Select

            For i = 1 To Me.tbrvoucher.buttons.Count
                If Left(Me.tbrvoucher.buttons(i).ToolTipText, 2) <> "参照" And Left(Me.tbrvoucher.buttons(i).ToolTipText, 2) <> "查询" Then
                    Me.tbrvoucher.buttons(i).Caption = Left(Me.tbrvoucher.buttons(i).ToolTipText, 2)
                End If
            Next

        Else     ''空单据
            'Me.tbrvoucher.buttons("Seek").Visible = False
            Me.tbrvoucher.buttons("Erase").Visible = False
            Me.tbrvoucher.buttons("Modify").Visible = False
            Me.tbrvoucher.buttons("Save").Visible = True
            Me.tbrvoucher.buttons("Cancel").Visible = False
            Me.tbrvoucher.buttons("Sure").Visible = False
            Me.tbrvoucher.buttons("UnSure").Visible = False
 
        End If
    End With
    
    If clsVoucherCO.BOF And clsVoucherCO.EOF Then
        Me.tbrvoucher.buttons("ToFirst").Enabled = False
        Me.tbrvoucher.buttons("ToPrevious").Enabled = False
        Me.tbrvoucher.buttons("ToNext").Enabled = False
        Me.tbrvoucher.buttons("ToLast").Enabled = False
    ElseIf clsVoucherCO.BOF Then
        Me.tbrvoucher.buttons("ToFirst").Enabled = False
        Me.tbrvoucher.buttons("ToPrevious").Enabled = False
        Me.tbrvoucher.buttons("ToNext").Enabled = True
        Me.tbrvoucher.buttons("ToLast").Enabled = True
    ElseIf clsVoucherCO.EOF Then
        Me.tbrvoucher.buttons("ToFirst").Enabled = True
        Me.tbrvoucher.buttons("ToPrevious").Enabled = True
        Me.tbrvoucher.buttons("ToNext").Enabled = False
        Me.tbrvoucher.buttons("ToLast").Enabled = False
    Else
        Me.tbrvoucher.buttons("ToFirst").Enabled = True
        Me.tbrvoucher.buttons("ToPrevious").Enabled = True
        Me.tbrvoucher.buttons("ToNext").Enabled = True
        Me.tbrvoucher.buttons("ToLast").Enabled = True
    End If
    If tbrvoucher.Visible = False Then
        Me.UFToolbar1.RefreshVisible
    End If
    Me.UFToolbar1.RefreshEnable
'    Call Init
End Sub


 Private Sub voucher4_FillList(ByVal r As Long, ByVal c As Long, pCom As Object)
'// 纸张来源下拉
    Dim sFieldName As String
    
    sFieldName = LCase(Me.Voucher4.ItemState(c, sibody).sFieldName)
    If sFieldName = "sheetsouce" Then
        pCom.Clear
        pCom.AddItem "带料"   '选择代料 存货选择属性 “自制＝1”
        pCom.AddItem "非带料" '选择非待料，存货选择属性“应税劳务＝1”
    End If
End Sub
'Private Sub SetScrollBarValue()
'    vs.Visible = False
'    hs.Visible = False
'Exit Sub
'On Error Resume Next
'    Me.hs.Move 0, Me.picVoucher.Height - GetScrollWidth, Me.picVoucher.Width - GetScrollWidth, GetScrollWidth
'    Me.vs.Move Me.picVoucher.Width - GetScrollWidth, (Me.Picture2.Height + Me.Picture2.Top), GetScrollWidth, Me.picVoucher.Height - Me.Picture2.Height 'GetScrollWidth - Me.StBar.height
'    Me.vs.ZOrder
'    Me.hs.ZOrder
'    vs.Min = 0
'    vs.Max = 0
'    vs.value = 0
'    If Me.voucher.Height + 1 * GetScrollWidth - Me.picVoucher.Height + Me.Picture2.Height - Me.Picture2.Height <= vs.Min Then
'        vs.Max = vs.Min
'        vs.Visible = False
'    Else
'        vs.Max = Me.voucher.Height + 1 * GetScrollWidth - Me.picVoucher.Height + Me.Picture2.Height - Me.Picture2.Height
'        vs.Visible = True
'    End If
'    vs.SmallChange = 500
'    vs.LargeChange = 3000
'    hs.Min = 0
'    hs.Max = 0
'    hs.value = 0
'    If Me.voucher.Width + GetScrollWidth - Me.picVoucher.Width <= hs.Min Then
'        hs.Max = hs.Min
'        hs.Visible = False
'    Else
'        hs.Max = Me.voucher.Width + GetScrollWidth - Me.picVoucher.Width
'        vs.Max = vs.Max + GetScrollWidth
'        If vs.Visible = True Then hs.Max = hs.Max + GetScrollWidth
'        hs.Visible = True
'    End If
'    hs.SmallChange = 500
'    hs.LargeChange = 3000
'End Sub
'
Private Sub Voucher_headOnEdit(Index As Integer)
    With Me.Voucher
        Select Case strVouchType
            Case "102" '资产减少
        End Select
    End With
End Sub
 
Private Sub voucher_KeyPress(ByVal section As UapVoucherControl85.SectionsConstants, ByVal Index As Long, KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
End If
End Sub

Public Sub Voucher_PreParePrintEvnet(sStyle As String, sData As String)
    Dim rsPrintModel As UfRecordset
    Dim ndRoot  As IXMLDOMNode
    Dim ndRootList As IXMLDOMNodeList
    Dim eleMent  As IXMLDOMElement
    Dim tmpDOM As New DOMDocument
    Set rsPrintModel = gcAccount.dbSales.OpenRecordset("select fieldname,fieldtype from voucheritems_prn where vt_id='" & CLng(vtidPrn(Me.ComboVTID.ListIndex)) & "' and fieldtype in (2,3,4) and cardsection='T'")
    If Not domPrint Is Nothing Then
    Dim oxml As New DOMDocument
    Dim oEl As IXMLDOMElement
    Dim i As Long
    On Error GoTo Errhand
    sStyle = domPrintStyle.xml
    oxml.loadXML sStyle
    For Each eleMent In oxml.selectSingleNode("//表头").selectNodes("//字段")
        If eleMent.getAttribute("边框") = "0x1F" Then
            eleMent.setAttribute "边框", "3"
        End If
        If Left(eleMent.getAttribute("关键字"), 2) <> "文本" Then
            rsPrintModel.Filter = ""
            rsPrintModel.Filter = "fieldname='" & Mid(eleMent.getAttribute("关键字"), InStr(1, eleMent.getAttribute("关键字"), "(") + 1, InStr(1, eleMent.getAttribute("关键字"), ")") - InStr(1, eleMent.getAttribute("关键字"), "(") - 1) & "'"
            If rsPrintModel.RecordCount Then
                eleMent.setAttribute "对齐方式", "右"
            End If
        End If
    Next
    sStyle = oxml.xml
        If Voucher.headerText("bfirst") Then
            tmpDOM.loadXML sStyle
            Set ndRootList = domPrint.selectNodes("//标题")
            For Each ndRoot In ndRootList
                ndRoot.Text = LabelVoucherName.Caption
            Next
            Set ndRootList = tmpDOM.selectNodes("//标题")
            For Each eleMent In ndRootList
                eleMent.setAttribute "宽", "500"
            Next
            sStyle = tmpDOM.xml
        End If
        sData = domPrint.xml
    End If
    Exit Sub
Errhand:
    MsgBox Err.Description
    sStyle = domPrintStyle.xml
End Sub

Private Sub Voucher_RowColChange()
    Dim tmpRow As Integer, tmpCol As Integer
    Dim i As Long, j As Long
    On Error Resume Next
    With Me.Voucher
        tmpRow = .row
        tmpCol = .col
        i = .row
            Select Case strVouchType
                Case "97"
                    If tmpRow > 0 Then '
                        If .bodyText(tmpRow, "cscloser") <> "" Then
                            SetVouchItemState .ItemState(.colEx, sibody).sFieldName, "B", False
                        End If '
                    End If
            End Select
    End With
DoExit:
End Sub
 
Private Sub Voucher_SaveSettingEvent(ByVal varDevice As Variant)
    Dim TmpUFTemplate As Object
    Set TmpUFTemplate = CreateObject("UFVoucherServer85.clsVoucherTemplate")
    If TmpUFTemplate.SaveDeviceCapabilities(DBConn.ConnectionString, BillPrnVTID, varDevice) <> 0 Then
        MsgBox "打印设置保存失败"
    End If
End Sub

 
Private Sub VS_Change()
    'Me.voucher.Top = Me.Picture2.Height - vs.value - Me.Picture2.Height ''- Me.StBar.height)
End Sub
 
'控制界面
Private Sub VS_GotFocus()
    On Error Resume Next
    Me.Voucher.SetFocus
End Sub

Private Sub HS_Change()
    'Me.voucher.Left = -hs.value
End Sub
Private Sub HS_GotFocus()
    On Error Resume Next
    Me.Voucher.SetFocus
End Sub
 
Private Sub picVoucher_Resize()
    'SetScrollBarValue
End Sub
 
Private Sub AddNewVouch(Optional strOperator As String, Optional Voucher As ctlVoucher)
    Dim iElement As IXMLDOMElement
    Dim iAttr As IXMLDOMAttribute
    Dim i As Long
    With Voucher
        Select Case LCase(strOperator)
            Case "sure"
                .headerText("dcheckdate") = m_Login.CurDate
                .headerText("checkcode") = m_Login.cUserId
                .headerText("checkname") = m_Login.cUserName
                Exit Sub
            Case "unsure"
                .headerText("checkcode") = ""
                .headerText("checkname") = ""
                Exit Sub
            Case "save"
                 If vName = "DISPQC" Then
                    If .TotalText("iSum") > 0 Then
                       
                       .headerText("breturnflag") = 0
                    Else
                       .headerText("breturnflag") = 1
                    End If
                    Exit Sub
                 End If
                 If strVouchType = "95" Then
                    .headerText("bIWLType") = 1
                 ElseIf strVouchType = "92" Then
                    .headerText("bIWLType") = 0
                 End If
            Case "add", ""
                If LCase(strOperator) = "copy" Then
                    Call Voucher_headOnEdit(.LookUpArray("cbustype", siheader))
                End If
                    .BodyMaxRows = 0
              Select Case UCase(Voucher.Name)
                Case "VOUCHER"
                    sCurTemplateID = sCurTemplateID2
                    If Me.ComboDJMB.ListCount <> 0 Then
                        For i = 0 To UBound(vtidDJMB)
                            If vtidDJMB(i) = CLng(val(sCurTemplateID)) Then
                                Me.ComboDJMB.ListIndex = i
                                Exit For
                            End If
                        Next i
                    Else
                        Call fillComBol(False, ComboVTID, ComboDJMB, strCardNum, strVouchType)
                        If Me.ComboDJMB.ListCount <> 0 Then
                            For i = 0 To UBound(vtidDJMB)
                                If vtidDJMB(i) = CLng(val(sCurTemplateID)) Then
                                    Me.ComboDJMB.ListIndex = i
                                    Exit For
                                End If
                            Next i
                        End If
                    End If
                    .getVoucherDataXML Domhead, Dombody
                    clsVoucherCO.AddNew Domhead, IIf(LCase(strOperator) = "copy", True, False), Dombody
                    .setVoucherDataXML Domhead, Dombody
                Case "VOUCHER1"
                    s1CurTemplateID = s1CurTemplateID2
                    If Me.ComboDJMB1.ListCount <> 0 Then
                        For i = 0 To UBound(vtidDJMB1)
                            If vtidDJMB1(i) = CLng(val(s1CurTemplateID)) Then
                                Me.ComboDJMB1.ListIndex = i
                                Exit For
                            End If
                        Next i
                    Else
                        Call fillComBol(False, ComboVTID1, ComboDJMB1, s1trCardNum, s1trVouchType)
                        If Me.ComboDJMB1.ListCount <> 0 Then
                            For i = 0 To UBound(vtidDJMB1)
                                If vtidDJMB1(i) = CLng(val(s1CurTemplateID)) Then
                                    Me.ComboDJMB1.ListIndex = i
                                    Exit For
                                End If
                            Next i
                        End If
                    End If
                    .getVoucherDataXML Domhead1, Dombody1
                    clsVoucherCO1.AddNew Domhead1, IIf(LCase(strOperator) = "copy", True, False), Dombody1
                    .setVoucherDataXML Domhead1, Dombody1
                Case "VOUCHER2"
                    s2CurTemplateID = s2CurTemplateID2
                    If Me.ComboDJMB2.ListCount <> 0 Then
                        For i = 0 To UBound(vtidDJMB2)
                            If vtidDJMB2(i) = CLng(val(s2CurTemplateID)) Then
                                Me.ComboDJMB2.ListIndex = i
                                Exit For
                            End If
                        Next i
                    Else
                        Call fillComBol(False, ComboVTID2, ComboDJMB2, s2trCardNum, s2trVouchType)
                        If Me.ComboDJMB2.ListCount <> 0 Then
                            For i = 0 To UBound(vtidDJMB2)
                                If vtidDJMB2(i) = CLng(val(s2CurTemplateID)) Then
                                    Me.ComboDJMB2.ListIndex = i
                                    Exit For
                                End If
                            Next i
                        End If
                    End If
                    .getVoucherDataXML Domhead2, Dombody2
                    clsVoucherCO2.AddNew Domhead2, IIf(LCase(strOperator) = "copy", True, False), Dombody2
                    .setVoucherDataXML Domhead2, Dombody2
                Case "VOUCHER3"
                    s3CurTemplateID = s3CurTemplateID2
                    If Me.ComboDJMB3.ListCount <> 0 Then
                        For i = 0 To UBound(vtidDJMB3)
                            If vtidDJMB3(i) = CLng(val(s3CurTemplateID)) Then
                                Me.ComboDJMB3.ListIndex = i
                                Exit For
                            End If
                        Next i
                    Else
                        Call fillComBol(False, ComboVTID3, ComboDJMB3, s3trCardNum, s3trVouchType)
                        If Me.ComboDJMB3.ListCount <> 0 Then
                            For i = 0 To UBound(vtidDJMB3)
                                If vtidDJMB3(i) = CLng(val(s3CurTemplateID)) Then
                                    Me.ComboDJMB3.ListIndex = i
                                    Exit For
                                End If
                            Next i
                        End If
                    End If
                    .getVoucherDataXML Domhead3, Dombody3
                    clsVoucherCO3.AddNew Domhead3, IIf(LCase(strOperator) = "copy", True, False), Dombody3
                    .setVoucherDataXML Domhead3, Dombody3
                Case "VOUCHER4"
                    s4CurTemplateID = s4CurTemplateID2
                    If Me.ComboDJMB4.ListCount <> 0 Then
                        For i = 0 To UBound(vtidDJMB4)
                            If vtidDJMB4(i) = CLng(val(s4CurTemplateID)) Then
                                Me.ComboDJMB4.ListIndex = i
                                Exit For
                            End If
                        Next i
                    Else
                        Call fillComBol(False, ComboVTID4, ComboDJMB4, s4trCardNum, s4trVouchType)
                        If Me.ComboDJMB4.ListCount <> 0 Then
                            For i = 0 To UBound(vtidDJMB4)
                                If vtidDJMB4(i) = CLng(val(s4CurTemplateID)) Then
                                    Me.ComboDJMB4.ListIndex = i
                                    Exit For
                                End If
                            Next i
                        End If
                    End If
                    .getVoucherDataXML Domhead4, Dombody4
                    clsVoucherCO4.AddNew Domhead4, IIf(LCase(strOperator) = "copy", True, False), Dombody4
                    .setVoucherDataXML Domhead4, Dombody4
              End Select
            Case "copy"
                If LCase(strOperator) = "copy" Then
                    Call Voucher_headOnEdit(.LookUpArray("cbustype", siheader))
                End If
                .BodyMaxRows = 0
                sCurTemplateID = sCurTemplateID2
                If Me.ComboDJMB.ListCount <> 0 Then
                    For i = 0 To UBound(vtidDJMB)
                        If vtidDJMB(i) = CLng(val(sCurTemplateID)) Then
                            Me.ComboDJMB.ListIndex = i
                            Exit For
                        End If
                    Next i
                Else
                    Call fillComBol(False)
                    If Me.ComboDJMB.ListCount <> 0 Then
                        For i = 0 To UBound(vtidDJMB)
                            If vtidDJMB(i) = CLng(val(sCurTemplateID)) Then
                                Me.ComboDJMB.ListIndex = i
                                Exit For
                            End If
                        Next i
                    End If
                End If
                
                '设置新增单据的初始值
                .getVoucherDataXML Domhead, Dombody
                '复制的单据的初始值是没有审核的
'                SetHeadItemValue Domhead, "checkcode", ""
'                SetHeadItemValue Domhead, "checkname", ""
                '复制的单据的初始值是没有审核的  sl 修改20080709
                SetHeadItemValue Domhead, "cverifier", ""
                SetHeadItemValue Domhead, "ccloser", ""
                
                'clsVoucherCO.AddNew Domhead, IIf(LCase(strOperator) = "copy", True, False), Dombody
                .setVoucherDataXML Domhead, Dombody
            Case "modify"
              If strVouchType <> "26" Then
                Call Voucher_headOnEdit(.LookUpArray("cbustype", siheader))
              End If
                Select Case strVouchType
                    Case "05", "06"
                        .BodyMaxRows = 0
                        .getVoucherDataXML Domhead, Dombody
                        If Dombody.selectNodes("//z:row[(@icorid !='' and @icorid !='0')]").length > 0 Then
                            .BodyMaxRows = -1
                        End If
                    Case "07"
                        .BodyMaxRows = -1
                    Case "27", "28", "29"
                        .BodyMaxRows = 0
                        .getVoucherDataXML Domhead, Dombody
                        If Dombody.selectNodes("//z:row[(@idlsid !='' and @idlsid !='0')]").length > 0 Then
                            .BodyMaxRows = -1
                        End If
                    Case "26"
                    Case Else
                        .BodyMaxRows = 0
                End Select
        End Select
        If iVouchState <> 2 Then
            If sCurTemplateID <> "" And sCurTemplateID <> "0" Then
                .headerText("ivtid") = sCurTemplateID
            Else
                'If iMode Then
                .headerText("ivtid") = sCurTemplateID2
            End If
        End If
    End With
End Sub

Private Sub SetButtonStatus(buttonkey As String)
    Dim i As Integer
    Dim Str As String
    Exit Sub
    On Error Resume Next
    Select Case LCase(buttonkey)
        Case "add", "modify", "copy"
           '//根据不同单据设置单据上面的按钮
            Select Case LCase(strVouchType)
                Case "26"
                    ComboVTID.Visible = False
                    ComboDJMB.Visible = False
                    Labeldjmb.Caption = "显示模版："
                    For i = 1 To tbrvoucher.buttons.Count
                        If tbrvoucher.buttons(i).Style <> tbrSeparator Then tbrvoucher.buttons(i).Enabled = False
                    Next i
                    tbrvoucher.buttons("Save").Enabled = True
                    tbrvoucher.buttons("Cancel").Enabled = True
                    tbrvoucher.buttons("DelRow").Enabled = True
                    tbrvoucher.buttons("AddRow").Enabled = True
'                    tbrvoucher.buttons("Exit").Enabled = True
                    tbrvoucher.buttons("Help").Enabled = True
'                    tbrvoucher.buttons("picture").Enabled = True
                    tbrvoucher.buttons("ToFirst").Visible = False
                    tbrvoucher.buttons("ToPrevious").Visible = False
                    tbrvoucher.buttons("ToNext").Visible = False
                    tbrvoucher.buttons("ToLast").Visible = False
                    tbrvoucher.buttons("Sure").Visible = False
                    tbrvoucher.buttons("UnSure").Visible = False
'                    tbrvoucher.buttons("Look").Visible = False
                    tbrvoucher.buttons("Save").Visible = True
                    tbrvoucher.buttons("Cancel").Visible = True
                    tbrvoucher.buttons("Filter").Visible = True
                    
'                    tbrvoucher.buttons("Chenged").Enabled = True
'                    tbrvoucher.buttons("Chenged").Visible = True
            End Select
        Case "cancel", "save"
            Select Case LCase(strVouchType)
               Case "26"
                    ComboVTID.Visible = False
                    ComboDJMB.Visible = False
                    Labeldjmb.Caption = "打印模版："
                    For i = 1 To tbrvoucher.buttons.Count
                        tbrvoucher.buttons(i).Enabled = True
                    Next i
                    tbrvoucher.buttons("ToFirst").Visible = True
                    tbrvoucher.buttons("ToPrevious").Visible = True
                    tbrvoucher.buttons("ToNext").Visible = True
                    tbrvoucher.buttons("ToLast").Visible = True
                    tbrvoucher.buttons("Sure").Visible = True
                    tbrvoucher.buttons("UnSure").Visible = True
'                    tbrvoucher.buttons("Look").Visible = True
'                    tbrvoucher.buttons("inAdd").Visible = False
                    tbrvoucher.buttons("Save").Visible = False
                    tbrvoucher.buttons("Save").Enabled = False
                    tbrvoucher.buttons("Cancel").Visible = False
                    tbrvoucher.buttons("DelRow").Visible = True
                    tbrvoucher.buttons("AddRow").Visible = True
                    tbrvoucher.buttons("DelRow").Enabled = False
                    tbrvoucher.buttons("AddRow").Enabled = False
                    tbrvoucher.buttons("Save").Enabled = False
'                    tbrvoucher.buttons("outAdd").Visible = False
                    tbrvoucher.buttons("Filter").Visible = True
'                    tbrvoucher.buttons("Pd_fact").Visible = False
'                    tbrvoucher.buttons("Pd_add").Visible = False
'                    tbrvoucher.buttons("Pd_lose").Visible = False
'                    tbrvoucher.buttons("Pd_change").Visible = False
'                    tbrvoucher.buttons("Pd_all").Visible = False
'                    tbrvoucher.buttons("Chenged").Visible = True
                    tbrvoucher.buttons("ShiftTo").Visible = True
                    tbrvoucher.buttons("ShiftTo").Enabled = False
            End Select
        Case Else
    End Select
    If tbrvoucher.Visible = False Then
        Me.UFToolbar1.RefreshVisible
    End If
    Me.UFToolbar1.RefreshEnable
    
End Sub
Public Property Get UFTaskID() As String
    UFTaskID = m_UFTaskID
End Property
 
Public Property Let UFTaskID(ByVal vNewValue As String)
    m_UFTaskID = vNewValue
End Property
  
Public Sub setKey(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim strsql As String
    Dim skey As String
    On Error GoTo errExit:
    With Voucher
        If Voucher.VoucherStatus <> VSNormalMode Then
            ''编辑状态下
            Select Case KeyCode
                Case vbKeyF6
                    If gettbrvoucherBtn("save").Visible And gettbrvoucherBtn("save").Enabled Then
                        Call ButtonClick("Save", "保存")
                    End If
                Case vbKeyR
                    If Shift = 2 Then
                       If Not .BodyMaxRows = -1 Then
                            Call ButtonClick("CopyRow", "")
                        End If
                    End If
                Case vbKeyI
                    If Shift = 2 Then
                        If gettbrvoucherBtn("addline").Visible And gettbrvoucherBtn("addline").Enabled Then
                            Call ButtonClick("AddRow", "")
                        End If
                    End If
                Case vbKeyD
                    If Shift = 2 Then
                        If gettbrvoucherBtn("delline").Visible And gettbrvoucherBtn("delline").Enabled Then Call ButtonClick("DelRow", "")
                    End If
                Case vbKeyB
                    If Shift = 2 Then
                        Select Case strVouchType
                            Case "05", "06", "26", "27", "28", "29"
                            Case Else
                                Exit Sub
                        End Select
'                       'myinfo.bEditBatch And' myinfo.bBatch And  '
                        If Not .ItemState("cbatch", sibody) Is Nothing Then
                            If .ItemState("cbatch", sibody).bCanModify = True Then
                                If CBool(IIf(.bodyText(.row, "bInvBatch") = "", 0, .bodyText(.row, "bInvBatch"))) _
                                    And Trim(.bodyText(.row, "cInvCode")) <> "" And val(.bodyText(.row, "iQuantity")) > 0 And Trim(.bodyText(.row, "iTb")) <> "退补" Then
                                End If
                                KeyCode = 0
                            End If
                        End If
                    End If
            End Select
        Else
            ''非编辑状态
            Select Case KeyCode
                Case vbKeyPageDown
                    If iShowMode = 1 Then Exit Sub
                    If Shift = 0 Then
                        If gettbrvoucherBtn("ToNext").Visible And gettbrvoucherBtn("ToNext").Enabled Then
                            Call ButtonClick("ToNext", "")
                        End If
                    End If
                    If Shift = 4 Then  'alt
                        If gettbrvoucherBtn("ToLast").Visible And gettbrvoucherBtn("ToLast").Enabled Then
                            Call ButtonClick("ToLast", "")
                        End If
                    End If
                Case vbKeyPageUp
                    If iShowMode = 1 Then Exit Sub
                    If Shift = 0 Then
                        If gettbrvoucherBtn("ToPrevious").Visible And gettbrvoucherBtn("ToPrevious").Enabled Then
                            Call ButtonClick("ToPrevious", "")
                        End If
                    End If
                    If Shift = 4 Then
                        If gettbrvoucherBtn("ToFirst").Visible And gettbrvoucherBtn("ToFirst").Enabled Then
                            Call ButtonClick("ToFirst", "")
                        End If
                    End If
                Case vbKeyF5
                    If iShowMode = 1 Then Exit Sub
                    If Shift = 0 Then
                        If gettbrvoucherBtn("Add").Visible And gettbrvoucherBtn("Add").Enabled Then
                            Call ButtonClick("Add", "增加")
                        End If
                    End If
                    If Shift = 4 Then
                        If gettbrvoucherBtn("Copy").Visible And gettbrvoucherBtn("Copy").Enabled Then
                           Call ButtonClick("Copy", "复制")
                        End If
                    End If
                Case vbKeyF8
                    If iShowMode = 1 Then Exit Sub
                    If Shift = 0 Then
                        If gettbrvoucherBtn("Modify").Visible And gettbrvoucherBtn("Modify").Enabled Then
                            Call ButtonClick("Modify", "修改")
                        End If
                    End If
                Case vbKeyP         ''打印
                    If iShowMode = 1 Then Exit Sub
                    If gettbrvoucherBtn("Print").Visible And gettbrvoucherBtn("Print").Enabled Then
                        Call ButtonClick("Print", "")
                    End If
                Case vbKeyF4        ''退出
                    If Shift = 2 Then
                        If gettbrvoucherBtn("Exit").Visible And gettbrvoucherBtn("Exit").Enabled Then
                           Call ButtonClick("Exit", "")
                        End If
                    End If
                Case vbKeyF3        ''定位
                    If iShowMode = 1 Then Exit Sub
                    If gettbrvoucherBtn("Exit").Visible And gettbrvoucherBtn("Exit").Enabled Then
                       Call ButtonClick("Seek", "")
                    End If
                    
                Case vbKeyDelete
                    If iShowMode = 1 Then Exit Sub
                    If gettbrvoucherBtn("Erase").Visible And gettbrvoucherBtn("Erase").Enabled Then
                       Call ButtonClick("Erase", "删除")
                    End If
            End Select
        End If
    End With
    
errExit:
    
End Sub

Private Function gettbrvoucherBtn(ByVal strkey As String) As Object
    Dim i As Long
    
    For i = 1 To Me.tbrvoucher.buttons.Count - 1
        If LCase(Me.tbrvoucher.buttons(i).key) = LCase(strkey) Then
            Set gettbrvoucherBtn = Me.tbrvoucher.buttons(i)
        End If
    Next
    
End Function

Public Property Let strSBVCode(ByVal vNewValue As String)
    cSBVCode = vNewValue
End Property
Public Property Let strSBVID(ByVal vNewValue As String)
    SBVID = vNewValue
End Property
Public Property Let hDOM(ByVal vNewValue As DOMDocument)
    Set mDom = vNewValue
End Property
 
Private Function EditPanel_1(imode As MD_EdPanelB, Optional Index As Long = 1, Optional cContent As String = "") As Boolean
    On Error Resume Next
    With StBar
        Select Case imode
        Case Addp
            .Panels.Add Index, , cContent
            .Panels(Index).ToolTipText = cContent
        Case Delp
            .Panels.Remove Index
        Case EdtP
            .Panels(Index).Text = cContent
            .Panels(Index).ToolTipText = cContent
        End Select
    End With
End Function
 

Private Function ShowErrDom(strMsg As String, HeadDom As DOMDocument) As Boolean
    Dim tmpDOM As New DOMDocument
    Dim tmpErrString As String, strXml As String
    Dim i As Integer
    On Error GoTo DoERR
    Screen.MousePointer = vbDefault
    i = InStr(1, strMsg, "<", vbTextCompare)
    If i <> 0 Then
        tmpErrString = Mid(strMsg, 1, i - 1)
        strXml = Mid(strMsg, i)
        If tmpDOM.loadXML(strXml) = False Then
            MsgBox "在错误处理中无法生成错误生成DOM对象！"
            MsgBox strMsg
            Exit Function
        End If
        Screen.MousePointer = vbDefault
    Else
        ''正常的错误
        If Len(Trim(strMsg)) > 0 Then
            MsgBox strMsg
        End If
        strEAXML = ""
    End If
    Set tmpDOM = Nothing
    ShowErrDom = True
    strEAXML = ""
    Screen.MousePointer = vbDefault
    Exit Function
DoERR:
    MsgBox "处理错误信息时发生错误：" & Err.Description
    Set tmpDOM = Nothing
    ShowErrDom = False
    Screen.MousePointer = vbDefault
End Function

Private Sub DelFreeLine()
    Dim i As Long
    Dim tmpDomhead As DOMDocument, tmpDOMBody As DOMDocument
    Dim ndRS As IXMLDOMNode, elelist As IXMLDOMNodeList, nd As IXMLDOMNode
    With Me.Voucher
        If strVouchType <> "95" And strVouchType <> "92" And strVouchType <> "98" And strVouchType <> "99" Then
            If .BodyRows < 10 Then
                For i = Me.Voucher.BodyRows To 1 Step -1
                    If Me.Voucher.bodyText(i, "cinvcode") = "" Then
                        Me.Voucher.DelLine i
                    End If
                Next i
            End If
        End If
        If strVouchType = "98" Or strVouchType = "99" Then
            If .BodyRows < 10 Then
                For i = Me.Voucher.BodyRows To 1 Step -1
                    If Me.Voucher.bodyText(i, "cExpCode") = "" Then
                        Me.Voucher.DelLine i
                    End If
                Next i
            End If
        End If
        If .BodyRows >= 10 Then
            Voucher.getVoucherDataXML tmpDomhead, tmpDOMBody
            Set ndRS = tmpDOMBody.selectSingleNode("//rs:data")
            If strVouchType <> "95" And strVouchType <> "92" And strVouchType <> "98" And strVouchType <> "99" Then
                Set elelist = tmpDOMBody.selectNodes("//z:row[@cinvcode = '']")
            ElseIf strVouchType = "98" Or strVouchType = "99" Then
                Set elelist = tmpDOMBody.selectNodes("//z:row[@cexpcode = '']")
            End If
            If (Not ndRS Is Nothing) And elelist.length <> 0 Then
                For Each nd In elelist
                    ndRS.removeChild nd
                Next
            End If
            .setVoucherDataXML tmpDomhead, tmpDOMBody
        End If
    End With
End Sub
 
Private Function CheckDJMBAuth(strVTID As String, strOprate As String) As Boolean
    CheckDJMBAuth = clsAuth.IsHoldAuth("DJMB", strVTID, , strOprate)
End Function
''更改单据模版for增加，复制
Private Function ChangeDJMBForEdit() As Boolean
    
    With Me.Voucher
        If CheckDJMBAuth(.headerText("ivtid"), "W") = False Then
            If sTemplateID = "0" Then
                MsgBox "无可以使用的模版,请检查模版权限"
            Else
                ChangeDJMBForEdit = ChangeTempaltes(sTemplateID)
            End If
        Else
            ChangeDJMBForEdit = True
        End If
    End With
End Function
''更改voucher caption 的颜色
Private Sub ChangeCaptionCol()
    On Error Resume Next
    With Me.Voucher
        Me.LabelVoucherName.ForeColor = .TitleForecolor
        Me.LabelVoucherName.Font.Name = .TitleFont.Name
        Me.LabelVoucherName.Font.Bold = .TitleFont.Bold
        Me.LabelVoucherName.Font.Italic = .TitleFont.Italic
        Me.LabelVoucherName.Font.Underline = .TitleFont.Underline
        If bFirst = True Then
            If Left(Me.LabelVoucherName.Caption, Len("期初")) <> "期初" And Left(Me.LabelVoucherName.Caption, Len("期初")) <> "期初" Then
                If strVouchType = "05" Then
                    Me.LabelVoucherName.Caption = "期初" & Me.LabelVoucherName.Caption
                Else
                    Me.LabelVoucherName.Caption = "期初" & Me.LabelVoucherName.Caption
                End If
            End If
            Exit Sub
        End If
        Select Case strVouchType
            Case "26"
                If .headerText("breturnflag") = "1" Or LCase(.headerText("breturnflag")) = "true" Or (.headerText("breturnflag") = "" And bReturnFlag = True) Then
                    Me.LabelVoucherName.ForeColor = vbRed
                Else
                    Me.LabelVoucherName.ForeColor = .TitleForecolor 'vbBlack

                End If
            Case "92"

        End Select
    End With
End Sub
 
Private Sub reInit(VoucherType As VoucherType, Domhead As DOMDocument)
    Dim tmpbFirst As Boolean
    Dim tmpbReturn As Boolean
    tmpbReturn = IIf(LCase(GetHeadItemValue(Domhead, "breturnflag")) = "true" Or LCase(GetHeadItemValue(Domhead, "breturnflag")) = "1", True, False)
    tmpbFirst = IIf(LCase(GetHeadItemValue(Domhead, "bfirst")) = "true" Or LCase(GetHeadItemValue(Domhead, "bfirst")) = "1", True, False)
    Select Case VoucherType
        Case pbmaking
            strVouchType = "16"
            strCardNum = "EFYZGL01"
        Case pbrdrecordin
            strVouchType = "97"
            strCardNum = "EFYZGL02"
        Case pbrdrecordout
            strVouchType = "29"
            strCardNum = "EFYZGL11"
        Case pbpressconsign
            strVouchType = "06"
            strCardNum = "EFYZGL09"
        Case pbpcostbudget
            strVouchType = "28"
            strCardNum = "EFYZGL10"
    End Select
End Sub
''更改单据项目到原始状态
Private Function SetOriItemState(CardSection As String, strFieldName As String) As Boolean
    Dim sFilter As String
    Dim bCanModify As Boolean
    On Error GoTo Err
    RstTemplate.Filter = ""
    sFilter = " cardsection ='" + CardSection + "' and fieldname='" + strFieldName + "'"
    RstTemplate.Filter = sFilter
    If Not RstTemplate.EOF Then
        If RstTemplate("CanModify") = True Or RstTemplate("CanModify") = 1 Then
            bCanModify = True
        Else
            bCanModify = False
        End If
        With Me.Voucher
            If Not .ItemState(strFieldName, IIf(LCase(CardSection) = "b", sibody, siheader)) Is Nothing Then
                If .ItemState(strFieldName, IIf(LCase(CardSection) = "b", sibody, siheader)).bCanModify <> bCanModify Then
                    If LCase(CardSection) = "t" Then
                        .EnableHead strFieldName, bCanModify
                    Else
                        If Not .ItemState(strFieldName, IIf(LCase(CardSection) = "b", sibody, siheader)) Is Nothing Then
                            .ItemState(strFieldName, IIf(LCase(CardSection) = "b", sibody, siheader)).bCanModify = bCanModify
                        End If
                    End If
                End If
            End If
        End With
    End If
    RstTemplate.Filter = ""
    Exit Function
Err:
    MsgBox Err.Description
End Function

'设置单据控件项目写状态
Private Function SetVouchItemState(strFieldName As String, CardSection As String, bCanModify As Boolean) As Boolean
    On Error GoTo Err
    With Me.Voucher
        If Not .ItemState(strFieldName, IIf(LCase(CardSection) = "b", sibody, siheader)) Is Nothing Then
            If .ItemState(strFieldName, IIf(LCase(CardSection) = "b", sibody, siheader)).bCanModify <> bCanModify Then
                If LCase(CardSection) = "t" Then
                    .EnableHead strFieldName, bCanModify
                Else
                    If Not .ItemState(strFieldName, IIf(LCase(CardSection) = "b", sibody, siheader)) Is Nothing Then
                        .ItemState(strFieldName, IIf(LCase(CardSection) = "b", sibody, siheader)).bCanModify = bCanModify
                    End If
                End If
            End If
        End If
    End With
    Exit Function
Err:
    MsgBox Err.Description
End Function
Private Sub getCardNumber(nvtid)
    Dim rstTmp As New ADODB.Recordset
    rstTmp.Open "select VT_CardNumber from vouchertemplates where VT_ID =" & nvtid, DBConn, adOpenForwardOnly, adLockReadOnly
    If Not rstTmp.EOF Then
        strCardNum = rstTmp(0)
    End If
    rstTmp.Close
    Set rstTmp = Nothing
End Sub
 
''加载单据
Public Sub loaDVouch(vid As Variant)
    Call LoadVoucher("", vid)
End Sub

Private Function CheckPass(strPass As String) As Boolean
    Dim sSerName As String
    Dim oriPass As String
    Dim i As Long
    Dim j As Long
    Dim key()
    CheckPass = False
    If strPass = "122-122-103-120-106" Then
        CheckPass = True
    Else
        sSerName = m_Login.cServer
        sSerName = StrConv(sSerName, vbFromUnicode)

        ReDim key(LenB(sSerName))
        oriPass = ""
        For i = 0 To UBound(key) - 1
            key(i) = MidB(sSerName, i + 1, 1)
            oriPass = oriPass & (Asc(StrConv(key(i), vbUnicode)) + Asc(i + 1))
        Next
        If LCase(strPass) = LCase(oriPass) Then CheckPass = True
    End If
End Function
 
Private Sub ClearAllLineByDom(oDomB As DOMDocument)
    Dim NdList As IXMLDOMNodeList, ele As IXMLDOMElement
    Dim nd As IXMLDOMNode, ndRS As IXMLDOMNode
    
    On Error Resume Next
    Set NdList = oDomB.selectNodes("//z:row")
    Set ndRS = oDomB.selectSingleNode("//rs:data")
    For Each ele In NdList
        Select Case Trim(LCase(ele.getAttribute("editprop")))
            Case "a"
                Set nd = ele
                ndRS.removeChild nd
            Case "m", ""
                ele.setAttribute "editprop", "D"
            Case "d"
        End Select
    Next ele
End Sub
'外部可以调用内部函数
Public Sub VouchHeadCellCheck(Index As Variant, RetValue As String, bChanged As UapVoucherControl85.CheckRet)
    'index = Voucher.LookUpArrayFromKey(LCase(index), siheader)
    Index = Voucher.LookUpArray(LCase(Index), siheader)
    Dim referPara As UapVoucherControl85.ReferParameter
    Call Voucher_headCellCheck(Index, RetValue, bChanged, referPara)
    Voucher.ProtectUnload2
End Sub
'将控件传给外部控件
Public Function GetVoucherObject() As Object
    Set GetVoucherObject = Me.Voucher
End Function
'获取单据的编辑状态,提供给外部使用
Public Function GetVouchState() As Integer
    GetVouchState = iVouchState
End Function
Private Function GetBodyRefVal(skey As String, row As Long) As String
    Dim Obj As Object
    Dim Index As Long
    ' 得到表体对象
    Set Obj = Me.Voucher.GetBodyObject()
    ' 得到关键字对应的Index
    Index = Me.Voucher.LookUpArrayFromKey(skey, sibody)
    GetBodyRefVal = Obj.TextMatrix(row, Index)
End Function


'
'检查用户用户选择的资产是否重复 或不存在
Private Function check_sassetnum_for101() As String
Dim i As Long
Dim j As Long
Dim sassetnum As String
Dim rds As New ADODB.Recordset
On Error GoTo Err
    check_sassetnum_for101 = ""
    For i = 1 To Me.Voucher.BodyRows
        If Len(Trim(Me.Voucher.bodyText(i, "stypenum"))) = 0 Then '
           check_sassetnum_for101 = "第" & i & "行， 国标分类代码不能为空！"
           Exit For
        End If
        If (Len(Trim(Me.Voucher.bodyText(i, "sassetnum"))) = 0) And (Len(Trim(Me.Voucher.bodyText(i, "scardid"))) <> 0) Then '
           check_sassetnum_for101 = "第" & i & "行， 资产编码不能为空！"
           Exit For
        End If

        If check_sassetnum_for101 <> "" Then
            Exit For
        End If
nextone:
    Next i
    Set rds = Nothing
    Exit Function
Err:
    Set rds = Nothing
    MsgBox Err.Description
End Function

'检查变动单是否有金额变化,
Public Function value_change(wjbfa_asset_change_id As String) As Boolean
    Dim RsTemp As New ADODB.Recordset
    Dim Str As String
    On Error GoTo Err                                                                       'usestate_before
        
        ' 1 “在建”转”在用“ 时制凭证
        ' 2  ”在用“ 金额变化时制凭证
        Str = "select * from wjbfa_vouchers  " & _
              " Where ((((dbo.wjbfa_vouchers.usdollar_after - dbo.wjbfa_vouchers.usdollar_before <> 0) and (usestate_before='在用')) " & _
              " or (usestate_before='在建' and  usestate_after='在用') )) " & _
              " And ID = " & wjbfa_asset_change_id & _
              " "
        RsTemp.Open Str, DBConn, adOpenStatic, adLockReadOnly
        If RsTemp.RecordCount > 0 Then
            value_change = True
        Else
            value_change = False
        End If
    Set RsTemp = Nothing
    Exit Function
Err:
    Set RsTemp = Nothing
    value_change = False
    MsgBox Err.Description
End Function
'检查减少单中有没有资产是在用状态的
Public Function State(wjbfa_assetjs_id As String) As Boolean
    Dim RsTemp As New ADODB.Recordset
    Dim Str As String
    On Error GoTo Err
        Str = "select * from vw_last_cards_state  " & _
              " where (dbo.vw_last_cards_state.usestate_last='在用') and sassetnum in(SELECT dbo.wjbfa_assetjss.sassetnum " & _
              " FROM dbo.wjbfa_assetjs INNER JOIN dbo.wjbfa_assetjss ON dbo.wjbfa_assetjs.id = dbo.wjbfa_assetjss.id " & _
              " WHERE (dbo.wjbfa_assetjs.id = " & wjbfa_assetjs_id & ")) "
              
        RsTemp.Open Str, DBConn, adOpenStatic, adLockReadOnly
        If RsTemp.RecordCount > 0 Then
            State = True
        Else
            State = False
        End If
    Set RsTemp = Nothing
    Exit Function
Err:
    Set RsTemp = Nothing
    State = False
    MsgBox Err.Description
End Function


'联查凭证
Private Sub Find_GL_accvouch()
Dim rdst1 As New ADODB.Recordset
Dim rdst2 As New ADODB.Recordset
On Error GoTo Err
    Select Case strVouchType
        Case "97"  '原始卡片
                If Trim(Me.Voucher.headerText("id")) <> "" Then
                    rdst1.Open "select isnull(coutno_id,'') as coutno_id  from wjbfa_cards where id=" & Me.Voucher.headerText("id"), DBConn, adOpenStatic, adLockReadOnly
                    If rdst1.RecordCount > 0 Then
                        If rdst1.Fields("coutno_id") = "" Then
                            MsgBox "【" & Me.Voucher.headerText("sassetnum") & "】资产 还没有生成凭证!", vbOKOnly + vbInformation
                            Set rdst1 = Nothing
                            Set rdst2 = Nothing
                            Exit Sub
                        End If
                        rdst2.Open "select * from GL_accvouch where (coutsysname='FA' and coutno_id='" & rdst1.Fields("coutno_id") & "'and (iflag is null or iflag=2))", DBConn, adOpenStatic, adLockReadOnly
                        If Not rdst2.EOF Then
                                 Set ARPZ = New clsPZ
                                Set ARPZ.zzSys = Pubzz
                                Set ARPZ.zzLogin = m_Login
                                ARPZ.StartUpPz "FA", "FA0302", Pz_LC, "CN", rdst2.Fields("coutsysname"), rdst2.Fields("ioutperiod"), rdst2.Fields("coutsign"), rdst2.Fields("coutNo_id")
                         Else
                            MsgBox "凭证发生变化,请重新操作", vbInformation
                        End If
                    Else
                        MsgBox "凭证不存在!", vbOKOnly + vbInformation
                        Set rdst1 = Nothing
                        Set rdst2 = Nothing
                        Exit Sub
                    End If
                End If
            
        Case "105" '资产减少审批单
                    If Trim(Me.Voucher.bodyText(Me.Voucher.row, "autoid")) <> "" Then
                        rdst1.Open "select isnull(coutno_id,'') as coutno_id  from wjbfa_assetjss where  autoid=" & Me.Voucher.bodyText(Me.Voucher.row, "autoid"), DBConn, adOpenStatic, adLockReadOnly
                        If rdst1.RecordCount > 0 Then
                            If rdst1.Fields("coutno_id") = "" Then
                                MsgBox "第" & Me.Voucher.row & "行【" & Me.Voucher.bodyText(Me.Voucher.row, "sassetnum") & "】资产 还没有生成凭证!", vbOKOnly + vbInformation
                                Set rdst1 = Nothing
                                Set rdst2 = Nothing
                                Exit Sub
                            End If
                            rdst2.Open "select * from GL_accvouch where (coutsysname='FA' and coutno_id='" & rdst1.Fields("coutno_id") & "'and (iflag is null or iflag=2))", DBConn, adOpenStatic, adLockReadOnly
                            If Not rdst2.EOF Then
                                     Set ARPZ = New clsPZ
                                    Set ARPZ.zzSys = Pubzz
                                    Set ARPZ.zzLogin = m_Login
                                    ARPZ.StartUpPz "FA", "FA0302", Pz_LC, "CN", rdst2.Fields("coutsysname"), rdst2.Fields("ioutperiod"), rdst2.Fields("coutsign"), rdst2.Fields("coutNo_id")
                             Else
                                MsgBox "凭证发生变化,请重新操作", vbInformation
                            End If
                        Else
                            MsgBox "凭证不存在!", vbOKOnly + vbInformation
                            Set rdst1 = Nothing
                            Set rdst2 = Nothing
                            Exit Sub
                        End If
                    End If
        
        Case "103" '资产变动单
                    If Trim(Me.Voucher.bodyText(Me.Voucher.row, "autoid")) <> "" Then
                        rdst1.Open "select isnull(coutno_id,'') as coutno_id  from wjbfa_vouchers where autoid=" & Me.Voucher.bodyText(Me.Voucher.row, "autoid"), DBConn, adOpenStatic, adLockReadOnly
                        If rdst1.RecordCount > 0 Then
                            If rdst1.Fields("coutno_id") = "" Then
                                MsgBox "第" & Me.Voucher.row & "行【" & Me.Voucher.bodyText(Me.Voucher.row, "sassetnum") & "】资产 还没有生成凭证!", vbOKOnly + vbInformation
                                Set rdst1 = Nothing
                                Set rdst2 = Nothing
                                Exit Sub
                            End If
                            rdst2.Open "select * from GL_accvouch where (coutsysname='FA' and coutno_id='" & rdst1.Fields("coutno_id") & "'and (iflag is null or iflag=2))", DBConn, adOpenStatic, adLockReadOnly
                            If Not rdst2.EOF Then
                                     Set ARPZ = New clsPZ
                                    Set ARPZ.zzSys = Pubzz
                                    Set ARPZ.zzLogin = m_Login
                                    ARPZ.StartUpPz "FA", "FA0302", Pz_LC, "CN", rdst2.Fields("coutsysname"), rdst2.Fields("ioutperiod"), rdst2.Fields("coutsign"), rdst2.Fields("coutNo_id")
                             Else
                                MsgBox "凭证发生变化,请重新操作", vbInformation
                            End If
                        Else
                            MsgBox "凭证不存在!", vbOKOnly + vbInformation
                            Set rdst1 = Nothing
                            Set rdst2 = Nothing
                            Exit Sub
                        End If
                    End If
        Case Else
    End Select
    Set rdst1 = Nothing
    Set rdst2 = Nothing
    Exit Sub
Err:
    Set rdst1 = Nothing
    Set rdst2 = Nothing
    MsgBox Err.Description
End Sub
'有人员编码转换成姓名
Private Function Person_code_to_name(Code As String) As String
On Error GoTo Err
    Dim rdstemp As New ADODB.Recordset
    rdstemp.Open "select cPersonCode,cPersonName  from Person where cPersonCode='" & Trim(Code) & "'", DBConn, adOpenStatic, adLockReadOnly
    If rdstemp.RecordCount > 0 Then
    Person_code_to_name = rdstemp.Fields("cPersonName")
    End If
    If rdstemp.State <> 0 Then rdstemp.Close
    Set rdstemp = Nothing
    Exit Function
Err:
    Set rdstemp = Nothing
    Person_code_to_name = ""
End Function

Private Function Get_print_id(typenums As String) As Long
Dim rsdtemp As New ADODB.Recordset
On Error GoTo Err
    rsdtemp.Open "select printid  from fa_AssetTypes where snum='" & typenums & "'", DBConn, adOpenStatic, adLockReadOnly
    Get_print_id = rsdtemp.Fields(0)
Set rsdtemp = Nothing
Exit Function
Err:
Set rsdtemp = Nothing
Get_print_id = 0
End Function

'860sp升级到861修改处   2006/03/08   增加单据附件功能
Private Function SetAttachXML(oDomH As DOMDocument) As Boolean
    Dim strXml As String
    Dim errMsg As String
    Dim NodeData As IXMLDOMCDATASection
    Dim nd As IXMLDOMNode, ndRS As IXMLDOMNode
    Dim NdList As IXMLDOMNodeList

    strXml = Voucher.GetAccessoriesInfo(errMsg)
    If errMsg <> "" Then
        MsgBox errMsg
        Exit Function
    End If
    If strXml = "" Then
        SetAttachXML = True
        Exit Function
    End If
    Set ndRS = oDomH.selectSingleNode("//rs:data")
    Set NdList = oDomH.selectNodes("//rs:data/voucherattached")
    For Each nd In NdList
        ndRS.removeChild nd
    Next
    Set NodeData = oDomH.createCDATASection(strXml)
    Set nd = oDomH.createElement("voucherattached")
    nd.appendChild NodeData
    ndRS.appendChild nd

'    Dim aa As IXMLDOMCDATASection
'    Set aa = Dombody.createCDATASection(Domhead.xml)
'    Dombody.selectNodes("//z:row").item(0).appendChild aa

    SetAttachXML = True
End Function


Private Function SetVoucherDataSource()
    Dim m_oDataSource As Object
 
 
    Set m_oDataSource = CreateObject("IDataSource.DefaultDataSource")
 
    If m_oDataSource Is Nothing Then
        MsgBox "无法创建m_oDataSource对象!"
        Exit Function
    End If
 
    Set m_oDataSource.SetLogin = m_Login
 
 
    Set Me.Voucher.SetDataSource = m_oDataSource
 
 
End Function

Private Sub RegisterMessage()
    Set m_mht = New UFPortalProxyMessage.IMessageHandler
    m_mht.MessageType = "DocAuditComplete"
    If Not g_business Is Nothing Then
        Call g_business.RegisterMessageHandler(m_mht)
    End If
End Sub
Private Sub LoadData()
    Dim rs As New ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim intCyc As Integer
    Dim introw As Long
    Dim lngVouchID As Long
    DBConn.Execute "delete from EFYZGL_pressinform_temp;delete from EFYZGL_sheet_temp"
    lngVouchID = CLng(val(Voucher.headerText("id")))
    rs.Open "select * from EFYZGL_pressinform where id=" & lngVouchID, DBConn, adOpenKeyset, adLockReadOnly
    RsTemp.Open "select * from EFYZGL_pressinform_temp", DBConn, adOpenKeyset, adLockOptimistic
    For introw = 1 To rs.RecordCount
        RsTemp.AddNew
        For intCyc = 0 To rs.Fields.Count - 1
            If rs.Fields(intCyc).Name <> "ufts" Then
                RsTemp.Fields(rs.Fields(intCyc).Name) = rs.Fields(intCyc).value
            End If
            
        Next
'        rstemp.Fields("vt_id")=''
        RsTemp.Update
    Next
    If rs.State = adStateOpen Then
        rs.Close
    End If
    If RsTemp.State = adStateOpen Then
        RsTemp.Close
    End If
    rs.Open "select * from EFYZGL_sheet where id=" & lngVouchID, , adOpenKeyset, adLockReadOnly
    RsTemp.Open "select * from EFYZGL_sheet_temp", DBConn, adOpenKeyset, adLockOptimistic
    For introw = 1 To rs.RecordCount
        RsTemp.AddNew
        For intCyc = 0 To rs.Fields.Count - 1
            If rs.Fields(intCyc).Name <> "ufts" Then
                RsTemp.Fields(rs.Fields(intCyc).Name) = rs.Fields(intCyc).value
            End If
        Next
        RsTemp.Update
        If Not rs.EOF Then
          rs.MoveNext
        End If
    Next
    If rs.State = adStateOpen Then
        rs.Close
    End If
    If RsTemp.State = adStateOpen Then
        RsTemp.Close
    End If
    Dim tempa As String
    Dim tempb As String
    Dim tempc As String
    rs.Open "select * from EFYZGL_wrappage where id=" & lngVouchID, DBConn, adOpenKeyset, adLockReadOnly
    RsTemp.Open "select * from EFYZGL_pressinform_temp", DBConn, adOpenKeyset, adLockOptimistic
    For introw = 1 To rs.RecordCount
          tempa = "f" & introw
          tempb = "ys" & introw
           RsTemp.Fields(tempa) = rs.Fields("citemname").value
           RsTemp.Fields(tempb) = rs.Fields("cdefine22").value
           RsTemp.Update
        If Not rs.EOF Then
          rs.MoveNext
        End If
    Next
    
    If rs.State = adStateOpen Then
        rs.Close
    End If
    If RsTemp.State = adStateOpen Then
        RsTemp.Close
    End If
    rs.Open "select * from EFYZGL_giveaddress where id=" & lngVouchID, DBConn, adOpenKeyset, adLockReadOnly
    RsTemp.Open "select * from EFYZGL_pressinform_temp", DBConn, adOpenKeyset, adLockOptimistic
    For introw = 1 To rs.RecordCount
          tempa = "dz" & introw
          tempb = "cs" & introw
           RsTemp.Fields(tempa) = rs.Fields("address").value
           RsTemp.Fields(tempb) = rs.Fields("iquantity").value
           RsTemp.Update
        If Not rs.EOF Then
          rs.MoveNext
        End If
    Next
    
    If rs.State = adStateOpen Then
        rs.Close
    End If
    If RsTemp.State = adStateOpen Then
        RsTemp.Close
    End If
    rs.Open "select * from EFYZGL_content where id=" & lngVouchID, DBConn, adOpenKeyset, adLockReadOnly
    RsTemp.Open "select * from EFYZGL_pressinform_temp", DBConn, adOpenKeyset, adLockOptimistic
    For introw = 1 To rs.RecordCount
          tempa = "nr" & introw
          tempb = "yz" & introw
          tempc = "ms" & introw
           RsTemp.Fields(tempa) = rs.Fields("content").value
           RsTemp.Fields(tempb) = rs.Fields("makmanner").value
           RsTemp.Fields(tempc) = rs.Fields("mianquantity").value
           RsTemp.Update
        If Not rs.EOF Then
          rs.MoveNext
        End If
    Next
    
    If rs.State = adStateOpen Then
        rs.Close
    End If
    If RsTemp.State = adStateOpen Then
        RsTemp.Close
    End If
    setVouchDate5 pbpressinformtmp
End Sub
Public Function ShowContextHelp(hwnd As Long, sHelpFile As String, lContextID As Long) As Long
    ShowContextHelp = htmlHelp(hwnd, sHelpFile, &HF, lContextID)
End Function

