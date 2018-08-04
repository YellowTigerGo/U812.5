VERSION 5.00
Object = "{456334B9-D052-4643-8F5F-2326B24BE316}#6.96#0"; "UAPvouchercontrol85.ocx"
Begin VB.Form Frm_voucher 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   20370
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin UAPVoucherControl85.ctlVoucher ctlVoucher1 
      Height          =   6015
      Left            =   720
      TabIndex        =   0
      Top             =   1680
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   10610
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10446406
      DisabledColor   =   16777215
      ColAlignment0   =   9
      Rows            =   20
      Cols            =   20
      TitleForecolor  =   16050403
      ControlScrollBars=   0
      ControlAutoScales=   0
      BaseOfVScrollPoint=   0
      ShowSorter      =   0   'False
      ShowFixColer    =   0   'False
   End
End
Attribute VB_Name = "Frm_voucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
