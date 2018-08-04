VERSION 5.00
Object = "{9FD12F62-6922-47E1-B1AC-3615BBD3D7A5}#1.0#0"; "UFLabel.ocx"
Object = "{86808282-58F4-4B17-BBCA-951931BB7948}#2.82#0"; "U8VouchList.ocx"
Begin VB.Form frmProgress 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1230
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9285
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "frmProgress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin UFLABELLib.UFLabel LblMsg 
      Height          =   1215
      Left            =   840
      TabIndex        =   1
      Top             =   0
      Width           =   8415
      _Version        =   65536
      _ExtentX        =   14843
      _ExtentY        =   2143
      _StockProps     =   111
   End
   Begin U8VouchList.ProgressAnimation ProgressAnimation1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   1296
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Property Let Msg(val As String)
    LblMsg.caption = val
End Property
