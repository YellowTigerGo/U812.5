VERSION 5.00
Begin VB.Form frmHost 
   BorderStyle     =   0  'None
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   600
   ControlBox      =   0   'False
   Icon            =   "frmHost.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   420
   ScaleWidth      =   600
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
End
Attribute VB_Name = "frmHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'Me.Left = 0
    'Me.Top = 0
    Me.Width = 0
    Me.Height = 0
    RegWrite CINSTLANGID & "Portal", "HostWindowHandle", CStr(Me.hWnd)

End Sub

Private Sub Form_Resize()
    'Me.Left = 0
    'Me.Top = 0
    Me.Width = 0
    Me.Height = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
