VERSION 5.00
Begin VB.Form frmFloat 
   BorderStyle     =   0  'None
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   1170
   ShowInTaskbar   =   0   'False
   Begin VB.Timer m_oTimer 
      Left            =   480
      Top             =   960
   End
End
Attribute VB_Name = "frmFloat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_nMaxWidth As Single

Private Sub Form_Load()
    SetHook AddressOf MouseProc, App.hInstance, App.ThreadID
    m_oTimer.Interval = 10
    m_oTimer.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_oTimer.Enabled = False
End Sub

Private Function IsOnMe() As Boolean
    Dim pnt As POINTAPI
    Dim X As Single, Y As Single
    pnt.X = g_x
    pnt.Y = g_y
    ScreenToClient gd_frmMain.hwnd, pnt
    X = pnt.X * Screen.TwipsPerPixelX
    Y = pnt.Y * Screen.TwipsPerPixelY + Me.Top
    If X <= Me.Left + Me.Width + 30 And _
        Y >= Me.Top And Y <= Me.Top + Me.Height Then
        IsOnMe = True
    Else
        IsOnMe = False
    End If
End Function

Private Sub m_oTimer_Timer()
    If IsOnMe Then
        Me.Left = 0
    Else
        Me.Left = Screen.TwipsPerPixelX - Me.m_nMaxWidth
    End If
End Sub


