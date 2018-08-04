Attribute VB_Name = "modMain"
Option Explicit

'
' Progarm Entry Point
'
Sub Main()
    Dim oNPASink As NPASink
    
    ' run vb runtime message loop
    frmHost.Show

    Set oNPASink = New NPASink
    oNPASink.Startup frmHost.hWnd

End Sub
