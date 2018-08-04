VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E532C1D0-3741-4753-ADA3-38891BC9C7FB}#5.1#0"; "U8Flow.ocx"
Begin VB.Form frmFlow 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "×ÀÃæ"
   ClientHeight    =   6180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10200
   ControlBox      =   0   'False
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   Begin FlowChrt.CFlowChart oFlowChart 
      Height          =   1185
      Left            =   705
      TabIndex        =   0
      Top             =   2475
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2090
   End
   Begin MSComctlLib.ImageList imgsTitle 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   800
      ImageHeight     =   62
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFlow.frx":0000
            Key             =   "Small1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFlow.frx":005E
            Key             =   "Large1"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFlow.frx":00BC
            Key             =   "Large0"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFlow.frx":011A
            Key             =   "Small0"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFlow.frx":0178
            Key             =   "Small2"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFlow.frx":01D6
            Key             =   "Large2"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFlow.frx":0234
            Key             =   "Small3"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFlow.frx":0292
            Key             =   "Large3"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblCaption 
      Caption         =   "Label1"
      Height          =   480
      Left            =   915
      TabIndex        =   2
      Top             =   1245
      Width           =   2190
   End
   Begin VB.Label ReturnValue 
      AutoSize        =   -1  'True
      Caption         =   "False"
      Height          =   180
      Left            =   6255
      TabIndex        =   1
      Top             =   3210
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgTitle 
      Height          =   930
      Left            =   -30
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "frmFlow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_cSubId As String
Private m_cUserId As String
Private m_cSysConnString As String
Private m_cAccId As String
Private m_oDomMenu As Object

Private Sub Form_Activate()
    ResizeFloatForm imgTitle.Height, True
    Me.Width = gd_frmMain.ScaleWidth
    Me.Height = gd_frmMain.ScaleHeight
End Sub

Private Sub Form_Deactivate()
    ResizeFloatForm 0, False
    Me.Width = gd_frmMain.ScaleWidth
    Me.Height = gd_frmMain.ScaleHeight
End Sub

 Private Function SplitCommand(ByVal cCmds As String, cMenuId As String, cMenuName As String, cAuthId As String, cCmdLine As String) As Boolean
     Dim cReturn As Variant
     cReturn = Split(cCmds, ",")
     If UBound(cReturn) > -1 Then cCmds = cReturn(0)
     If UBound(cReturn) > 0 Then cCmdLine = cReturn(1)
     
     cReturn = Split(cCmds, Chr(9))
     If UBound(cReturn) > -1 Then cMenuId = cReturn(0)
     If UBound(cReturn) > 0 Then cMenuName = cReturn(1)
     If UBound(cReturn) > 1 Then cAuthId = cReturn(2)
     'If UBound(cReturn) > 2 Then cCmdLine = cReturn(3)
End Function
 
Private Sub Form_LinkExecute(cCmdStr As String, Cancel As Integer)
    Dim cMenuId As String
    Dim cMenuName As String
    Dim cAuthId As String
    Dim cCmdLine As String
    SplitCommand cCmdStr, cMenuId, cMenuName, cAuthId, cCmdLine
    If cMenuId <> "QueryState" Then
        oFlowChart_OnCommand cMenuId, cMenuName, cAuthId, cCmdLine
        g_bCommandFromPortal = True
        Debug.Assert False
    End If
    Cancel = 0
End Sub

Private Sub Form_LinkOpen(Cancel As Integer)
    Cancel = 0
End Sub

Private Sub ChangeCaption()
    If g_typeOther.bUserDef Then
        On Error Resume Next
        Dim pic As IPictureDisp
        Set pic = LoadPicture(g_typeOther.cPicPath)
        If Not pic Is Nothing Then
            Me.BackColor = g_typeOther.nBackColor
            Set imgTitle.Picture = pic
            Set pic = Nothing
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
    End If

    If Screen.Width / Screen.TwipsPerPixelX < 1024 Then
        imgTitle.Picture = imgsTitle.ListImages("Small" & g_nStyle).Picture
    Else
        imgTitle.Picture = imgsTitle.ListImages("Large" & g_nStyle).Picture
    End If
    
End Sub

Private Sub Form_Resize()
    GetApperance
    ChangeCaption
    
    imgTitle.Left = 0
    imgTitle.Top = 0
    imgTitle.Height = 64 * Screen.TwipsPerPixelY
    imgTitle.Width = Screen.Width
    imgTitle.Visible = True
        
    oFlowChart.Left = 0
    oFlowChart.Top = imgTitle.Height
    oFlowChart.Width = Me.ScaleWidth
    oFlowChart.Height = IIf(Me.ScaleHeight - imgTitle.Height > 0, Me.ScaleHeight - imgTitle.Height, 0)
    
    If g_typeOther.bUserDef Then
        lblCaption.FontBold = True
        lblCaption.AutoSize = True
        lblCaption.Caption = g_typeOther.ctext
        lblCaption.BackStyle = 0
        lblCaption.ForeColor = g_typeOther.nTextColor
        lblCaption.FontSize = g_typeOther.nTextSize
        lblCaption.FontName = g_typeOther.cTextFont
        lblCaption.Left = 0
        lblCaption.Top = (imgTitle.Height - lblCaption.Height) / 2
    End If
    lblCaption.Visible = g_typeOther.bUserDef
End Sub

Public Function GetTitleHeight()
    GetTitleHeight = imgTitle.Height + imgTitle.Top
End Function


Private Sub Form_Load()
    Me.LinkMode = vbLinkSource
    Me.LinkTopic = "SuperLink"
End Sub

Public Function SaveChart(ByVal cSubId As String, ByVal cUserId As String, ByVal cSysConnString As String)
        oFlowChart.SaveFlowChart cSubId, cUserId, cSysConnString
End Function

Public Function LoadChart(ByVal cSubId As String, ByVal cUserId As String, ByVal cSysConnString As String, ByVal oDomMenu As Object, ByVal cAccId As String)
    m_cSubId = cSubId
    m_cUserId = cUserId
    m_cSysConnString = cSysConnString
    m_cAccId = cAccId
    oFlowChart.SetAccId cAccId
    oFlowChart.LoadFlowChart cSubId, cUserId, cSysConnString
    Me.Refresh
    Set m_oDomMenu = oDomMenu
    oFlowChart.SetMenu oDomMenu
End Function
 
Private Sub oFlowChart_OnCommand(ByVal cMenuId As String, ByVal cMenuName As String, ByVal cAuthId As String, ByVal cCmdLine As String)
    Select Case cMenuId
    Case "SaveMe"
        SaveChart m_cSubId, m_cUserId, m_cSysConnString
    Case "CancelMe"
        oFlowChart.LoadFlowChart m_cSubId, m_cUserId, m_cSysConnString
    Case Else
        gd_frmMain.m_oMenu_OnCommand cMenuId, cMenuName, cAuthId, cCmdLine
    End Select
End Sub
