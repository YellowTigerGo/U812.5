VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSort 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "≈≈–Ú…Ë÷√"
   ClientHeight    =   4920
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '∆¡ƒª÷––ƒ
   Begin MSFlexGridLib.MSFlexGrid msgSort 
      Height          =   4260
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4956
      _ExtentX        =   8731
      _ExtentY        =   7514
      _Version        =   393216
      RowHeightMin    =   300
      BackColorBkg    =   12632256
      ScrollBars      =   2
      Appearance      =   0
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "»∑»œ"
      Default         =   -1  'True
      Height          =   300
      Left            =   2700
      TabIndex        =   1
      Top             =   4500
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "»°œ˚"
      Height          =   300
      Left            =   3945
      TabIndex        =   0
      Top             =   4500
      Width           =   1125
   End
   Begin ComctlLib.ImageList imgSort 
      Left            =   840
      Top             =   4440
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSort.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSort.frx":005E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSort.frx":00BC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cOrder As String
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSure_Click()
    Dim i As Integer
    On Error Resume Next
    cOrder = ""
    With msgSort
        For i = 1 To .rows - 1
            If .TextMatrix(i, 2) <> "" Then
                If .TextMatrix(i, 2) = "…˝–Ú" Then
                    cOrder = IIf(Trim(cOrder) = "", .TextMatrix(i, 1), cOrder & ", " & .TextMatrix(i, 1))
                End If
                If .TextMatrix(i, 2) = "Ωµ–Ú" Then
                    cOrder = IIf(Trim(cOrder) = "", .TextMatrix(i, 1) & " DESC", cOrder & ", " & .TextMatrix(i, 1) & " DESC")
                End If
            End If
        Next i
    End With
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    With msgSort
        .rows = UBound(vSort) + 1
        .FixedCols = 1
        .FixedRows = 1
        .cols = 4
        .FormatString = "<œÓƒø√˚≥∆|<œÓƒø¥˙¬Î|<≈≈–Ú∑Ω Ω"
        For i = 1 To UBound(vSort)
            .TextMatrix(i, 0) = vSort(i - 1).sName
            .TextMatrix(i, 1) = vSort(i - 1).sCode
            Select Case vSort(i - 1).iSort
                Case 0
                    .TextMatrix(i, 2) = ""
                Case 1
                    .TextMatrix(i, 2) = "…˝–Ú"
                Case 2
                    .TextMatrix(i, 2) = "Ωµ–Ú"
            End Select
            .RowHeight(i) = imgSort.ListImages(1).Picture.Height
        Next i
        .FixedAlignment(0) = 4: .colwidth(0) = .Width / 2
        .FixedAlignment(1) = 4: .colwidth(1) = 0
        .FixedAlignment(2) = 4: .colwidth(2) = .Width / 2
        InitGrdCol msgSort
    
    End With
End Sub

Private Sub msgSort_DblClick()
    Select Case msgSort.TextMatrix(msgSort.RowSel, 2)
        Case ""
            msgSort.TextMatrix(msgSort.RowSel, 2) = "…˝–Ú"
        Case "…˝–Ú"
            msgSort.TextMatrix(msgSort.RowSel, 2) = "Ωµ–Ú"
        Case "Ωµ–Ú"
            msgSort.TextMatrix(msgSort.RowSel, 2) = ""
    End Select
End Sub

