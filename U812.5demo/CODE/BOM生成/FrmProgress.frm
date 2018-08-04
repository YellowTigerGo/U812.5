VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmProgress 
   Caption         =   "请稍后......."
   ClientHeight    =   555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   555
   ScaleWidth      =   6720
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label Label1 
      Caption         =   "数据正在进行处理，请稍后....."
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   6615
   End
End
Attribute VB_Name = "FrmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_Cancel, m_UnloadMode As Integer
Attribute m_UnloadMode.VB_VarUserMemId = 1073938432
'每个窗体都需要这个方法。Cancel与UnloadMode的参数的含义与QueryUnload的参数相同。
Public Sub ExitForm(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    Cancel = m_Cancel
    UnloadMode = m_UnloadMode

End Sub

Private Sub Form_Load()
    Me.Icon = LoadResPicture(166, vbResIcon)
End Sub
