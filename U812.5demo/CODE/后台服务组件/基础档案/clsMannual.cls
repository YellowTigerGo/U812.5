VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMannual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------
' 文 件 名: clsMannual.cls
'
' 功 能: GSP分类
'
' 作 者: 王良彬    时  间:     2002
' 北就用友软件集团 版权所有 Copyright(c) 2002
'--------------------------------------------

Option Base 0
Option Explicit

'<属性
Private WithEvents frm          As frmType                      '分类窗体类
Attribute frm.VB_VarHelpID = -1

Private mstrTable               As String                       '表名
Private mstrArTable             As String                       '档案表名
Private mstrCode                As String                       '分类编号
Private mstrArVouchType         As String                       '对应档案的单据编号
Private mstrArCaption           As String                       '分类标题
Private mstrUfts                As String                       '时间戳
Private mstrTxtBeforeChange     As String                       'textbox修改前的值

'auth
Private mstrAdd_auth            As String                       '增加


Private mlngRsCount             As Long                         '当前纪录数
Private Opt                     As Long                         '操作类型：1、添加；2、修改；其它：无操作
Private mlngIndex               As Long                         '树当前操作节点
Private mlngBColor              As Long                         'TreeView当前节点的颜色


Private mbolEnd                 As Boolean                      '是否是末节点
Private mbolCollapse            As Boolean                      '是否展开节点
Private mbolAddCheck            As Boolean                      '控制增加时是否判断可否增加
Private mbolChange              As Boolean
'是否已作修改
'/属性>



Public Property Get Caption() As String
    Caption = mstrCaption
End Property
Public Property Let Caption(V As String)
    mstrCaption = V
End Property




'-----------------------------------------------------------
'功能：初始化属性，并显示分类
'
'参数：lType档案类型(3、标准档案，4、管理制度档案....)，sCaption 档案标题， tbl toolbar显示类型，sAdd_auth编辑权限
'
'返回：
'
'-----------------------------------------------------------

Public Sub DOModal(lType As Long, sCaption As String, tbl As Long, sAdd_auth As String)
    '标题
    mstrCaption = sCaption
      Set frm = New frmType
    
'    Select Case tbl
'    Case 0
'        TBLStyle = TBLNormal
'    Case Else
'        TBLStyle = TBLText
'    End Select
    TBLStyle = tbl
    '权限编码
    mstrAdd_auth = sAdd_auth
    ''''
    Select Case lType
    Case 1                                      '开本档案
        mstrTable = "EFBWGL_DBFormat"           '档案表名
        mstrArVouchType = "080"                 '档案单据编号
        mbolAddCheck = True
        frm.Label4.Visible = False
        frm.Check1.Visible = False
        frm.Label5.Visible = False
        frm.Check2.Visible = False
        mstrHelpID = 20090101                   '设置帮助编号
        frm.HelpContextID = 20090101
        frm.Picture2.HelpContextID = 20090101
        mstrArCaption = "开本档案"
    
    Case 2                                      '装帧档案
        mstrTable = "EFBWGL_DBBookBinding"      '档案表名
        mstrArVouchType = "080"                 '档案单据编号
        mbolAddCheck = True
        frm.Label4.Visible = False
        frm.Check1.Visible = False
        frm.Label5.Visible = False
        frm.Check2.Visible = False
        mstrHelpID = 20090102                   '设置帮助编号
        frm.HelpContextID = 20090102
        frm.Picture2.HelpContextID = 20090102
        mstrArCaption = "装帧档案"
    Case 3                                      '文种档案
        mstrTable = "EFBWGL_DBBookLanguage"      '档案表名
        mstrArVouchType = "080"                 '档案单据编号
        mbolAddCheck = True
        frm.Label4.Visible = False
        frm.Check1.Visible = False
        frm.Label5.Visible = False
        frm.Check2.Visible = False
        mstrHelpID = 20090103                   '设置帮助编号
        frm.HelpContextID = 20090103
        frm.Picture2.HelpContextID = 20090103
        mstrArCaption = "文种档案"
    
    Case 4                                      '图书分类档案
        mstrTable = "EFBWGL_DBBookClass"      '档案表名
        mstrArVouchType = "080"                 '档案单据编号
        mbolAddCheck = True
        frm.Label4.Visible = False
        frm.Check1.Visible = False
        frm.Label5.Visible = False
        frm.Check2.Visible = False
        mstrHelpID = 20090104                   '设置帮助编号
        frm.HelpContextID = 20090104
        frm.Picture2.HelpContextID = 20090104
        mstrArCaption = "图书分类档案"
    
    Case 5                                      '合同条款档案
        mstrTable = "EFBWGL_DBCBHT"      '档案表名
        mstrArVouchType = "080"                 '档案单据编号
        mbolAddCheck = True
        frm.Label4.Visible = False
        frm.Check1.Visible = False
        frm.Label5.Visible = False
        frm.Check2.Visible = False
        mstrHelpID = 20090105                   '设置帮助编号
        frm.HelpContextID = 20090105
        frm.Picture2.HelpContextID = 20090105
        mstrArCaption = "合同条款档案"
        
    Case 6                                     '著作方式档案
        mstrTable = "EFBWGL_DBBMMode"      '档案表名
        mstrArVouchType = "080"                 '档案单据编号
        mbolAddCheck = True
        frm.Label4.Visible = False
        frm.Check1.Visible = False
        frm.Label5.Visible = False
        frm.Check2.Visible = False
        mstrHelpID = 20090106                   '设置帮助编号
        frm.HelpContextID = 20090106
        frm.Picture2.HelpContextID = 20090106
        mstrArCaption = "著作方式档案"
          
    Case 7                                     '类型词档案
        mstrTable = "EFBWGL_DBTypeWord"      '档案表名
        mstrArVouchType = "080"                 '档案单据编号
        mbolAddCheck = True
        frm.Label4.Visible = False
        frm.Check1.Visible = False
        frm.Label5.Visible = False
        frm.Check2.Visible = False
        mstrHelpID = 20090107                   '设置帮助编号
        frm.HelpContextID = 20090107
        frm.Picture2.HelpContextID = 20090107
        mstrArCaption = "类型词档案"
             
    Case 8                                     '选题类别档案
        mstrTable = "EFBWGL_DBBookSelStress"      '档案表名
        mstrArVouchType = "080"                 '档案单据编号
        mbolAddCheck = True
        frm.Label4.Visible = False
        frm.Check1.Visible = False
        frm.Label5.Visible = False
        frm.Check2.Visible = False
        mstrHelpID = 20090108                   '设置帮助编号
        frm.HelpContextID = 20090108
        frm.Picture2.HelpContextID = 20090108
        mstrArCaption = "选题类别档案"
               
    Case 9                                     ' 进程档案
        mstrTable = "EFBWGL_DBTenor"      '档案表名
        mstrArVouchType = "080"                 '档案单据编号
        mbolAddCheck = True
        frm.Label4.Visible = False
        frm.Check1.Visible = False
        frm.Label5.Visible = False
        frm.Check2.Visible = False
        mstrHelpID = 20090109                   '设置帮助编号
        frm.HelpContextID = 20090109
        frm.Picture2.HelpContextID = 20090109
        mstrArCaption = "进程档案"
                  
    Case 10                                     ' 申报类型档案
        mstrTable = "EFBWGL_DBBookDeclClass"      '档案表名
        mstrArVouchType = "080"                 '档案单据编号
        mbolAddCheck = True
        frm.Label4.Visible = False
        frm.Check1.Visible = False
        frm.Label5.Visible = False
        frm.Check2.Visible = False
        mstrHelpID = 20090110                   '设置帮助编号
        frm.HelpContextID = 20090110
        frm.Picture2.HelpContextID = 20090110
        mstrArCaption = "申报类型档案"
                     
    Case 11                                     ' 合同类型档案
        mstrTable = "EFBWGL_DBHTClass"      '档案表名
        mstrArVouchType = "080"                 '档案单据编号
        mbolAddCheck = True
        frm.Label4.Visible = False
        frm.Check1.Visible = False
        frm.Label5.Visible = False
        frm.Check2.Visible = False
        mstrHelpID = 20090111                   '设置帮助编号
        frm.HelpContextID = 20090111
        frm.Picture2.HelpContextID = 20090111
        mstrArCaption = "合同类型档案"
                         
    Case 12                                     ' 装订方式档案
        mstrTable = "EFBWGL_DBbindmanner"      '档案表名
        mstrArVouchType = "080"                 '档案单据编号
        mbolAddCheck = True
        frm.Label4.Visible = False
        frm.Check1.Visible = False
        frm.Label5.Visible = False
        frm.Check2.Visible = False
        mstrHelpID = 20090112                   '设置帮助编号
        frm.HelpContextID = 20090112
        frm.Picture2.HelpContextID = 20090112
        mstrArCaption = "装订方式档案"
                             
    Case 13                                     ' 装订方法档案
        mstrTable = "EFBWGL_DBbindway"      '档案表名
        mstrArVouchType = "080"                 '档案单据编号
        mbolAddCheck = True
        frm.Label4.Visible = False
        frm.Check1.Visible = False
        frm.Label5.Visible = False
        frm.Check2.Visible = False
        mstrHelpID = 20090113                   '设置帮助编号
        frm.HelpContextID = 20090113
        frm.Picture2.HelpContextID = 20090113
        mstrArCaption = "装订方法档案"
                                 
    Case 14                                    ' 出版方式档案
        mstrTable = "EFBWGL_DBBookPubMode"      '档案表名
        mstrArVouchType = "080"                 '档案单据编号
        mbolAddCheck = True
        frm.Label4.Visible = False
        frm.Check1.Visible = False
        frm.Label5.Visible = False
        frm.Check2.Visible = False
        mstrHelpID = 20090114                   '设置帮助编号
        frm.HelpContextID = 20090114
        frm.Picture2.HelpContextID = 20090114
        mstrArCaption = "出版方式档案"
    End Select
    
    '--------------------------------------------
    Call g_oPub.FrmLoadInit(frm, mstrTable, g_oLogin, SrvDB)                '初始化g_oPub
    frm.Label3.Caption = sCaption
    frm.Stb.Panels(3).text = "当前记录数：" + Trim(mlngRsCount) + "条"
    Load frm
    'frm.Show 1
    
    If (g_business Is Nothing) Then
        frm.Show 1
    Else
        Call g_business.ShowForm(frm, "GS", CreateGUID, True, True)
    End If
    
End Sub

'-----------------------------------------------------------
'功能：加载树节点
'
'参数：
'
'返回：
'
'-----------------------------------------------------------
Public Sub LoadTree()
    Dim i               As Long
    Dim recR            As New ADODB.Recordset      '树根节点
    
    DeleteRec mstrTable
    mlngRsCount = 0
    recR.CursorLocation = adUseClient
    With frm.TREE1
        .Nodes.Clear
        .Nodes.Add , , "K", mstrCaption
        recR.Open "SELECT * FROM " & mstrTable & " WHERE ISNULL(CPARENTNODE,'')='' AND ISNULL(CCODE,'')<>'' order by ccode,cparentnode", AdoCnn, adOpenKeyset, adLockOptimistic
        If mstrTable = "GSP_DEVICETYPE" Then AdoCnn.Execute "UPDATE " & mstrTable & " SET BSYSTEM=0 WHERE BSYSTEM IS NULL"
        If recR Is Nothing Then Exit Sub
        '先填充根结点
        Do While Not recR.EOF
            mlngRsCount = mlngRsCount + 1
            If mstrTable = "GSP_DEVICETYPE" Then
                .Nodes.Add "K", tvwChild, "K" & Trim(recR.Fields("CCODE").value), "(" & Trim(recR.Fields("CCODE").value) & ") " & Trim(recR.Fields("CNAME").value) & IIf(recR.Fields("BSYSTEM") = True, "[系统]", "")
            Else
                .Nodes.Add "K", tvwChild, "K" & Trim(recR.Fields("CCODE").value), "(" & Trim(recR.Fields("CCODE").value) & ") " & Trim(recR.Fields("CNAME").value)
            End If
            FillTree Trim(recR.Fields("CCODE").value), "K" & Trim(recR.Fields("CCODE").value)
            recR.MoveNext
        Loop
        recR.Close
        Set recR = Nothing
    End With
    '设置默认选项
    If mlngIndex = 1 Or mstrCode = "" Then
        frm.TREE1.Nodes(1).Selected = True
        mlngIndex = 1
        frm.TREE1.Nodes(1).Expanded = True
    Else
        For i = 1 To frm.TREE1.Nodes.Count
            If InStr(Trim(frm.TREE1.Nodes(i).text), "(" & mstrCode & ")") = 1 Then
                mlngIndex = frm.TREE1.Nodes(i).Index
                Exit For
            End If
        Next
        '展开树节点
        If i > frm.TREE1.Nodes.Count Then       '没找到刷新前的节点
            mlngIndex = 1
        End If
        i = mlngIndex
        frm.TREE1.Nodes(i).Expanded = True
        Do While i <> 1
            i = frm.TREE1.Nodes(i).Parent.Index
            frm.TREE1.Nodes(i).Expanded = True
        Loop
        frm.TREE1.Nodes(mlngIndex).Selected = True
        frm.TREE1.Nodes(mlngIndex).BackColor = mlngBColor + 10
    End If
    frm.Stb.Panels(3).text = "当前记录数：" + Trim(mlngRsCount) + "条"
End Sub

'-----------------------------------------------------------
'功能：填充节点的子节点
'
'参数：sCode要填充的节点，sParent要填充的节点的父节点
'
'返回：
'
'-----------------------------------------------------------
Sub FillTree(sCode As String, sParent As String)
    Dim rec     As New ADODB.Recordset
    
    '查询编码为sCode节点的子节点
    rec.CursorLocation = adUseClient
    rec.Open "SELECT * FROM " & mstrTable & " WHERE CPARENTNODE='" & sCode & "' AND ISNULL(CCODE,'')<>'' order by ccode", AdoCnn, adOpenKeyset, adLockOptimistic
    If rec Is Nothing Then Exit Sub
    '递归增加子节点
    Do While Not rec.EOF
        With frm.TREE1
            mlngRsCount = mlngRsCount + 1
            .Nodes.Add "K" & sCode, tvwChild, "K" & Trim(rec.Fields("CCODE").value), "(" & Trim(rec.Fields("CCODE").value) & ") " & Trim(rec.Fields("CNAME").value)
            FillTree Trim(rec.Fields("CCODE").value), Trim(rec.Fields("CPARENTNODE").value) & Trim(rec.Fields("CCODE").value)
        End With
        rec.MoveNext
    Loop
    rec.Close
    Set rec = Nothing
End Sub
Private Sub Class_Initialize()
    Opt = 0
    mlngIndex = 1
    mlngBColor = -1
    mbolChange = False
End Sub


'-----------------------------------------------------------
'功能：判断当前编号的分类是否存在
'
'参数：
'
'返回：判断档案是否存在
'
'-----------------------------------------------------------
Public Function Exist() As Boolean
    Dim rec As New ADODB.Recordset
    
    Exist = True
    rec.CursorLocation = adUseClient
    rec.Open "select * from " & mstrTable & " where ccode='" & frm.txtCode & "' AND ( convert(char,Convert(money,Ufts),2) ='" & mstrUfts & "'" & IIf(mstrUfts <> "", "", " OR UFTS=null ") & ")", AdoCnn, adOpenKeyset, adLockPessimistic
    If rec.RecordCount <= 0 Then
        rec.Close
        Exist = False
        ShowMsg "该分类已被他人更新或删除，请刷新后再试！", 1
        Exit Function
    End If
    rec.Close
    Set rec = Nothing
End Function

'-----------------------------------------------------------
'功能：点击toolbar按钮
'
'参数：sKey toolbar 按钮的Key
'
'返回：
'
'-----------------------------------------------------------
Private Sub frm_ButtonClick(sKey As String)
    Dim lngTemp     As Long
    Dim rec         As New ADODB.Recordset
    On Error GoTo ErrorLine
    Select Case sKey
    Case "SetUp", "Print", "Preview", "SaveFile"                '打印、预览、输出
        PrintAll sKey, frm.Prn, frm, mstrTable, mstrCaption
    Case "Add"
        If mbYearEnd Then                                       '是否已年结
            ShowMsg msg_YearEnd, 0
            Exit Sub
        End If
        If Not UA_Task(mstrAdd_auth) Then Exit Sub              '申请权限
        If mlngIndex > 1 Then                                   '查询当前节点是否已被删除
            rec.Open "select * from " & mstrTable & " where ccode='" & Mid(Trim(frm.TREE1.SelectedItem.text), 2, InStr(Trim(frm.TREE1.SelectedItem.text), ")") - 2) & "'", AdoCnn, adOpenKeyset, adLockPessimistic
            If rec.RecordCount <= 0 Then
                rec.Close
                Call UA_FreeTask(mstrAdd_auth)                   '释放权限
                ShowMsg "该分类已被别人删除，不能添加子分类，请刷新后再试！", 1
                Exit Sub
            End If
            rec.Close
        End If
        
        Opt = 1
        Call SetTlbButton(sKey)                                     '设置toolbar按钮的enable状态
    Case "Modify"
        If mbYearEnd Then
            ShowMsg msg_YearEnd, 0
            Exit Sub
        End If
        If bBaseSysValue(mstrTable, frm.txtCode.text, "") Then
            ShowMsg "系统标准档案数据不能更新或删除！", 1
            Exit Sub
        End If
        If Not UA_Task(mstrAdd_auth) Then Exit Sub                  '申请权限
        
        If Not Exist Then
            Call UA_FreeTask(mstrAdd_auth)                          '释放权限
            Exit Sub
        End If
        Opt = 2
        Call SetTlbButton(sKey)
    Case "Delete"
        If mbYearEnd Then
            ShowMsg msg_YearEnd, 0
            Exit Sub
        End If
        If bBaseSysValue(mstrTable, frm.txtCode.text, "") Then
            ShowMsg "系统标准档案数据不能更新或删除！", 1
            Exit Sub
        End If
        If Not UA_Task(mstrAdd_auth) Then Exit Sub         '申请权限
        Delete
        Call UA_FreeTask(mstrAdd_auth)                     '释放权限
    Case "Back"
        Call UA_FreeTask(mstrAdd_auth)
        Back
'        Call SetTlbButton(sKey)
        mbolChange = False
    Case "Help"
'       SendKeys "{F1}"
         On Error Resume Next
         ShowContextHelp frm.hwnd, App.HelpFile, frm.HelpContextID
    Case "SaveRs"
        If Save(sKey) <> "" Then Exit Sub
        Call SetTlbButton(sKey)
        Opt = 0
        mbolChange = False
        Call UA_FreeTask(mstrAdd_auth)                     '释放权限
    Case "Refresh"
        If Opt = 1 Or Opt = 2 Then
            If MsgBox("是否放弃当前操作", vbQuestion + vbYesNo, "提示") = vbNo Then Exit Sub
        End If
        frm_ButtonClick "Back"
        LoadTree
        lngTemp = mlngIndex
        mlngIndex = 1
        frm_TreeNodeClick frm.TREE1.Nodes(lngTemp)
    Case "btnSep4"

    Case "TBL"
        If frm.mnuTBL.Caption = "文本按钮(&T)" Then
            TBLStyle = TBLText
            frm.CTBCtrl1.SetToolbar frm.Tlb
            frm.CTBCtrl1.SetDisplayStyle TextOnly
            frm.CTBCtrl1.RefreshVisible
            frm.Tlb.Visible = False
            frm.CTBCtrl1.Visible = True
            frm.CTBCtrl1.ZOrder 0
            frm.mnuTBL.Caption = "图标按钮(&T)"
        Else
            frm.mnuTBL.Caption = "文本按钮(&T)"
            TBLStyle = TBLNormal
            frm.Tlb.Visible = True
            frm.CTBCtrl1.Visible = False
            frm.Tlb.ZOrder 0
        End If
    Case "Status"
        frm.mnuStatus.Checked = Not frm.mnuStatus.Checked
        frm.Stb.Visible = frm.mnuStatus.Checked
        Resize frm
    Case "Exit"
        If frm.Tlb.Buttons("SaveRs").Enabled = True And mbolChange = True Then
            If ShowMsg("是否保存当前记录的编辑？", 2) = 6 Then
                If Save(sKey) <> "" Then Exit Sub
            End If
        End If
        Call UA_FreeTask(mstrAdd_auth)
        frm.Hide
    End Select
    Exit Sub
ErrorLine:
    Err.Clear
End Sub


'-----------------------------------------------------------
'功能：窗体的checkbox引发的事件
'
'参数：
'
'返回：
'
'-----------------------------------------------------------
Private Sub frm_Check()
    EndCheck
    mbolChange = True
    setCTBEnable frm.CTBCtrl1, frm.Tlb, "SaveRs", True
    frm.mnuSave.Enabled = True
End Sub


'-----------------------------------------------------------
'功能：末级节点和非末级节点转换的合法判断，frm.Check2.Value = 1 表示要变成末级节点
'
'参数：
'
'返回：
'
'-----------------------------------------------------------
Private Function EndCheck() As Boolean
    Dim rec As New ADODB.Recordset
    EndCheck = True
    If frm.txtCode = "" Then Exit Function
    If mlngIndex = 1 Then Exit Function
    If Opt = 1 Then Exit Function                       '增加不用判断
    If Not mbolEnd And frm.Check2.value = 1 Then        '从飞末级节点改成末级节点
        If frm.TREE1.Nodes(mlngIndex).children > 0 Then
            frm.Check2.value = IIf(mbolEnd, 1, 0)
            EndCheck = False
            ShowMsg "该节点不能设为末节点！", 1
        End If
    End If
    If Not mbolAddCheck Then Exit Function              '不用做增加按钮判断
    If mbolEnd And frm.Check2.value = 0 Then            '如果该分类有档案，不能从末级节点改为非末级节点
        rec.CursorLocation = adUseClient
        If mstrTable = "GSP_DEVICETYPE" Then
            rec.Open "SELECT * FROM gsp_deviceinfo WHERE cdevicetypeid='" & Mid(Trim(frm.TREE1.Nodes(mlngIndex).text), 2, InStr(Trim(frm.TREE1.Nodes(mlngIndex).text), ")") - 2) & "'", AdoCnn, adOpenKeyset, adLockOptimistic
        Else
            rec.Open "SELECT * FROM " & mstrArTable & " WHERE CPARENTNODE='" & Mid(Trim(frm.TREE1.Nodes(mlngIndex).text), 2, InStr(Trim(frm.TREE1.Nodes(mlngIndex).text), ")") - 2) & "' AND CVOUCHTYPE='" & mstrArVouchType & "'", AdoCnn, adOpenKeyset, adLockOptimistic
        End If
        If rec.RecordCount > 0 Then
            EndCheck = False
            rec.Close
            frm.Check2.value = IIf(mbolEnd, 1, 0)
            ShowMsg "该末节点已经有档案，不能改为非末级节点！", 1
            
        Else
            rec.Close
        End If
    End If
    Set rec = Nothing
End Function

Private Sub frm_FormKey(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
'        Case 112    'F1
'            If frm.Tlb.Buttons("Help").Enabled Then frm_ButtonClick "Help"
        Case 113    'F2
'            If cmdRef.Visible Then cmdRef_Click
        Case 114    'F3
        
        Case 115    'F4
            If Shift = 2 Then
                If frm.Tlb.Buttons("Exit").Enabled Then
                    frm_ButtonClick "Exit"
                Else
'                    MsgBox "现在不可以退出！", vbInformation, Msg_Title
                End If
            End If
'        Case 116    'F5
'            If frm.Tlb.Buttons("Add").Enabled Then
'                frm_ButtonClick "Add"
'            Else
''                MsgBox "是末级节点不能增加！", vbInformation, Msg_Title
'            End If
'        Case 117    'F6
'            If frm.Tlb.Buttons("SaveRs").Enabled Then
'                frm_ButtonClick "SaveRs"
'            Else
''                MsgBox "不处在增加或修改状态不能保存！", vbInformation, Msg_Title
'            End If
'        Case 119    'F8
'            If frm.Tlb.Buttons("Modify").Enabled Then
'                frm_ButtonClick "Modify"
'            Else
''                MsgBox "当前状态不能修改或该节点不能修改！", vbInformation, Msg_Title
'            End If
'        Case 46     'Delete
'            If frm.Tlb.Buttons("Delete").Enabled Then
'                frm_ButtonClick "Erase"
'            Else
'                If frm.TREE1.Enabled = True Then
''                    MsgBox "当前状态不能删除或该节点不能删除！", vbInformation, Msg_Title
'                Else
'                    If frm.ActiveControl.Enabled = True And frm.Tlb.Buttons("Back").Enabled = True Then
'                        mbolChange = True
'                        setCTBEnable frm.CTBCtrl1, frm.Tlb, "SaveRs", True
'                    End If
'                End If
'            End If
'
'            If Not frm.ActiveControl Is frm.TREE1 Then Exit Sub
'            If frm.Tlb.Buttons("Delete").Enabled Then
'                frm_ButtonClick "Erase"
'            Else
'                'MsgBox "当前状态不能删除或该节点不能删除！", vbInformation, Msg_Title
'            End If
                
        Case 80     'Ctrl + P
            If Shift = 2 Then
                If frm.Tlb.Buttons("Print").Enabled Then frm_ButtonClick "Print"
            End If
        Case 13
            If Shift = 0 Then SendKeys "{Tab}"
    End Select
End Sub

'-----------------------------------------------------------
'功能：窗体的QueryUnload事件
'
'参数：
'
'返回：
'
'-----------------------------------------------------------
Private Sub frm_FrmQuryExit(Cancel As Integer)
    On Error Resume Next
    If frm.Tlb.Buttons("SaveRs").Enabled = True And mbolChange = True Then
        If ShowMsg("是否保存当前记录的编辑？", 2) = 6 Then
            If Save("Exit") <> "" Then
                Cancel = 1
                Exit Sub
            End If
        End If
    End If
    Call UA_FreeTask(mstrAdd_auth)
    frm.Hide
End Sub

'-----------------------------------------------------------
'功能：窗体的load事件
'
'参数：
'
'返回：
'
'-----------------------------------------------------------
Private Sub frm_Load()
    setStb frm                                          '设置状态栏
    frm.Caption = mstrCaption                           '设置标题
    frm.Label1.Caption = mstrCaption + "编码"
    frm.Label2.Caption = mstrCaption + "名称"
    LoadTree                                            '加载树节点
    mlngBColor = frm.TREE1.Nodes(1).BackColor
    frm.TREE1.Nodes(mlngIndex).BackColor = mlngBColor + 10
   ' frm.Icon = LoadResPicture(120, vbResIcon)
    SetTlbButton                                        '设置toolbar的按钮
    If TBLStyle = TBLNormal Then                        '设置toolbar的风格
        frm.mnuTBL.Caption = "文本按钮(&T)"
        TBLStyle = TBLNormal
        frm.Tlb.Visible = True
        frm.CTBCtrl1.Visible = False
        frm.Tlb.ZOrder 0
    Else
        frm.mnuTBL.Caption = "图标按钮(&T)"
        TBLStyle = TBLText
        frm.CTBCtrl1.SetToolbar frm.Tlb
        frm.CTBCtrl1.SetDisplayStyle TextOnly
        frm.CTBCtrl1.RefreshVisible
        frm.Tlb.Visible = False
        frm.CTBCtrl1.Visible = True
        frm.CTBCtrl1.ZOrder 0
    End If
End Sub


'-----------------------------------------------------------
'功能：设置Toolbar按钮状态
'
'参数：sKey 按钮的key，=''即为初始化；bEnd是否是末级节点；bSystem 是否是系统默认
'
'返回：
'
'-----------------------------------------------------------
Private Sub SetTlbButton(Optional sKey As String = "", Optional bEnd As Boolean, Optional bSystem As Boolean)
    Select Case sKey
    Case "Add", "Modify"
        frm.TREE1.Enabled = False
        setCTBEnable frm.CTBCtrl1, frm.Tlb, "Add", False
        setCTBEnable frm.CTBCtrl1, frm.Tlb, "Modify", False
        setCTBEnable frm.CTBCtrl1, frm.Tlb, "Delete", False
        setCTBEnable frm.CTBCtrl1, frm.Tlb, "Back", True
        setCTBEnable frm.CTBCtrl1, frm.Tlb, "SaveRs", False
        'setCTBEnable frm.CTBCtrl1, frm.Tlb, "Refresh", False
                
        frm.txtName.Enabled = True
        frm.Check2.Enabled = True
        frm.Check1.value = 0
        If sKey = "Add" Then
            frm.txtCode.Enabled = True
            frm.Check2.value = 1
            frm.txtCode.SetFocus
            frm.txtCode.text = ""
'            frm.txtCode.SelStart = Len(frm.txtCode.text)
            frm.txtName.text = ""
        Else
            frm.txtName.SetFocus
        End If
    Case Else
        frm.TREE1.Enabled = True
        '\\\\\\\zhupeibi未设置级次关系时增加按钮可用
'        setCTBEnable frm.CTBCtrl1, frm.Tlb, "Add", True
        setCTBEnable frm.CTBCtrl1, frm.Tlb, "Add", IIf(CanAdd And Not bEnd, True, False)
        
        If mlngIndex > 1 Then
            setCTBEnable frm.CTBCtrl1, frm.Tlb, "Modify", True Xor bSystem
            setCTBEnable frm.CTBCtrl1, frm.Tlb, "Delete", True Xor bSystem
        Else
            
            setCTBEnable frm.CTBCtrl1, frm.Tlb, "Modify", False
            setCTBEnable frm.CTBCtrl1, frm.Tlb, "Delete", False
        End If
        setCTBEnable frm.CTBCtrl1, frm.Tlb, "Back", False
        setCTBEnable frm.CTBCtrl1, frm.Tlb, "SaveRs", False
        setCTBEnable frm.CTBCtrl1, frm.Tlb, "Refresh", True
        
        frm.txtCode.Enabled = False
        frm.txtName.Enabled = False
        frm.Check2.Enabled = False
    End Select
    '设置菜单
    frm.mnuAdd.Enabled = frm.Tlb.Buttons("Add").Enabled
    frm.mnuModify.Enabled = frm.Tlb.Buttons("Modify").Enabled
    frm.mnuDelete.Enabled = frm.Tlb.Buttons("Delete").Enabled
    frm.mnuBack.Enabled = frm.Tlb.Buttons("Back").Enabled
    frm.mnuSave.Enabled = frm.Tlb.Buttons("SaveRs").Enabled
    frm.mnuRefresh.Enabled = frm.Tlb.Buttons("Refresh").Enabled
End Sub

'-----------------------------------------------------------
'功能：当前控件事去焦点事件
'
'参数：obj为当前焦点的控件
'
'返回：
'
'-----------------------------------------------------------
Private Sub frm_LostFocus(obj As Control)
    Lose obj
End Sub


'-----------------------------------------------------------
'功能：当前控件事去焦点时处理
'
'参数：obj为当前焦点的控件
'
'返回：
'
'-----------------------------------------------------------
Private Function Lose(obj As Control) As Boolean
    On Error Resume Next
    Lose = True
    If Not IsQualify(obj.text) Then
        Lose = False
        obj.SetFocus
    End If
End Function

'-----------------------------------------------------------
'功能：保存打印设置
'
'参数：
'
'返回：
'
'-----------------------------------------------------------
Private Sub frm_SettingChanged(ByVal varLocalSettings As Variant, ByVal varModuleSettings As Variant)
    Call g_oPub.SavePrnStyle(SrvDB, mstrTable, CStr(varLocalSettings) + CStr(varModuleSettings))
End Sub


'-----------------------------------------------------------
'功能：点击树节点
'
'参数：node当前点击节点
'
'返回：
'
'-----------------------------------------------------------
Private Sub frm_TreeNodeClick(ByVal Node As MSComctlLib.Node)
    Dim rec    As ADODB.Recordset
    If Not mbolCollapse Then Node.Expanded = True
    If mlngIndex = Node.Index Then Exit Sub         '选择同一个节点
    If mlngIndex <> Node.Index And mlngIndex <> 0 Then
        If mlngIndex <= frm.TREE1.Nodes.Count Then frm.TREE1.Nodes(mlngIndex).BackColor = mlngBColor
    End If
    mlngIndex = Node.Index
    Node.BackColor = mlngBColor + 10
    Node.Selected = True
    
    mstrCode = ""
    frm.txtCode = ""
    frm.txtName = ""
    frm.Check1.value = 0
    frm.Check2.value = 0
    If Node.Index = 1 Then
        SetTlbButton
        mbolEnd = False
        Exit Sub
    End If
    '''
    Set rec = New ADODB.Recordset
    rec.CursorLocation = adUseClient
    mstrCode = Mid(Node.text, 2, InStr(Node.text, ")") - 2)
    rec.Open "select CCODE,CNAME,BEND,convert(char,Convert(money,Ufts),2) AS UFTS FROM " & mstrTable & " WHERE CCODE='" & mstrCode & "'", AdoCnn, adOpenKeyset, adLockOptimistic
    If rec.RecordCount = 1 Then
        If IsNull(rec.Fields("UFTS")) Then
            mstrUfts = ""
        Else
            mstrUfts = rec.Fields("UFTS")
        End If
    Else
        ShowMsg "该分类已被他人更新或删除，请刷新后再试！", 1
        Exit Sub
    End If
    frm.txtCode = Trim(Mid(Node.text, 2, InStr(Node.text, ")") - 2))
    frm.txtName = Trim(rec.Fields("CNAME"))
    mbolEnd = GetEnd(Trim(frm.txtCode))
    frm.Check2.value = IIf(mbolEnd = True, 1, 0)
    '如果是设施类别定义
    If frm.Check1.Visible Then
        Dim strTemp As String
        frm.Check1.value = IIf(InStr(Node.text, "[系统]") > 0, 1, 0)         '是系统默认节点
        If frm.Check1.value = 1 Then
            frm.TREE1.Nodes(frm.TREE1.SelectedItem.Index).text = "(" & Trim(frm.txtCode) & ") " & Trim(frm.txtName) & "[系统]"
'            frm.txtName = Mid(Node.text, InStr(Node.text, ")") + 2, InStr(Node.text, "[") - InStr(Node.text, ")") - 2)
        Else
            frm.TREE1.Nodes(frm.TREE1.SelectedItem.Index).text = "(" & Trim(frm.txtCode) & ") " & Trim(frm.txtName)
'            frm.txtName = Right(Node.text, Len(Node.text) - InStr(Node.text, ")") - 1)
        End If
    Else
        frm.TREE1.Nodes(frm.TREE1.SelectedItem.Index).text = "(" & Trim(frm.txtCode) & ") " & Trim(frm.txtName)
'        frm.txtName = Right(Node.text, Len(Node.text) - InStr(Node.text, ")") - 1)
    End If
    
    rec.Close
    Set rec = Nothing
    SetTlbButton "", mbolEnd, IIf(frm.Check1.value = 1, True, False)
End Sub


'-----------------------------------------------------------
'功能：直接判断当前编辑的strCode是否存在左边树节点
'
'参数：
'
'返回：ExistNode<>''说明存在
'
'-----------------------------------------------------------
Public Function ExistNode() As String
    Dim i           As Long
    Dim sText       As String
    
    ExistNode = ""
    For i = 1 To frm.TREE1.Nodes.Count
        sText = frm.TREE1.Nodes(i).text
        If InStr(sText, "(" & Trim(frm.txtCode) & ")") = 1 Then
            ExistNode = "该分类已存在左边列表，请刷新后再试！"
            Exit Function
        End If
    Next
End Function


'-----------------------------------------------------------
'功能：保存档案
'
'参数：skey//现无用
'
'返回：Save<>''保存失败
'
'-----------------------------------------------------------
Private Function Save(ByVal sKey As String) As String
    Dim sParent     As String
    Dim sErrMsg     As String
    Dim lngTemp     As Long
    Dim lngReturn   As Long
    Dim rec         As New ADODB.Recordset
    
    rec.CursorLocation = adUseClient
    '验证输入控件值的合法性
    On Error GoTo ErrorLine
    frm_TxtKeyUp frm.txtCode
    frm_TxtKeyUp frm.txtName
    If Not Lose(frm.txtCode) Then GoTo ErrorLine
    If Not Lose(frm.txtName) Then GoTo ErrorLine
    '必填项控制
    If mstrTable = "GSP_DEVICETYPE" And Right(Trim(frm.txtName.text), 4) = "[系统]" Then
        MsgBox mstrCaption + "名称末尾不能为""" & "[系统]" & """！", vbOKOnly, "出版管理"
        frm.txtName.SetFocus
        Save = "必填项"
        Exit Function
    End If
    If Trim(frm.txtCode.text) = "" Then
        MsgBox mstrCaption + "编码为必填项！", vbOKOnly, "出版管理"
        frm.txtCode.SetFocus
        Save = "必填项"
        Exit Function
    End If
    If Trim(frm.txtName.text) = "" Then
        MsgBox mstrCaption + "名称为必填项！", vbOKOnly, "出版管理"
        frm.txtName.SetFocus
        Save = "必填项"
        Exit Function
    End If
    If Len(frm.txtCode.text) > 10 Then
        MsgBox mstrCaption + "编码长度过长！", vbOKOnly, "出版管理"
        frm.txtCode.SetFocus
        Save = "必填项"
        Exit Function
    End If
    If Len(frm.txtName.text) > 40 Then
        MsgBox mstrCaption + "名称长度过长！", vbOKOnly, "出版管理"
        frm.txtName.SetFocus
        Save = "必填项"
        Exit Function
    End If
    '取父节点的key\text
   'If mlngIndex = 1 Then
    If mlngIndex = mlngIndex Then
        sParent = ""
    Else
        If Opt = 1 Then '增加
            sParent = Trim(frm.TREE1.Nodes(mlngIndex).text)
'        Else
'            sParent = Trim(frm.TREE1.Nodes(frm.TREE1.Nodes(mlngIndex).Parent.Index).Text)
        End If
    End If
    If sParent <> "" Then sParent = Mid(sParent, InStr(sParent, "(") + 1, InStr(sParent, ")") - InStr(sParent, "(") - 1)
    
    'adocnn.Execute "UPDATE " & mstrTable & " SET CPARENTNODE='' WHERE CPARENTNODE IS NULL"
    If Opt = 1 Then
        AdoCnn.BeginTrans
        AdoCnn.Execute "UPDATE GSP_DEVICETYPE SET BSYSTEM=BSYSTEM WHERE BSYSTEM=1"   '锁定数据库
        'rec.CursorLocation = adUseServer
        rec.Open "SELECT COUNT(CCODE) AS AA FROM " & mstrTable & " WHERE CCODE in('" & Trim(frm.txtCode) & "')", AdoCnn, adOpenStatic, adLockPessimistic
        If Not rec.EOF Then
            If rec.Fields("AA") > 0 Then
                rec.Close
                sErrMsg = "添加失败，" & mstrCaption & "编码不能相同，请刷新后再做添加！"
                
                GoTo ErrorLine
            End If
        End If
        rec.Close
        If sParent <> "" Then
            'rec.CursorLocation = adUseServer
            rec.Open "SELECT COUNT(CCODE) AS AA FROM " & mstrTable & " WHERE CCODE in('" & sParent & "') and isnull(bend,0)=0", AdoCnn, adOpenStatic, adLockPessimistic
            If Not rec.EOF Then
                If rec.Fields("AA") <= 0 Then
                    rec.Close
                    sErrMsg = "添加失败，父结点被人删除或改成末级结点，请刷新后再做添加！"
                    
                    GoTo ErrorLine
                End If
            Else
                rec.Close
                sErrMsg = "添加失败，父结点被人删除或改成末级结点，请刷新后再做添加！"
                
                GoTo ErrorLine
            End If
            rec.Close
        End If
        sErrMsg = ExistNode
        If sErrMsg <> "" Then GoTo ErrorLine
        AdoCnn.Execute "INSERT INTO " & mstrTable & "(CCODE,CNAME,CPARENTNODE,BEND) VALUES('" & Trim(frm.txtCode) & "','" & Trim(frm.txtName) & "','" & sParent & "'," & frm.Check2.value & ")"
        AdoCnn.CursorLocation = adUseClient
        frm.TREE1.Nodes.Add "K" & sParent, tvwChild, "K" & Trim(frm.txtCode), "(" & Trim(frm.txtCode) & ") " & Trim(frm.txtName)
        '设置选者节点
        frm.TREE1.Nodes(mlngIndex).BackColor = mlngBColor
        lngTemp = frm.TREE1.Nodes(frm.TREE1.Nodes.Count).Index
        AdoCnn.CommitTrans
        mlngIndex = 1
        frm_TreeNodeClick frm.TREE1.Nodes(lngTemp)
        mlngRsCount = mlngRsCount + 1
        frm.Stb.Panels(3).text = "当前记录数：" + Trim(mlngRsCount) + "条"
    Else
        '修改
        If Not EndCheck Then GoTo ErrorLine
        AdoCnn.BeginTrans
        AdoCnn.Execute "UPDATE " & mstrTable & " SET CNAME='" & Trim(frm.txtName) & "',BEND=" & frm.Check2.value & " WHERE CCODE='" & Trim(frm.txtCode) & "' and (convert(char,Convert(money,Ufts),2) ='" & mstrUfts & "'  " & IIf(mstrUfts = "", " OR UFTS=NULL", "") & ")", lngReturn
        If lngReturn <= 0 Then
            sErrMsg = "保存失败，该分类被人更新或删除，请刷新后再做修改！"
            
            GoTo ErrorLine
        End If
        AdoCnn.CommitTrans
        frm.TREE1.Nodes(mlngIndex).text = "(" & Trim(frm.txtCode) & ") " & Trim(frm.txtName) & IIf(frm.Check1.value = 1, "[系统]", "")
        frm.TREE1.Nodes(mlngIndex).Selected = True
        
        lngTemp = mlngIndex
        mlngIndex = 1
        frm_TreeNodeClick frm.TREE1.Nodes(lngTemp)
    End If
    Set rec = Nothing
    Exit Function
ErrorLine:
    On Error Resume Next
    If rec.State Then Set rec = Nothing
    AdoCnn.CursorLocation = adUseClient
    AdoCnn.RollbackTrans
    LoadTree
    frm.txtCode.SetFocus
    'If sErrMsg = "" Then sErrMsg = "保存失败，请重试！"
    If sErrMsg <> "" Then ShowMsg sErrMsg
    Save = "保存失败！"
'    Err.Raise vbObjectError + 1002, "出版管理", "保存失败！"
End Function


'-----------------------------------------------------------
'功能：是否可以删除
'
'参数：
'
'返回：
'
'-----------------------------------------------------------
'Private Function CanDelete() As Boolean
'  Dim strCode         As String
'  Dim rec             As ADODB.Recordset
'
'    strCode = frm.txtCode
'    Set rec = New ADODB.Recordset
'    rec.CursorLocation = adUseClient
'    CanDelete = True
'    If mstrTable = "GSP_DEVICETYPE" Then
'        rec.Open "SELECT COUNT(*) FROM GSP_DEVICEINFO WHERE CDEVICETYPEID='" & strCode & "'", AdoCnn, adOpenStatic, adLockPessimistic
'    ElseIf UCase(mstrTable) = UCase("GSP_certifcateCategory") Then
'        rec.Open "SELECT COUNT(*) FROM EF_GSP_CusLicences WHERE ccode='" & strCode & "'", AdoCnn, adOpenStatic, adLockPessimistic
'        If rec.RecordCount > 0 Then
'          If IsNull(rec.Fields(0).Value) Then
'          Else
'            If Val(rec.Fields(0).Value) > 0 Then
'                CanDelete = False
'                rec.Close
'                Set rec = Nothing
'                Exit Function
'            End If
'          End If
'        End If
'        Set rec = New ADODB.Recordset
'        rec.CursorLocation = adUseClient
'        CanDelete = True
'        rec.Open "SELECT COUNT(*) FROM EF_GSP_VenLicences WHERE ccode='" & strCode & "'", AdoCnn, adOpenStatic, adLockPessimistic
'    ElseIf UCase(mstrTable) = UCase("GSP_MedicineCategory") Then
'        rec.Open "SELECT COUNT(*) FROM inventory WHERE cinvdefine1='" & strCode & "'", AdoCnn, adOpenStatic, adLockPessimistic
'    Else
'        rec.Open "SELECT COUNT(*) FROM " & mstrTable & " WHERE ccode='" & strCode & "'", AdoCnn, adOpenStatic, adLockPessimistic
'    End If
'    If rec.RecordCount > 0 Then
'        If IsNull(rec.Fields(0).Value) Then
'        Else
'            If Val(rec.Fields(0).Value) > 0 Then
'                CanDelete = True
'            End If
'        End If
'    End If
'    rec.Close
'    Set rec = Nothing
'End Function

'-----------------------------------------------------------
'功能：删除节点
'
'参数：
'
'返回：
'
'-----------------------------------------------------------
Private Sub Delete()
    Dim lngTemp         As Long
    Dim lngReturn       As Long                 '影响行数
    Dim strParent       As String
    On Error GoTo ErrorLine
    ''
    If frm.TREE1.Nodes(mlngIndex).children > 0 Then
        MsgBox "有下级分类，不可删除！", vbOKOnly + vbInformation, "出版管理"
        Exit Sub
    End If
'    If Not CanDelete Then
'        MsgBox "此""" & mstrCaption & """" & "在""" & mstrArCaption & """中使用，不可删除！", vbOKOnly + vbInformation, "出版管理"
'        Exit Sub
'    End If
    If MsgBox("确信要删除该分类？", vbYesNo, "出版管理") = vbNo Then Exit Sub
    
    
    strParent = frm.TREE1.Nodes(mlngIndex).Key
'    DeleteNode Mid(frm.TREE1.Nodes(mlngIndex).Text, InStr(frm.TREE1.Nodes(mlngIndex).Text, "(") + 1, InStr(frm.TREE1.Nodes(mlngIndex).Text, ")") - InStr(frm.TREE1.Nodes(mlngIndex).Text, "(") - 1), strParent
'zhupeibi
    If CheckCCodeState(Trim(frm.txtCode)) = False Then
         ShowMsg "该" & mstrArCaption & "已经在业务系统中应用，不可以删除！", 0
       Exit Sub
    End If
    
    AdoCnn.Execute "DELETE FROM " & mstrTable & " WHERE  CCODE='" & Trim(frm.txtCode) & "' AND ( convert(char,Convert(money,Ufts),2) ='" & mstrUfts & "'" & IIf(mstrUfts <> "", "", " OR UFTS=null ") & ")", lngReturn
    If lngReturn <= 0 Then
        ShowMsg "该分类已被他人更新或删除，请刷新后再试！", 1
        Exit Sub
    End If
    
    lngTemp = frm.TREE1.Nodes(mlngIndex).Parent.Index
    frm.TREE1.Nodes.Remove mlngIndex
    'mlngIndex = lngTemp
    If frm.TREE1.Nodes(lngTemp) Is Nothing Then lngTemp = 1
    frm_TreeNodeClick frm.TREE1.Nodes(lngTemp)
    mlngRsCount = mlngRsCount - 1
    frm.Stb.Panels(3).text = "当前记录数：" + Trim(mlngRsCount) + "条"
    
    Exit Sub
ErrorLine:
    AdoCnn.RollbackTrans
    Err.Raise vbObjectError + 1003, "出版管理", "删除失败！"
End Sub
'判断档案是否已在业务中使用
Public Function CheckCCodeState(ByVal ccode As String) As Boolean
        Dim strSQL As String
        Dim rst_c As New ADODB.Recordset
        CheckCCodeState = True
        Dim coltables As Collection
        Set coltables = New Collection
       Select Case mstrTable
        Case "EFBWGL_DBFormat"      '开本档案
        '选题论证
        strSQL = "select cformatcode from EFBWGL_SelDeclare where cformatcode='" & ccode & "'"
         rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
        '选题登录
        strSQL = "select cformatcode from EFBWGL_SelRegister where cformatcode ='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4    ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
        '发稿记录
        strSQL = "select cformatcode from EFBWGL_DistRecord where cformatcode='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
        
        'cip数据
        strSQL = "select cformatcode from EFBWGL_CipData where cformatcode='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
        '制版单表头
        strSQL = "select cformatcode from EFYZGL_Pmaking where cformatcode='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
        
        '付印通知单表头
        strSQL = "select cformatcode from EFYZGL_pressinform where cformatcode='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
        '印刷委托书
        strSQL = "select cformatcode from EFYZGL_pressconsign where cformatcode='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
        
        '费用预估单表头
       strSQL = "select cformatcode from EFFYGL_Pcostbudget  where cformatcode='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
        
        '费用结算单表
        strSQL = "select cformatcode from EFFYGL_SettleVouch  where cformatcode='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
        
    Case "EFBWGL_DBBookBinding"      '装帧档案
    
        coltables.Add "EFBWGL_SelDeclare"
        coltables.Add "EFBWGL_SelRegister"
        coltables.Add "EFBWGL_DistRecord"   '待验证
        coltables.Add "EFBWGL_CipData"
        coltables.Add "EFYZGL_pressinform"
        If CheckIsReferByTable(coltables, "cbookbindingcode", ccode) Then CheckCCodeState = False: Exit Function
      
        
    Case "EFBWGL_DBBookLanguage"     '文种档案
        coltables.Add "EFBWGL_SelRegister"
        coltables.Add "EFBWGL_DistRecord"
        coltables.Add "EFBWGL_CipData"
        If CheckIsReferByTable(coltables, "clanguagecode", ccode) Then CheckCCodeState = False: Exit Function
    Case "EFBWGL_DBBookClass"        '图书分类档案
    
        coltables.Add "EFBWGL_SelRegister"
        coltables.Add "EFBWGL_DistRecord"
        'coltables.Add "EFBWGL_CipData"
        If CheckIsReferByTable(coltables, "selclasscode", ccode) Then CheckCCodeState = False: Exit Function
        
    Case "EFBWGL_DBCBHT"             '合同条款档案
        strSQL = "select htid from EFBWGL_BookHts where htid='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
        
        strSQL = "select htid from EFYZGL_pressinforms where htid='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
    Case "EFBWGL_DBBMMode"           '著作方式档案
        strSQL = "select cbmcode from EFBWGL_DistRecord where cbmcode='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
        
        strSQL = "select firstbmcode from EFBWGL_CipData where firstbmcode='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
        
        strSQL = "select otherbmcode from EFBWGL_CipData where otherbmcode='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
    Case "EFBWGL_DBTypeWord"         '类型词档案
        strSQL = "select cstylecode from EFBWGL_SelRegister where cstylecode ='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4    ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
        
        strSQL = "select cstylecode from EFBWGL_DistRecord where cstylecode='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
    Case "EFBWGL_DBBookSelStress"    '选题类别档案
        strSQL = "select cselstresscode from EFBWGL_DistRecord where cselstresscode='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close

        strSQL = "select cselstresscode from EFBWGL_SelRegister where cselstresscode='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
        
    Case "EFBWGL_DBTenor"           '进程档案
        strSQL = "select ctenorcode from EFBWGL_CheckNote where ctenorcode='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
        
        strSQL = "select crdcode from EFYZGL_RdRecord where crdcode='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
    Case "EFBWGL_DBBookDeclClass"   '申报类型档案
        strSQL = "select cdeclcode from EFBWGL_SelDeclare where cdeclcode='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
        
        strSQL = "select cdeclcode from EFBWGL_DistRecord where cdeclcode='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
    Case "EFBWGL_DBHTClass"         '合同类型档案
        strSQL = "select htclasscode from EFBWGL_DistRecord where htclasscode='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
        
        strSQL = "select htclasscode from EFBWGL_BookHt where htclasscode='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
    Case "EFBWGL_DBbindmanner"      '装订方式档案
        strSQL = "select cbookbindingcode1 from EFYZGL_pressinform where cbookbindingcode1='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
        
    Case "EFBWGL_DBbindway"        '装订方法档案
        strSQL = "select cbookbindingcode from EFYZGL_Pmaking where cbookbindingcode='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
        
        strSQL = "select bindmanner from EFYZGL_pressinform where bindmanner='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
        
        strSQL = "select cbindway from EFYZGL_pressconsign where cbindway='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4   ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
    Case "EFBWGL_DBBookPubMode"    '出版方式档案
        strSQL = "select cpubmodecode from EFBWGL_SelRegister where cpubmodecode ='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4    ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
        
        strSQL = "select cpubmodecode from EFBWGL_DistRecord where cpubmodecode ='" & ccode & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4    ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckCCodeState = False
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
   End Select
End Function

'-----------------------------------------------------------
'功能：删除树节点节点，及其子节点
'
'参数：
'
'返回：
'
'-----------------------------------------------------------
Private Sub DeleteNode(ByVal sCode As String, ByVal strParent As String)
    Dim rec As New ADODB.Recordset
    
    rec.CursorLocation = adUseClient
    rec.Open "SELECT * FROM " & mstrTable & " WHERE CPARENTNODE='" & strParent & "'", AdoCnn, adOpenStatic, adLockPessimistic
    If rec.RecordCount > 0 Then
        Do While Not rec.EOF
            DeleteNode Trim(rec!ccode), Trim(rec!cparentnode) & Trim(rec!ccode)
            If mstrTable <> "GSP_DEVICETYPE" Then AdoCnn.Execute "DELETE FROM " & mstrArTable & " WHERE CPARENTNODE='" & strParent & "'"
            AdoCnn.Execute "DELETE FROM " & mstrTable & " WHERE CCODE='" & sCode & "'"
            rec.MoveNext
        Loop
    Else
        If mstrTable <> "GSP_DEVICETYPE" Then AdoCnn.Execute "DELETE FROM " & mstrArTable & " WHERE CPARENTNODE='" & strParent & "'"
        AdoCnn.Execute "DELETE FROM " & mstrTable & " WHERE CCODE='" & sCode & "'"
    End If
    rec.Close
    Set rec = Nothing
End Sub

'-----------------------------------------------------------
'功能：放弃增加、修改
'
'参数：
'
'返回：
'
'-----------------------------------------------------------
Private Sub Back()
    Dim Node   As Node
    
    Set Node = frm.TREE1.Nodes(mlngIndex)
    mlngIndex = 0
    frm_TreeNodeClick Node
    
End Sub

'-----------------------------------------------------------
'功能：刷新树节点
'
'参数：
'
'返回：
'
'-----------------------------------------------------------
Private Sub Refresh()
    LoadTree
End Sub


'-----------------------------------------------------------
'功能：Tree节点展开事件
'
'参数：
'
'返回：
'
'-----------------------------------------------------------
Private Sub frm_TreeNodeCollapse(ByVal Node As MSComctlLib.Node)
    mbolCollapse = True
    Call frm_TreeNodeClick(Node)
    mbolCollapse = False
End Sub


'-----------------------------------------------------------
'功能：textbox控件Change事件
'
'参数：txt当前焦点textbox,lLen允许输入长度
'
'返回：
'
'-----------------------------------------------------------
Private Sub frm_TxtChange(txt As TextBox, lLen As Long)
'    Dim lSelStart   As Long
'    lSelStart = txt.SelStart
    mbolChange = True
    setCTBEnable frm.CTBCtrl1, frm.Tlb, "SaveRs", True
    frm.mnuSave.Enabled = True
    txt.text = Replace(txt.text, "[档案]", "")
    txt.text = Replace(txt.text, "[系统]", "")
    If EnterLen(txt.text) > lLen Then txt.text = mstrTxtBeforeChange
'    txt.SelStart = lSelStart
    'txt.text = GetString(txt.text, lLen)
    mstrTxtBeforeChange = txt.text
End Sub

'-----------------------------------------------------------
'功能：textbox控件GetFocus事件
'
'参数：txt当前焦点textbox
'
'返回：
'
'-----------------------------------------------------------
Private Sub frm_TxtGetFocus(txt As Control)
    mstrTxtBeforeChange = txt.text
End Sub

'-----------------------------------------------------------
'功能：textbox控件KeyPress事件
'
'参数：txt当前焦点textbox，asc输入的ascii，lLen允许输入的长度，objName空间名称
'
'返回：
'
'-----------------------------------------------------------
Private Sub frm_TxtKeyPress(txt As TextBox, ByRef asc As Integer, lLen As Long, objName As String)
    '判断是否超过长度
    If EnterLen(GetText(asc, txt)) > lLen Then  'And asc <> 8 And txt.SelLength <= 0 Then
        asc = 0
        Exit Sub
    End If
    '不允许直接输入空格(但是在SelLength>0时可以)，日期格式在GetDateAsc里控制
    If asc = 32 Then
        If txt.SelLength <= 0 Then asc = 0
        Exit Sub
    End If
    If asc = 24 Or asc = 22 Or asc = 3 Or asc = 26 Then 'ctrl+v,c,x,z
        Exit Sub
    End If
    GetAsc asc, IIf(objName = "txtCode", True, False)
End Sub

Private Function GetValue(ByRef sValue) As String
    GetValue = Replace(sValue, "'", "'''")
End Function


'-----------------------------------------------------------
'功能：设置CTBCtrl1按钮是否可用
'
'参数：CTB CTBCtrl1控件， Tlb CTB对应的toolbar， sKey 当前要设置的buttons.key， bEnabled enable值
'
'返回：
'
'-----------------------------------------------------------
Public Sub setCTBEnable(CTB As Object, Tlb As Object, sKey As String, bEnabled As Boolean)
    Dim i           As Long
    Dim lIndex      As Long

    If TBLStyle = TBLNormal Then
        Tlb.Buttons(sKey).Enabled = bEnabled
    Else
        lIndex = -1
        For i = 1 To Tlb.Buttons.Count
            If UCase(sKey) = UCase(Tlb.Buttons(i).Key) Then
                lIndex = i
                Exit For
            End If
        Next
        If lIndex <> -1 Then
            Tlb.Buttons(lIndex).Enabled = bEnabled
            CTB.SetButtonEnable lIndex, bEnabled
        End If
    End If
End Sub


'-----------------------------------------------------------
'功能：判断该节点是否还可以增加；如果已经有档案了就不允许增加
'
'参数：
'
'返回：是否还可以增加
'
'-----------------------------------------------------------
Private Function CanAdd() As Boolean
    Dim strCode     As String
    
    CanAdd = True
    strCode = frm.TREE1.Nodes(mlngIndex).text
    
    If InStr(strCode, "(") Then
        strCode = Mid(strCode, 2, InStr(strCode, ")") - 2)
    Else
        strCode = ""
    End If
    
    If strCode = "" Then Exit Function
    If frm.Check2.value = 1 Then CanAdd = False
End Function

'-----------------------------------------------------------
'功能：判断是否是末节点
'
'参数：要判断节点的text
'
'返回：
'
'-----------------------------------------------------------
Function GetEnd(strCode As String) As Boolean

    Dim rec         As ADODB.Recordset
    
    Set rec = New ADODB.Recordset
    rec.CursorLocation = adUseClient
    rec.Open "SELECT ISNULL(BEND,0) FROM " & mstrTable & " WHERE CCODE='" & strCode & "'", AdoCnn, adOpenStatic, adLockReadOnly
    If rec.RecordCount <= 0 Then
        ShowMsg "该节点可能已被别人删除，请刷新后再做操作！"
        LoadTree
        GetEnd = False
    Else
        GetEnd = rec.Fields(0).value
    End If
    rec.Close
    Set rec = Nothing
End Function


'-----------------------------------------------------------
'功能：textbox的keyup 事件
'
'参数：
'
'返回：
'
'-----------------------------------------------------------
Private Sub frm_TxtKeyUp(txt As TextBox)
    GetPressUp txt, IIf(txt.Name = "txtCode", True, False)
    mstrTxtBeforeChange = txt.text
End Sub
Public Function ShowContextHelp(hwnd As Long, sHelpFile As String, lContextID As Long) As Long
    ShowContextHelp = htmlHelp(hwnd, sHelpFile, &HF, lContextID)
End Function

'检查某个字段的值被是否引用
Public Function CheckIsReferByTable(tablenames As Collection, fieldname As String, value As String) As Boolean
    Dim i As Integer
    Dim rst_c As ADODB.Recordset
    Dim strSQL As String
    Set rst_c = New ADODB.Recordset
    For i = 1 To tablenames.Count
        strSQL = "select " & fieldname & " from " & tablenames.item(i) & " where " & fieldname & " ='" & value & "'"
        rst_c.Open strSQL, AdoCnn, 3, 4    ' adOpenStatic, adLockReadOnly
        If Not rst_c.EOF Then
           CheckIsReferByTable = True
           rst_c.Close
           Set rst_c = Nothing
           Exit Function
        End If
        rst_c.Close
    Next i
    Set rst_c = Nothing
End Function
