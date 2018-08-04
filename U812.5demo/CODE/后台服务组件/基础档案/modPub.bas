Attribute VB_Name = "modPub"
Option Base 0
Option Explicit

Public SrvDB        As Object 'U8SrvTrans.IClsCommon        'U8SrvTrans.IClsCommon '设置公共数据库事务处理对象
Public g_oLogin     As U8Login.clsLogin             'UFLoginSQL.Login  '设置Login对象
Public g_oPub       As U8Pub.IPub                   'U8Pub.IPub '公共组件模块
Public g_DbGSP      As UfDatabase
Public clsAuth      As U8RowAuthsvr.clsRowAuth
Public ctlDate      As CalendarAPP.ICaleCom

Public AdoCnn       As ADODB.Connection

Public TBLStyle     As TBLType
Public AppPath      As String                       '帮助文件路径
Public mhwndMain    As Long                         'MDI主窗体的句柄

Public mlngType     As Long

Public mbolChangeOther      As Boolean              '是否可以修改别人的纪录
Public mbolAuditOwner       As Boolean              '是否可以审核自己的纪录
Public mstrOperator         As String               '操作员

Public mstrHelpID           As String               '帮助号

Public mstrRef              As String
Public mstrCaption          As String

Public mbYearEnd            As Boolean              '是否年结
Public mbCanModifyOther     As Boolean
Public mbCanAuditOwn        As Boolean

Public Const msg_YearEnd = "出版系统已封账，不能进行业务处理操作！"
Public Const Msg_Title = "出版管理"
Public m_login As U8Login.clsLogin
'toolbar,ctblctrl
Enum TBLType
    TBLText
    TBLPicture
    TBLNormal
End Enum

'API声明
'发送消息
Public Const VK_F1 = &H70
Public Const WM_KEYDOWN = &H100
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
  ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As String) As Long

 ''帮助文件
Public Declare Function htmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public Const HH_DISPLAY_topic = &H0
Public Const HH_HELP_CONTEXT = &HF

'-----------------------------------------------------------
'功能：公共消息接口
'
'参数：sMsg消息内容，lType msgbox类型
'
'返回：
'
'-----------------------------------------------------------
Public Function ShowMsg(ByVal sMsg As String, Optional ByVal lType As Long = 0) As Long
    
    Dim lReturn     As Long
    
    Select Case lType
    Case 0: lReturn = MsgBox(sMsg, vbInformation + vbOKOnly, "出版管理")
    Case 1: lReturn = MsgBox(sMsg, vbCritical + vbOKOnly, "出版管理")
    Case 2: lReturn = MsgBox(sMsg, vbQuestion + vbYesNo, "出版管理")
    End Select
    ShowMsg = lReturn
End Function


'-----------------------------------------------------------
'功能：申请任务
'
'参数：TaskID  任务号
'
'返回：
'
'-----------------------------------------------------------
Public Function UA_Task(ByVal TaskID As String) As Boolean
    On Error GoTo errHandle
    Dim sStr As String
    
    
    If Not g_oLogin Is Nothing Then
        g_oLogin.ClearError
        If g_oLogin.TaskExec(Trim(TaskID), -1, g_oLogin.cIYear) Then
            UA_Task = True
            Exit Function
        Else
            If g_oLogin.ShareString <> "" Then
                MsgBox g_oLogin.ShareString, 64, Msg_Title
            Else
                MsgBox "共享(网络)冲突或没有此项操作的权限，请稍后再试。", 64, Msg_Title
            End If
            g_oLogin.ClearError
            UA_Task = False
            Exit Function
        End If
    Else
        MsgBox "系统管理或注册服务程序工作异常,不能进行功能申请,请检查网络环境。", vbCritical, Msg_Title
        UA_Task = False
        Exit Function
    End If
'    UA_Task = True
    Exit Function
 
errHandle:
    MsgBox Err.Description, vbExclamation, Msg_Title
  
End Function


'-----------------------------------------------------------
'功能：释放任务
'
'参数：TaskID  任务号
'
'返回：
'
'-----------------------------------------------------------
Public Function UA_FreeTask(ByVal TaskID As String) As Boolean
 On Error GoTo errHandle
 
 If Not g_oLogin Is Nothing Then
    g_oLogin.ClearError
     If g_oLogin.TaskExec(TaskID, 0, g_oLogin.cIYear) Then
        UA_FreeTask = True
     Else
        g_oLogin.ClearError
        UA_FreeTask = False
     End If
 Else
     MsgBox "系统管理或注册服务程序工作异常,不能进行功能释放,请检查网络环境。", vbCritical, Msg_Title
     UA_FreeTask = False
     Exit Function
 End If
'    UA_FreeTask = True
     Exit Function

errHandle:
  MsgBox Err.Description, vbExclamation, Msg_Title
End Function




'-----------------------------------------------------------
'功能：打印接口
'
'参数：Key(打印、预览、输出)
'      Prn指某个窗体上的控件
'      frm Prn所在的窗体
'      mstrTable 打印单据表名
'      sCaption打印单据标题
'
'返回：
'
'-----------------------------------------------------------
Public Sub PrintAll(ByVal Key As String, ByRef Prn As Control, frm As Form, ByVal mstrTable As String, ByVal sCaption As String)
    Dim Rs      As ADODB.Recordset
    
    If frm.TREE1.Nodes.Count < 2 Then
        ShowMsg "没有可用数据！"
        Exit Sub
    End If
    
    Prn.Visible = False
    
    Dim sData       As String                   '打印数据xml脚本
    Dim sStyle      As String                   '打印格式xml脚本
    
    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseClient
    
    '设置打印的格式和数据脚本
    If UCase(mstrTable) = "EFBWGL_DBCBHT" Then             '合同条款档案
        Rs.Open "select CCODE,CNAME,HTCONTENT,HTMEMO,(CASE WHEN ISNULL(BEND,0)=0 THEN '否' ELSE '是' END) AS BEND from " & mstrTable & " order by CCODE,CPARENTNODE ASC", AdoCnn, adOpenStatic, adLockReadOnly
        WriteSytle sStyle, Rs, mstrTable
        WriteData sData, Rs, sCaption
    ElseIf UCase(mstrTable) = "GSP_QSTANFILELIST" Then      '标准档案
        'rs.Open "select GSP_QSTANFILELIST.CCODE,GSP_QSTANFILELIST.CNAME,GSP_QSTANFILELIST.CPARENTNODE,GSP_STANDARDTYPE.CNAME AS CPARENTNAME,GSP_QSTANFILELIST.CDEPCODE,DEPARTMENT.CDEPNAME,GSP_QSTANFILELIST.DDATE,GSP_QSTANFILELIST.CMAKER,GSP_QSTANFILELIST.CVERIFIER,GSP_QSTANFILELIST.CAPPROVER from GSP_QSTANFILELIST LEFT JOIN GSP_STANDARDTYPE ON GSP_QSTANFILELIST.CPARENTNODE=GSP_STANDARDTYPE.CCODE LEFT JOIN DEPARTMENT ON GSP_QSTANFILELIST.CDEPCODE=DEPARTMENT.CDEPCODE order by GSP_QSTANFILELIST.CPARENTNODE,GSP_QSTANFILELIST.CCODE  ASC", g_DbGSP.DbConnect, adOpenStatic, adLockReadOnly
        Rs.Open "select GSP_QSTANFILELIST.CCODE,GSP_QSTANFILELIST.CNAME,GSP_STANDARDTYPE.CNAME AS CPARENTNAME,DEPARTMENT.CDEPNAME,GSP_QSTANFILELIST.DDATE,GSP_QSTANFILELIST.CMAKER,GSP_QSTANFILELIST.CVERIFIER,GSP_QSTANFILELIST.CAPPROVER from GSP_QSTANFILELIST LEFT JOIN GSP_STANDARDTYPE ON GSP_QSTANFILELIST.CPARENTNODE=GSP_STANDARDTYPE.CCODE LEFT JOIN DEPARTMENT ON GSP_QSTANFILELIST.CDEPCODE=DEPARTMENT.CDEPCODE order by GSP_QSTANFILELIST.CCODE,GSP_QSTANFILELIST.CPARENTNODE  ASC", AdoCnn, adOpenStatic, adLockReadOnly
        WriteSytle1 sStyle, Rs, mstrTable
        WriteData1 sData, Rs, sCaption
    ElseIf UCase(mstrTable) = "GSP_QMANAFILELIST" Then      '管理制度档案
        'rs.Open "select GSP_QMANAFILELIST.CCODE,GSP_QMANAFILELIST.CNAME,GSP_QMANAFILELIST.CPARENTNODE,GSP_MANASYSTYPE.CNAME AS CPARENTNAME,GSP_QMANAFILELIST.CDEPCODE,DEPARTMENT.CDEPNAME,GSP_QMANAFILELIST.DDATE,GSP_QMANAFILELIST.CMAKER,GSP_QMANAFILELIST.CVERIFIER,GSP_QMANAFILELIST.CAPPROVER from GSP_QMANAFILELIST LEFT JOIN GSP_MANASYSTYPE ON GSP_QMANAFILELIST.CPARENTNODE=GSP_MANASYSTYPE.CCODE LEFT JOIN DEPARTMENT ON GSP_QMANAFILELIST.CDEPCODE=DEPARTMENT.CDEPCODE order by GSP_QMANAFILELIST.CPARENTNODE,GSP_QMANAFILELIST.CCODE  ASC", g_DbGSP.DbConnect, adOpenStatic, adLockReadOnly
        Rs.Open "select GSP_QMANAFILELIST.CCODE,GSP_QMANAFILELIST.CNAME,GSP_MANASYSTYPE.CNAME AS CPARENTNAME,DEPARTMENT.CDEPNAME,GSP_QMANAFILELIST.DDATE,GSP_QMANAFILELIST.CMAKER,GSP_QMANAFILELIST.CVERIFIER,GSP_QMANAFILELIST.CAPPROVER from GSP_QMANAFILELIST LEFT JOIN GSP_MANASYSTYPE ON GSP_QMANAFILELIST.CPARENTNODE=GSP_MANASYSTYPE.CCODE LEFT JOIN DEPARTMENT ON GSP_QMANAFILELIST.CDEPCODE=DEPARTMENT.CDEPCODE order by GSP_QMANAFILELIST.CCODE,GSP_QMANAFILELIST.CPARENTNODE  ASC", AdoCnn, adOpenStatic, adLockReadOnly
        WriteSytle1 sStyle, Rs, mstrTable
        WriteData1 sData, Rs, sCaption
    Else
        Rs.Open "select CCODE,CNAME,(CASE WHEN ISNULL(BEND,0)=0 THEN '否' ELSE '是' END)AS BEND from " & mstrTable & " order by CCODE,CPARENTNODE ASC", AdoCnn, adOpenStatic, adLockReadOnly
        WriteSytle2 sStyle, Rs, mstrTable
        WriteData2 sData, Rs, sCaption
    End If

    
    If Prn.SetDataStyleXML(sData, False, sStyle, False, "Default") <> 0 Then Exit Sub
    Select Case Key
        Case "SetUp"                            '打印设置
            Prn.PageSetup
             Call Prn.TriggerEvent(0)
        Case "Print"                            '打印
            Prn.DoPrint
        Case "Preview"                          '预览
            Prn.SetOwner (Prn.Parent.hwnd)
            Prn.PrintPreview
        Case "SaveFile"                         '输出
            Dim sTypeList As String
            Dim sSizeList As String
            Dim i As Long
            Dim e As Long
            i = 0
            Call GetTypeSize(sTypeList, sSizeList, Rs)
            e = Prn.ExportToFile(i, sTypeList, sSizeList, "", "")
            If e = 3021 Then
                MsgBox "没有数据，不能输出！", vbInformation, Msg_Title
            Else
                If e <> 0 And e <> 3999 And e <> 3006 Then
                    MsgBox "输出文件不成功！", vbCritical, Msg_Title
                End If
            End If
    End Select
    If Rs.State = 1 Then Rs.Close
    Set Rs = Nothing
End Sub

'-----------------------------------------------------------
'功能：填写打印设施类型格式字符串
'
'参数：sXML符合XML格式字符串
'      rst 打印数据
'      mstrTable 打印单据表名
'
'返回：
'
'-----------------------------------------------------------
Public Sub WriteSytle(ByRef sXML As String, ByVal rst As ADODB.Recordset, ByVal mstrTable As String)
    Dim i           As Integer
    Dim iWidth      As Integer
    Dim sColWidth   As String
    
'    sColWidth = ""
'    For i = 0 To rst.Fields.Count - 1 + 1 '加1表示添加序号
'        iWidth = 3000
'        iWidth = CInt(iWidth * 25.4 * 10 / 1440)
'        If i = 0 Then
'            sColWidth = sColWidth + CStr(CInt(iWidth - 2 * iWidth / 3)) + ","
'        ElseIf i = 2 Then
'            sColWidth = sColWidth + CStr(CInt(iWidth + 1 * iWidth / 3) + 200) + ","
'        ElseIf i = 1 Then
'            sColWidth = sColWidth + CStr(iWidth - 200) + ","
'        Else
'            sColWidth = sColWidth + CStr(CInt(iWidth * 1 / 6 + 50)) + ","
'        End If
'
'    Next i
    sColWidth = "176,176,176,800,320"
    'sColWidth = Left(sColWidth, Len(sColWidth) - 1)
'sColWidth = "1,1,1,1,1"
    Call g_oPub.GetNewPrnStyle(sColWidth, sXML, mstrTable, SrvDB, Nothing, "", rst)
End Sub


'-----------------------------------------------------------
'功能：填写打印档案格式字符串
'
'参数：sXML符合XML格式字符串
'      rst 打印数据
'      mstrTable 打印单据表名
'
'返回：
'
'-----------------------------------------------------------
Public Sub WriteSytle1(ByRef sXML As String, ByVal rst As ADODB.Recordset, ByVal mstrTable As String)
    Dim i           As Integer
    Dim iWidth      As Integer
    Dim sColWidth   As String
    
    sColWidth = ""
    iWidth = 3000
    sColWidth = CStr("100,200,500,200,200,200,100,100,100")
'
'    For i = 0 To Rst.Fields.Count - 1 + 1 '加1表示添加序号
'        iWidth = 3000
'        iWidth = CInt(iWidth * 25.4 * 10 / 1440)
'        If i = 0 Then
'            sColWidth = sColWidth + CStr(CInt(iWidth - 4 * iWidth / 5)) + ","
'        ElseIf i = 2 Then
'            sColWidth = sColWidth + CStr(CInt(iWidth + 1 * iWidth / 4) + 200) + ","
'        ElseIf i = 1 Then
'            sColWidth = sColWidth + CStr(iWidth - 200) + ","
'        Else
'            sColWidth = sColWidth + CStr(iWidth * 1 / 6 + 50) + ","
'        End If
'
'    Next i
'    sColWidth = Left(sColWidth, Len(sColWidth) - 1)
    
    Call g_oPub.GetNewPrnStyle(sColWidth, sXML, mstrTable, SrvDB, Nothing, "", rst)
End Sub


'-----------------------------------------------------------
'功能：填写打印其他分类、档案格式字符串
'
'参数：sXML符合XML格式字符串
'      rst 打印数据
'      mstrTable 打印单据表名
'
'返回：
'
'-----------------------------------------------------------
Public Sub WriteSytle2(ByRef sXML As String, ByVal rst As ADODB.Recordset, ByVal mstrTable As String)
    Dim i           As Integer
    Dim iWidth      As Integer
    Dim sColWidth   As String
    
    sColWidth = ""
    For i = 0 To rst.Fields.Count - 1 + 1 '加1表示添加序号
        iWidth = 3000
        iWidth = CInt(iWidth * 25.4 * 10 / 1440)
        If i = 0 Then
            sColWidth = sColWidth + CStr(CInt(iWidth - 2 * iWidth / 3)) + ","
        ElseIf i = 2 Then
            sColWidth = sColWidth + CStr(CInt(iWidth + 1 * iWidth / 3) + 200) + ","
        ElseIf i = 1 Then
            sColWidth = sColWidth + CStr(iWidth - 200) + ","
        Else
            sColWidth = sColWidth + CStr(CInt(iWidth * 1 / 3 + 50)) + ","
        End If
        
    Next i
    sColWidth = Left(sColWidth, Len(sColWidth) - 1)
    
    Call g_oPub.GetNewPrnStyle(sColWidth, sXML, mstrTable, SrvDB, Nothing, "", rst)
    
End Sub

'-----------------------------------------------------------
'功能：填写打印设施类型类容
'
'参数：sXML符合XML格式字符串
'      rst 打印数据
'     sCaption打印标题
'
'返回：
'
'-----------------------------------------------------------
Public Sub WriteData(ByRef sXML As String, ByVal Rs As ADODB.Recordset, ByVal sCaption)
    Dim sTitle  As String
    Dim sBit    As String
    Dim sChFld  As String

    sBit = "BSYSTEM,BEND"
    sTitle = sCaption
    sChFld = "序号," & sCaption & "编码," & sCaption & "名称," & sCaption & "内容," & sCaption & "备注"
'    sChFld = "序号," & sCaption & "编码," & sCaption & "名称,是否末级,系统默认"
    Call g_oPub.GetData(sXML, sChFld, sBit, sTitle, g_oLogin.cAcc_Id, g_oLogin.cUserId, g_oLogin.cUserName, Rs, SrvDB)

End Sub
'-----------------------------------------------------------
'功能：填写打印档案类容
'
'参数：sXML符合XML格式字符串
'      rst 打印数据
'     sCaption打印标题
'
'返回：
'
'-----------------------------------------------------------
Public Sub WriteData1(ByRef sXML As String, ByVal Rs As ADODB.Recordset, ByVal sCaption)
    Dim sTitle  As String
    Dim sBit    As String
    Dim sChFld  As String

    sBit = ""
    sTitle = sCaption
    sChFld = "序号," & sCaption & "编码," & sCaption & "名称," & sCaption & "分类名称,部门名称,编写日期,编写人,审核人,审批人"
    Call g_oPub.GetData(sXML, sChFld, sBit, sTitle, g_oLogin.cAcc_Id, g_oLogin.cUserId, g_oLogin.cUserName, Rs, SrvDB)

End Sub
'-----------------------------------------------------------
'功能：填写打印其他分类、档案内容
'
'参数：sXML符合XML格式字符串
'      rst 打印数据
'     sCaption打印标题
'
'返回：
'
'-----------------------------------------------------------
Public Sub WriteData2(ByRef sXML As String, ByVal Rs As ADODB.Recordset, ByVal sCaption)
    Dim sTitle  As String
    Dim sBit    As String
    Dim sChFld  As String

    sBit = "BEND"
    sTitle = sCaption
    sChFld = "序号," & sCaption & "编码," & sCaption & "名称"
'    sChFld = "序号," & sCaption & "编码," & sCaption & "名称,是否末级"
    Call g_oPub.GetData(sXML, sChFld, sBit, sTitle, g_oLogin.cAcc_Id, g_oLogin.cUserId, g_oLogin.cUserName, Rs, SrvDB)

End Sub


'-----------------------------------------------------------
'功能：获得字段数据类型何大小
'
'参数：sTypeList 数据类型字符串
'      sSizeList 字段大小字符串
'      RsType字段记录集
'
'返回：
'
'-----------------------------------------------------------
Private Sub GetTypeSize(ByRef sTypeList As String, ByRef sSizeList As String, ByVal RsType As ADODB.Recordset)
    Call g_oPub.GetTypeSize(sTypeList, sSizeList, RsType)
End Sub

'-----------------------------------------------------------
'功能：显示帮助
'
'参数：frm 显示帮助的窗体

'
'返回：
'
'-----------------------------------------------------------
Public Sub ShowHelpConText(frm As Form)
   
    On Error Resume Next
    Err.Clear
    If Trim(AppPath) = "" Then
        MsgBox "无法显示帮助目录，该工程没有相关联的帮助。", vbInformation, Msg_Title
        Exit Sub
    End If
    Screen.MousePointer = 11
'    frm.HelpContextID = mstrHelpID
    htmlHelp mhwndMain, AppPath, IIf(mstrHelpID = 0, HH_DISPLAY_topic, HH_HELP_CONTEXT), mstrHelpID
    Screen.MousePointer = 1
    If Err Then
        MsgBox Err.Description, vbInformation, Msg_Title
    End If
End Sub

'-----------------------------------------------------------
'功能：窗体resize
'
'参数：frm
'
'返回：
'
'-----------------------------------------------------------
Public Sub Resize(frm As Object)
    On Error Resume Next
    If frm.WindowState = 1 Then Exit Sub
    If frm.Width < 10155 Then frm.Width = 10155
    If frm.Height < 7065 Then frm.Height = 7065
    
    'frm.Tlb.Width = frm.Width
    frm.CTBCtrl1.Width = frm.Width
    frm.CTBCtrl1.Left = 0
    frm.CTBCtrl1.Top = 0
  
    frm.Picture1.Width = frm.Width
    frm.Picture1.Height = frm.Picture2.Top - frm.CTBCtrl1.Height
    frm.Picture1.Top = frm.CTBCtrl1.Height
    frm.Label3.Caption = mstrCaption
    frm.Label3.Left = (frm.Width - frm.Label3.Width) / 2
    
    frm.Line1.X1 = frm.Label3.Left - 1560 - 500
    frm.Line1.X2 = frm.Line1.X1 + 1560
    frm.Line2.X1 = frm.Label3.Left + frm.Label3.Width + 500
    frm.Line2.X2 = frm.Line2.X1 + 1560
    
    frm.TREE1.Width = frm.Width * (2 / 5)
    frm.Picture2.Left = frm.TREE1.Width - 20
    frm.Picture2.Width = frm.Width - frm.TREE1.Width
    
    frm.TREE1.Height = frm.Height - frm.CTBCtrl1.Height - frm.Picture1.Height - IIf(frm.Stb.Visible = True, frm.Stb.Height, 0) - 650
    frm.Picture2.Height = frm.TREE1.Height + 10

    '设置状态栏宽度
'    setStb frm
    With frm.Stb
         .Panels(1).Width = frm.Width * 3 / 9 - 200
         .Panels(2).Width = frm.Width * 2 / 9
         .Panels(3).Width = frm.Width * 2 / 9 - 100
         .Panels(4).Width = frm.Width * 1 / 9
         .Panels(5).Width = frm.Width * 1 / 9 + 100
         .ZOrder 0
    End With
    frm.STBTimer.Panels(1).Width = frm.Stb.Width
End Sub

'-----------------------------------------------------------
'功能：设置窗体状态栏
'
'参数：frm 状态栏所在的窗体
'
'返回：
'
'-----------------------------------------------------------
Public Sub setStb(frm As Form)
    On Error Resume Next
    frm.Stb.Panels.Clear
    frm.Stb.Panels.Add 1, "k1"
    frm.Stb.Panels.Add 2, "k2"
    frm.Stb.Panels.Add 3, "k3"
    frm.Stb.Panels.Add 4, "k4"
    frm.Stb.Panels.Add 5, "k5"
    '设置状态栏宽度
    With frm.Stb
         .Panels(1).Width = frm.Width * 3 / 9 - 200
         .Panels(1).Alignment = sbrLeft
         .Panels(1).text = "账套：[" & g_oLogin.cAcc_Id & "]" & g_oLogin.cAccName
         .Panels(2).Width = frm.Width * 2 / 9
         .Panels(2).text = "操作员：" & g_oLogin.cUserName & IIf(g_oLogin.IsAdmin = True, "(账套主管)", "")
         .Panels(3).Width = frm.Width * 2 / 9 - 100
         .Panels(3).text = "当前记录数："
         .Panels(4).Width = frm.Width * 1 / 9
         .Panels(4).text = g_oLogin.CurDate
         .Panels(5).Width = frm.Width * 1 / 9 + 100
         .Panels(5).text = "【用友软件】"
    End With
'    frm.STBTimer.Height = frm.Stb.Height
    frm.STBTimer.Panels.Clear
    frm.STBTimer.Panels.Add 1, "k1"
    frm.STBTimer.Panels(1).Width = frm.Width
    frm.STBTimer.Panels(1).text = Time()
    frm.Stb.Visible = True
    frm.STBTimer.Visible = False
End Sub

'-----------------------------------------------------------
'功能：编码或名称输入控制
'
'参数：asc输入的ascii，bcode是否是编码控件
'
'返回：
'
'-----------------------------------------------------------
Public Sub GetAsc(ByRef asc As Integer, bcode As Boolean)
    On Error Resume Next
    '编码只能输入a-z、A-Z、0-9、backspace
    If bcode Then
        If asc >= 48 And asc <= 57 Then
        ElseIf asc >= 65 And asc <= 90 Then
        ElseIf asc >= 97 And asc <= 122 Then
        ElseIf asc = 8 Then
        Else
            asc = 0
        End If
    Else
        '下面是名称不能输入的符号
        Select Case Chr(asc)
        Case "“", "”", "’", "‘"
            asc = 0
        Case " ", "`", "~", "!", "@", "#", "$", "^", "&", "(", ")", ":", "|", "<", ">", "?", "'", """", "\", "/"
            asc = 0
        End Select
    End If
End Sub

'-----------------------------------------------------------
'功能：textbox的keyup后验证输入串
'
'参数：txt 要验证的TextBox， bcode是否是编码
'
'返回：
'
'-----------------------------------------------------------
Public Sub GetPressUp(txt As TextBox, bcode As Boolean)
    On Error Resume Next
    Dim i       As Long
    Dim item    As Integer
    Dim text    As String
    
    text = ""
    For i = 1 To Len(txt.text)
        item = asc(Mid(txt.text, i, 1))
        Call GetAsc(item, bcode)
        If item <> 0 Then text = text & Chr(item)
    Next
    text = Replace(text, "[档案]", "")
    text = Replace(text, "[系统]", "")
    txt.text = text
End Sub


'-----------------------------------------------------------
'功能：计算字符串长度
'
'参数：要计算长度的字符串
'
'返回：字符串长度
'
'-----------------------------------------------------------
Public Function EnterLen(cSource As String) As Integer

    Dim dLen        As Double
    Dim maxLen      As Integer
    Dim i           As Integer
        
    dLen = 0
    maxLen = Len(cSource)
    
    For i = 1 To maxLen
        If asc(Mid(cSource, i, 1)) > 0 And asc(Mid(cSource, i, 1)) < 256 Then
           dLen = dLen + 1
        Else
           dLen = dLen + 2
        End If
    Next i
    
    EnterLen = dLen
Exit Function
End Function


'-----------------------------------------------------------
'功能：在keypress事件中，验证输入后的字符串
'
'参数：asc当前输入的ascii，txt要输入的textbox
'
'返回：输入后的合法的字符串
'
'-----------------------------------------------------------

Public Function GetText(asc As Integer, txt As TextBox) As String
    On Error Resume Next
    Dim sText        As String
    Dim lSelStart    As Long
    
    
    lSelStart = txt.SelStart
    sText = txt.text
    If txt.SelLength > 0 Then
        If asc = 22 Then        '粘贴
            sText = Left(sText, lSelStart) & Clipboard.GetText & Right(sText, Len(sText) - lSelStart - txt.SelLength)
        ElseIf asc = 8 Or asc = 46 Then
            sText = Left(sText, lSelStart) & Right(sText, Len(sText) - lSelStart - txt.SelLength)
        Else
            sText = Left(sText, lSelStart) & Chr(asc) & Right(sText, Len(sText) - lSelStart - txt.SelLength)
        End If
    Else
        If asc = 22 Then        '粘贴
            sText = Left(sText, lSelStart) & Clipboard.GetText & Right(sText, Len(sText) - lSelStart)
        ElseIf asc = 8 Then
            If txt.SelStart > 0 Then sText = Left(sText, lSelStart - 1) & Right(sText, Len(sText) - lSelStart)
        ElseIf asc = 46 Then
            If txt.SelStart < Len(sText) Then sText = Left(sText, lSelStart) & Right(sText, Len(sText) - lSelStart - 1)
        Else
            sText = Left(sText, lSelStart) & Chr(asc) & Right(sText, Len(sText) - lSelStart)
        End If
    End If
    GetText = sText
End Function

'-----------------------------------------------------------
'功能：返回定长字符串(Unicode)
'
'参数：cSource原字符串，lLen 要返回的长度
'
'返回：返回定长字符串
'
'-----------------------------------------------------------
Public Function GetString(cSource As String, ByVal lLen As Long) As String

    Dim dLen        As Long
    Dim lTemp       As Long
    Dim maxLen      As Long
    Dim i           As Long
    Dim sTemp       As String
    Dim sChr        As String
    
    dLen = 0
    sTemp = ""
    maxLen = Len(cSource)
    
    For i = 1 To maxLen
        sChr = Mid(cSource, i, 1)
        If asc(sChr) > 0 And asc(sChr) < 256 Then
            lTemp = 1
        Else
            lTemp = 2
        End If
        If lTemp + dLen > lLen Then
            Exit For
        Else
            dLen = dLen + lTemp
            sTemp = sTemp & sChr
        End If
    Next i
    GetString = sTemp
End Function

'-----------------------------------------------------------
'功能：检查cSource中的非法字符
'
'参数：cSource要检查的字符串
'
'返回：是否有非法字符
'
'-----------------------------------------------------------
Public Function IsQualify(ByVal cSource As String) As Boolean
    Dim maxLen As Integer
    Dim i As Integer
        
    maxLen = Len(cSource)
    IsQualify = True
    For i = 1 To maxLen
        Select Case Mid(cSource, i, 1)
        Case "“", "”", "‘", "’"
            IsQualify = False
            MsgBox "非法字符""" & Mid(cSource, i, 1) & """！", vbCritical + vbOKOnly
            Exit For
        Case " ", "`", "~", "!", "@", "#", "$", "^", "&", "(", ")", ":", "|", "<", ">", "?", "'", """", "\", "/"
            IsQualify = False
            MsgBox "非法字符""" & Mid(cSource, i, 1) & """！", vbCritical + vbOKOnly
            Exit For
        End Select
    Next
    
End Function

'-----------------------------------------------------------
'功能：检查是否年度封账
'
'参数：
'
'返回：是否年度封账
'
'-----------------------------------------------------------
Public Function bYearEnd() As Boolean
    Dim sSql As String
    Dim rst  As ADODB.Recordset
On Error Resume Next
    bYearEnd = False
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    sSql = "Select isnull(bflag_gs,0) as bflag_gs from gl_mend where iperiod=12"
    rst.Open sSql, AdoCnn, adOpenStatic, adLockReadOnly, adCmdText
    If rst.Fields("bflag_gs").Value = True Then bYearEnd = True
    rst.Close
    Set rst = Nothing
End Function

'-----------------------------------------------------------
'功能：设置toolbar的tooltip
'
'参数：tlb要设置的toolbar
'
'返回：
'
'-----------------------------------------------------------
Public Sub SetTooltip(Tlb As Control)
'    Select Case mlngType
'    Case 0:
        Tlb.Buttons("Exit").ToolTipText = "退出"
        Tlb.Buttons("Help").ToolTipText = "帮助"
End Sub


'-----------------------------------------------------------
'功能：在keypress事件中，验证日期输入后的字符串
'
'参数：asc当前输入的ascii，txt要输入的textbox
'
'返回：
'
'-----------------------------------------------------------
Public Sub GetDateAsc(ByRef asc As Integer, txt As TextBox)
    On Error Resume Next
    Dim lSelStart       As Long
    Dim sText           As String
    Dim a()             As String
    Dim i               As Long
    
    lSelStart = txt.SelStart
    sText = Trim(txt)
    txt = sText
    If (asc >= 48 And asc <= 57) Or asc = 8 Or asc = 46 Then
        If sText = "" Then Exit Sub
        If txt.SelLength > 0 Then
            If asc = 8 Or asc = 46 Then
                sText = Left(sText, lSelStart) & Right(sText, Len(sText) - lSelStart - txt.SelLength)
            Else
                sText = Left(sText, lSelStart) & Chr(asc) & Right(sText, Len(sText) - lSelStart - txt.SelLength)
            End If
        Else
            If asc = 8 Then
                If txt.SelStart > 0 Then sText = Left(sText, lSelStart - 1) & Right(sText, Len(sText) - lSelStart)
            ElseIf asc = 46 Then
                If txt.SelStart < Len(sText) Then sText = Left(sText, lSelStart) & Right(sText, Len(sText) - lSelStart - 1)
            Else
                sText = Left(sText, lSelStart) & Chr(asc) & Right(sText, Len(sText) - lSelStart)
            End If
        End If
        a = Split(sText, "-")
        For i = 0 To UBound(a)
            Select Case i
            Case 0:
                If Len(a(i)) > 4 Then           '输入年'年只能四位
                    asc = 0
                    Exit Sub
                End If
            Case 1:
                If Val(a(1)) > 12 Or Len(a(1)) > 2 Then          '输入月
                    asc = 0
                    Exit Sub
                End If
            Case 2:
                If Val(a(2)) > 31 Or Len(a(2)) > 2 Then        '输入日期
                    asc = 0
                    Exit Sub
                End If
            End Select
         Next
    ElseIf asc = 45 Then
        If sText = "" Or txt.SelStart = 0 Then                        '第一个不能输入-
            asc = 0
        Else                                        '不能输入两个以上-
            a = Split(sText, "-")
            If UBound(a) >= 2 Then asc = 0
        End If
    ElseIf asc = 39 Or asc = 37 Then
    Else
        asc = 0
    End If
End Sub


'-----------------------------------------------------------
'功能：删除分类和档案的冗余纪录
'
'参数：strtable 要删除纪录的表名，strartable要删除纪录的档案的表名
'
'返回：
'
'-----------------------------------------------------------
Public Sub DeleteRec(ByVal strTable As String, Optional ByVal strArTable As String = "")
    
    '删除多余分类和档案
    
    AdoCnn.BeginTrans
    '删除分类
    AdoCnn.Execute "delete from " & strTable & " where isnull(cparentnode,'')<>'' and cparentnode not in(select ccode from " & strTable & ")"
    '删除档案
    If strArTable <> "" Then
        AdoCnn.Execute "delete from GSP_ARCHIVE where(cparentnode not in(select ccode from GSP_STANDARDTYPE where isnull(ccode,'')<>'' and isnull(bend,0)=1)) and (cparentnode not in(select ccode from GSP_MANASYSTYPE where isnull(ccode,'')<>'' and isnull(bend,0)=1))"
    End If
    AdoCnn.CommitTrans

End Sub


Public Function bBaseSysValue(ByVal sTable As String, ByVal cCode As String, ByVal sOpt As String) As Boolean
    Dim strSql As String:   Dim Rs As Object
    
    strSql = "select * from EFBWGL_BaseSysvalue where cBaseTableName='" & VBA.Replace(sTable, "'", "''") & "' and cBasecCode='" & VBA.Replace(cCode, "'", "''") & "' "
    strSql = strSql & ""
    Set Rs = AdoCnn.Execute(strSql)
    bBaseSysValue = Not Rs.EOF
    Exit Function
End Function




