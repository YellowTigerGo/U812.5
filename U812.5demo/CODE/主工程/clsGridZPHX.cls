VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGridZPHX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'报账单支票批量核销功能  Create by Wuyp 2008.03
Option Explicit
Implements IPlugInExCtl

Private Const ROOT      As String = "V_MT_budget04" '表名
Private Const PK        As String = "id"            '主键
Private Const COLNAME   As String = "序号|选择|ID|支票号|支票金额|ufts|"
Private Const COLFIELDS As String = "NULL_Order|sel|id|zph|zpje|ufts|NULL_Hide_Flag"

Private Const I_COL_NULL_Order            As Long = 0   '序号
Private Const I_COL_sel                   As Long = 1   '选择
Private Const I_COL_id                    As Long = 2   'id
Private Const I_COL_zph                   As Long = 3   '支票号
Private Const I_COL_zpje                  As Long = 4   '支票金额
Private Const I_COL_ufts                  As Long = 5   'ufts
Private Const I_COL_NULL_Hide_Flag        As Long = 6   '标志列

Private m_oGrid                 As Object   'm_oGrid要比m_Grid多出一些控件本身的属性，比如:Height,Width,...
Private WithEvents m_Grid       As MsSuperGrid.SuperGrid
Attribute m_Grid.VB_VarHelpID = -1
Private WithEvents m_ToolBar    As UFToolBarCtrl.UFToolbar
Attribute m_ToolBar.VB_VarHelpID = -1
Private m_IF                    As IPlugInExIFx
Private m_Form                  As Form
Private m_Conn                  As ADODB.Connection
Private m_oLogin                As Object
Private m_FilterXML             As String

Private Sub InitGridCol()
    With m_Grid
        .Rows = 1
        .FormatString = COLNAME
        
        '数据类型
        .ColDataType(I_COL_sel) = EditStr
        .ColDataType(I_COL_id) = EditStr
        .ColDataType(I_COL_zph) = EditStr
        .ColDataType(I_COL_zpje) = EditDbl
        .ColDataType(I_COL_ufts) = EditStr
        
        '列宽度
        .colwidth(0) = 500
        .colwidth(I_COL_sel) = 700
        .colwidth(I_COL_id) = 0
        .colwidth(I_COL_zph) = 2000
        .colwidth(I_COL_zpje) = 1200
        .colwidth(I_COL_ufts) = 0
        .colwidth(I_COL_NULL_Hide_Flag) = 0     '隐藏列
        
        '对齐方式
        .ColFormat(I_COL_zpje) = "#,##0.00"
        .ColAlignment(I_COL_zpje) = 7   '右对齐
        .ColAlignment(I_COL_sel) = 4  '居中对齐
        .ColAlignment(I_COL_NULL_Order) = 4   '居中对齐
        
    End With
End Sub

Private Sub Class_Terminate()
    Set m_Grid = Nothing
    Set m_oGrid = Nothing
    Set m_ToolBar = Nothing
    Set m_IF = Nothing
    Set m_Form = Nothing
    Set m_oLogin = Nothing
    Set m_oLogin = Nothing
    
End Sub

Private Sub IPlugInExCtl_Clone()

End Sub

Private Property Get IPlugInExCtl_DispName() As String
    IPlugInExCtl_DispName = "报账单支票核销"
End Property

Private Function IPlugInExCtl_DoOtherOperation(ByVal sOptName As String, Optional vPara As Variant, Optional sErrMsg As String, Optional oSender As Object, Optional vResult As Variant) As Boolean
    Select Case UCase(sOptName)
        Case UCase$("SetButton")
            Dim Tlb As MSComctlLib.Toolbar
            Set Tlb = vPara
            Call SetButtons(Tlb)  '设置工具栏按钮
        
        Case UCase("SetToolBar")    '传入工具栏对象
            Set m_ToolBar = vPara
            
        Case UCase("SetForm")       '传入窗体对象
            Set m_Form = vPara
        
        Case UCase("SetConnection") '传入数据库连接
            Set m_Conn = vPara
    End Select
End Function

Private Sub IPlugInExCtl_EndAdd()
End Sub

Private Sub IPlugInExCtl_EndModify()
End Sub

Private Function IPlugInExCtl_GetData() As String

End Function

Private Function IPlugInExCtl_Init(Optional oLogin As Object, _
                                   Optional ByVal vPara As Variant, _
                                   Optional oSender As Object) As Boolean
    Set m_oLogin = oLogin
    Set m_Grid = vPara
    Set m_oGrid = vPara
    Call InitGrid(m_Grid) '通用Grid初始化
    Call InitGridCol
End Function

Private Function IPlugInExCtl_SetData(ByVal sXML As String) As Boolean
    m_FilterXML = sXML
    
    Dim sDepCode As String
    Dim sItemCode As String
    Dim sItem_Class As String
    sDepCode = GetXMLAttrValue(sXML, "//", "cdepcode")
    sItem_Class = GetXMLAttrValue(sXML, "//", "citem_class")
    sItemCode = GetXMLAttrValue(sXML, "//", "citemcode")
    
    Dim sMeta As String
    sMeta = "<ROOT>" & vbCrLf & _
            "<V_MT_budget04 Format='RsDom' Select='SELECT [id],zph,isnull(zpje,0)+isnull(cdefine7,0) as zpje,ufts FROM V_MT_budget04 where citem_class=&apos;" & sItem_Class & "&apos; and citemcode=&apos;" & sItemCode & "&apos; and isnull(cverifier,&apos;&apos;) !=&apos;&apos;  and isnull(chandler,&apos;&apos;) =&apos;&apos; '>" & vbCrLf & _
            "</V_MT_budget04></ROOT>"
    
    Dim sResultXML As String
    Dim sErr As String
    
    Dim m_cDAL      As CPlugInExDAL     '数据操作类
    Dim m_IDAL      As IPlugInExDAL     '数据操作接口
    
    Set m_cDAL = New CPlugInExDAL
    Set m_IDAL = m_cDAL
    
    If m_IDAL.Load(m_Conn, sMeta, sResultXML, sErr) Then
        With XmlStringToDom(sResultXML)
            m_Grid.Rows = 1
            Call SetGridData(m_Grid, COLFIELDS, .selectSingleNode("//V_MT_budget04").xml)
            m_Grid.Rows = m_Grid.Rows + 1
            Call SetSumData
        End With
    End If
    
    Set m_cDAL = Nothing
    Set m_IDAL = Nothing
    m_Grid.ReadOnly = True
End Function

Private Sub SetSumData()
    With m_Grid
        Dim iRow As Long
        Dim iCol As Long
                
        Dim iSum As Currency, iSelSum As Currency
        iSum = 0: iSelSum = 0
        
        For iRow = 1 To .Rows - 1
            iSum = iSum + val(.TextMatrix(iRow, I_COL_zpje))
            If .TextMatrix(.Rows - 1, I_COL_sel) = "Y" Then
                iSelSum = iSelSum + val(.TextMatrix(iRow, I_COL_zpje))
            End If
        Next
        
        .TextMatrix(.Rows - 1, I_COL_zpje) = iSum
        .TextMatrix(.Rows - 1, I_COL_zph) = "合计"
    End With
End Sub

Private Sub IPlugInExCtl_StartAdd()
End Sub

Private Sub IPlugInExCtl_StartModify()
End Sub

Private Function IPlugInExCtl_Talk(ByVal sXML As String) As String
End Function

Private Function IPlugInExCtl_VerifyData(Optional sErrMsg As String, Optional ByVal bShowMsgbox As Boolean = False) As Boolean
End Function

Private Sub m_Grid_DblClick()
    With m_Grid
        If .row > 0 Then
            If .TextMatrix(.row, I_COL_id) <> "" Then
                If .TextMatrix(.row, I_COL_sel) = "Y" Then
                    .TextMatrix(.row, I_COL_sel) = ""
                Else
                    .TextMatrix(.row, I_COL_sel) = "Y"
                End If
            End If
        End If
    End With
End Sub

Private Sub m_Grid_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(" ") Then
        Call m_Grid_DblClick
    End If
End Sub

Private Sub m_ToolBar_OnCommand(ByVal enumType As UFToolBarCtrl.ENUM_MENU_OR_BUTTON, ByVal cButtonId As String, ByVal cmenuid As String)
    Call Operating(cButtonId)
End Sub

Private Sub SetButtons(ByRef Tlb As MSComctlLib.Toolbar)
    Dim btnX As MSComctlLib.Button
    With Tlb
        Set btnX = .buttons.Add(, "DoVerify", "核销", tbrDefault)
        btnX.ToolTipText = "核销"
        btnX.Description = btnX.ToolTipText
        btnX.Visible = True
        
        '//支票借款单审核权限
        If Not LockItem("MT300403", True, False) Then
            btnX.Enabled = False
        End If
        Call LockItem("MT300403", False, False)
        
        Set btnX = .buttons.Add(, "Refresh", "刷新", tbrDefault)
        btnX.ToolTipText = "刷新"
        btnX.Description = btnX.ToolTipText
        btnX.Visible = True
        
        Set btnX = .buttons.Add(, "Exit", "返回", tbrDefault)
        btnX.ToolTipText = "返回"
        btnX.Description = btnX.ToolTipText
        btnX.Visible = True
        
        

    End With
End Sub

Private Function Operating(ByVal sKey As String) As Boolean

    Select Case LCase(sKey)
        Case LCase("DoVerify")  '核销
            Call DoVerify
            
        Case LCase("Refresh")   '刷新
            Call IPlugInExCtl_SetData(m_FilterXML)
            
        Case LCase("Exit")      '退出
            m_Form.Hide
            
    End Select

End Function

'核销
Private Function DoVerify()
    Dim sSelectXML As String
    sSelectXML = GetSelectXML()

    With XmlStringToDom(sSelectXML).selectNodes("//V_MT_budget04")

        If .length > 0 Then
            Dim i            As Long
            Dim AffectedLine As Long     'SQL执行后受影响的行数.AffectedLine=0 表示没有修改成功
            Dim strSql       As String
            Dim sUserID      As String   '当前操作员
            Dim sID          As String   '当前借款单ID
            Dim errMsg       As String
            Dim bOK          As Boolean
            Dim sDepCode     As String
            Dim sItem_Class  As String
            Dim sItemCode    As String
            
            sUserID = m_oLogin.cUserId
            sDepCode = GetXMLAttrValue(m_FilterXML, "//", "cdepcode")
            sItem_Class = GetXMLAttrValue(m_FilterXML, "//", "citem_class")
            sItemCode = GetXMLAttrValue(m_FilterXML, "//", "citemcode")
            
            m_Conn.BeginTrans
            bOK = True

            For i = 0 To .length - 1
                
                sID = GetXMLElementValue(.Item(i).xml, "//id")
                strSql = "UPDATE MT_budget set chandler='" & sUserID & "' where id=" & sID & " and replace(convert(nchar,convert(money,ufts),2),' ','')='" & GetXMLElementValue(.Item(i).xml, "//ufts") & "'"
                m_Conn.Execute strSql, AffectedLine
                
                errMsg = ""
                GetMT_sum m_Conn, sID, sDepCode, sItem_Class, sItemCode, errMsg

                If AffectedLine = 0 Then
                    bOK = False
                    errMsg = "支票号:" & GetXMLElementValue(.Item(i).xml, "//zph") & "的借款单已经被其他人改动,核销不成功,请刷新后重试!"
                    Exit For
                ElseIf errMsg <> "" Then
                    errMsg = "支票号:" & GetXMLElementValue(.Item(i).xml, "//zph") & "的借款单核销失败:" & errMsg
                    bOK = False
                    Exit For
                End If

            Next

            If bOK Then
                m_Conn.CommitTrans
                MsgBox "核销成功！", vbInformation
                IPlugInExCtl_SetData m_FilterXML '刷新Grid
            Else
                m_Conn.RollbackTrans
                MsgBox errMsg
            End If

        Else
            MsgBox "请先选择要核销的支票!", vbInformation
        End If

    End With
    
End Function


'@MT_id nvarchar(50),             --当前单据ID
'@cDepCode nvarchar(50),          --部门编码
'@cItemCode nvarchar(50),         --项目编码
'@cExpCode nvarchar(50)='',       --预算费用
'@ErrStr nvarchar(500) OUTPUT     --出错信息
Private Sub GetMT_sum(conn As ADODB.Connection, MT_id As String, cDepCode As String, cItem_Class As String, cItemCode As String, ErrStr As String, Optional cExpCode As String)
    Dim AdoComm As ADODB.Command
    On Error GoTo DoErr
    Set AdoComm = New ADODB.Command
    With AdoComm
        .ActiveConnection = conn
        .CommandText = "MT_sum"
        .CommandType = adCmdStoredProc
        .Prepared = False
        .Parameters.Append .CreateParameter("MT_id", adVarChar, adParamInput, 50, MT_id)
        .Parameters.Append .CreateParameter("cDepCode", adVarChar, adParamInput, 50, cDepCode)
        .Parameters.Append .CreateParameter("cItem_class", adVarChar, adParamInput, 50, cItem_Class)
        .Parameters.Append .CreateParameter("cItemCode", adVarChar, adParamInput, 50, cItemCode)
        .Parameters.Append .CreateParameter("cExpCode", adVarChar, adParamInput, 50, cExpCode)
        .Parameters.Append .CreateParameter("ErrStr", adVarChar, adParamOutput, 500)
        .Execute
        ErrStr = CStr(.Parameters("ErrStr"))
    End With
    Set AdoComm = Nothing
    Exit Sub
DoErr:
    ErrStr = "发生错误：" & Err.Description
    Set AdoComm = Nothing
End Sub

'提取所选行的XML
Private Function GetSelectXML() As String

    Dim arrColName()   As String
    Dim arrColFields() As String
    
    arrColName = Split(COLNAME, "|")
    arrColFields = Split(COLFIELDS, "|")
    
    Dim iRow       As Long
    Dim sRowXML    As String
    Dim sSelectXML As String
    
    With m_Grid

        For iRow = 1 To .Rows - 1

            If .TextMatrix(iRow, I_COL_sel) = "Y" Then
                sRowXML = GridRowToXML(m_Grid, iRow, arrColName, arrColFields, ROOT, PK)
                sSelectXML = sSelectXML & sRowXML & vbCrLf
            End If

        Next

    End With

    GetSelectXML = "<ROOT>" & vbCrLf & sSelectXML & "</ROOT>"
    
End Function

