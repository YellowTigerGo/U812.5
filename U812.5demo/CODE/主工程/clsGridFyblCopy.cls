VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGridFyblCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'费用比例和费用分组比例批量复制功能 Create by Wuyp 2008.03
Option Explicit
Implements IPlugInExCtl

Private Const ROOT      As String = "v_mt_items" '表名
Private Const PK        As String = "citemcode"            '主键
Private Const COLNAME   As String = "序号|选择|项目编号|项目名称|部门编号|部门名称|"
Private Const COLFIELDS As String = "NULL_Order|sel|citemcode|citemname|citemccode|citemcname|NULL_Hide_Flag"

Private Const I_COL_NULL_Order            As Long = 0   '序号
Private Const I_COL_sel                   As Long = 1   '选择
Private Const I_COL_citemcode             As Long = 2   '项目编号
Private Const I_COL_citemname             As Long = 3   '项目名称
Private Const I_COL_citemccode            As Long = 4   '部门编号
Private Const I_COL_citemcname            As Long = 5   '部门名称
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
Private m_DataXML               As String
Private m_HeadXML               As String
Private m_BodyXML               As String
Private m_VoucherCO             As EFPBVoucherCo.ClsVoucherCO_GDZC

Private Sub InitGridCol()
    With m_Grid
        .Rows = 1
        .FormatString = COLNAME
        
        '数据类型
        .ColDataType(I_COL_sel) = EditStr
        .ColDataType(I_COL_citemcode) = EditStr
        .ColDataType(I_COL_citemname) = EditStr
        .ColDataType(I_COL_citemccode) = EditStr
        .ColDataType(I_COL_citemcname) = EditStr
        
        '列宽度
        .colwidth(0) = 500
        .colwidth(I_COL_sel) = 700
        .colwidth(I_COL_citemcode) = 900
        .colwidth(I_COL_citemname) = 2000
        .colwidth(I_COL_citemccode) = 900
        .colwidth(I_COL_citemcname) = 1600
        .colwidth(I_COL_NULL_Hide_Flag) = 0

        '对齐
        .ColAlignment(I_COL_sel) = 4            '居中对齐
        .ColAlignment(I_COL_NULL_Order) = 4     '居中对齐
        .ColAlignment(I_COL_citemcode) = 1      '居中对齐
        .ColAlignment(I_COL_citemccode) = 1     '居中对齐
        
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
    Set m_VoucherCO = Nothing
End Sub

Private Sub IPlugInExCtl_Clone()

End Sub

Private Property Get IPlugInExCtl_DispName() As String
    IPlugInExCtl_DispName = "批量复制"
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
            
        Case UCase("SetDomHead")
            With XmlStringToDom(vPara)
                .selectSingleNode("//z:row").Attributes.removeNamedItem("editprop").nodeTypedValue = "A"
                m_HeadXML = .xml
            End With
        Case UCase("SetDomBody")
            With XmlStringToDom(vPara)
                With .selectNodes("//z:row")
                    If .length > 0 Then
                        Dim i As Long
                        For i = 0 To .length - 1
                            .Item(i).Attributes.getNamedItem("editprop").nodeTypedValue = "A"
                        Next
                    End If
                End With
                m_BodyXML = .xml
            End With
            
        Case UCase("SetVoucherCO")
            Set m_VoucherCO = vPara
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
    m_DataXML = sXML
    
    Dim sDepCode As String
    Dim sItemCode As String
    Dim sItem_Class As String
    sDepCode = GetXMLAttrValue(m_HeadXML, "//z:row", "cdepcode")
    sItem_Class = GetXMLAttrValue(m_HeadXML, "//z:row", "citem_class")
    sItemCode = GetXMLAttrValue(m_HeadXML, "//z:row", "citemcode")
    
    Dim sVouchType As String
    sVoucherType = GetXMLAttrValue(m_HeadXML, "//z:row", "cvouchtype")
    
    Dim sExpCode As String
    sExpCode = GetXMLAttrValue(m_HeadXML, "//z:row", "cexpcode")
    
    Dim sCCode As String
    sCCode = GetXMLAttrValue(m_BodyXML, "//z:row", "ccode")

    
    
        
        
    Dim sMeta As String
    If sVoucherType = "01" Then
  'shanlan delete --------原来的代码---------------------------------
'        sMeta = "<ROOT>" & vbCrLf & _
'               "<V_MT_budget04 Format='RsDom' Select='select citemcode,citemname,a.citemccode,citemcname from fitemss" & sItem_Class & " a left outer join fitemss" & sItem_Class & "class b on a.citemccode=b.citemccode " & vbCrLf & _
'               " where &apos;" & sItem_Class & "&apos;+citemcode not in (select citem_class+citemcode from V_MT_baseset01 where cvouchtype=&apos;01&apos; and cexpcode=&apos;" & sExpCode & "&apos;)  " & _
'               " order by a.citemccode,citemcode'>" & vbCrLf & _
'               "</V_MT_budget04></ROOT>"
        'shanlan add ------------批量复制时不需要判断在费用项目科目对照中已经存在的不能复制--------------------------------
         sMeta = "<ROOT>" & vbCrLf & _
               "<V_MT_budget04 Format='RsDom' Select='select citemcode,citemname,a.citemccode,citemcname from fitemss" & sItem_Class & " a left outer join fitemss" & sItem_Class & "class b on a.citemccode=b.citemccode " & vbCrLf & _
               " where &apos;" & sItem_Class & "&apos;+citemcode not in (select citem_class+citemcode from V_MT_baseset01 where cvouchtype=&apos;01&apos; and cexpcode=&apos;" & sExpCode & "&apos; and citem_class=&apos;" & sItem_Class & "&apos; and citemcode=&apos;" & sItemCode & "&apos;)  " & _
               " order by a.citemccode,citemcode'>" & vbCrLf & _
               "</V_MT_budget04></ROOT>"
    ElseIf sVoucherType = "11" Then
         sMeta = "<ROOT>" & vbCrLf & _
               "<V_MT_budget04 Format='RsDom' Select='select citemcode,citemname from fitemss" & sItem_Class & vbCrLf & _
               " where &apos;" & sItem_Class & "&apos;+citemcode not in (select citem_class+citemcode from V_MT_basesets11 where cvouchtype=&apos;" & sVoucherType & "&apos; and citem_class=&apos;" & sItem_Class & "&apos;)  " & _
               " order by citemccode,citemcode'>" & vbCrLf & _
               "</V_MT_budget04></ROOT>"
    ElseIf sVoucherType = "12" Then
         sMeta = "<ROOT>" & vbCrLf & _
               "<V_MT_budget04 Format='RsDom' Select='select citemcode,citemname from fitemss" & sItem_Class & vbCrLf & _
               " where &apos;" & sItem_Class & "&apos;+citemcode not in (select citem_class+citemcode from V_MT_basesets12 where cvouchtype=&apos;" & sVoucherType & "&apos; and citem_class=&apos;" & sItem_Class & "&apos;)  " & _
               " order by citemccode,citemcode'>" & vbCrLf & _
               "</V_MT_budget04></ROOT>"
    Else
        sMeta = "<ROOT>" & vbCrLf & _
                "<V_MT_budget04 Format='RsDom' Select='select citemcode,citemname,a.citemccode,citemcname from fitemss" & sItem_Class & " a left outer join fitemss" & sItem_Class & "class b on a.citemccode=b.citemccode " & vbCrLf & _
                " where &apos;" & sItem_Class & "apos;+citemcode not in (select citem_class+citemcode from MT_baseset where cvouchtype=&apos;" & sVoucherType & "&apos;)  " & _
                " order by a.citemccode,citemcode'>" & vbCrLf & _
                "</V_MT_budget04></ROOT>"
    End If
    
    Dim sResultXML As String
    Dim sErr As String
    
    Dim m_cDAL      As CPlugInExDAL      '数据操作类
    Dim m_IDAL      As IPlugInExDAL     '数据操作接口
    
    Set m_cDAL = New CPlugInExDAL
    Set m_IDAL = m_cDAL
    
    If m_IDAL.Load(m_Conn, sMeta, sResultXML, sErr) Then
        With XmlStringToDom(sResultXML)
            m_Grid.Rows = 1
            Call SetGridData(m_Grid, COLFIELDS, .selectSingleNode("//V_MT_budget04").xml)
        End With
    End If
    
    Set m_cDAL = Nothing
    Set m_IDAL = Nothing
    m_Grid.ReadOnly = True
End Function

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
            If .TextMatrix(.row, I_COL_citemcode) <> "" Then
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
        Set btnX = .buttons.Add(, "DoBatchCopy", "复制", tbrDefault)
        btnX.ToolTipText = "复制"
        btnX.Description = btnX.ToolTipText
        btnX.Visible = True
        
        Set btnX = .buttons.Add(, "s1", "|", tbrSeparator)
        btnX.ToolTipText = "|"
        btnX.Description = btnX.ToolTipText
        btnX.Visible = True
        
        Set btnX = .buttons.Add(, "SelectAll", "全选", tbrDefault)
        btnX.ToolTipText = "全选"
        btnX.Description = btnX.ToolTipText
        btnX.Visible = True
        
        Set btnX = .buttons.Add(, "SelectNone", "刷新", tbrDefault)
        btnX.ToolTipText = "刷新"
        btnX.Description = btnX.ToolTipText
        btnX.Visible = True
                
        Set btnX = .buttons.Add(, "s2", "|", tbrSeparator)
        btnX.ToolTipText = "|"
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
        Case LCase("DoBatchCopy")  '批量复制
            Call DoBatchCopy
            IPlugInExCtl_SetData m_DataXML
            
        Case LCase("SelectAll")
            Call DoSelect("Y")
            
        Case LCase("SelectNone")   '刷新
            IPlugInExCtl_SetData m_DataXML
            
        Case LCase("Exit")      '退出
            m_Form.Hide
            
    End Select

End Function

Private Sub DoSelect(ByVal sFlag As String)
    With m_Grid
        Dim iRow As Long
        For iRow = 1 To .Rows - 1
            .TextMatrix(iRow, I_COL_sel) = sFlag
        Next
    End With
End Sub

'批量复制
Private Function DoBatchCopy()
    Dim sSelectXML As String
    'LDX    2009-08-02  Add Beg
    Dim sCCode As String
    Dim Rst As New ADODB.Recordset
    Dim strSql As String
    'LDX    2009-08-02  Add Beg
    Dim sDepCode As String
    Dim sItem_Class As String
    Dim sItemCode As String
    Dim sVouchType As String
    Dim sExpCode As String
    Dim Domhead As DOMDocument
    Dim Dombody As DOMDocument
    Dim i As Long
    
    sSelectXML = GetSelectXML()

    With XmlStringToDom(sSelectXML).selectNodes("//v_mt_items")

        If .length > 0 Then
            Screen.MousePointer = vbHourglass
            sVouchType = GetXMLAttrValue(m_HeadXML, "//z:row", "cvouchtype")
            sExpCode = GetXMLAttrValue(m_HeadXML, "//z:row", "cexpcode")
            sItem_Class = GetXMLAttrValue(m_HeadXML, "//z:row", "citem_class")
            'LDX    2009-08-02  Add Beg
            sCCode = GetXMLAttrValue(m_HeadXML, "//z:row", "ccode")
            'LDX    2009-08-02  Add End
            
            For i = 0 To .length - 1
                sDepCode = .Item(i).selectSingleNode("citemccode").nodeTypedValue
                sItemCode = .Item(i).selectSingleNode("citemcode").nodeTypedValue
                
                '检查是否已经存在了相应部门、项目的设置表，如果有则忽略
                If CheckVoucher(sVouchType, sItem_Class, sItemCode, "", sExpCode) Then
                    Set Domhead = New DOMDocument
                    Set Dombody = New DOMDocument
                    Domhead.loadXML m_HeadXML
                    Dombody.loadXML m_BodyXML
                    
                    If sVouchType <> "01" Then
                        'Domhead.selectSingleNode("//z:row").Attributes.getNamedItem("cdepcode").nodeTypedValue = sDepCode
                    End If
                    
                    Domhead.selectSingleNode("//z:row").Attributes.getNamedItem("citemcode").nodeTypedValue = sItemCode
                    
                    'LDX    2009-08-02  Add Beg
                    strSql = "select max(ccode) as ccode from  mt_baseset where cvouchtype=" & sVouchType & ""
                    If Rst.State <> 0 Then Rst.Close
                    Rst.Open strSql, DBConn, adOpenStatic, adLockReadOnly
                    If Not Rst.EOF Then
                        sCCode = Rst.Fields("ccode").value + 1
                    End If
                    sCCode = Format(sCCode, "000000")
                    Domhead.selectSingleNode("//z:row").Attributes.getNamedItem("ccode").nodeTypedValue = sCCode
                    'LDX    2009-08-02  Add End
                    Call MakeVoucher(Domhead, Dombody)
                End If
            Next
            Screen.MousePointer = vbDefault
            MsgBox "复制完毕！", vbInformation
        Else
            MsgBox "请先复制到的项目!", vbInformation
        End If

    End With
End Function

Private Function MakeVoucher(ByRef Domhead As DOMDocument, ByRef Dombody As DOMDocument)
    Dim strError As String, vNewID As Long
    strError = m_VoucherCO.Save(Domhead, Dombody, 0, vNewID, XmlStringToDom("<EAI>0</EAI>"))

    If strError <> "" Then
        MsgBox strError
    End If
End Function

Private Function CheckVoucher(ByVal sVouchType As String, Optional ByVal sItem_Class As String = "", Optional ByVal sItemCode As String = "", Optional ByVal sDepCode As String = "", Optional ByVal sExpCode As String = "") As Boolean
    CheckVoucher = False
    Dim sSQL As String
        ''shanlan add--------------------------------------------
    Dim strError As String
    Dim domheadtemp As DOMDocument
    Dim ssqltemp As String
    Dim RsTemp As ADODB.Recordset
    '----------------------------------------------------------
    
    If sVouchType = "01" And sItem_Class <> "" And sItemCode <> "" And sExpCode <> "" Then
        If sDepCode = "" Then
            sSQL = "select count(*) as iCount from v_MT_basesets01 where citem_class='" & sItem_Class & "' and citemcode='" & sItemCode & "' and cvouchtype='" & sVouchType & "' and cexpcode='" & sExpCode & "'"
        ElseIf sItem_Class = "" And sItemCode = "" Then
            sSQL = "select count(*) as iCount from v_MT_basesets01 where cdepcode='" & sDepCode & "' and citemcode='" & sItemCode & "' and cvouchtype='" & sVouchType & "' and cexpcode='" & sExpCode & "'"
        Else
            GoTo DoExit
        End If
    Else
        'sSQL = "select count(*) as iCount from MT_baseset where cdepcode='" & sDepCode & "' and citemcode='" & sItemCode & "' and cvouchtype='" & sVouchType & "'"
        CheckVoucher = True
        Exit Function
    End If
    
    Dim rs As ADODB.Recordset
    Set rs = m_Conn.Execute(sSQL)
    If rs.BOF And rs.EOF Then
        CheckVoucher = True
    Else
        If rs("iCount") = 0 Then
            CheckVoucher = True
         'shanlan add ---------------------------------------------------------------------
        ElseIf rs("icount") > 0 Then
            If MsgBox("项目" & sItemCode & "已经设定费用科目对照，是否覆盖?", vbYesNo) = vbYes Then
              '首先调用单据的删除方法删除此单据
                  Set domheadtemp = New DOMDocument
                  ssqltemp = "select *  from v_MT_basesets01 where citem_class='" & sItem_Class & "' and citemcode='" & sItemCode & "' and cvouchtype='" & sVouchType & "' and cexpcode='" & sExpCode & "'"
                  Set RsTemp = New ADODB.Recordset
                  If RsTemp.State <> 0 Then RsTemp.Close
                  RsTemp.Open ssqltemp, m_Conn, adOpenForwardOnly, adLockReadOnly
                  RsTemp.Save domheadtemp, adPersistXML
                  
                  Set domheadtemp = New DOMDocument
                  ssqltemp = "select *  from v_MT_baseset01 where citem_class='" & sItem_Class & "' and citemcode='" & sItemCode & "' and cvouchtype='" & sVouchType & "' "
                  Set RsTemp = New ADODB.Recordset
                  If RsTemp.State <> 0 Then RsTemp.Close
                  RsTemp.Open ssqltemp, m_Conn, adOpenForwardOnly, adLockReadOnly
                  RsTemp.Save domheadtemp, adPersistXML
                  
                  If IIf(IsNull(RsTemp.Fields("cverifier")), "", RsTemp.Fields("cverifier")) <> "" Then
                   strError = m_VoucherCO.VerifyVouch(domheadtemp, False)
                  Else
                    strError = m_VoucherCO.Delete(domheadtemp)
                  End If
                    If strError <> "" Then
                        MsgBox strError
                    Else
                        CheckVoucher = True
                    End If
                              
            End If
        '-------------------------------------------------------------------------------------
        End If
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
DoExit:
    CheckVoucher = False
End Function

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



