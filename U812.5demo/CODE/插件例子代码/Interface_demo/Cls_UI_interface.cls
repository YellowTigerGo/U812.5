VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_UI_interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'接口类说明
'本插件主要是实现在用户界面上的接口插件
Option Explicit

Dim m_login As Object           '当前登陆对象
Dim DBconn As Object            '当前数据库连接对象
Dim StrCardNum As String           '当前单据的CardNum
'Dim Vouchers As UAPVoucherControl85.ctlVoucher
Dim Vouchers As Object          '当前单据对象
Dim Vouchers_Frm As Object      '当前窗体
Dim DomH As New DOMDocument     '
Dim DomB As New DOMDocument     '
Dim g_busines1 As Object


' '初始化
'Public Function Init(ByVal objLogin As Object, ByVal objForm As Object, ByVal objVoucher As Object, msbar As Object)
''MsgBox "Init"
'End Function

'单据数据装载后事件
Public Sub LoadAfter_VoucherData(strUserErr As String, bSuc As Boolean)
 On Error GoTo Errhandle
     Dim i As Integer
     Vouchers_Frm.picSure.Visible = True
     
     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
'     bSuc = True
ExitSub:
     Exit Sub
Errhandle:
     bSuc = False
     strUserErr = Err.Description
End Sub


'初始化用户插件
'mLogin 当前登陆对象
'DBcon 当前数据库连接对象
'Frm 当前单据的CardNum
'CardNum 当前单据对象
'strUserErr  错误信息
'bSuc  =true 表示初始化成功   =false 表示初始化失败
Public Sub Userdll_Init(g_busines As Object, mLogin As Object, DBcon As Object, Frm As Object, CardNum As String, strUserErr As String, bSuc As Boolean)
    Set g_busines1 = g_busines
    Set m_login = mLogin
    Set DBconn = DBcon
    Set Vouchers_Frm = Frm
    If Frm.Name = "frmVouchNew" Then
        Set Vouchers = Vouchers_Frm.Voucher
    End If
    StrCardNum = CardNum
    bSuc = True
'    MsgBox " Userdll_Init 插件初始化配置成功！"
End Sub

 
'Voucher   当前单据对象
'ButtonKey 按扭关键字
'strUserErr  错误信息
'bSuc  =true 表示成功并终止系统流程   =false表示失败并继续系统流程
Public Sub Before_ButtonClick(Voucher As Object, ButtonKey As String, strUserErr As String, bSuc As Boolean)
    Dim id As String
    Dim strsql As String
    Dim i As Long
    On Error GoTo hErr
    bSuc = False
    Select Case ButtonKey
        Case "createbom"
            DoCreateBom Voucher
            bSuc = True
        Case "btnopen"
            DoClose False, Voucher
            bSuc = True
        Case "btnclose"
            DoClose True, Voucher
            bSuc = True
        Case "doAction"
            doAction Voucher
    End Select
    Exit Sub
hErr:
    MsgBox Err.Description, vbCritical, "提示"
End Sub

Private Function doAction(Voucher As Object) As Boolean
    On Error GoTo hErr
    Dim doBom As New UF_Public_base.clsBom
    If MsgBox("是否执行此操作？", vbOKCancel, "提示") = vbCancel Then
        GoTo hFinish
    End If
    doBom.init DBconn, m_login
    doBom.DoUpdate Voucher
    doAction = True
    GoTo hFinish
hErr:
    MsgBox "执行操作异常：" & Err.Description, vbCritical, "提示"
hFinish:
    Set doBom = Nothing
End Function


Private Sub DoCreateBom(VchLst As Object)
    Dim id As String
    Dim strsql As String
    Dim sError As String
    Dim i As Long
    Dim cInvCode As String
    Dim citemcode As String
    Dim cabinetnumber As String
    Dim ConfigFree As Integer
    Dim ConfigFree3 As Integer
    Dim billtype As String
    Dim cFree1 As String
    Dim cFree2 As String
    Dim iQty As Double
    On Error GoTo hErr
    
    frmProgress.Show 0
    frmProgress.Msg = "BOM生成..."
    
    For i = 1 To VchLst.rows - 1
        DoEvents
        frmProgress.Msg = "正在处理第" & i & "行数据..."
        If VchLst.TextMatrix(i, VchLst.GridColIndex("selcol")) = "Y" Then 'And VchLst.TextMatrix(i, VchLst.GridColIndex("bomcreated")) = "否"
            id = VchLst.TextMatrix(i, VchLst.GridColIndex("id"))
            cInvCode = VchLst.TextMatrix(i, VchLst.GridColIndex("stockcode"))
            citemcode = VchLst.TextMatrix(i, VchLst.GridColIndex("projectnumber"))
            cabinetnumber = VchLst.TextMatrix(i, VchLst.GridColIndex("cabinetnumber"))
            ConfigFree = GetInventoryConfigFree(cInvCode, "bConfigFree1")
            iQty = VchLst.TextMatrix(i, VchLst.GridColIndex("num"))
            '存货自由项处理
            If ConfigFree = -1 Then
                MsgBox "存货编码" & cInvCode & "在U8系统中不存在。", vbCritical, "提示"
                    GoTo hFinish
            ElseIf ConfigFree = 1 Then
                If citemcode = "" Then
                    MsgBox "存货启用了结构性自由项1，项目编码不能为空", vbCritical, "提示"
                    GoTo hFinish
                Else
                    If Not InsertInventoryFreeValue(cInvCode, "cFree1", citemcode, cabinetnumber, sError) Then
                        MsgBox sError, vbCritical, "提示"
                        GoTo hFinish
                    End If
                    cFree1 = citemcode
                End If
            End If
            ConfigFree = GetInventoryConfigFree(cInvCode, "bConfigFree2")
            If ConfigFree = -1 Then
                MsgBox "存货编码" & cInvCode & "在U8系统中不存在。", vbCritical, "提示"
                    GoTo hFinish
            ElseIf ConfigFree = 1 Then
                If cabinetnumber = "" Then
                    MsgBox "存货启用了结构性自由项2，屏体编码不能为空", vbCritical, "提示"
                    GoTo hFinish
                Else
                    If Not InsertInventoryFreeValue(cInvCode, "cFree2", citemcode, cabinetnumber, sError) Then
                        MsgBox sError, vbCritical, "提示"
                        GoTo hFinish
                    End If
                    cFree2 = cabinetnumber
                End If
            End If
            
            ConfigFree3 = GetInventoryConfigFree(cInvCode, "bConfigFree3")
            If Not Insertbas_partFree(cInvCode, citemcode, cabinetnumber, ConfigFree3, sError) Then
                MsgBox sError, vbCritical, "提示"
                GoTo hFinish
            End If
            'BOM生成
            '校验
            billtype = VchLst.TextMatrix(i, VchLst.GridColIndex("billtype"))
            If BomExist(cInvCode, cFree1, cFree2, ConfigFree3) Then
                If Not AddBomComponent(id, cInvCode, cFree1, cFree2, ConfigFree3, iQty, sError) Then
                    MsgBox "第" & i & "行物料BOM子件追加失败:" & sError, vbCritical, "提示"
                    GoTo hFinish
                End If
                If Not AddMom_moallocate(id, cInvCode, cFree1, cFree2, ConfigFree3, iQty, sError) Then
                    MsgBox "第" & i & "行物料生产订单子件材料追加失败:" & sError, vbCritical, "提示"
                    GoTo hFinish
                End If
            Else
                If Not CallU8Api4bom(id, cInvCode, cFree1, cFree2, ConfigFree3, iQty, sError) Then
                    MsgBox "第" & i & "行物料BOM创建失败:" & sError, vbCritical, "提示"
                    GoTo hFinish
                Else
                    If billtype = "冷作清单" Then
                        DBconn.Execute "update hzland_lzdetail set bomcreated='是' where id='" & id & "'"
                    ElseIf billtype = "材料清单" Then
                        DBconn.Execute "update hzland_materialsdetail set bomcreated='是' where id='" & id & "'"
                    End If
                End If
            End If
        End If
    Next
    MsgBox "物料清单生成完毕.", vbInformation, "提示"
    GoTo hFinish
hErr:
    MsgBox Err.Description, vbCritical, "提示"
hFinish:
    frmProgress.Hide
End Sub

'判断BOM是否存在
Private Function BomExist(ByVal cInvCode As String, cFree1 As String, cFree2 As String, ConfigFree3 As Integer) As Boolean
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    On Error GoTo hErr
    strsql = "select * from bas_part  a inner join bom_parent b on a.PartId =b.ParentId " & _
            " where InvCode='" & cInvCode & "' and Free1='" & cFree1 & "' and Free2='" & cFree2 & "'"
    If ConfigFree3 = 1 Then
        strsql = strsql & " AND Free3='空'"
    End If
    rs.Open strsql, DBconn
    If Not rs.BOF And Not rs.EOF Then
        BomExist = True
    End If
hErr:
hFinish:
    If rs.State <> 0 Then
        rs.Close
    End If
    Set rs = Nothing
End Function

Private Function AddMom_moallocate(ByVal id As String, ByVal cInvCode As String, cFree1 As String, cFree2 As String, ConfigFree3 As Integer, iQty As Double, ByRef strErr As String) As Boolean
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    
    Dim BomId As String
    Dim MoDid  As String
    Dim mom_orderdetailSortSeq As Double
    Dim AllocateId As String
    Dim AllocateIds As String
    Dim mom_moallocateSortSeq As Double
    Dim StartDate As String
    Dim DueDate As String
    Dim MoCode As String
    Dim stockcode As String
    Dim ComponentId As String
    Dim OpComponentId As String
    Dim cWhcode As String
    Dim csubsysbarcode As String
    Dim iMQty As Double
    
    On Error GoTo ErrHandler
    strsql = "select bomid from v_bom_head where invcode='" & cInvCode & "' and InvFree_1='" & cFree1 & "' and InvFree_2='" & cFree2 & "'"
    If ConfigFree3 = 1 Then
        strsql = strsql & " and InvFree_3='空'"
    End If
    strsql = strsql & " order by CreateDate desc"
    rs.Open strsql, DBconn
    If Not rs.EOF And Not rs.BOF Then
        BomId = rs!BomId
    End If
    strsql = "select d.MoDid,d.SortSeq,t.MoCode,d.Qty from mom_orderdetail d left outer join mom_order t on d.MoId =t.MoId where d.bomid='" & BomId & "'"
    If rs.State <> 0 Then rs.Close
    rs.Open strsql, DBconn
    While Not rs.EOF
        MoCode = rs!MoCode
        MoDid = rs!MoDid & ""
        iMQty = str2Dbl(rs!Qty & "")
        mom_orderdetailSortSeq = rs!SortSeq
        mom_moallocateSortSeq = GetSortSeq("mom_moallocate", "MoDId", MoDid)
        strsql = "select * from mom_morder where modid='" & MoDid & "'"
        If rsTmp.State <> 0 Then rsTmp.Close
        rsTmp.Open strsql, DBconn
        If Not rsTmp.EOF And Not rsTmp.BOF Then
            StartDate = rsTmp!StartDate
            DueDate = rsTmp!DueDate
        End If
        strsql = "select * from EF_V_HZDesignList where _parentid='" & id & "' and id in(select Define28 from bom_opcomponent where bomid='" & BomId & "')" & _
                " and id not IN( select define28 from mom_moallocate where modid='" & MoDid & "')"
                
        If rsTmp.State <> 0 Then rsTmp.Close
        rsTmp.Open strsql, DBconn
        
        While Not rsTmp.EOF
            stockcode = rsTmp!stockcode & ""
            ComponentId = GetPartId(stockcode)
            cWhcode = GetInvWhCode(stockcode)
            mom_moallocateSortSeq = mom_moallocateSortSeq + 10
            AllocateId = GetVouchID("mom_moallocate", AllocateIds, 1, strErr)
            OpComponentId = GetOpComponentId(rsTmp!id)
            csubsysbarcode = "||MO21|" & MoCode & "|" & mom_orderdetailSortSeq & "|" & mom_moallocateSortSeq
            strsql = "insert into mom_moallocate ([AllocateId],[MoDId],[SortSeq],[OpSeq],[ComponentId],[FVFlag],[BaseQtyN],[BaseQtyD],[ParentScrap],[CompScrap],[Qty],[IssQty],[DeclaredQty],[StartDemDate],[EndDemDate],[WhCode],[LotNo],[WIPType],[ByproductFlag],[QcFlag],[Offset],[InvCode],[Free1],[Free2],[Free3],[Free4],[Free5],[Free6],[Free7],[Free8],[Free9],[Free10],[OpComponentId],[Define22],[Define23],[Define24],[Define25],[Define26],[Define27],[Define28],[Define29],[Define30],[Define31],[Define32],[Define33],[Define34],[Define35],[Define36],[Define37],[AuxUnitCode],[ChangeRate],[AuxBaseQtyN],[AuxQty],[ReplenishQty],[Remark],[TransQty],[ProductType],[SoType],[SoDId],[SoCode],[SoSeq],[DemandCode],[QmFlag],[OrgQty],[OrgAuxQty],[CostItemCode],[CostItemName],[RequisitionFlag],[RequisitionQty],[RequisitionIssQty],[CostWIPRel],[MoallocateSubId],[cSubSysBarCode],[PickingQty],[PickingAuxQty],[UpperMoQty],[InvAlloeFlag])" & _
                    "values(" & AllocateIds & "," & MoDid & "," & mom_moallocateSortSeq & ",'0000'," & ComponentId & ",1,'" & rsTmp!Num & "','" & iQty & "','0','0','" & (rsTmp!Num / iQty) * iMQty & "','0','0'" & _
                    ",'" & StartDate & "','" & DueDate & "','" & cWhcode & "',Null,3,'0','0',0,'" & stockcode & "','','','','','','','','','',''" & _
                    "," & OpComponentId & ",Null,Null,Null,Null,Null,Null,'" & rsTmp!id & "',Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,'0',Null,'0',0,0,Null,Null,Null,Null,'0','0','0',Null,Null,'1','0','0','0',Null" & _
                    ",'" & csubsysbarcode & "','0','0','1','0')"
            DBconn.Execute strsql
            rsTmp.MoveNext
        Wend
        rs.MoveNext
    Wend
    
    AddMom_moallocate = True
    GoTo hFinish
ErrHandler:
    strErr = Err.Description
hFinish:
    If rs.State <> 0 Then rs.Close
    Set rs = Nothing
End Function

Private Function AddBomComponent(ByVal id As String, ByVal cInvCode As String, cFree1 As String, cFree2 As String, ConfigFree3 As Integer, iQty As Double, ByRef strErr As String) As Boolean
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    
    Dim BomId As String
    Dim optionsid As String
    Dim optionsids As String
    Dim ComponentId As String
    Dim OpComponentId As String
    Dim OpComponentIds As String
    Dim SortSeq As Integer
    Dim cDepCode As String
    Dim cWhcode As String
    Dim stockcode As String
    Dim billtype As String
    
On Error GoTo ErrHandler
    strsql = "select bomid from v_bom_head where invcode='" & cInvCode & "' and InvFree_1='" & cFree1 & "' and InvFree_2='" & cFree2 & "'"
    If ConfigFree3 = 1 Then
        strsql = strsql & " and InvFree_3='空'"
    End If
    strsql = strsql & " order by CreateDate desc"
    rs.Open strsql, DBconn
    If Not rs.EOF And Not rs.BOF Then
        BomId = rs!BomId
    End If
    cDepCode = getAccinformation("PD", "cDepCode", DBconn)
    SortSeq = GetSortSeq("bom_opcomponent", "BomId", BomId)
    '根据设计清单id过滤掉已生成BOM的记录
    strsql = "select * from EF_V_HZDesignList where _parentid='" & id & "' and id not in(select Define28 from bom_opcomponent where bomid='" & BomId & "')"
    If rs.State <> 0 Then rs.Close
    rs.Open strsql, DBconn
    
    While Not rs.EOF
        billtype = rs!billtype & ""
        stockcode = rs!stockcode & ""
        Call Insertbas_partFree(stockcode, "", "", 0, strErr)
        ComponentId = GetPartId(stockcode)
        cWhcode = GetInvWhCode(stockcode)
        optionsid = GetVouchID("bom_opcomponentopt", optionsids, 1, strErr)
        strsql = "insert into bom_opcomponentopt ([OptionsId],[Offset],[WIPType],[AccuCostFlag],[DrawDeptCode],[Whcode],[OptionalFlag],[MutexRule],[PlanFactor],[CostWIPRel],[FeatureRel])" & _
                "values(" & optionsids & ",0,3,'1','" & cDepCode & "','" & cWhcode & "','0',2,'100','0','0')"
        DBconn.Execute strsql
        OpComponentId = GetVouchID("bom_opcomponent", OpComponentIds, 1, strErr)
        SortSeq = SortSeq + 10
        strsql = "insert into bom_opcomponent ([OpComponentId],[BomId],[SortSeq],[OpSeq],[ComponentId],[EffBegDate],[EffEndDate],[FVFlag],[BaseQtyN],[BaseQtyD],[CompScrap],[ByproductFlag],[OptionsId],[AuxUnitCode],[ChangeRate],[AuxBaseQtyN],[ProductType],[Define22],[Define23],[Define24],[Define25],[Define26],[Define27],[Define29],[Define30],[Define31],[Define32],[Define33],[Define34],[Define35],[Define36],[Define37],[Remark],[RecursiveFlag],[Free1],[Free2],[Free3],[Free4],[Free5],[Free6],[Free7],[Free8],[Free9],[Free10],[bProcessProduct],[bProcessMaterial],[Define28])" & _
                "values(" & OpComponentIds & "," & BomId & "," & SortSeq & ",'0000'," & ComponentId & ",'2000-01-01','2099-12-31',1,'" & rs!Num & "','" & iQty & "','0','0'," & optionsids & ",Null,Null,Null,0,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,'0','','','','','','','','','','','0','0','" & rs!id & "')"
        DBconn.Execute strsql
        If billtype = "冷作清单" Then
            DBconn.Execute "update hzland_lzdetail set bomcreated='是' where id='" & rs!id & "'"
        ElseIf billtype = "材料清单" Then
            DBconn.Execute "update hzland_materialsdetail set bomcreated='是' where id='" & rs!id & "'"
        End If
        rs.MoveNext
    Wend
    AddBomComponent = True
    GoTo hFinish
ErrHandler:
    strErr = Err.Description
hFinish:
    If rs.State <> 0 Then rs.Close
    Set rs = Nothing
End Function

Private Function GetInvWhCode(cInvCode As String) As String
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    On Error GoTo hErr
    strsql = "select cDefWareHouse from inventory where cinvcode='" & cInvCode & "'"
    rs.Open strsql, DBconn
    If Not rs.EOF And Not rs.BOF Then
        GetInvWhCode = rs!cDefWareHouse
    End If
    GoTo hFinish
hErr:

hFinish:
    If rs.State <> 0 Then rs.Close
    Set rs = Nothing
End Function


Private Function GetOpComponentId(Define28 As String) As Long
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    On Error GoTo hErr
    strsql = "select OpComponentId  from bom_opcomponent  where Define28='" & Define28 & "'"
    rs.Open strsql, DBconn
    If Not rs.EOF And Not rs.BOF Then
        GetOpComponentId = rs!OpComponentId
    End If
    GoTo hFinish
hErr:

hFinish:
    If rs.State <> 0 Then rs.Close
    Set rs = Nothing
End Function

Private Function GetPartId(cInvCode As String) As Long
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    On Error GoTo hErr
    strsql = "select PartId from bas_part where InvCode='" & cInvCode & "'"
    rs.Open strsql, DBconn
    If Not rs.EOF And Not rs.BOF Then
        GetPartId = rs!partid
    End If
    GoTo hFinish
hErr:

hFinish:
    If rs.State <> 0 Then rs.Close
    Set rs = Nothing
End Function

Private Function GetSortSeq(cTable As String, cFieldName As String, cFieldValue As String) As Integer
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    On Error GoTo hErr
    strsql = "select max(SortSeq) as SortSeq from " & cTable & " where " & cFieldName & "='" & cFieldValue & "'"
    rs.Open strsql, DBconn
    If Not rs.EOF And Not rs.BOF Then
        GetSortSeq = rs!SortSeq
    End If
hErr:

hFinish:
    If rs.State <> 0 Then rs.Close
End Function

Private Function CallU8Api4bom(ByVal id As String, ByVal cInvCode As String, cFree1 As String, cFree2 As String, ConfigFree3 As Integer, iQty As Double, ByRef strErr As String) As Boolean
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim n As Long
On Error GoTo ErrHandler
        
        '第二步：构造环境上下文对象，传入login，并按需设置其它上下文参数
        Dim u8EnvCtx As U8EnvContext
        '第三步：构造ApiBroker对象,调用Connect,传入Api的地址标识(Url)，传入上下文
        Dim u8apiBroker As U8ApiComBroker
        '第四步：API参数赋值
        '首先，获取扩展BO实体（通过扩展BO参数名）
        Dim extbo As ExtensionBusinessEntity
        
        '设置扩展BO实体的初始容量（可以自由设置扩展BO实体项目数为大于零的整数，也可以"不设置"而自动增加）
        'extbo.ItemCount = 10 '该步骤可省略
        '然后，使扩展BO实体增加一个新的数据项
        Dim newItem As ExtensionItem
        Dim Bom_Component As ExtensionBusinessEntity
        '以下代码示例只设置新增扩展BO实体第一项的值。扩展BO对象的各字段定义详见API服务接口定义
        Dim i As Long
        Dim j As Long
        
        ' 每次都重新初始化 api
        Set u8EnvCtx = New U8EnvContext
        Set u8EnvCtx.U8Login = m_login
        Set u8apiBroker = New U8ApiComBroker
        u8apiBroker.Connect "U8API/BOM/BomAdd", u8EnvCtx
        Set extbo = u8apiBroker.GetExtBoEntity("extbo")
        Set newItem = extbo.newItem()
        '----------------------------------- 必输字段 --------------------------------
        newItem.setValue "BomId", "1000000003"  '主键，Integer类型
        '----获取存货相关信息
        strsql = "select b.cComUnitName,c.cGroupName ,c.iGroupType ,a.* from inventory a left join ComputationUnit b on a.cComUnitCode =b.cComunitCode" _
                & " left join ComputationGroup c on a.cGroupCode =c.cGroupCode  where cinvcode ='" & cInvCode & "'"
        Set rs = New ADODB.Recordset
        rs.Open strsql, DBconn, adOpenStatic, adLockReadOnly
        
        
        If Not rs.EOF Then
            extbo(0).setValue "InvCode", cInvCode ' "001"  '母件编码，String类型
            extbo(0).setValue "InvName", rs.fields("cinvname").Value & ""  '母件名称(导出用)，String类型
            extbo(0).setValue "InvStd", rs.fields("cinvstd").Value & ""  '规格型号(导出用)，String类型
            extbo(0).setValue "InvUnitName", rs.fields("cComUnitName").Value & ""   '计量单位(导出用)，String类型
            extbo(0).setValue "InvUnit", rs.fields("cComUnitCode").Value & ""   '计量单位编码(导出用)，String类型
            extbo(0).setValue "InvGroupType", rs.fields("iGroupType").Value & ""   '存货计量单位组类型(导出用)，Integer类型
            extbo(0).setValue "InvGroupCode", rs.fields("cGroupCode").Value & ""   '存货量单位组编码(导出用)，String类型
            extbo(0).setValue "InvGroupName", rs.fields("cGroupName").Value & ""   '存货计量单位名称(导出用)，String类型
        Else
            strErr = "存货" & cInvCode & "在U8系统存货档案中不存在."
            Exit Function
        End If
        
        extbo(0).setValue "CreateUser", m_login.cUserName  ' "demo"  '建档人(导出用)，String类型
        extbo(0).setValue "ModifyUser", ""  '修改人(导出用)，String类型
        extbo(0).setValue "CreateDate", m_login.CurDate   '建档日期(导出用)，Date类型
        extbo(0).setValue "ModifyDate", ""  '修改日期(导出用)，Date类型
        extbo(0).setValue "ModifyTime", ""  '修改时间(导出用)，Date类型
        extbo(0).setValue "CreateTime", m_login.CurDate & " " & Time   '创建时间(导出用)，Date类型
        extbo(0).setValue "RelsUser", ""  '审核人(导出用)，String类型
        extbo(0).setValue "RelsDate", ""  '审核日期(导出用)，Date类型
        extbo(0).setValue "RelsTime", ""  '审核时间(导出用)，Date类型
        extbo(0).setValue "CloseUser", ""  '停用人(导出用)，String类型
        extbo(0).setValue "CloseDate", ""  '停用日期(导出用)，Date类型
        extbo(0).setValue "CloseTime", ""  '停用时间(导出用)，Date类型
        '---------------------------------- 非必输字段 -------------------------------
'        extbo(0).setValue "PartId", ""  'ParentId(导出用)，Integer类型
        extbo(0).setValue "BomType", "1"  'BOM类别(1主2替代)，Integer类型
'        extbo(0).setValue "BasEngineerFigNo", ""  '工程图号(导出用)，String类型
        extbo(0).setValue "ParentScrap", "0"  '母件损耗率(％0~99)，Double类型
        
        extbo(0).setValue "Version", "10"  '版本代号，Integer类型
        extbo(0).setValue "VersionEffDate", "2000-01-01" ' m_login.CurDate  '版本日期(YYYY-MM-DD)，Date类型
        
        extbo(0).setValue "VersionDesc", "设计清单生成" '版本说明，String类型
        
'        extbo(0).setValue "IdentCode", ""  '替代标识，String类型
'        extbo(0).setValue "IdentDesc", ""  '替代说明，String类型
'        extbo(0).setValue "InvAddCode", ""  '母件代号(导出用)，String类型
        extbo(0).setValue "InvFree_1", cFree1  '母件自由项1，String类型
        extbo(0).setValue "InvFree_2", cFree2  '母件自由项2，String类型
        If ConfigFree3 = 1 Then
            extbo(0).setValue "InvFree_3", "空"  '母件自由项3，String类型
        End If
        extbo(0).setValue "Define_1", ""  'BOM自定义项1，String类型
        
        extbo(0).setValue "Define_14", ""   'BOM自定义项14，String类型
        
        extbo(0).setValue "Define_16", ""  'BOM自定义项16，Double类型
        extbo(0).setValue "InvDefine_1", "" '本次倒入唯一码  用于失败后删除已导入的数据  '母件自定义项1，String类型
        extbo(0).setValue "InvDefine_16", ""  '母件自定义项16，Date类型
        extbo(0).setValue "BomState", "1"  '状态(导出用)，Integer类型
        extbo(0).setValue "VersionEndDate", "2099-12-31"  '版本失效日期(导出用)，Date类型   实在不行就在导入完毕后更新下实效日期

        Dim DDeptCode As String
        DDeptCode = getAccinformation("PD", "cDepCode", DBconn)
        '----处理子件
        strsql = "select * from EF_V_HZDesignList where _parentid='" & id & "'"
        If rs.State <> 0 Then rs.Close
        rs.Open strsql, DBconn
        While Not rs.EOF
            Set Bom_Component = extbo(0).GetSubEntity("Bom_Component")
            '----------------------------------- 必输字段 --------------------------------
            Bom_Component(n).setValue "DSortSeq", (n + 1) * 10 '子件行号(默认系统自增)，Integer类型
            Bom_Component(n).setValue "DOpSeq", "0000" 'Format((n + 1) * 10, "0000") '工序行号(默认0000)，String类型
            
            strsql = "select b.cComUnitName,c.cGroupName ,c.iGroupType ,a.* from inventory a left join ComputationUnit b on a.cComUnitCode =b.cComunitCode" _
                & " left join ComputationGroup c on a.cGroupCode =c.cGroupCode  where cinvcode ='" & rs!stockcode & "'"
            If rsTmp.State <> 0 Then rsTmp.Close
            rsTmp.Open strsql, DBconn, adOpenStatic, adLockReadOnly
            
            
            If Not rsTmp.EOF Then
                Bom_Component(n).setValue "DInvCode", rsTmp.fields("cinvcode").Value & ""   '子件编码，String类型
                Bom_Component(n).setValue "DInvName", rsTmp.fields("cinvname").Value & ""  '子件名称(导出用)，String类型
                Bom_Component(n).setValue "DInvStd", rsTmp.fields("cinvstd").Value & ""   '子件规格(导出用)，String类型
                Bom_Component(n).setValue "DInvUnit", rsTmp.fields("cComUnitCode").Value & ""  '计量单位编码(子件)，String类型
                Bom_Component(n).setValue "DInvUnitName", rsTmp.fields("cComUnitName").Value & ""  '计量单位(导出用)，String类型
                Bom_Component(n).setValue "DInvGroupCode", rsTmp.fields("cGroupCode").Value & ""  '计量单位组(导出用)，String类型
                Bom_Component(n).setValue "DInvGroupName", rsTmp.fields("cGroupName").Value & ""  '存货计量单位名称(导出用)，String类型
                Bom_Component(n).setValue "DInvGroupType", rsTmp.fields("iGroupType").Value & ""   '计量单位组类别(导出用)，Integer类型
            Else
                strErr = "存货" & rs!stockcode & "在U8存货档案中不存在."
                Exit Function
            End If
            
            
            Bom_Component(n).setValue "DBaseQtyN", str2Dbl(rs!Num) '基本用量(默认1)，Double类型
            Bom_Component(n).setValue "DBaseQtyD", iQty  '基础数量(默认1)，Double类型
            
            Bom_Component(n).setValue "DCompScrap", "0"  '子件损耗率％(默认0)，Double类型
            Bom_Component(n).setValue "DFVFlag", "1"  '固定用量(默认固定)，Integer类型
            Bom_Component(n).setValue "DWIPType", "3"  '供应类型(默认领用)，Integer类型
            Bom_Component(n).setValue "DEffBegDate", "2000-01-01" ' m_login.CurDate   '生效日期(默认母件生效日)，Date类型
            Bom_Component(n).setValue "DEffEndDate", "2099-12-31"  '失效日期(默认2099-12-31)，Date类型
            Bom_Component(n).setValue "DPlanRate", "100"  '计划比例％(默认100)，Double类型
            Bom_Component(n).setValue "DByproductFlag", "0"  '产出品(默认否)，Integer类型
            Bom_Component(n).setValue "DAccuCostFlag", "1"  '成本相关，Integer类型
            Bom_Component(n).setValue "DOptionalFlag", "0"  '是否可选(默认全选)，Integer类型
            Bom_Component(n).setValue "DMutexRule", "2"  '选择规则(默认全选)，Integer类型
            Bom_Component(n).setValue "DProductType", "0"  '产出类型(默认空，受是否产出品约束)，Integer类型
            If iQty = 0 Then
                Bom_Component(n).setValue "DQty", str2Dbl(rs!Num)
            Else
                Bom_Component(n).setValue "DQty", str2Dbl(rs!Num) / iQty '使用数量(默认1)，Double类型
            End If
    '        Bom_Component(n).setValue "DWhName", ""  '仓库名称(导出用)，String类型
    '        Bom_Component(n).setValue "DDeptName", ""  '部门名称(导出用)，String类型
    '        Bom_Component(n).setValue "DSubFlag", ""  '替代标志(导出用)，Integer类型
    '        Bom_Component(n).setValue "DAuxUnitName", ""  '辅助单位(导出用)，String类型
    '        Bom_Component(n).setValue "DAuxQty", ""  '辅助使用量，Double类型
            Bom_Component(n).setValue "DOpDesc", ""  '工序名称(导出用)，String类型
            '---------------------------------- 非必输字段 -------------------------------
            Bom_Component(n).setValue "DRemark", "" '备注，String类型
    '        Bom_Component(n).setValue "DInvAddCode", ""  '子件代号(导出用)，String类型
    '        Bom_Component(n).setValue "DOffset", ""  '偏置期(默认不偏置)，Integer类型
    '        Bom_Component(n).setValue "DWhCode", ""  '仓库编码(默认子件存货档案)，String类型
            Bom_Component(n).setValue "DDeptCode", DDeptCode  '领料部门(默认为子件存货档案)，String类型
            Bom_Component(n).setValue "DInvFree_1", ""  '子件自由项1，String类型
            Bom_Component(n).setValue "DInvFree_2", ""  '子件自由项2，String类型
'            Bom_Component(n).setValue "DInvFree_10", ""  '子件自由项10，String类型
            Bom_Component(n).setValue "DInvDefine_1", ""  '子件自定义项1，String类型
            Bom_Component(n).setValue "DInvDefine_16", ""  '子件自定义项16，Date类型
    '        Bom_Component(n).setValue "DSubDate", ""  '替换日期(导出用)，Date类型
    '        Bom_Component(n).setValue "DAuxUnitCode", ""  '辅助单位编码，String类型
    '        Bom_Component(n).setValue "DChangeRate", ""  '换算率，Double类型
    '        Bom_Component(n).setValue "DAuxBaseQtyN", ""  '辅助基本用量，Double类型
    '        Bom_Component(n).setValue "DBasEngineerFigNo", ""  '工程图号(导出用)，String类型
            Bom_Component(n).setValue "DDefine_28", rs!id & "" '
            Bom_Component(n).setValue "DDefine_37", ""  'BOM表体自定义项16，Date类型
'            Bom_Component(n).setValue "cBDefine14", ""  '屏体编码
'            Bom_Component(n).setValue "cBDefine15", ""  '屏体名称
'            Bom_Component(n).setValue "cBDefine24", ""  '项目编码
'            Bom_Component(n).setValue "cBDefine25", ""  '项目名称
        
    '        Bom_Component(n).setValue "DPartId", ""  '子件物料ID，Integer类型
            n = n + 1
            rs.MoveNext
        Wend
        
        '---- 有子件的时候才插入
        If n > 0 Then
            '第五步：调用API
            If u8apiBroker.InvokeApi() = False Then
                '第六步：错误处理
                strErr = u8apiBroker.GetLastError()
                If u8apiBroker.ErrorType = ExceptionType_Business Then
                    '处理API业务错误
                ElseIf u8apiBroker.ErrorType = ExceptionType_System Then
                    '处理系统错误
                End If
                CallU8Api4bom = False
            Else
                '第七步：获取返回结果
    
                '获取返回值
                '获取普通返回值。此返回值数据类型为Boolean，此参数按值传递，表示返回值: true:成功, false: 失败
                Dim result As Boolean
                result = CBool(u8apiBroker.GetReturnValue())
                CallU8Api4bom = result
                DBconn.Execute "update hzland_lzdetail set bomcreated='是' where _parentid='" & id & "'"
                DBconn.Execute "update hzland_materialsdetail set bomcreated='是' where _parentid='" & id & "'"
            End If
        Else
            CallU8Api4bom = True
        End If

        
    GoTo hFinish
ErrHandler:
    strErr = Err.Description
hFinish:
    '结束本次调用，释放API资源
    u8apiBroker.Disconnect
    Set u8apiBroker = Nothing
    If rs.State <> 0 Then rs.Close
    Set rs = Nothing
    If rsTmp.State <> 0 Then rsTmp.Close
    Set rsTmp = Nothing
End Function

'物料增加结构
Private Function Insertbas_partFree(cInvCode As String, citemcode As String, cabinetnumber As String, ConfigFree3 As Integer, ByRef sError As String) As Boolean
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    Dim partid As String
    Dim Free3Value As String
    On Error GoTo hErr
    strsql = "select * from bas_part where InvCode='" & cInvCode & "' and Free1='" & citemcode & "' and Free2='" & cabinetnumber & "'"
    If ConfigFree3 = 1 Then
        strsql = strsql & " AND Free3='空'"
        Free3Value = "'空'"
    Else
        Free3Value = "''"
    End If
    rs.Open strsql, DBconn
    If rs.BOF Or rs.EOF Then
        partid = GetVouchID("bas_part", "", 1, sError)
        strsql = "insert into bas_part(partid,InvCode,bVirtual,free1,free2,free3,mulqty,safeqty,fbasmaxsupply,minqty,fixqty,cbasengineerfigno,isurenesstype,idatetype,idatesum,ibestrowsum,ipercentumsum,bfreestop )" & _
                "values(" & partid & ",N'" & cInvCode & "',0,N'" & citemcode & "',N'" & cabinetnumber & "'," & Free3Value & ",NULL,NULL,NULL,NULL,NULL,NULL,N'1',NULL,NULL,NULL,NULL,N'0')"
        DBconn.Execute strsql
    End If
    Insertbas_partFree = True
    GoTo hFinish
hErr:
    sError = Err.Description
hFinish:
    If rs.State <> 0 Then
        rs.Close
    End If
    Set rs = Nothing
End Function

Private Function GetVouchID(strTableName As String, lngIDs As String, lngsTableCount As Long, errMsg As String) As String
    Dim AdoComm As ADODB.Command
    On Error GoTo DoErr
    Set AdoComm = New ADODB.Command
    With AdoComm
        .ActiveConnection = DBconn
        .CommandText = "sp_GetID"
        .CommandType = adCmdStoredProc
        .Prepared = False
        .Parameters.Append .CreateParameter("RemoteId", adVarChar, adParamInput, 3, "00")
        .Parameters.Append .CreateParameter("cAcc_Id", adVarChar, adParamInput, 3, m_login.cAcc_Id)
        .Parameters.Append .CreateParameter("VouchType", adVarChar, adParamInput, 50, strTableName)
        .Parameters.Append .CreateParameter("iAmount", adInteger, adParamInput, 8, lngsTableCount)
        .Parameters.Append .CreateParameter("MaxID", adBigInt, adParamOutput)
        .Parameters.Append .CreateParameter("MaxIDs", adBigInt, adParamOutput)
        .Execute
        GetVouchID = CStr(.Parameters("MaxID"))
        lngIDs = .Parameters("MaxIDs") - lngsTableCount + 1
    End With
    Set AdoComm = Nothing
    Exit Function
DoErr:
    errMsg = "获取单据ID发生错误：" & Err.Description
    Set AdoComm = Nothing
End Function


'写入存货自由项
Private Function InsertInventoryFreeValue(cInvCode As String, cFree As String, citemcode As String, cabinetnumber As String, ByRef sError As String) As Boolean
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    Dim bTrans As Boolean
    On Error GoTo hErr
    DBconn.BeginTrans
    bTrans = True
    If cFree = "cFree1" Then
        '自由项档案
        strsql = "select  *  from userdefine where cId=N'20' and cvalue='" & citemcode & "' "
        rs.Open strsql, DBconn
        If rs.BOF Or rs.EOF Then
            strsql = "insert into userdefine(cID,cValue)values('20','" & citemcode & "')"
            DBconn.Execute strsql
        End If
        '存货自由项取值范围
        strsql = "select * from AA_InvFreeContrapose where cinvcode='" & cInvCode & "' and cudid='20' and cudvalue='" & citemcode & "'"
        If rs.State <> 0 Then rs.Close
        rs.Open strsql, DBconn
        If rs.BOF Or rs.EOF Then
            strsql = "insert into AA_InvFreeContrapose(cInvFreeGUID,cInvCode,cUDID,cUDValue)values(newid(),'" & cInvCode & "','20','" & citemcode & "')"
            DBconn.Execute strsql
        End If
    ElseIf cFree = "cFree2" Then
        '自由项档案
        strsql = "select  *  from userdefine where cId=N'21' and cvalue='" & cabinetnumber & "' "
        rs.Open strsql, DBconn
        If rs.BOF Or rs.EOF Then
            strsql = "insert into userdefine(cID,cValue)values('21','" & cabinetnumber & "')"
            DBconn.Execute strsql
        End If
        '存货自由项取值范围
        strsql = "select * from AA_InvFreeContrapose where cinvcode='" & cInvCode & "' and cudid='21' and cudvalue='" & cabinetnumber & "'"
        If rs.State <> 0 Then rs.Close
        rs.Open strsql, DBconn
        If rs.BOF Or rs.EOF Then
            strsql = "insert into AA_InvFreeContrapose(cInvFreeGUID,cInvCode,cUDID,cUDValue)values(newid(),'" & cInvCode & "','21','" & cabinetnumber & "')"
            DBconn.Execute strsql
        End If
    End If
    If bTrans Then DBconn.CommitTrans: bTrans = False
    InsertInventoryFreeValue = True
    GoTo hFinish
hErr:
    If bTrans Then DBconn.RollbackTrans: bTrans = False
    sError = Err.Description
hFinish:
    If rs.State <> 0 Then
        rs.Close
    End If
    Set rs = Nothing
End Function

'获取存货结构性自由项属性
Private Function GetInventoryConfigFree(cInvCode As String, ConfigFree As String) As Integer
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    On Error GoTo hErr
    strsql = "select " & ConfigFree & " as ConfigFree from inventory where cinvcode='" & cInvCode & "'"
    rs.Open strsql, DBconn
    If Not rs.EOF And Not rs.BOF Then
        GetInventoryConfigFree = IIf(rs!ConfigFree, 1, 0)
    Else
        GetInventoryConfigFree = -1
    End If
    GoTo hFinish
hErr:

hFinish:
    If rs.State <> 0 Then
        rs.Close
    End If
    Set rs = Nothing
End Function

'Voucher   当前单据对象
'ButtonKey 按扭关键字
'strUserErr  错误信息
'bSuc  =true 表示成功并终止系统流程   =false表示失败并继续系统流程
Public Sub After_ButtonClick(Voucher As Object, ButtonKey As String, strUserErr As String, bSuc As Boolean)
    bSuc = False
    Select Case ButtonKey
        Case "refervouch"
            If StrCardNum = "PD010401" Then
                Vouchers_Frm.Voucher.getVoucherDataXML DomH, DomB
                AddNew_Voucher "PD010301", "PD010401", DomH, DomB, strUserErr, bSuc
            ElseIf StrCardNum = "PD010501" Then
                Vouchers_Frm.Voucher.getVoucherDataXML DomH, DomB
                AddNew_Voucher "PD010301", "PD010501", DomH, DomB, strUserErr, bSuc
            End If
        Case "calc"
            bSuc = calc(Voucher, strUserErr)
            If Not bSuc Then
                MsgBox strUserErr, vbInformation, "提示"
            End If
        Case "modify"
            If StrCardNum = "EP0205" Then
                Voucher.ItemState("cpmrpcode", 0).bCanModify = False
            End If
        Case Else
'            If StrCardNum = "EP0205" Then
'                Voucher.ItemState("cpmrpcode", 0).bCanModify = True
'            End If
    End Select
    
End Sub

'关闭、打开
Private Function DoClose(bClose As Boolean, VchLst As Object) As Boolean
    Dim i As Long
    Dim strsql As String
    Dim id As String
    On Error GoTo hErr
    For i = 1 To VchLst.rows - 1
        If VchLst.TextMatrix(i, VchLst.GridColIndex("selcol")) = "Y" Then
            id = VchLst.TextMatrix(i, VchLst.GridColIndex("id"))
            If bClose Then
                strsql = "update hzland_ChangedTask set status='关闭' where id='" & id & "' and isnull(status,'')=''"
            Else
                strsql = "update hzland_ChangedTask set status='' where id='" & id & "' and status='关闭'"
            End If
            DBconn.Execute strsql
        End If
    Next
    DoClose = True
    Exit Function
hErr:
    MsgBox Err.Description, vbCritical, "提示"
End Function

'采购计划展算
Private Function calc(Voucher As Object, ByRef strUserErr As String) As Boolean
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    Dim i As Long
    Dim cpmrpcode As String
    Dim id As String
    Dim cInvCode As String
    Dim fSupplyMulti As Double
    Dim dnverifytime As String
    On Error GoTo hErr
    cpmrpcode = Voucher.headerText("cpmrpcode")
    id = Voucher.headerText("id")
    If cpmrpcode = "" Then
        strUserErr = "请先录入项目需求计划单号。"
        GoTo hErr
    End If
    
    strsql = "select 1 from EF_ProcurementPlan where cpmrpcode='" & cpmrpcode & "'"
    If id <> "" Then
        strsql = strsql & " and id<>'" & id & "'"
    End If
    rs.Open strsql, DBconn
    If Not rs.EOF And Not rs.BOF Then
        strUserErr = "项目需求计划单已展算。"
        GoTo hErr
    End If
    '清空表体
    For i = Voucher.rows To 1 Step -1
        Voucher.DelLine i
    Next
    If rs.State <> 0 Then
        rs.Close
    End If
    strsql = "select p.*,i.iSafeNum,i.fSupplyMulti,i.fMinSupply,N.cPersonCode,N.cPersonName " & _
        " from V_List_EF_ProjectMRP p left outer join inventory i on p.b_cinvcode=i.cInvCode " & _
        " left outer join Person N on i.cPurPersonCode=N.cPersonCode" & _
        " where p.ccode='" & cpmrpcode & "' and ISNULL(p.cverifier,'')<>'' and ISNULL(p.ccloser,'')='' and ISNULL(p.cbcloser,'')='' "
    rs.Open strsql, DBconn
    i = 1
    While Not rs.EOF
        dnverifytime = rs!dnverifytime & ""
        Voucher.bodyText(i, "cpmrpautoid") = rs!autoid & ""
        cInvCode = rs!b_cinvcode & ""
        Voucher.bodyText(i, "b_cinvcode") = cInvCode
        Voucher.bodyText(i, "b_cinvname") = rs!b_cinvname & ""
        Voucher.bodyText(i, "b_cinvstd") = rs!b_cinvstd & ""
        Voucher.bodyText(i, "b_ccomunitname") = rs!b_ccomunitname & ""
        Voucher.bodyText(i, "coutsourced") = rs!coutsourced & ""
        Voucher.bodyText(i, "cperform") = rs!cperform & ""
        Voucher.bodyText(i, "b_cpersoncode") = rs!cPersonCode & ""
        Voucher.bodyText(i, "b_cpersonname") = rs!cPersonName & ""
        Voucher.bodyText(i, "cbmemo") = rs!cbmemo & ""
        Dim m As Integer
        For m = 1 To 16
            Voucher.bodyText(i, "cInvDefine" & m) = rs.fields("cInvDefine" & m).Value & ""
        Next
        Voucher.bodyText(i, "imrpqty") = rs!iQty & ""
        Voucher.bodyText(i, "imrpqtyl") = Getimrpqtyl(cInvCode, dnverifytime)
        Voucher.bodyText(i, "iinqty") = GetiInQty(cInvCode, id)
        Voucher.bodyText(i, "istock") = GetiStock(cInvCode)
        '可用量=现存量+预计入库量-预计出库量
        Voucher.bodyText(i, "ikyl") = str2Dbl(Voucher.bodyText(i, "istock")) + str2Dbl(Voucher.bodyText(i, "iinqty"))
        Voucher.bodyText(i, "isafenum") = str2Dbl(rs!iSafeNum & "")
        '计划请购数量=累计(项目需求计划量)-可用量+安全库存
        Voucher.bodyText(i, "ijhqty") = str2Dbl(Voucher.bodyText(i, "imrpqtyl")) - str2Dbl(Voucher.bodyText(i, "ikyl")) + str2Dbl(Voucher.bodyText(i, "isafenum"))
        If str2Dbl(Voucher.bodyText(i, "ijhqty")) < 0 Then
            Voucher.bodyText(i, "ijhqty") = 0
        End If
        '建议采购量=向上取整(采购计划量/采购经济批量)* 采购经济批量
        fSupplyMulti = str2Dbl(rs!fSupplyMulti & "")
        Voucher.bodyText(i, "iminqty") = fSupplyMulti
        If fSupplyMulti > 0 Then
            Voucher.bodyText(i, "ijyqty") = UpLng(str2Dbl(Voucher.bodyText(i, "ijhqty")) / fSupplyMulti) * fSupplyMulti
        Else
            Voucher.bodyText(i, "ijyqty") = str2Dbl(Voucher.bodyText(i, "ijhqty"))
        End If
        Voucher.bodyText(i, "isjqty") = Voucher.bodyText(i, "ijyqty")
        i = i + 1
        Voucher.AddLine
        rs.MoveNext
    Wend
    Voucher.RemoveEmptyRow
    calc = True
hErr:
    If rs.State <> 0 Then
        rs.Close
    End If
    Set rs = Nothing
    strUserErr = strUserErr & Err.Description
End Function


'现存量
Private Function GetiStock(cInvCode As String) As Double
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    On Error GoTo hErr
    strsql = "select SUM(iQuantity) as iQuantity  from CurrentStock where cInvCode='" & cInvCode & "'"
    rs.Open strsql, DBconn
    If Not rs.BOF And Not rs.EOF Then
        GetiStock = str2Dbl(rs!iQuantity & "")
    End If
hErr:
    If rs.State <> 0 Then
        rs.Close
    End If
    Set rs = Nothing
End Function


'预计入库量=累计(已计划未请购)+累计(已请购未采购)+ 累计(已采购未入库)
Private Function GetiInQty(cInvCode As String, id As String) As Double
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    On Error GoTo hErr
    strsql = "select SUM(iqty) as iqty from (select sum(iwqgqty) as iqty from V_List_EF_ProcurementPlan where b_cinvcode='" & cInvCode & "' and id<>'" & id & "'  union all " & _
            " select (b.fQuantity-b.iReceivedQTY) as iqty   from PU_AppVouch m left outer join PU_AppVouchs b on m.ID=b.ID " & _
            " where isnull(m.cCloser,'')='' " & _
            " and b.cInvCode='" & cInvCode & "' and (b.fQuantity-b.iReceivedQTY)>0 union all" & _
            " select b.iQuantity-isnull(b.fPoArrQuantity,0)+isnull(b.fPoRetQuantity,0) as iqty   from PO_Pomain m left outer join PO_Podetails  b on m.POID =b.POID" & _
            " where  isnull(m.cCloser,'')='' and b.iQuantity-isnull(b.fPoArrQuantity,0)+isnull(b.fPoRetQuantity,0)>0 and b.cInvCode='" & cInvCode & "' union all" & _
            " select iQuantity-fValidInQuan  from PU_ArrivalVouch p left outer join PU_ArrivalVouchs s on p.ID=s.ID where cInvCode='" & cInvCode & "' union all " & _
            "select b.iQuantity from RdRecord01 t left outer join RdRecords01 b on t.ID=b.ID where cInvCode='" & cInvCode & "' and isnull(t.cHandler,'')='' ) t"

    rs.Open strsql, DBconn
    If Not rs.EOF And Not rs.BOF Then
        GetiInQty = str2Dbl(rs!iQty & "")
    End If
    Exit Function
hErr:
    If rs.State <> 0 Then
        rs.Close
    End If
    Set rs = Nothing
End Function

'累计需求计划量
Private Function Getimrpqtyl(cInvCode As String, dnverifytime As String) As Double
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    On Error GoTo hErr
    strsql = "select sum(iqty) as iqty from V_List_EF_ProjectMRP where b_cinvcode='" & cInvCode & "' and ISNULL(cverifier,'')<>'' and ISNULL(ccloser,'')='' and ISNULL(cbcloser,'')='' and dnverifytime<='" & dnverifytime & "'"
    rs.Open strsql, DBconn
    If Not rs.BOF And Not rs.EOF Then
        Getimrpqtyl = str2Dbl(rs!iQty & "")
    End If
hErr:
    If rs.State <> 0 Then
        rs.Close
    End If
    Set rs = Nothing
    
End Function

'Voucher    当前单据对象
'sKey       当前操作关键字
'sRet       返回值
'strUserErr  错误信息
'bSuc  =true 表示成功并终止系统流程   =false表示失败并继续系统流程
Public Sub Voucher_headBrowUser(ByVal Voucher As Object, ByVal Skey As String, sRet As Variant, strUserErr As String, bSuc As Boolean)

    bSuc = False
End Sub


'初始化用户插件
'Voucher    当前单据对象
'sKey       当前操作关键字
'sRet       返回值
'strUserErr  错误信息
'bChanged
'bSuc  =true 表示成功并终止系统流程   =false表示失败并继续系统流程
Public Sub Voucher_headCellCheck(ByVal Voucher As Object, Skey As String, RetValue As String, bChanged As Long, strUserErr As String, bSuc As Boolean)
    Dim strsql As String
    Dim rs As ADODB.Recordset
    Dim i As Long
    On Error GoTo hErr
    Select Case StrCardNum
        Case "EP0207"
            Select Case Skey
                Case "citemcode"
                    '清空表体
                    For i = Voucher.rows To 1 Step -1
                        Voucher.DelLine i
                    Next
                    strsql = "select SUM(p.iSum) as isum,i.cInvDefine2 from PurBillVouchs p " & _
                            " left outer join inventory i on p.cInvCode=i.cInvCode " & _
                            " where p.citemcode='" & RetValue & "' group by i.cInvDefine2 "
                    Set rs = New ADODB.Recordset
                    rs.Open strsql, DBconn
                    i = 1
                    While Not rs.EOF
                        Voucher.bodyText(i, "cmaterialclass") = rs!cInvDefine2 & ""
                        Voucher.bodyText(i, "icb") = str2Dbl(rs!isum & "")
                        Voucher.AddLine
                        rs.MoveNext
                        i = i + 1
                    Wend
                    Voucher.RemoveEmptyRow
            End Select
        
    End Select
    
    bSuc = False
    GoTo Finish
hErr:
    strUserErr = Err.Description
    bSuc = True
Finish:
    If Not rs Is Nothing Then
        If rs.State <> 0 Then
            rs.Close
        End If
        Set rs = Nothing
    End If
End Sub

 
'Voucher    当前单据对象
'sKey       当前操作关键字
'R          当前行数
'sRet       返回值
'strUserErr  错误信息
'bSuc  =true 表示成功并终止系统流程   =false表示失败并继续系统流程
Public Sub Voucher_bodyBrowUser(ByVal Voucher As Object, Skey As String, ByVal R As Long, sRet As Variant, strUserErr As String, bSuc As Boolean)
    bSuc = False

'    Set Vouchers = Voucher


'        MsgBox " Voucher_bodyBrowUser 插件配置成功！"

 
End Sub


'初始化用户插件
'Voucher    当前单据对象
'sRet       返回值
'sKey       当前操作关键字
'R          当前行数
'strUserErr  错误信息
'bChanged
'bSuc  =true 表示成功并终止系统流程   =false表示失败并继续系统流程
Public Sub Voucher_bodyCellCheck(ByVal Voucher As Object, RetValue As Variant, bChanged As Long, Skey As String, ByVal R As Long, strUserErr As String, bSuc As Boolean)
    bSuc = False
    Select Case Skey
        Case "iunitqty"
            If RetValue = "" Then
                Voucher.bodyText(R, "iqty") = ""
            Else
                Voucher.bodyText(R, "iqty") = str2Dbl(Voucher.headerText("ipqty")) * str2Dbl(RetValue)
            End If
    End Select
    

End Sub

'初始化用户插件
'Voucher        当前单据对象
'sKey           当前操作关键字
'pCom           下拉控件
'strUserErr     错误信息
'bChanged
'bSuc  =true 表示成功并终止系统流程   =false表示失败并继续系统流程
Public Sub Voucher_FillHeadComboBox(ByVal Voucher As Object, Skey As String, pCom As Object, strUserErr As String, bSuc As Boolean)
    bSuc = False
    
    
    Set Vouchers = Voucher
'    Dim i As Integer
'    Dim rds As New ADODB.Recordset
'    Select Case LCase(Me.Voucher.ItemState(Index, siheader).sFieldName)
'
'        Case "iyear" '编制年度
'                pCom.Clear
'                For i = 0 To 4
'                  pCom.AddItem CStr(Val(m_login.cIYear) + i)
'                Next

'    End Select
'        MsgBox " Voucher_FillHeadComboBox 插件配置成功！"




End Sub
 
 '初始化用户插件
'Voucher        当前单据对象
'sKey           当前操作关键字
'R          当前行数
'pCom           下拉控件
'strUserErr     错误信息
'bChanged
'bSuc  =true 表示成功并终止系统流程   =false表示失败并继续系统流程
Public Sub Voucher_FillList(ByVal Voucher As Object, Skey As String, ByVal R As Long, pCom As Object, strUserErr As String, bSuc As Boolean)
    bSuc = False
    
    Set Vouchers = Voucher
    'Dim sFieldName As String
    'sFieldName = LCase(Me.Voucher.ItemState(C, sibody).sFieldName)
    'Select Case sFieldName
    '    Case "adds", "lenssen" '控制方向
    '        pCom.Clear
    '        pCom.AddItem ""
    '        pCom.AddItem "借方"
    '        pCom.AddItem "贷方"
 
    '


'        MsgBox " Voucher_FillList 插件配置成功！"
    
End Sub



''
'Source_Cardnumber    来源单句
'Dest_Cardnumber      目标单据
'Dest_domHead         目标单据表头数据
'Dest_domBody         目标单据表体数据
'strUserErr           错误信息
'bSuc               =true 表示成功    =false 表示失败
Public Sub AddNew_Voucher(Source_Cardnumber As String, Dest_Cardnumber As String, Dest_domHead As DOMDocument, Dest_domBody As DOMDocument, strUserErr As String, bSuc As Boolean)
On Error GoTo ExitSub
    Dim FrmR As New frmRefernew ' frmRefer
    bSuc = False
    With FrmR
        Set FrmR.Dest_Voucher = Vouchers
        Call .init(m_login, DBconn, FrmR, Dest_domHead, Dest_domBody, Source_Cardnumber, Dest_Cardnumber)
        If .filter Then
            FrmR.SetData
            FrmR.Show 1
'            Set Dest_domHead = .DomHead_Dest.cloneNode(True)
'            Set Dest_domBody = .DomBody_Dest.cloneNode(True)
'            Vouchers_Frm.Voucher.setVoucherDataXML Dest_domHead, Dest_domBody
        End If
    End With
    bSuc = True
    Set FrmR = Nothing
    Exit Sub
ExitSub:
    bSuc = False
    strUserErr = strUserErr & Err.Description
End Sub


'个性化生单填充方法
'Source_Cardnumber  来源参照类型
'Dest_Cardnumber    目标单据的cardnumber
'Source_domHead     参照生单界面上半部分数据对象
'Source_domBody     参照生单界面上半部分数据对象
'Dest_domHead       目标单据 表头数据对象
'Dest_domHead       目标单据 表头数据对象
'strUserErr         错误信息
'bSuc               =true 表示个性化填充       =false 表示系统自动填中
Public Sub Full_Voucher(Voucher As Object, Source_Cardnumber As String, Dest_Cardnumber As String, Source_domHead As DOMDocument, Source_domBody As DOMDocument, Dest_domHead As DOMDocument, Dest_domBody As DOMDocument, strUserErr As String, Optional bSuc As Boolean = False)
    Dim list As IXMLDOMNodeList
    Dim ele As IXMLDOMElement
    Dim iRow As Integer
    On Error GoTo ExitSub
    bSuc = False
    '处理表头填充
'    Voucher.headerText("citemcode") = GetHeadItemValue(Source_domHead, "projectnumber")
'    Voucher.headerText("citemname") = GetHeadItemValue(Source_domHead, "projectname")
    
    
    '清除当前单据表体数据
    '处理表体填充
    Set list = Source_domHead.selectNodes("//z:row")
    Voucher.RemoveEmptyRow
    iRow = Voucher.BodyRows + 1
    For Each ele In list
        Voucher.AddLine Voucher.BodyRows + 1
        Voucher.SimulateInput SectionsConstants.sibody, Voucher.BodyRows, "b_cinvcode", GetNodeValue(ele, "stockcode"), 0
        If Voucher.bodyText(iRow, "b_cinvcode") <> "" Then
    '        Voucher.bodyText(iRow, "b_cinvcode") = GetNodeValue(ele, "stockcode")
    '        Voucher.bodyText(iRow, "b_cinvname") = GetNodeValue(ele, "stockname")
    '        Voucher.bodyText(iRow, "b_cinvstd") = GetNodeValue(ele, "specification")
    '        Voucher.bodyText(iRow, "b_ccomunitname") = GetNodeValue(ele, "unit")
            Voucher.bodyText(iRow, "iqty") = GetNodeValue(ele, "iwsendqty")
            Voucher.bodyText(iRow, "hzid") = GetNodeValue(ele, "id")
            Voucher.bodyText(iRow, "hzbilltype") = GetNodeValue(ele, "billtype")
            Voucher.bodyText(iRow, "b_citemcode") = GetNodeValue(ele, "projectnumber")
            Voucher.bodyText(iRow, "b_citemname") = GetNodeValue(ele, "projectname")
            Voucher.bodyText(iRow, "cabinetnumber") = GetNodeValue(ele, "cabinetnumber")
            Voucher.bodyText(iRow, "cabinetname") = GetNodeValue(ele, "cabinetname")
            iRow = iRow + 1
        End If
    Next
    bSuc = True
    Exit Sub
ExitSub:
    bSuc = False
    strUserErr = strUserErr & Err.Description
End Sub


 '提示：单据单击事件
Public Sub Voucher_Click(ByVal Voucher As Object, section As Variant, Skey As String, R As Long, strUserErr As String, bSuc As Boolean)
On Error GoTo Errhandle

'     Dim i As Integer
'     bSuc = False
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("Voucher_Click", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.Voucher_Click Voucher, section, Skey, R, strUserErr, bSuc
'            If bSuc = True Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
''     bSuc = True
''    bSuc = False
ExitSub:
     Exit Sub
Errhandle:
     bSuc = False
     strUserErr = Err.Description
End Sub


 '提示：单据单击事件
Public Sub Voucher_DblClick(ByVal Voucher As Object, section As Variant, Skey As String, R As Long, strUserErr As String, bSuc As Boolean)
On Error GoTo Errhandle

'     Dim i As Integer
'     bSuc = False
'     For i = 0 To colInterface.Count - 1
'         '判断用户DLL是否有这个方法
'         strTemp = colInterface.Keys(i)
'         If bHasFunction("Voucher_DblClick", strTemp) Then
'            Set clsInterface = colInterface.Item(strTemp)
'            clsInterface.Voucher_DblClick Voucher, section, Skey, R, strUserErr, bSuc
'            If bSuc = True Then
'                GoTo ExitSub
'            End If
'         End If
'     Next i
'     '注意：如果没有任何的用户自定义动态库，则需要返回TRUE
''     bSuc = True
''    bSuc = False
ExitSub:
     Exit Sub
Errhandle:
     bSuc = False
     strUserErr = Err.Description
End Sub


