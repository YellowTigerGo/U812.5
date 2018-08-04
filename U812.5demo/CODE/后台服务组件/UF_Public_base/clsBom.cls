VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private DBconn As ADODB.Connection
Private M_login As Object

Public Sub Init(db As ADODB.Connection, g_Login As Object)
    Set DBconn = db
    Set M_login = g_Login
End Sub

'设计清单变更处理
Public Sub DoUpdate(VchLst As Object)
    Dim id As String
    Dim strsql As String
    Dim sError As String
    Dim sMsg As String
    Dim rs As New ADODB.Recordset
'    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    Dim actionType As String
    Dim cInvCode As String
    Dim cInvName As String
    Dim specification As String
    Dim citemcode As String
    Dim citemname As String
    Dim cabinetnumber As String
    Dim cabinetname As String
    Dim cabinetid As String
    Dim parentid As String
    Dim pid As String
    Dim iQty As Double
    Dim stockqty As Double
    On Error GoTo hErr
    frmProgress.Show 0
    frmProgress.Msg = "BOM生成..."
    
    For i = 1 To VchLst.Rows - 1
        DoEvents
        frmProgress.Msg = "正在处理第" & i & "行数据..."
        If VchLst.TextMatrix(i, VchLst.GridColIndex("selcol")) = "Y" Then
            id = VchLst.TextMatrix(i, VchLst.GridColIndex("id"))
            actionType = VchLst.TextMatrix(i, VchLst.GridColIndex("actiontype"))
            cInvCode = VchLst.TextMatrix(i, VchLst.GridColIndex("stockcode"))
            cInvName = VchLst.TextMatrix(i, VchLst.GridColIndex("stockname"))
            specification = VchLst.TextMatrix(i, VchLst.GridColIndex("specification"))
            citemcode = VchLst.TextMatrix(i, VchLst.GridColIndex("projectnumber"))
            citemname = VchLst.TextMatrix(i, VchLst.GridColIndex("projectname"))
            cabinetnumber = VchLst.TextMatrix(i, VchLst.GridColIndex("cabinetnumber"))
            cabinetid = VchLst.TextMatrix(i, VchLst.GridColIndex("cabinetid"))
            iQty = CDbl(VchLst.TextMatrix(i, VchLst.GridColIndex("num")))
            Select Case actionType
                Case "减少屏体产量"
                    sError = "减少屏体产量，请手工处理。"
                    GoTo hFinish
                Case "删除材料"
                    '判断是否已领料
                    strsql = "select ccode from kcmaterialapplist  where cDefine28='" & id & "'"
                    rs.open strsql, DBconn
                    If Not rs.EOF And Not rs.BOF Then
                        sError = "已存在材料出库单" & rs!ccode & "，材料不能删除。"
                        GoTo hFinish
                    End If
                    '删除生产订单子件
                    strsql = "select m.mocode,d.SortSeq,c.AllocateId from mom_orderdetail d" & _
                            " left outer join mom_order m on m.MoId=d.moid left outer join mom_moallocate c on d.modid=c.modid" & _
                            " where c.Define28 ='" & id & "'"
                    If rs.State <> 0 Then rs.Close
                    rs.open strsql, DBconn
                    While Not rs.EOF
                        strsql = "delete mom_moallocate_extradefine where AllocateId='" & rs!AllocateId & "'"
                        DBconn.Execute strsql
                        strsql = "delete mom_moallocate where AllocateId='" & rs!AllocateId & "'"
                        DBconn.Execute strsql
                        '消息：m.mocode,d.SortSeq
                    Wend
                    '删除BOM子件
                    strsql = "select b.InvCode,p.OpComponentId from bom_parent m left outer join bom_opcomponent p on m.BomId=p.BomId" & _
                            " left outer join bas_part b on m.ParentId=b.PartId" & _
                            " where p.Define28='" & id & "'"
                    If rs.State <> 0 Then rs.Close
                    rs.open strsql, DBconn
                    While Not rs.EOF
                        strsql = "delete bom_opcomponent where OpComponentId='" & rs!OpComponentId & "'"
                        DBconn.Execute strsql
                        '消息：b.InvCode
                    Wend
                    '删除设计清单材料
                    strsql = "delete hzland_materialsdetail where id='" & id & "'"
                    DBconn.Execute strsql
                    strsql = "delete hzland_lzdetail where id='" & id & "'"
                    DBconn.Execute strsql
                Case "追加材料"
                    '判断是否有关闭的生产订单
                    strsql = "select m.mocode,d.SortSeq,c.AllocateId,d.CloseUser from mom_orderdetail d" & _
                            " left outer join mom_order m on m.MoId=d.moid left outer join mom_moallocate c on d.modid=c.modid" & _
                            " where c.Define28 ='" & id & "' "
                    If rs.State <> 0 Then rs.Close
                    rs.open strsql, DBconn
                    While Not rs.EOF
                        If rs!CloseUser & "" <> "" Then
                            sError = "生产订单" & rs!MoCode & "行号" & rs!SortSeq & "已经关闭，不能追加材料。"
                            GoTo hFinish
                        End If
                        rs.MoveNext
                    Wend
                    
                    parentid = VchLst.TextMatrix(i, VchLst.GridColIndex("parentid"))
                    strsql = "select * from EF_V_HZDesignList where _parentid='" & parentid & "'"
                    If rs.State <> 0 Then rs.Close
                    rs.open strsql, DBconn
                    If Not rs.EOF And Not rs.BOF Then
                        pid = rs!pid
                        stockcode = rs!stockcode & ""
                        '设计清单追加材料
                        strsql = "insert into hzland_materialsdetail(id,pid,_parentid,cabinetname,cabinetnumber,stockname,stockcode,specification,num,cabinetid)values(" & _
                                "'" & id & "','" & pid & "','" & parentid & "','" & cabinetname & "','" & cabinetnumber & "','" & cInvName & "','" & cInvCode & "'," & _
                                "'" & specification & "','" & iQty & "','" & cabinetid & "')"
                        DBconn.Execute strsql
                        'BOM追加材料
                        If Not AddBomComponent(parentid, stockcode, citemcode, cabinetnumber, 1, iQty, sError) Then
                            MsgBox "第" & i & "行物料BOM子件追加失败:" & sError, vbCritical, "提示"
                            GoTo hFinish
                        End If
                        '生产订单追加子件
                        If Not AddMom_moallocate(parentid, stockcode, citemcode, citemname, cabinetnumber, cabinetname, citemcode, cabinetnumber, 1, iQty, sError) Then
                            MsgBox "第" & i & "行物料生产订单子件材料追加失败:" & sError, vbCritical, "提示"
                            GoTo hFinish
                        End If
                    Else
                        sError = "在设计清单里未找到对应的屏体信息"
                        GoTo hFinish
                    End If
                Case "减少材料用量"
                    stockqty = CDbl(VchLst.TextMatrix(i, VchLst.GridColIndex("stockqty")))
                    strsql = "select m.mocode,d.SortSeq,c.AllocateId,d.CloseUser,c.InvCode,c.Qty,c.RequisitionQty ,c.IssQty ,c.RequisitionFlag from mom_orderdetail d" & _
                            " left outer join mom_order m on m.MoId=d.moid left outer join mom_moallocate c on d.modid=c.modid" & _
                            " where c.Define28 ='" & id & "' "
                    If rs.State <> 0 Then rs.Close
                    rs.open strsql, DBconn
                    While Not rs.EOF
                        If rs!CloseUser & "" <> "" Then
                            sError = "生产订单" & rs!MoCode & "行号" & rs!SortSeq & "已经关闭，不能减少材料用量。"
                            GoTo hFinish
                        End If
                        If rs!RequisitionFlag = 1 Then '领料申请
                            If stockqty + CDbl(rs!Qty - rs!RequisitionQty) <= 0 Then
                                sError = "生产订单" & rs!MoCode & "行号" & rs!SortSeq & "材料" & rs!InvCode & "剩余申请量为" & CDbl(rs!Qty - rs!RequisitionQty) & ",材料用量减少" & stockqty & "，操作失败。"
                                GoTo hFinish
                            End If
                        Else
                            If stockqty + CDbl(rs!Qty - rs!IssQty) <= 0 Then
                                sError = "生产订单" & rs!MoCode & "行号" & rs!SortSeq & "材料" & rs!InvCode & "未领量为" & CDbl(rs!Qty - rs!RequisitionQty) & ",材料用量减少" & stockqty & "，操作失败。"
                                GoTo hFinish
                            End If
                        End If
                        strsql = "update mom_moallocate set Qty=Qty+" & stockqty & " where AllocateId='" & rs!AllocateId & "'"
                        DBconn.Execute strsql
                        strsql = "update hzland_materialsdetail set num=num+" & stockqty & " where id='" & id & "'"
                        DBconn.Execute strsql
                        '消息：m.mocode,d.SortSeq,c.InvCode
                        rs.MoveNext
                    Wend
            End Select
        End If
    Next
    MsgBox "设计清单变更处理完毕.", vbInformation, "提示"
    GoTo hFinish
hErr:
    MsgBox "设计清单变更处理异常：" & i & Err.Description, vbCritical, "提示"
hFinish:
    If rs.State <> 0 Then
        rs.Close
    End If
    Set rs = Nothing
    frmProgress.Hide
    If Len(sError) > 0 Then
        MsgBox sError, vbInformation, "提示"
    End If
End Sub

'是否有生产订单
Private Function HasMomOrder() As Boolean

End Function

'设计清单生成BOM,追加BOM子件，追加生产订单子件
Public Sub DoCreateBom(VchLst As Object)
    Dim id As String
    Dim strsql As String
    Dim sError As String
    Dim i As Long
    Dim cInvCode As String
    Dim citemcode As String
    Dim citemname As String
    Dim cabinetnumber As String
    Dim cabinetname As String
    Dim cabinetid As String
    Dim ConfigFree As Integer
    Dim ConfigFree3 As Integer
    Dim billtype As String
    Dim cFree1 As String
    Dim cFree2 As String
    Dim iQty As Double
'    Dim bTran As Boolean
    On Error GoTo hErr
    
    frmProgress.Show 0
    frmProgress.Msg = "BOM生成..."
    
    For i = 1 To VchLst.Rows - 1
        DoEvents
        frmProgress.Msg = "正在处理第" & i & "行数据..."
        If VchLst.TextMatrix(i, VchLst.GridColIndex("selcol")) = "Y" Then 'And VchLst.TextMatrix(i, VchLst.GridColIndex("bomcreated")) = "否"
            If VchLst.TextMatrix(i, VchLst.GridColIndex("status")) = "关闭" Then
                MsgBox "第" & i & "行已关闭，不能生成BOM。", vbCritical, "提示"
                GoTo hFinish
            End If
            id = VchLst.TextMatrix(i, VchLst.GridColIndex("id"))
            cInvCode = VchLst.TextMatrix(i, VchLst.GridColIndex("stockcode"))
            If Left(cInvCode, 2) <> "01" Then
                MsgBox "第" & i & "行存货编码为" & cInvCode & "，不是01开头的存货，不能生成BOM。", vbCritical, "提示"
                GoTo hFinish
            End If
            If GetChildCount(id) <= 0 Then
                MsgBox "第" & i & "行存货编码为" & cInvCode & "，没有子件信息。", vbCritical, "提示"
                GoTo hFinish
            End If
            citemcode = VchLst.TextMatrix(i, VchLst.GridColIndex("projectnumber"))
            citemname = VchLst.TextMatrix(i, VchLst.GridColIndex("projectname"))
            cabinetnumber = VchLst.TextMatrix(i, VchLst.GridColIndex("cabinetnumber"))
            cabinetname = VchLst.TextMatrix(i, VchLst.GridColIndex("cabinetname"))
            cabinetid = VchLst.TextMatrix(i, VchLst.GridColIndex("cabinetid"))
            ConfigFree = GetInventoryConfigFree(cInvCode, "bConfigFree1")
            iQty = str2Dbl(VchLst.TextMatrix(i, VchLst.GridColIndex("num")))
            If iQty <= 0 Then
                MsgBox "第" & i & "数量不能为零或负数。", vbCritical, "提示"
                GoTo hFinish
            End If
'            DBconn.BeginTrans
'            bTran = True
            '存货自由项处理
            If ConfigFree = -1 Then
                MsgBox "第" & i & "存货编码" & cInvCode & "在U8系统中不存在。", vbCritical, "提示"
                GoTo hFinish
            ElseIf ConfigFree = 1 Then
                If citemcode = "" Then
                    MsgBox "第" & i & "存货启用了结构性自由项1，项目编码不能为空", vbCritical, "提示"
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
                MsgBox "第" & i & "存货编码" & cInvCode & "在U8系统中不存在。", vbCritical, "提示"
                    GoTo hFinish
            ElseIf ConfigFree = 1 Then
                If cabinetnumber = "" Then
                    MsgBox "第" & i & "存货启用了结构性自由项2，屏体编码不能为空", vbCritical, "提示"
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
                '追加BOM子件
                If Not AddBomComponent(id, cInvCode, cFree1, cFree2, ConfigFree3, iQty, sError) Then
                    MsgBox "第" & i & "行物料BOM子件追加失败:" & sError, vbCritical, "提示"
                    GoTo hFinish
                End If
                '追加生产订单子件
                If Not AddMom_moallocate(id, cInvCode, citemcode, citemname, cabinetnumber, cabinetname, cFree1, cFree2, ConfigFree3, iQty, sError) Then
                    MsgBox "第" & i & "行物料生产订单子件材料追加失败:" & sError, vbCritical, "提示"
                    GoTo hFinish
                End If
                If billtype = "冷作清单" Then
                    '同步累计生产订单量
                    DBconn.Execute "update hzland_lzdetail set imqty=s.imqty from hzland_materialsdetail s where hzland_lzdetail.cabinetid=s.cabinetid and hzland_lzdetail.stockcode=s.stockcode and hzland_lzdetail.id='" & id & "' and isnull(s._parentid,'')='' and s.bomcreated='是'"
                    'TODO:
                    DBconn.Execute "update m set m.imqty=s.imqty from hzland_lzdetail m, hzland_lzdetail s where m.cabinetid=s.cabinetid and m.stockcode=s.stockcode and m.id='" & id & "' and isnull(s._parentid,'')='' and s.bomcreated='是'"
                    DBconn.Execute "update hzland_lzdetail set bomcreated='是' where id='" & id & "'"
                ElseIf billtype = "材料清单" Then
                    '同步累计生产订单量
                    DBconn.Execute "update hzland_materialsdetail set imqty=s.imqty from hzland_lzdetail s where hzland_materialsdetail.cabinetid=s.cabinetid and hzland_materialsdetail.stockcode=s.stockcode and hzland_materialsdetail.id='" & id & "' and isnull(s._parentid,'')='' and s.bomcreated='是'"
                    'TODO:同步累计生产订单量,追加材料
                    DBconn.Execute "update m set m.imqty=s.imqty from hzland_materialsdetail m, hzland_materialsdetail s where m.cabinetid=s.cabinetid and m.stockcode=s.stockcode and m.id='" & id & "' and isnull(s._parentid,'')='' and s.bomcreated='是'"
                    DBconn.Execute "update hzland_materialsdetail set bomcreated='是' where id='" & id & "'"
                End If
            Else '新增BOM
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
'            If bTran Then
'                DBconn.CommitTrans
'                bTran = False
'            End If
        End If
    Next
    MsgBox "物料清单生成完毕.", vbInformation, "提示"
    GoTo hFinish
hErr:
    MsgBox "物料清单生成异常：" & i & Err.Description, vbCritical, "提示"
hFinish:
'    If bTran Then
'        DBconn.RollbackTrans
'    End If
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
    rs.open strsql, DBconn
    If Not rs.BOF And Not rs.EOF Then
        BomExist = True
    End If
    GoTo hFinish
hErr:
    MsgBox "查询bas_part异常：" & Err.Description, vbCritical, "提示"
hFinish:
    If rs.State <> 0 Then
        rs.Close
    End If
    Set rs = Nothing
End Function

'生产订单追加子件
Private Function AddMom_moallocate(ByVal id As String, ByVal cInvCode As String, citemcode As String, citemname As String, cabinetnumber As String, cabinetname As String, cFree1 As String, cFree2 As String, ConfigFree3 As Integer, iQty As Double, ByRef strErr As String) As Boolean
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
    
    On Error GoTo Errhandler
    strsql = "select bomid from v_bom_head where invcode='" & cInvCode & "' and InvFree_1='" & cFree1 & "' and InvFree_2='" & cFree2 & "'"
    If ConfigFree3 = 1 Then
        strsql = strsql & " and InvFree_3='空'"
    End If
    strsql = strsql & " order by CreateDate desc"
    rs.open strsql, DBconn
    If Not rs.EOF And Not rs.BOF Then
        BomId = rs!BomId
    End If
    strsql = "select d.MoDid,d.SortSeq,t.MoCode,d.Qty from mom_orderdetail d left outer join mom_order t on d.MoId =t.MoId where d.bomid='" & BomId & "'"
    If rs.State <> 0 Then rs.Close
    rs.open strsql, DBconn
    While Not rs.EOF
        MoCode = rs!MoCode
        MoDid = rs!MoDid & ""
        iMQty = str2Dbl(rs!Qty & "")
        mom_orderdetailSortSeq = rs!SortSeq
        mom_moallocateSortSeq = GetSortSeq("mom_moallocate", "MoDId", MoDid)
        strsql = "select * from mom_morder where modid='" & MoDid & "'"
        If rsTmp.State <> 0 Then rsTmp.Close
        rsTmp.open strsql, DBconn
        If Not rsTmp.EOF And Not rsTmp.BOF Then
            StartDate = rsTmp!StartDate
            DueDate = rsTmp!DueDate
        End If
        strsql = "select * from EF_V_HZDesignList where _parentid='" & id & "' and id in(select Define28 from bom_opcomponent where bomid='" & BomId & "' AND ISNULL(define28,'')<>'')" & _
                " and id not IN( select define28 from mom_moallocate where modid='" & MoDid & "' AND ISNULL(define28,'')<>'')"
                
        If rsTmp.State <> 0 Then rsTmp.Close
        rsTmp.open strsql, DBconn
        
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
            strsql = "insert into mom_moallocate_extradefine(allocateid,cbdefine24,cbdefine25,cbdefine14,cbdefine15)values(" & _
                AllocateIds & ",'" & citemcode & "','" & citemname & "','" & cabinetnumber & "','" & cabinetname & "')"
            DBconn.Execute strsql
            rsTmp.MoveNext
        Wend
        rs.MoveNext
    Wend
    
    AddMom_moallocate = True
    GoTo hFinish
Errhandler:
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
    Dim bTran As Boolean
On Error GoTo Errhandler
    DBconn.BeginTrans
    bTran = True
    strsql = "select bomid from v_bom_head where invcode='" & cInvCode & "' and InvFree_1='" & cFree1 & "' and InvFree_2='" & cFree2 & "'"
    If ConfigFree3 = 1 Then
        strsql = strsql & " and InvFree_3='空'"
    End If
    strsql = strsql & " order by CreateDate desc"
    rs.open strsql, DBconn
    If Not rs.EOF And Not rs.BOF Then
        BomId = rs!BomId
    End If
    cDepCode = getAccinformation("PD", "cDepCode", DBconn)
    SortSeq = GetSortSeq("bom_opcomponent", "BomId", BomId)
    '根据设计清单id过滤掉已生成BOM的记录
    strsql = "select * from EF_V_HZDesignList where _parentid='" & id & "' and id not in(select Define28 from bom_opcomponent where bomid='" & BomId & "' AND ISNULL(define28,'')<>'')"
    If rs.State <> 0 Then rs.Close
    rs.open strsql, DBconn
    
    While Not rs.EOF
        If rs!Num <= 0 Then
            strErr = "材料出现数量为零或负数的情况。"
            GoTo hFinish
        End If
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
        strsql = "insert into bom_opcomponent ([OpComponentId],[BomId],[SortSeq],[OpSeq],[ComponentId],[EffBegDate],[EffEndDate],[FVFlag],[BaseQtyN],[BaseQtyD],[CompScrap],[ByproductFlag],[OptionsId],[AuxUnitCode],[ChangeRate],[AuxBaseQtyN],[ProductType],[Define22],[Define23],[Define24],[Define25],[Define26],[Define27],[Define29],[Define30],[Define31],[Define32],[Define33],[Define34],[Define35],[Define36],[Define37],[RecursiveFlag],[Free1],[Free2],[Free3],[Free4],[Free5],[Free6],[Free7],[Free8],[Free9],[Free10],[bProcessProduct],[bProcessMaterial],[Define28],[Remark])" & _
                "values(" & OpComponentIds & "," & BomId & "," & SortSeq & ",'0000'," & ComponentId & ",'2000-01-01','2099-12-31',1,'" & rs!Num & "','" & iQty & "','0','0'," & optionsids & ",Null,Null,Null,0,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,'0','','','','','','','','','','','0','0','" & rs!id & "','" & Replace(rs!remark & "", """", "“") & "')"
        DBconn.Execute strsql
        
        If billtype = "冷作清单" Then
            DBconn.Execute "update hzland_lzdetail set bomcreated='是' where id='" & rs!id & "'"
        ElseIf billtype = "材料清单" Then
            DBconn.Execute "update hzland_materialsdetail set bomcreated='是' where id='" & rs!id & "'"
        End If
        rs.MoveNext
    Wend
    If bTran Then
        DBconn.CommitTrans
    End If
    bTran = False
    AddBomComponent = True
    GoTo hFinish
Errhandler:
    If bTran Then
        DBconn.RollbackTrans
    End If
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
    rs.open strsql, DBconn
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
    rs.open strsql, DBconn
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
    rs.open strsql, DBconn
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
    rs.open strsql, DBconn
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
On Error GoTo Errhandler
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
        Set u8EnvCtx.U8Login = M_login
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
        rs.open strsql, DBconn, adOpenStatic, adLockReadOnly
        
        
        If Not rs.EOF Then
            extbo(0).setValue "InvCode", cInvCode ' "001"  '母件编码，String类型
'            extbo(0).setValue "InvName", rs.fields("cinvname").Value & ""  '母件名称(导出用)，String类型
'            extbo(0).setValue "InvStd", rs.fields("cinvstd").Value & ""  '规格型号(导出用)，String类型
            extbo(0).setValue "InvUnitName", rs.fields("cComUnitName").Value & ""   '计量单位(导出用)，String类型
            extbo(0).setValue "InvUnit", rs.fields("cComUnitCode").Value & ""   '计量单位编码(导出用)，String类型
            extbo(0).setValue "InvGroupType", rs.fields("iGroupType").Value & ""   '存货计量单位组类型(导出用)，Integer类型
            extbo(0).setValue "InvGroupCode", rs.fields("cGroupCode").Value & ""   '存货量单位组编码(导出用)，String类型
            extbo(0).setValue "InvGroupName", rs.fields("cGroupName").Value & ""   '存货计量单位名称(导出用)，String类型
        Else
            strErr = "存货" & cInvCode & "在U8系统存货档案中不存在."
            Exit Function
        End If
        
        extbo(0).setValue "CreateUser", M_login.cUserName  ' "demo"  '建档人(导出用)，String类型
        extbo(0).setValue "ModifyUser", ""  '修改人(导出用)，String类型
        extbo(0).setValue "CreateDate", M_login.CurDate   '建档日期(导出用)，Date类型
        extbo(0).setValue "ModifyDate", ""  '修改日期(导出用)，Date类型
        extbo(0).setValue "ModifyTime", ""  '修改时间(导出用)，Date类型
        extbo(0).setValue "CreateTime", M_login.CurDate & " " & Time   '创建时间(导出用)，Date类型
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
'        extbo(0).setValue "Define_1", ""  'BOM自定义项1，String类型
        
'        extbo(0).setValue "Define_14", ""   'BOM自定义项14，String类型
        
'        extbo(0).setValue "Define_16", ""  'BOM自定义项16，Double类型
'        extbo(0).setValue "InvDefine_1", "" '本次倒入唯一码  用于失败后删除已导入的数据  '母件自定义项1，String类型
'        extbo(0).setValue "InvDefine_16", ""  '母件自定义项16，Date类型
        extbo(0).setValue "BomState", "1"  '状态(导出用)，Integer类型
        extbo(0).setValue "VersionEndDate", "2099-12-31"  '版本失效日期(导出用)，Date类型   实在不行就在导入完毕后更新下实效日期

        Dim DDeptCode As String
        DDeptCode = getAccinformation("PD", "cDepCode", DBconn)
        '----处理子件
        strsql = "select * from EF_V_HZDesignList where _parentid='" & id & "'"
        If rs.State <> 0 Then rs.Close
        rs.open strsql, DBconn
        While Not rs.EOF
            Set Bom_Component = extbo(0).GetSubEntity("Bom_Component")
            '----------------------------------- 必输字段 --------------------------------
            Bom_Component(n).setValue "DSortSeq", (n + 1) * 10 '子件行号(默认系统自增)，Integer类型
            Bom_Component(n).setValue "DOpSeq", "0000" 'Format((n + 1) * 10, "0000") '工序行号(默认0000)，String类型
            
            strsql = "select b.cComUnitName,c.cGroupName ,c.iGroupType ,a.* from inventory a left join ComputationUnit b on a.cComUnitCode =b.cComunitCode" _
                & " left join ComputationGroup c on a.cGroupCode =c.cGroupCode  where cinvcode ='" & rs!stockcode & "'"
            If rsTmp.State <> 0 Then rsTmp.Close
            rsTmp.open strsql, DBconn, adOpenStatic, adLockReadOnly
            
            
            If Not rsTmp.EOF Then
                Bom_Component(n).setValue "DInvCode", rsTmp.fields("cinvcode").Value & ""   '子件编码，String类型
'                Bom_Component(n).setValue "DInvName", rsTmp.fields("cinvname").Value & ""  '子件名称(导出用)，String类型
'                Bom_Component(n).setValue "DInvStd", rsTmp.fields("cinvstd").Value & ""   '子件规格(导出用)，String类型
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
            Bom_Component(n).setValue "DRemark", Replace(rs!remark & "", """", "“") '备注，String类型
    '        Bom_Component(n).setValue "DInvAddCode", ""  '子件代号(导出用)，String类型
    '        Bom_Component(n).setValue "DOffset", ""  '偏置期(默认不偏置)，Integer类型
    '        Bom_Component(n).setValue "DWhCode", ""  '仓库编码(默认子件存货档案)，String类型
            Bom_Component(n).setValue "DDeptCode", DDeptCode  '领料部门(默认为子件存货档案)，String类型
'            Bom_Component(n).setValue "DInvFree_1", ""  '子件自由项1，String类型
'            Bom_Component(n).setValue "DInvFree_2", ""  '子件自由项2，String类型
'            Bom_Component(n).setValue "DInvFree_10", ""  '子件自由项10，String类型
'            Bom_Component(n).setValue "DInvDefine_1", ""  '子件自定义项1，String类型
'            Bom_Component(n).setValue "DInvDefine_16", ""  '子件自定义项16，Date类型
    '        Bom_Component(n).setValue "DSubDate", ""  '替换日期(导出用)，Date类型
    '        Bom_Component(n).setValue "DAuxUnitCode", ""  '辅助单位编码，String类型
    '        Bom_Component(n).setValue "DChangeRate", ""  '换算率，Double类型
    '        Bom_Component(n).setValue "DAuxBaseQtyN", ""  '辅助基本用量，Double类型
    '        Bom_Component(n).setValue "DBasEngineerFigNo", ""  '工程图号(导出用)，String类型
            Bom_Component(n).setValue "DDefine_28", rs!id & "" '
'            Bom_Component(n).setValue "DDefine_37", ""  'BOM表体自定义项16，Date类型
'            Bom_Component(n).setValue "cBDefine14", rs!cabinetnumber & ""  '屏体编码
'            Bom_Component(n).setValue "cBDefine15", rs!cabinetname & ""  '屏体名称
'            Bom_Component(n).setValue "cBDefine24", rs!projectnumber & ""  '项目编码
'            Bom_Component(n).setValue "cBDefine25", rs!projectname & ""  '项目名称
        
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
                    strErr = "系统错误："
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
    '结束本次调用，释放API资源
    u8apiBroker.Disconnect
    Set u8apiBroker = Nothing
    GoTo hFinish
Errhandler:
    strErr = Err.Description
hFinish:
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
    rs.open strsql, DBconn
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

Private Function GetVouchID(strTableName As String, lngIDs As String, lngsTableCount As Long, ErrMsg As String) As String
    Dim AdoComm As ADODB.Command
    On Error GoTo DoErr
    Set AdoComm = New ADODB.Command
    With AdoComm
        .ActiveConnection = DBconn
        .CommandText = "sp_GetID"
        .CommandType = adCmdStoredProc
        .Prepared = False
        .Parameters.Append .CreateParameter("RemoteId", adVarChar, adParamInput, 3, "00")
        .Parameters.Append .CreateParameter("cAcc_Id", adVarChar, adParamInput, 3, M_login.cAcc_Id)
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
    ErrMsg = "获取单据ID发生错误：" & Err.Description
    Set AdoComm = Nothing
End Function


'写入存货自由项
Private Function InsertInventoryFreeValue(cInvCode As String, cFree As String, citemcode As String, cabinetnumber As String, ByRef sError As String) As Boolean
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    Dim bTrans As Boolean
    On Error GoTo hErr
'    DBconn.BeginTrans
'    bTrans = True
    If cFree = "cFree1" Then
        '自由项档案
        strsql = "select  *  from userdefine where cId=N'20' and cvalue='" & citemcode & "' "
        rs.open strsql, DBconn
        If rs.BOF Or rs.EOF Then
            strsql = "insert into userdefine(cID,cValue)values('20','" & citemcode & "')"
            DBconn.Execute strsql
        End If
        '存货自由项取值范围
        strsql = "select * from AA_InvFreeContrapose where cinvcode='" & cInvCode & "' and cudid='20' and cudvalue='" & citemcode & "'"
        If rs.State <> 0 Then rs.Close
        rs.open strsql, DBconn
        If rs.BOF Or rs.EOF Then
            strsql = "insert into AA_InvFreeContrapose(cInvFreeGUID,cInvCode,cUDID,cUDValue)values(newid(),'" & cInvCode & "','20','" & citemcode & "')"
            DBconn.Execute strsql
        End If
    ElseIf cFree = "cFree2" Then
        '自由项档案
        strsql = "select  *  from userdefine where cId=N'21' and cvalue='" & cabinetnumber & "' "
        rs.open strsql, DBconn
        If rs.BOF Or rs.EOF Then
            strsql = "insert into userdefine(cID,cValue)values('21','" & cabinetnumber & "')"
            DBconn.Execute strsql
        End If
        '存货自由项取值范围
        strsql = "select * from AA_InvFreeContrapose where cinvcode='" & cInvCode & "' and cudid='21' and cudvalue='" & cabinetnumber & "'"
        If rs.State <> 0 Then rs.Close
        rs.open strsql, DBconn
        If rs.BOF Or rs.EOF Then
            strsql = "insert into AA_InvFreeContrapose(cInvFreeGUID,cInvCode,cUDID,cUDValue)values(newid(),'" & cInvCode & "','21','" & cabinetnumber & "')"
            DBconn.Execute strsql
        End If
    End If
'    If bTrans Then DBconn.CommitTrans: bTrans = False
    InsertInventoryFreeValue = True
    GoTo hFinish
hErr:
'    If bTrans Then DBconn.RollbackTrans: bTrans = False
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
    rs.open strsql, DBconn
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

Private Function GetChildCount(id As String) As Integer
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    On Error GoTo hErr
    strsql = "select count(1) as count from EF_V_HZDesignList where _parentid='" & id & "'"
    rs.open strsql, DBconn
    If Not rs.EOF And Not rs.BOF Then
        GetChildCount = str2Dbl(rs!Count & "")
    End If
    GoTo hFinish
hErr:

hFinish:
    If rs.State <> 0 Then
        rs.Close
    End If
    Set rs = Nothing
End Function
