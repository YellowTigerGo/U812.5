Attribute VB_Name = "modZwdZDFunc_PZ"
Option Explicit
Dim cOrderNo As String
Public md_ccode As String '默认借方科目
Public mc_ccode As String '默认贷方科目

Function AP_ZD(CZDList As clsZDList, Optional ByVal bcheck As Boolean = False, Optional ByRef bHide As Boolean) As Boolean
    Dim cTmpTable1 As String, cTmpTable2 As String
    Dim bRet As Boolean
    Dim i As Integer
    Dim cUser As String, cComp As String
    Dim iLen As Long, cLockType As String, cOp As String
    Dim tmprst As UfDbKit.UfRecordset ', GRst As UfDbKit.UfRecordset
    Dim cSql As String
    Dim rst1 As UfDbKit.UfRecordset
    Dim ikk As Long
    Dim cPZid$
    Dim Rst2 As New ADODB.Recordset
    cTmpTable1 = TmpZD1
    cTmpTable2 = TmpZD2
    On Error Resume Next
    Account.dbMData.Execute "Drop Table " & cTmpTable1
    Account.dbMData.Execute "Drop Table " & cTmpTable2
    'add by shenbo for 860
    Account.dbMData.Execute "delete " & Pubzz.WbTableName
    CreateZDTmp cTmpTable1
    bRet = Fill_VouchOther(CZDList, cTmpTable1, cTmpTable2, bcheck, bHide)
    If bRet And Not bcheck Then
        FrmZDKJ.LoadZZPz
        Unload FrmZDKJ
    End If
ExitFunc:
    AP_ZD = bRet
End Function
Public Function getPZID(ByVal cProcStyle As String, ByVal cCancelNo As String) As String
    Dim rs      As New ADODB.Recordset
    Dim cSql    As String
    
    Select Case cProcStyle

    End Select
    rs.Open cSql, Account.dbMData
    If rs.BOF Then
    'Result:Row=120 Col=17  Content="已进行取消操作"        ID=1d0fafc8-6b01-43f8-83b0-8c8af2701031
        getPZID = "已进行取消操作" 'GetResStringNoParam("U8.CW.APAR.ARAPMain.Canceled")
    Else
        getPZID = NullToStr(rs(0))
    End If
    
    rs.Close
    Set rs = Nothing
End Function
' 功能: 把原始单据转换为凭证格式
' 参数: cZD - 制单类
' 返回: 无
' 说明：在 {GL_vouchother} 表的 [coutsign] 字段填写值是为了从凭证到单据的查询中使用，明确单据来源
Function BillToPZ(CZD As clsZD, CPZ As clsFLList) As Boolean
On Error Resume Next
    Dim cKm1 As String, ckm2 As String, ckm3 As String, ckm4 As String
    Dim Rst As New ADODB.Recordset, rst1 As New ADODB.Recordset
    Dim cCust As String, cVend As String, cXmClass As String
    Dim iSum As Currency, iSum_f As Currency, iSum_s As Double, iSJ As Currency, iSJ_f As Currency
    Dim CtmpFL As clsAccount
    Dim i As Integer, cLink As String, iLine As Long ', bDirect As Boolean
    Dim cProcStyle As String, cCancelNo As String, cPzlb As String
    Dim cVouchType As String, cvouchid As String
    Dim cSql As String, cInv As String, ctmpID As String, ctmptype As String    'MD tqz 5.31
    Dim cSSCode As String
 
    Dim cexchName As String
    Dim KmSql As String, TmpSql As String
    Dim KmRst As New ADODB.Recordset, tmprst As New ADODB.Recordset
    Dim CMEM As String
    Dim m_BaseValue As Double, m_FValue As Double
    Dim cItem_Class As String, cItemCode As String
    Dim oDom    As DOMDocument, sTemp$
    
    cProcStyle = CZD.cProcStyle
    cCancelNo = CZD.cCancelNo
    Rst.CursorLocation = adUseClient
    Rst.LockType = adLockOptimistic
    Select Case cProcStyle
        Case "费用预估"
            GDZC_INC_FL CZD, CPZ
        Case "借款单"
            GDZC_INC_FL CZD, CPZ
    End Select
    BillToPZ = True
End Function
' 功能: 获取科目备查簿设置信息、科目指定信息
' 参数: oDefArch:科目备查簿设置信息, oCodeRemark:科目指定信息
Private Sub getCodeRemark(oDefArch As DOMDocument, oCodeRemark As DOMDocument)
    Dim cSql    As String
    Dim rs      As ADODB.Recordset
    
On Error Resume Next
    cSql = "select cfieldename,cfieldcname,isnull(cfielddefault,N'') as cfielddefault," & _
        "isnull(carchivename,N'') as carchivename,isnull(carchivefield,N'') as carchivefield " & _
        "from AutoSetFieldInf where cTableCode=N'CodeRemark' and isDefined=0 order by AutoID"
    Set rs = New ADODB.Recordset
    rs.Open cSql, Account.dbMData
    If oDefArch Is Nothing Then Set oDefArch = New DOMDocument
    rs.Save oDefArch, adPersistXML
    rs.Close
    cSql = "select ccode,remarkitem from gl_coderemarkitem order by ccode,i_id"
    rs.Open cSql
    If oCodeRemark Is Nothing Then Set oCodeRemark = New DOMDocument
    rs.Save oCodeRemark, adPersistXML
    rs.Close
    Set rs = Nothing
End Sub
' 功能: 把原始单据以凭证格式填入中间临时表中,整理好后的数据装入GL_VouchOther
' 参数: cZDList - 制单集合, cTmpTable1, cTmpTable2 - 临时表
' 返回: True: 成功，False: 失败
Function Fill_VouchOther(CZDList As clsZDList, ByVal cTmpTable1 As String, ByVal cTmpTable2 As String, Optional ByVal bcheck As Boolean = False, Optional ByRef bHide As Boolean) As Boolean
    Dim rstS As UfDbKit.UfRecordset, rstD As UfDbKit.UfRecordset
    Dim CTmpPZ As clsFLList, CtmpFL As clsFL, CTmpZD As clsZD
    Dim i As Long, j As Long, iCurrYear As Integer, iPeriod As Byte
    Dim bRet As Boolean
    Dim tdf As UfRecordset
    Dim cSql1 As String, cSql2 As String
    Dim cTmpSign As String, ctmpID As String, bUnite As Boolean
    Dim bMerge As Boolean, cPZid As String, iLine As Long, iFixedFields As Integer
    Dim cFieldList1 As String, cFieldList2 As String, cSumFieldList As String, cTmpFld As String
    Dim Code As String
    Dim iTmp_id As Long
    Dim cFld1() As String, cFld2() As String
    Dim cZdType As String
    Dim cTmpDigset As String, cTmpVTP As String, cTmpVID As String
    Dim sClass$, cItem$
    Dim oDefArch    As New DOMDocument
    Dim oCodeRemark As New DOMDocument
    Dim cProcStyle As String, cCancelNo As String
    
    getCodeRemark oDefArch, oCodeRemark
    iCurrYear = Year(m_Login.CurDate)
    iPeriod = Pubzz.DateToPeriod(CZDList.Item(1).dBillDate)
    
    Set rstD = Account.UfDbData.OpenRecordset(cTmpTable1, dbOpenDynaset)
    
    Dim iPos As Long
    Dim oldPH As String
    For i = 1 To CZDList.Count
        If bMerge Then Exit For
        If oldPH = CZDList.Item(i).cPH Then
            bMerge = True
        End If
        oldPH = CZDList.Item(i).cPH
    Next i
    
    For i = 1 To CZDList.Count
        Set CTmpPZ = New clsFLList
        Set CTmpZD = CZDList.Item(i)
        cZdType = CTmpZD.cProcSign
        
        '填充凭证类
        bRet = BillToPZ(CTmpZD, CTmpPZ)
        
        If Not bRet Then
            Set CTmpPZ = Nothing
            rstD.oClose
            Fill_VouchOther = False
            Exit Function
        End If
        With rstD                                               '凭证数据
            For j = 1 To CTmpPZ.Count
                Set CtmpFL = CTmpPZ(j)
                If CtmpFL.mc <> 0 Or CtmpFL.mc_f <> 0 Or CtmpFL.md <> 0 Or CtmpFL.md_f <> 0 Or CtmpFL.nd_s <> 0 Or CtmpFL.nc_s <> 0 Then
                    .AddNew
                    !coutaccset = m_Login.cAcc_Id               '外部凭证账套号√
                    !ioutyear = m_Login.cIYear                  '外部凭证会计年度√
                    !coutsysname = "AP"                         '外部凭证系统名称(ID)√
                    !doutbilldate = CTmpZD.dBillDate
                    !ioutperiod = iPeriod                       '外部凭证会计期间√
                    !cbill = m_Login.cUserName                   '制单人
                    !coutsign = "AP"
                    !coutno_id = CtmpFL.coutno_id
                    !inid = CtmpFL.inid
                    !doutdate = m_Login.CurDate  'CtmpFL.doutdate
                    !coutbillsign = CtmpFL.coutsign
                    !coutid = CtmpFL.coutid
                    !cSign = CtmpFL.cSign               'cPzlb
                    !idoc = CtmpFL.idoc                 '附加单据数 shanlan modify 20090728
                    cTmpDigset = CtmpFL.cDigest
                    cTmpVTP = CtmpFL.coutbillsign
                    cTmpVID = CtmpFL.coutid
                    !cDigest = IIf(CtmpFL.cDigest = "", " ", Pubzz.GenTrim(CtmpFL.cDigest, 60))
                    !cCode = StrToNull(CtmpFL.cCode)
                    !md = CtmpFL.md
                    !mc = CtmpFL.mc
                    !nfrat = CtmpFL.nfrat
                    If !nfrat <> 0 Then
                        !md_f = CtmpFL.md_f
                        !mc_f = CtmpFL.mc_f
                    Else
                        !md_f = 0
                        !mc_f = 0
                    End If
                    !nd_s = CtmpFL.nd_s
                    !nc_s = CtmpFL.nc_s
                    !csettle = StrToNull(CtmpFL.csettle)
                    !cn_id = StrToNull(CtmpFL.cn_id)
                    !dt_date = IIf(CtmpFL.dt_date = 0, Null, CtmpFL.dt_date)
                    !cdept_id = StrToNull(CtmpFL.cdept_id)
                    !ccus_id = StrToNull(CtmpFL.ccus_id)
                    !csup_id = StrToNull(CtmpFL.csup_id)
                    !citem_id = StrToNull(CtmpFL.citem_id)
                    !cItem_Class = StrToNull(CtmpFL.cItem_Class)
                    !cname = StrToNull(CtmpFL.cname)
                    !cPerson_id = StrToNull(CtmpFL.cPerson_id)
                    !bvouchedit = CtmpFL.bvouchedit
                    !bvouchAddordele = CtmpFL.bvouchAddordele
                    !bvouchmoneyhold = CtmpFL.bvouchmoneyhold
                    !bvalueedit = CtmpFL.bvalueedit
                    !bcodeedit = CtmpFL.bcodeedit
                    !ccodecontrol = CtmpFL.ccodecontrol
                    'modify by wgz 2002-12-26
                    If !bvalueedit Then
                        !bPCSedit = 1
                        !bCusSupInput = 1
                    Else
                        !bPCSedit = 0
                        !bCusSupInput = 0
                    End If
                    !bDeptedit = CtmpFL.bDeptedit
                    !bItemedit = CtmpFL.bItemedit
                    !cDefine1 = StrToLen(CtmpFL.cDefine1, 20)
                    !cDefine2 = StrToLen(CtmpFL.cDefine2, 20)
                    !cDefine3 = StrToLen(CtmpFL.cDefine3, 20)
                    !cDefine4 = StrToNull(CtmpFL.cDefine4)
                    !cDefine5 = StrToNull(CtmpFL.cDefine5)
                    !cDefine6 = StrToNull(CtmpFL.cDefine6)
                    !cDefine7 = StrToNull(CtmpFL.cDefine7)
                    !cDefine8 = StrToLen(CtmpFL.cDefine8, 4)
                    !cDefine9 = StrToLen(CtmpFL.cDefine9, 8)
                    !cDefine10 = StrToLen(CtmpFL.cDefine10, 60)
                    !cDefine11 = StrToLen(CtmpFL.cDefine11, 120)
                    !cDefine12 = StrToLen(CtmpFL.cDefine12, 120)
                    !cDefine13 = StrToLen(CtmpFL.cDefine13, 120)
                    !cDefine14 = StrToLen(CtmpFL.cDefine14, 120)
                    !cDefine15 = StrToNull(CtmpFL.cDefine15)
                    !cDefine16 = StrToNull(CtmpFL.cDefine16)
                    ' 以下字段用于整理数据时使用
                    !cBlueID = CtmpFL.cBlueID
                    !cTableName = CtmpFL.cTableName
                    !cFieldName = CtmpFL.cFieldName
                    !ibillno_id = CtmpFL.ibillno_id
                    !cProcNo = CtmpFL.cProcNo                 'cPH
                    !bTaxFlag = CtmpFL.bTaxFlag
                    !cProcStyle = CtmpFL.cProcStyle
                    !cCancelNo = CtmpFL.cCancelNo
                    !bPrepay = CtmpFL.bPrepay
                    !iLink = CtmpFL.iLink
                    !cvouchid = CtmpFL.cvouchid
                    'add by wgz 2003-7-30
'                    If !cCode <> "" Then ConvertDef rstD, oDefArch, oCodeRemark
                    'add by shenbo 2004-01-15
                    cProcStyle = CtmpFL.cProcStyle
                    cCancelNo = CtmpFL.cCancelNo
                    iPos = iPos + 1
                    !iPos = iPos
                    
                    !cmergeno = CtmpFL.cmergeno
            
                    .Update '将每条分录记入TMPZD1内
                End If
            Next j
        End With
        Set CTmpPZ = Nothing
    Next i
    rstD.oClose '关闭TMPZD1
 
    Set tdf = Account.UfDbData.OpenRecordset(Pubzz.WbTableName, dbOpenDynaset)
    iFixedFields = tdf.Fields.Count
    ReDim cFld1(iFixedFields), cFld2(iFixedFields)
    '拼写 GL_Vouchother 字段串,以","隔开。
    cFieldList1 = ""
    cFieldList2 = ""
    cSumFieldList = ""
    For j = 1 To iFixedFields
        cTmpFld = tdf.Fields(j - 1).Name
        If cTmpFld <> "i_id" Then
            cFieldList1 = cFieldList1 & "," & cTmpFld
            cFieldList2 = cFieldList2 & ",t_" & cTmpFld
            If cTmpFld = "md" Or cTmpFld = "mc" Or cTmpFld = "md_f" Or cTmpFld = "mc_f" Or cTmpFld = "nd_s" Or cTmpFld = "nc_s" Then
                cSumFieldList = cSumFieldList & ",Sum(" & cTmpFld & ") as t_" & cTmpFld
            ElseIf tdf.Fields(j - 1).Type = adBoolean Then
                cSumFieldList = cSumFieldList & ",Max(Convert(int," & cTmpFld & ")) as t_" & cTmpFld
'            ElseIf cTmpFld = "cn_id" Or cTmpFld = "inid" Then 'shanlan 090813 票号不同不合并 090922 完全不合并
'                cSumFieldList = cSumFieldList & "," & cTmpFld & " as t_" & cTmpFld
            Else
                cSumFieldList = cSumFieldList & ",Max(" & cTmpFld & ") as t_" & cTmpFld
            End If
        End If
    Next j
    
    cSumFieldList = cSumFieldList & ",min(cProcNo) as t_cProcNo"
    cSumFieldList = cSumFieldList & ",min(iPos) as t_iPos"
    cFieldList1 = Mid(cFieldList1, 2)
    cFieldList2 = Mid(cFieldList2, 2)
    cSumFieldList = Mid(cSumFieldList, 2)
'    If oAcc.MakeShtFs = 3 Then
'        Account.dbMData.Execute "Update " & cTmpTable1 & " set ccus_id=null from code where " & cTmpTable1 & ".cCode=code.cCode and code.bcus=0"
'        Account.dbMData.Execute "Update " & cTmpTable1 & " set csup_id=null from code where " & cTmpTable1 & ".cCode=code.cCode and code.bsup=0"
'    End If
    
'    Account.dbMData.Execute "Update " & cTmpTable1 & " set cdept_id=null from code where " & cTmpTable1 & ".cCode=code.cCode and (code.bdept=0 And bperson=0)"
'    'add by wgz 2001/11/21 (个人核算)
'    Account.dbMData.Execute "Update " & cTmpTable1 & " set cperson_id=null from code where " & cTmpTable1 & ".cCode=code.cCode and bperson=0"
'    Account.dbMData.Execute "Update " & cTmpTable1 & " set citem_class=null,citem_id=null from code where " & cTmpTable1 & ".cCode=code.cCode and code.cass_item is null"
    
    Dim bCtrl As Boolean
    Dim cCode As String
    Dim bd_c As Boolean
    Dim cCusCode As String
    Dim cVenCode As String
    Dim cdept_id$, cPerson_id$, citem_id$, cExch_Name$
    Dim nfrat As Double
    Dim coutbillsign$, coutid$, cProcNo$
    Dim bTmpd_c As Boolean
    Dim coutno_id$, sTemp$
    Dim oElm    As IXMLDOMElement
       
    cSql2 = "Select all cProcStyle,cCancelNo,inid,bvalueedit,ccode,md,mc,nd_s,nc_s,coutno_id,cMergeno,ccus_id,csup_id," & _
    "cdept_id,cperson_id,citem_id,cexch_name,nfrat,coutbillsign,coutid,cprocno,cn_id"
 
    cSql2 = cSql2 & sTemp & " From " & cTmpTable1 & _
    " order by inid, ccode,case when md=0 then 0 else 1 end,ccus_id,csup_id," & _
    "cdept_id,cperson_id,citem_id,cexch_name,nfrat,coutbillsign,coutid,cprocno,cn_id"
    Set rstS = Account.UfDbData.OpenRecordset(cSql2)
    i = 0
    bCtrl = False
    With rstS
        Do While Not .EOF
            bTmpd_c = !md <> 0
            If cCode = NullToStr(!cCode) And _
                (bd_c = bTmpd_c) And _
                cCusCode = NullToStr(!ccus_id) And _
                cVenCode = NullToStr(!csup_id) And _
                cdept_id = NullToStr(!cdept_id) And _
                cPerson_id = NullToStr(!cPerson_id) And _
                citem_id = NullToStr(!citem_id) And _
                cExch_Name = NullToStr(!cExch_Name) And _
                (coutbillsign = !coutbillsign) And _
                cProcNo = !cProcNo And nfrat = !nfrat Then
            Else
                cCode = NullToStr(!cCode)
                bd_c = bTmpd_c
                cCusCode = NullToStr(!ccus_id)
                cVenCode = NullToStr(!csup_id)
                cdept_id = NullToStr(!cdept_id)
                cPerson_id = NullToStr(!cPerson_id)
                citem_id = NullToStr(!citem_id)
                coutbillsign = NullToStr(!coutbillsign)
                cProcNo = NullToStr(!cProcNo)
                coutno_id = NullToStr(!coutno_id)
                nfrat = NullToZero(!nfrat)
                i = i + 1
            End If
            .Edit
            '!cmergeno = i
            !coutno_id = coutno_id
            .Update
            .MoveNext
        Loop
        .oClose
    End With
        
    cSql2 = "SELECT DISTINCT cMergeno," & cSumFieldList & _
            " INTO " & cTmpTable2 & _
            " FROM " & cTmpTable1 & " GROUP BY cMergeno,ccode" ',nfrat"
    '2001-3-20 add by wgz(end)
    
    '生成合并表
    Account.dbMData.Execute cSql2
    '删除借贷金额都为零的分录(add by wgz 2002-12-3)
    Account.dbMData.Execute "Delete " & cTmpTable2 & " Where t_md=0 and t_mc=0"
    '整理合并不同方向分录的借贷方
    Account.dbMData.Execute "Update " & cTmpTable2 & _
        " set t_md=case when t_md>t_mc then t_md-t_mc else 0 end," & _
        "t_mc=case when t_mc>t_md then t_mc-t_md else 0 end," & _
        "t_md_f=case when t_md_f>t_mc_f then t_md_f-t_mc_f else 0 end," & _
        "t_mc_f=case when t_mc_f>t_md_f then t_mc_f-t_md_f else 0 end," & _
        "t_nd_s=case when t_nd_s>t_nc_s then t_nd_s-t_nc_s else 0 end," & _
        "t_nc_s=case when t_nc_s>t_nd_s then t_nc_s-t_nd_s else 0 end " & _
        "Where t_md<>0 and t_mc<>0"
     
    Set rstS = Account.UfDbData.OpenRecordset("SELECT * FROM " & cTmpTable2 & " ORDER BY t_cprocno,case when t_md=0 then 1 else 0 end,t_inid,t_iPos", 2) 'val(cmergeno)", dbOpenDynaset)
    iLine = 1
        rstS.oClose
    '如果合并制单则业务类型为系统名
    If bUnite Then Account.dbMData.Execute "Update " & cTmpTable2 & _
        " set t_coutsign=N'" & "AP" & "' Where t_coutno_id=N'" & cPZid & "'"
    
    Set rstS = Account.UfDbData.OpenRecordset(cTmpTable2, dbOpenDynaset)
    Set rstD = Account.UfDbData.OpenRecordset(Pubzz.WbTableName, dbOpenDynaset)   '  Account.ufdbdata.OpenRecordset("GL_Vouchother", dbOpenDynaset)
    For j = 0 To iFixedFields - 1
        cFld2(j) = rstD.Fields(j).Name
        cFld1(j) = "t_" & cFld2(j)
    Next j
'    rstS.MoveFirst
    Do While Not rstS.EOF
        rstD.AddNew
        For j = 0 To iFixedFields - 1
            If cFld2(j) <> "i_id" Then
                If rstS.Fields(cFld1(j)).Type = adDBTimeStamp Then
                    If IsNull(rstS.Fields(cFld1(j))) Then rstD.Fields(cFld2(j)) = Null Else rstD.Fields(cFld2(j)) = CDate(rstS.Fields(cFld1(j)))
                Else
                    rstD.Fields(cFld2(j)) = rstS.Fields(cFld1(j))
                End If
            Else
                iTmp_id = rstD!i_id
            End If
        Next j
        rstD.Fields("coutid") = LeftEx(NullToStr(rstS.Fields("t_coutID")) + Space(30), 30) & rstS.Fields("cmergeno")
        rstD.Update
        rstS.MoveNext
    Loop
    rstS.oClose
    rstD.oClose
    Set rstS = Nothing
    Set rstD = Nothing
    Fill_VouchOther = True
End Function


'超长截断
Private Function StrToLen(ByVal Str As String, ByVal nLen As Integer) As Variant
    If Str = "" Then
       StrToLen = Null
    Else
       StrToLen = Left(Str, nLen)
    End If
End Function

' 功能: 创建临时凭证表
' 参数: cTmpName - 表名
' 返回: 无
Sub CreateZDTmp(ByVal cTmpName As String)
    Dim cTmpTable As String, cSql As String
    Dim cDBDataName As String
    '对于临时凭证表操作改成 tempdb 数据库
    On Error Resume Next
    Err = 0
    With Account.dbMData
        On Error Resume Next
        .Execute "DROP TABLE " & cTmpName
        On Error GoTo 0
        'Err.Clear
        '生成表结构
        .Execute "SELECT * INTO " & cTmpName & " FROM " + Pubzz.WbTableName + " WHERE 1=0"
        .Execute "ALTER TABLE " & cTmpName & " ADD cBlueID nvarchar(20) "
        .Execute "ALTER TABLE " & cTmpName & " ADD ctag nvarchar(1)"
        .Execute "ALTER TABLE " & cTmpName & " ADD cTableName nvarchar(30)"
        .Execute "ALTER TABLE " & cTmpName & " ADD cFieldname nvarchar(20)"
        .Execute "ALTER TABLE " & cTmpName & " ADD iBillno_id nvarchar(40)"
        .Execute "ALTER TABLE " & cTmpName & " ADD cprocstyle nvarchar(10)"
        .Execute "ALTER TABLE " & cTmpName & " ADD ccancelno nvarchar(40)"
        .Execute "ALTER TABLE " & cTmpName & " ADD cprocno int"
        .Execute "ALTER TABLE " & cTmpName & " ADD cmergeno nvarchar(20)"
        .Execute "ALTER TABLE " & cTmpName & " ADD btaxflag bit"
        .Execute "ALTER TABLE " & cTmpName & " ADD iLink int"
        .Execute "ALTER TABLE " & cTmpName & " ADD bprepay bit"
        .Execute "ALTER TABLE " & cTmpName & " ADD cItemName nVarchar(60)"
        .Execute "ALTER TABLE " & cTmpName & " ADD iPos int" '物理顺序2004-07-22
        'by zzc
        .Execute "ALTER TABLE " & cTmpName & " ADD cvouchID nvarchar(40)"
        
        '创建索引(by wgz 2002/11/25)
        .Execute "Create Index iBillno_id ON " & cTmpName & " (iBillno_id)"
        .Execute "Create Index coutid ON " & cTmpName & " (coutbillsign,coutid,cCancelNo,ccus_id,iLink,ino_id,coutno_id)"
        .Execute "Create Index cTableName ON " & cTmpName & "(cTableName)"
        .Execute "Create Index cmergeno ON " & cTmpName & "(cmergeno)"
        .Execute "Create Index cprocstyle ON " & cTmpName & "(cprocstyle)"
        
        If Err > 0 Then
            MsgBox Err.Description, vbCritical
            Exit Sub
        End If
    End With
    On Error GoTo 0
End Sub

Function GDZC_INC_FL(CZD As clsZD, CPZ As clsFLList) As Boolean
Dim cKm1 As String, ckm2 As String, ckm3 As String, ckm4 As String
    Dim Rst As UfRecordset
    Dim rst1 As UfRecordset
    Dim cCust As String, cVend As String, cXmClass As String
    Dim iSum As Currency, iSum_f As Currency, iSum_s As Double
    Dim CtmpFL As clsFL
    Dim i As Integer
    Dim cProcStyle As String, cCancelNo As String
    Dim cVouchType As String, cvouchid As String
    Dim cSql As String, cInv As String, ctmpID As String, ctmptype As String
'    Dim iBillType As EnumVouchType
    Dim tmpdei As String
    Dim iSJ As Currency, iSJ_f As Currency
    Dim imcSum As Currency, imcSum_f As Currency
    Dim imdSum As Currency, imdSum_f As Currency
    Dim imcSj As Currency, imcSj_f As Currency
    Dim imdSj As Currency, imdSj_f As Currency
    Dim iLine As Integer
    Dim RstSj As UfRecordset
    Dim cSSCode As String
    Dim CMEM As String
    Dim cexchName As String
    Dim cKm(1 To 4) As String, iVal(1 To 4) As Currency, bd_c(1 To 4) As Boolean
    Dim cDigest(1 To 4) As String
    cProcStyle = CZD.cProcStyle
    cCancelNo = CZD.cCancelNo
    cVouchType = cProcStyle
    cvouchid = CZD.cvouchid
    iLine = 1
    Set CtmpFL = New clsFL
    With CtmpFL    '借方
        .iPeriod = CZD.iPeriod
        .cSign = CZD.cPzlb
        .dBillDate = CZD.dBillDate
        .idoc = CZD.idoc
        .cbill = CZD.cbill
        .ccheck = CZD.ccheck
        .cbook = CZD.cbook
        .ibook = CZD.ibook
        .iflag = CZD.iflag
        .ctext1 = CZD.ctext1
        .ctext2 = CZD.ctext2
        .cDigest = CZD.cDigest
        .cCode = CZD.cCode
        .ccashier = CZD.ccashier
        .cExch_Name = CZD.cExch_Name
        .md = CZD.md
        .mc = 0
        .md_f = CZD.md_f
        .mc_f = 0
        .nfrat = CZD.nfrat
        .nd_s = CZD.nd_s
        .nc_s = CZD.nc_s
        .csettle = CZD.csettle
        .cn_id = CZD.cn_id
        .dt_date = CZD.dt_date
        .cdept_id = CZD.cdept_id
        .cPerson_id = CZD.cPerson_id
        .ccus_id = CZD.ccus_id
        .csup_id = CZD.csup_id
        .citem_id = CZD.citem_id
        .cItem_Class = CZD.cItem_Class
        .cname = CZD.cname
        .ccode_equal = CZD.ccode_equal
        .iflagbank = CZD.iflagbank
        .iflagPerson = CZD.iflagPerson
        .bdelete = CZD.bdelete
        .coutaccset = CZD.coutaccset
        .ioutyear = CZD.ioutyear
        .coutsysname = CZD.coutsysname
        .coutsysver = CZD.coutsysver
        .doutbilldate = CZD.doutbilldate
        .ioutperiod = CZD.ioutperiod
        .coutsign = CZD.cProcSign
        .coutno_id = cCancelNo
        .doutdate = CZD.doutdate
        .coutbillsign = CZD.coutbillsign
        .coutid = CZD.coutid
        .ccodecontrol = "AP,#"
        .cDefine1 = CZD.cDefine1                                     '自定义项1
        .cDefine2 = CZD.cDefine2                                     '自定义项2
        .cDefine3 = CZD.cDefine3                                     '自定义项3
        .cDefine4 = CZD.cDefine4                                     '自定义项4
        .cDefine5 = CZD.cDefine5                                     '自定义项5
        .cDefine6 = CZD.cDefine6                                     '自定义项6
        .cDefine7 = CZD.cDefine7                                     '自定义项7
        .cDefine8 = CZD.cDefine8                                     '自定义项8
        .cDefine9 = CZD.cDefine9                                     '自定义项9
        .cDefine10 = CZD.cDefine10                                   '自定义项10
        .cDefine11 = CZD.cDefine11                                   '自定义项11
        .cDefine12 = CZD.cDefine12                                   '自定义项12
        .cDefine13 = CZD.cDefine13                                   '自定义项13
        .cDefine14 = CZD.cDefine14                                   '自定义项14
        .cDefine15 = CZD.cDefine15                                   '自定义项15
        .cDefine16 = CZD.cDefine16                                   '自定义项16
        .inid = CZD.inid                      '//自动生成的行号
        .coutsign = CZD.cProcSign         '//业务类型
        .coutbillsign = CZD.cProcStyle    '//处理方式
        .cTableName = ""
        .cFieldName = ""
        .cProcNo = CZD.cPH                '//制单批号
        .cProcStyle = cProcStyle          '//处理方式
        .cCancelNo = CZD.cCancelNo        '//记录ID号
        
        .cmergeno = CZD.cmergeno
        'by ahzzd 20060911 凭证摘要需要修改
        .bvouchedit = True
        .bvouchAddordele = False
        .bvouchmoneyhold = False
        .bvalueedit = True
        .bcodeedit = True
        .bPCSedit = True
        .bDeptedit = True
        .bItemedit = True
        .bCusSupInput = True
        '自定义项
    End With
    
    CPZ.AddFL CtmpFL
    Set CtmpFL = Nothing
    iLine = iLine + 1
    Set CtmpFL = New clsFL
    With CtmpFL    '贷方
        .iPeriod = CZD.iPeriod
        .cSign = CZD.cPzlb
        .dBillDate = CZD.dBillDate
        .idoc = CZD.idoc
        .cbill = CZD.cbill
        .ccheck = CZD.ccheck
        .cbook = CZD.cbook
        .ibook = CZD.ibook
        .iflag = CZD.iflag
        .ctext1 = CZD.ctext1
        .ctext2 = CZD.ctext2
        .cDigest = CZD.cDigest
        .cCode = CZD.dCode
        .ccashier = CZD.ccashier
        .cExch_Name = CZD.cExch_Name
        .md = 0
        .mc = CZD.mc
        .md_f = 0
        .mc_f = CZD.mc_f
        .nfrat = CZD.nfrat
        .nd_s = CZD.nd_s
        .nc_s = CZD.nc_s
        .csettle = CZD.csettle
        .cn_id = CZD.cn_id
        .dt_date = CZD.dt_date
        .cdept_id = CZD.cdept_id
        .cPerson_id = CZD.cPerson_id
        .ccus_id = CZD.ccus_id
        .csup_id = CZD.csup_id
        .citem_id = CZD.citem_id
        .cItem_Class = CZD.cItem_Class
        .cname = CZD.cname
        .ccode_equal = CZD.ccode_equal
        .iflagbank = CZD.iflagbank
        .iflagPerson = CZD.iflagPerson
        .bdelete = CZD.bdelete
        .coutaccset = CZD.coutaccset
        .ioutyear = CZD.ioutyear
        .coutsysname = CZD.coutsysname
        .coutsysver = CZD.coutsysver
        .doutbilldate = CZD.doutbilldate
        .ioutperiod = CZD.ioutperiod
        .coutsign = CZD.cProcSign
        .coutno_id = cCancelNo
        .doutdate = CZD.doutdate
        .coutbillsign = CZD.coutbillsign
        .coutid = CZD.coutid
        .ccodecontrol = "AP,#"
        .cDefine1 = CZD.cDefine1                                     '自定义项1
        .cDefine2 = CZD.cDefine2                                     '自定义项2
        .cDefine3 = CZD.cDefine3                                     '自定义项3
        .cDefine4 = CZD.cDefine4                                     '自定义项4
        .cDefine5 = CZD.cDefine5                                     '自定义项5
        .cDefine6 = CZD.cDefine6                                     '自定义项6
        .cDefine7 = CZD.cDefine7                                     '自定义项7
        .cDefine8 = CZD.cDefine8                                     '自定义项8
        .cDefine9 = CZD.cDefine9                                     '自定义项9
        .cDefine10 = CZD.cDefine10                                   '自定义项10
        .cDefine11 = CZD.cDefine11                                   '自定义项11
        .cDefine12 = CZD.cDefine12                                   '自定义项12
        .cDefine13 = CZD.cDefine13                                   '自定义项13
        .cDefine14 = CZD.cDefine14                                   '自定义项14
        .cDefine15 = CZD.cDefine15                                   '自定义项15
        .cDefine16 = CZD.cDefine16                                   '自定义项16
        .inid = CZD.inid                      '//自动生成的行号
        .coutsign = CZD.cProcSign         '//业务类型（固定资产大修计提）
        .coutbillsign = CZD.cProcStyle    '//处理方式（资产增加/资产减少）
        .cTableName = ""
        .cFieldName = ""
        .cProcNo = CZD.cPH                '//制单批号
        .cProcStyle = cProcStyle          '//处理方式
        .cCancelNo = CZD.cCancelNo        '//记录ID号
        
        .cmergeno = CZD.cmergeno
        
        'by ahzzd 20060911 凭证摘要需要修改
        .bvouchedit = True
        .bvouchAddordele = False
        .bvouchmoneyhold = False
        .bvalueedit = True
        .bcodeedit = True
        .bPCSedit = True
        .bDeptedit = True
        .bItemedit = True
        .bCusSupInput = True
        '自定义项
    End With
    CPZ.AddFL CtmpFL
    'iLine = iLine + 1
End Function
