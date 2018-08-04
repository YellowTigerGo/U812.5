VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public m_Conn As New ADODB.Connection
'为了web调用不出现冲突
Public clsSys As Object 'New clsSystem
Public strVouchType As String
Public o_cac As Object
Enum UseMode
    BS
    CS
End Enum
Public m_UseMode As UseMode
'Public menuVouchtype As Vouchertype
Enum ReferType
    Inventory
    warehouse
    Customer
    Department
    Person
    SaleType
    Rd_Style
    PayCondition
    Digest
    ForeignCurrency
    userdefine
    ItemClass
    Memo
    ExpenseItem
    Batch
    VendorTrack     '供应商追踪
    InvTrack        '存货追踪
    CurrentStock
    vendor '供应商 代管仓
    demandclass
End Enum
'Dim clsVouchCheckbill As New clsSBVFunc
Private sSql As String
Private sKeyOld As String
'检查原始单据是否有批号等
Private Function GetBatchProperty(CN As ADODB.Connection, ele As IXMLDOMElement) As String
    Dim objSrv As Object
    Dim strXMl As String
    Dim rst As New ADODB.Recordset
    Dim i As Integer
    
    rst.CursorLocation = adUseClient
    Set objSrv = CreateObject("U8SrvTrans.IClsCommon")
    objSrv.initnew clsSys.objlogin
    strXMl = "<AA_BatchProperty>" & _
        "<cInvCode>" & GetAttrValue(ele, "cinvcode") & "</cInvcode>" & _
        "<cBatch>" & GetAttrValue(ele, "cbatch") & "</cBatch>" & _
        "<cFree1>" & GetAttrValue(ele, "cfree1") & "</cFree1>" & _
        "<cFree2>" & GetAttrValue(ele, "cfree2") & "</cFree2>" & _
        "<cFree3>" & GetAttrValue(ele, "cfree3") & "</cFree3>" & _
        "<cFree4>" & GetAttrValue(ele, "cfree4") & "</cFree4>" & _
        "<cFree5>" & GetAttrValue(ele, "cfree5") & "</cFree5>" & _
        "<cFree6>" & GetAttrValue(ele, "cfree6") & "</cFree6>" & _
        "<cFree7>" & GetAttrValue(ele, "cfree7") & "</cFree7>" & _
        "<cFree8>" & GetAttrValue(ele, "cfree8") & "</cFree8>" & _
        "<cFree9>" & GetAttrValue(ele, "cfree9") & "</cFree9>" & _
        "<cFree10>" & GetAttrValue(ele, "cfree10") & "</cFree10>" & _
        "</AA_BatchProperty>"
    Set rst = objSrv.getrstbyxml("AA_BatchProperty", strXMl)
    If Not rst Is Nothing Then
        If rst.EOF Then
            For i = 1 To 10
                ele.setAttribute "cbatchproperty" & CStr(i), ""
            Next
        Else
            For i = 1 To 10
                ele.setAttribute "cbatchproperty" & CStr(i), IIf(IsNull(rst.Fields("cbatchproperty" & CStr(i))), "", rst.Fields("cbatchproperty" & CStr(i)))
            Next
        End If
        rst.Close
    End If
    Set rst = Nothing
    Set objSrv = Nothing
End Function
Public Function CheckOriBatch(CN As ADODB.Connection, domBody As DOMDocument, sKey As String, ErrMsg As String, R As Long, sVouchType As String) As Boolean  '
    Dim rst As ADODB.Recordset
    Dim strSQL As String, strOriID As String
    
    On Error GoTo DoErr
    
    If sVouchType <> "05" And sVouchType <> "06" Then
        CheckOriBatch = False
        Exit Function
    End If
    
    strOriID = GetBodyItemValue(domBody, "iCorID", 0)
    If Trim(strOriID) = "" Or Val(strOriID) = 0 Then
        CheckOriBatch = False
        Exit Function
    End If
    strSQL = "select " + sKey + " from dispatchlists where idlsid=" & strOriID
    Set rst = New ADODB.Recordset
    rst.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    If rst.EOF Then
        CheckOriBatch = False
        Exit Function
    Else
        If IsNull(rst(0)) Then
            rst.Close
            Set rst = Nothing
            CheckOriBatch = False
            Exit Function
        Else
            If rst(0) <> "" Then
                rst.Close
                Set rst = Nothing
                CheckOriBatch = True
                Exit Function
            Else
                rst.Close
                Set rst = Nothing
                CheckOriBatch = True
                Exit Function
            End If
        End If
    End If
    Exit Function
DoErr:
    ErrMsg = GetStringPara("U8.SA.USSASERVER.clscommcheck.00377", err.Description) 'Para zh-CN：检查对应单据项目状态发生错误{0}
    If CN.Errors.Count > 0 Then
        ErrMsg = GetStringPara("U8.SA.USSASERVER.clscommcheck.00377", CN.Errors(0).Description) 'Para zh-CN：检查对应单据项目状态发生错误{0}
    End If
    Set rst = Nothing
    CheckOriBatch = True
    Exit Function
End Function

Private Function CheckPass(strPass As String) As Boolean
    Dim sSerName As String
    Dim oriPass As String
    Dim i As Long
    Dim j As Long
    Dim key()
    CheckPass = False
    If strPass = "122-122-103-120-106" Then
        CheckPass = True
    Else
        sSerName = clsSys.objlogin.cServer
        sSerName = StrConv(sSerName, vbFromUnicode)

        ReDim key(LenB(sSerName))
        oriPass = ""
        For i = 0 To UBound(key) - 1
            key(i) = MidB(sSerName, i + 1, 1)
            oriPass = oriPass & (Asc(StrConv(key(i), vbUnicode)) + Asc(i + 1))
        Next
        If LCase(strPass) = LCase(oriPass) Then CheckPass = True
        
        
    End If
End Function

Private Function GetSQLToValid(ByVal enmReferType As ReferType, Optional domBody As DOMDocument, Optional domHead As DOMDocument, Optional ByVal R) As String
    Dim strWhCode As String
    Dim strInvCode As String
    Dim strFree1 As String
    Dim strFree2 As String
    Dim strFree3 As String
    Dim strFree4 As String
    Dim strFree5 As String
    Dim strFree6 As String
    Dim strFree7 As String
    Dim strFree8 As String
    Dim strFree9 As String
    Dim strFree10 As String
    Dim strBatch As String
    Dim strCode As String
    Dim strVenCode As String
    Dim i As Integer
    Dim strTaxRate As String
    Dim siSOsID As String
    Dim cvmivencode As String
'     Dim blnWhAuth As Boolean
     Dim strAuth As String
    'Dim sDate As String, sWhere As String
    
    If IsMissing(R) Then R = 0
    Select Case enmReferType
        Case demandclass
            GetSQLToValid = "SELECT cRClassCode as cdemandcode,cRClassName as cdemandmemo from AA_RequirementClass"
        Case Customer
            
            'GetSQLToValid = "SELECT cCusCode,cCusAbbName,ccusbank FROM customer "
            GetSQLToValid = "SELECT Person.cPersonCode as chandler,cCusAddress as cshipaddress,department.cdepcode as cdepcode,cDepName,cPersonName,person.cPersonCode as cpersoncode, Customer.cCusCode, Customer.cCusName, Customer.cCusAbbName, "
            GetSQLToValid = GetSQLToValid & "Customer.cCCCode, Customer.cDCCode, Customer.cTrade,"
            GetSQLToValid = GetSQLToValid & "Customer.cCusAddress, Customer.cCusPostCode,"
            GetSQLToValid = GetSQLToValid & "Customer.cCusRegCode, Customer.cCusBank, Customer.cCusAccount,"
            GetSQLToValid = GetSQLToValid & "Customer.dCusDevDate, Customer.cCusLPerson, Customer.cCusEmail,"
            GetSQLToValid = GetSQLToValid & "Customer.cCusPerson, Customer.cCusPhone, Customer.cCusFax,"
            GetSQLToValid = GetSQLToValid & "Customer.cCusBP, Customer.cCusHand, Customer.cCusPPerson,"
            GetSQLToValid = GetSQLToValid & "Customer.iCusDisRate, Customer.cCusCreGrade, Customer.iCusCreLine,"
            GetSQLToValid = GetSQLToValid & "Customer.iCusCreDate, Customer.cCusPayCond,"
            GetSQLToValid = GetSQLToValid & "Customer.cCusOAddress, Customer.cCusOType,"
            GetSQLToValid = GetSQLToValid & "Customer.cCusHeadCode, Customer.cCusWhCode,"
            GetSQLToValid = GetSQLToValid & "Customer.cCusDepart, Customer.iARMoney, Customer.dLastDate,"
            GetSQLToValid = GetSQLToValid & "Customer.iLastMoney, Customer.dLRDate, Customer.iLRMoney,"
            GetSQLToValid = GetSQLToValid & "Customer.dEndDate, Customer.iFrequency, Customer.cCusDefine1,"
            GetSQLToValid = GetSQLToValid & "Customer.cCusDefine2, Customer.cCusDefine3, Customer.iCostGrade,"
            GetSQLToValid = GetSQLToValid & "Customer.cCreatePerson, Customer.cModifyPerson,"
            GetSQLToValid = GetSQLToValid & "Customer.dModifyDate, Customer.cRelVendor, Customer.iId,"
            GetSQLToValid = GetSQLToValid & "Customer.cPriceGroup, Customer.cOfferGrade, Customer.iOfferRate,"
            GetSQLToValid = GetSQLToValid & "Customer.cCusDefine4, Customer.cCusDefine5, Customer.cCusDefine6,"
            GetSQLToValid = GetSQLToValid & "Customer.cCusDefine7, Customer.cCusDefine8, Customer.cCusDefine9,"
            GetSQLToValid = GetSQLToValid & "Customer.cCusDefine10, Customer.cCusDefine11,"
            GetSQLToValid = GetSQLToValid & "Customer.cCusDefine12, Customer.cCusDefine13,"
            GetSQLToValid = GetSQLToValid & "Customer.cCusDefine14, Customer.cCusDefine15,"
            GetSQLToValid = GetSQLToValid & "Customer.cCusDefine16, isnull(cPayCode,N'') as cPayCode, isnull(cPayName,N'') as cPayName,"
            GetSQLToValid = GetSQLToValid & " isnull(ShippingChoice.cSCCode,N'') as cSCCode,isnull(ShippingChoice.cSCName,N'') as cSCName,N'' as ccontactcode,N'' as ccontactname FROM customer left join person on customer.cCusPPerson=person.cpersoncode left join department on customer.cCusDepart=department.cdepcode left join ShippingChoice on customer.cCusOType=ShippingChoice.cSCCode left join  PayCondition on customer.cCusPayCond=PayCondition.cPayCode"
            GetSQLToValid = GetSQLToValid & " left join  PayCondition on customer.cCusPayCond=PayCondition.cPayCode"
        Case Inventory
            ''GetSQLToValid = "SELECT * FROM Inventory "
            
'            blnWhAuth = CBool(clsSys.bAuth_Wh)
                If clsSys.bAuth_Wh Then
'                    clsAuth.Init clsSys.objlogin.UfDbName, clsSys.objlogin.cUserId, False, "SA"
                    strAuth = clsSys.clsAuth.GetAuthStringCode("WAREHOUSE", , "W")
                    If strAuth = "1=2" Then
                        strAuth = "and 1=2 "
                    ElseIf strAuth <> "" Then
                        strAuth = "and cwhcode in(" + strAuth + ")"
                    End If
                Else
                    strAuth = ""
                End If
            Dim strWareHouse As String
            Dim strVouchType As String

            If Not IsMissing(domHead) Then
                strTaxRate = GetHeadItemValue(domHead, "itaxrate")
                strVouchType = GetHeadItemValue(domHead, "cvouchtype")
            End If
            If Not (domBody Is Nothing Or GetHeadItemValue(domHead, "cbustype") = "直运销售" Or GetBodyItemValue(domBody, "itb", 0) = "退补") Then
                strWareHouse = Trim(GetBodyItemValue(domBody, "cwhcode", 0))
            Else
                strWareHouse = ""
            End If
            GetSQLToValid = "Select cVenAbbName,cInvCode,cInvAddCode,cInvName,cInvStd,cInvCCode,Inventory.cVenCode,inventory.cbarcode,case  inventory.cmassunit when N'1' then N'年' when N'2' then N'月' when N'3' then N'日' else N'' end as cmassunit,"
            GetSQLToValid = GetSQLToValid & "case when bAtoModel=1 then N'是' else N'否' end as batomodel,case when bPtoModel=1 then N'是' else N'否' end as bptomodel,N'否' as borderbom," + IIf(strVouchType = "", "case isnull(iRequireTrackStyle,0) when 1 then '5' when 2 then '1' when 3 then '4' when 0 then N'' end ", " case when isnull(inventory.cSRPolicy,N'PE')=N'PE' then '' else '0' end ") & " as idemandtype,N'' as cdemandcode,N'' as cdemandmemo,"
            GetSQLToValid = GetSQLToValid & "case when bInvModel=1 then N'是' else N'否' end as binvmodel,'' as icusbomid,'未选配' as cconfigstatus,'0' as natostatus,"
            GetSQLToValid = GetSQLToValid & "bService,bAccessary," & IIf(strVouchType = "28", " '0' as iTaxRateOLd ", " isnull(iTaxRate,'" & strTaxRate & "') as iTaxRate ") & ",iInvWeight,iVolume,bInvBatch,"
'            GetSQLToValid = GetSQLToValid & "bService,bAccessary," & IIf(strVouchType = "28", " isnull(iTaxRate,'" & strTaxRate & "') as iTaxRateOLd ", " isnull(iTaxRate,'" & strTaxRate & "') as iTaxRate ") & ",iInvWeight,iVolume,bInvBatch,"
            GetSQLToValid = GetSQLToValid & " cInvDefine1,cInvDefine2,cInvDefine3,cInvDefine4,cInvDefine5,cInvDefine6,cInvDefine7,"
            GetSQLToValid = GetSQLToValid & "cInvDefine8,cInvDefine9,cInvDefine10,cInvDefine11,cInvDefine12,cInvDefine13,cInvDefine14,cInvDefine15,cInvDefine16,"
            GetSQLToValid = GetSQLToValid & "bInvType , cGroupCode,cComunitCode, iMassDate, iGroupType, cSAComUnitCode,cAssComUnitCode,bInvquality,bTrack,iinvlscost,inventory.cSRPolicy,N'' as cbatchproperty1,N'' as cbatchproperty2,N'' as cbatchproperty3,N'' as cbatchproperty4,N'' as cbatchproperty5,N'' as cbatchproperty6,N'' as cbatchproperty7,N'' as cbatchproperty8,N'' as cbatchproperty9,N'' as cbatchproperty10,"
            GetSQLToValid = GetSQLToValid & " case when (isnull(bPropertyCheck,0)=1 and iGroupType <>2 )  then N'是' else N'否' end as bgsp ,N'' as dvdate,N'' as dmdate,iExpiratDateCalcu,N'' as dExpirationdate,N'' as cExpirationdate,N'' as cbatch,N''as ibatch,N'' as ccode,N'' as inufts,0 as iSettlePrice,N'' as cvmivenname,N'' as cvmivencode,isnull(bbatchcreate,0) as bbatchcreate"
'            For i = 1 To 10
'                GetSQLToValid = GetSQLToValid & ", N'' as cfree" & i
'            Next i
            For i = 1 To 10
                GetSQLToValid = GetSQLToValid & ", isnull(bfree" + CStr(i) + ",0) as bfree" & CStr(i) & ",isnull(bsalepricefree" + CStr(i) + ",0) as bsalepricefree" & CStr(i) & ", isnull(bBatchProperty" + CStr(i) + ",0) as bBatchProperty" & CStr(i)
            Next i
            '----support 73588 预完工日期=单据日期+存货累计提前期
            GetSQLToValid = GetSQLToValid & ",iAdvanceDate"
            '----------------------------------------------------
            GetSQLToValid = GetSQLToValid & IIf(strWareHouse <> "", "", ", case when bService=0 and bInvType=0 and '" + GetHeadItemValue(domHead, "cbustype") + "'<>'直运销售' and '" + GetBodyItemValue(domBody, "itb", 0) + "'<>'退补' " + strAuth + " then warehouse.cwhcode else N'' end as cwhcode,case when bService=0 and bInvType=0  and '" + GetHeadItemValue(domHead, "cbustype") + "'<>'直运销售' and '" + GetBodyItemValue(domBody, "itb", 0) + "'<>'退补' " + strAuth + " then warehouse.cwhname else N'' end as cwhname,case when bService=0 and bInvType=0 " + strAuth + " then isnull(warehouse.bproxywh,0) else N'' end as bproxywh ")
            GetSQLToValid = GetSQLToValid               '& ",Vendor.cvenname as cvmivenname,Vendor.cvencode as cvmivencode"
            GetSQLToValid = GetSQLToValid & " From inventory with (nolock) left join vendor with (nolock) on Inventory.cVenCode=Vendor.cVenCode left join inventory_sub on inventory.cinvcode=inventory_sub.cinvsubcode"
            GetSQLToValid = GetSQLToValid & IIf(strWareHouse <> "", "", " left join warehouse with (nolock) on inventory.cDefWarehouse= warehouse.cwhcode ")
        Case CurrentStock
            If Not IsMissing(domHead) Then strTaxRate = GetHeadItemValue(domHead, "itaxrate")
            GetSQLToValid = "Select cVenAbbName,cInvCode,cInvAddCode,cInvName,cInvStd,cInvCCode,Inventory.cVenCode,case  N'" + GetBodyItemValue(domBody, "cmassunit", 0) + "' when N'1' then N'年' when N'2' then N'月' when N'3' then N'日' else  N'" + GetBodyItemValue(domBody, "cmassunit", 0) + "' end as cmassunit,"
            GetSQLToValid = GetSQLToValid & "case when bAtoModel=1 then N'是' else N'否' end as batomodel,case when bPtoModel=1 then N'是' else N'否' end as bptomodel,"
            GetSQLToValid = GetSQLToValid & "case when bInvModel=1 then N'是' else N'否' end as binvmodel,"
            GetSQLToValid = GetSQLToValid & "bService,bAccessary," & IIf(strVouchType = "28", " '0' as iTaxRateOld ", "isnull(iTaxRate,'" & strTaxRate & "') as iTaxRate ") & ",iInvWeight,iVolume,bInvBatch,"
'            GetSQLToValid = GetSQLToValid & "bService,bAccessary," & IIf(strVouchType = "28", " isnull(iTaxRate,'" & strTaxRate & "') as iTaxRateOld ", "isnull(iTaxRate,'" & strTaxRate & "') as iTaxRate ") & ",iInvWeight,iVolume,bInvBatch,"
            GetSQLToValid = GetSQLToValid & " cInvDefine1,cInvDefine2,cInvDefine3,cInvDefine4,cInvDefine5,cInvDefine6,cInvDefine7,"
            GetSQLToValid = GetSQLToValid & "cInvDefine8,cInvDefine9,cInvDefine10,cInvDefine11,cInvDefine12,cInvDefine13,cInvDefine14,cInvDefine15,cInvDefine16,"
            GetSQLToValid = GetSQLToValid & "bInvType , cGroupCode,cComunitCode, iMassDate, iGroupType, cSAComUnitCode,cAssComUnitCode,bInvquality,bTrack,iinvlscost,0 as iSettlePrice,"
            GetSQLToValid = GetSQLToValid & " case when isnull(bPropertyCheck,0)=0 then N'否' else N'是' end as bgsp ,N''as ibatch,N'' as ccode,N'' as inufts,inventory.cSRPolicy"      ',Vendor.cvenname as cvmivenname,cvmivencode
            GetSQLToValid = GetSQLToValid & " From inventory with (nolock) left join vendor with (nolock) on Inventory.cVenCode=Vendor.cVenCode"
        Case warehouse
            ''GetSQLToValid = "SELECT * FROM WareHouse "
            GetSQLToValid = "SELECT cWhCode,cWhName,cDepCode,cWhAddress,cWhPhone,cWhPerson,cWhValueStyle,bWhPos,bProxyWh,N'' as cvmivencode,N'' as cvmivenname,"   'N'' as bProxyWh,"
            GetSQLToValid = GetSQLToValid & " iWhFundQuota,cMonth,cWhMemo,bFreeze,cBarCode "
            If Not Val(GetBodyItemValue(domBody, "iquantity", 0)) < 0 Then
                GetSQLToValid = GetSQLToValid & " ,N'' as dvdate,N'' as dmdate,N'' as cbatch,N'' as dExpirationdate,N'' as cExpirationdate "
            End If
            GetSQLToValid = GetSQLToValid & ",N''as ibatch,N'' as ccode,N'' as inufts  FROM WAREHOUSE"
        Case userdefine
            GetSQLToValid = "SELECT cAlias,cValue FROM UserDefine"
        Case Department
            GetSQLToValid = "SELECT cDepCode,cDepName,idepgrade FROM Department"
        Case Person
            If Not domHead Is Nothing Then
                If GetHeadItemValue(domHead, "cdepcode") = "" Then
                    GetSQLToValid = "SELECT person.cdepcode as oricdepcode,cPersonCode,cPersonName ,cPersonCode as cHandler ,isnull(department.cdepcode,N'') as cdepcode,isnull(cdepname,N'') as cdepname FROM Person left join department on person.cdepcode=department.cdepcode"
                Else
                    GetSQLToValid = "SELECT person.cdepcode  as oricdepcode,cPersonCode,cPersonName ,cPersonCode as cHandler FROM Person"
                End If
            Else
                GetSQLToValid = "SELECT person.cdepcode  as oricdepcode,cPersonCode,cPersonName ,cPersonCode as cHandler FROM Person"
            End If
            
        Case PayCondition
            GetSQLToValid = "SELECT cPayCode,cPayName FROM PayCondition"
        Case SaleType
          GetSQLToValid = "SELECT cSTCode,cSTName FROM SaleType"
        Case ForeignCurrency
            GetSQLToValid = "SELECT cexch_name,cexch_code FROM ForeignCurrency"
        Case ExpenseItem
            GetSQLToValid = "SELECT cExpCode, cExpName  FROM ExpenseItem"
        Case vendor '代管仓
            GetSQLToValid = "SELECT cvencode as cvmivencode, cvenname as cvmivenname FROM Vendor"
        Case Batch, VendorTrack, InvTrack  '批次
            strWhCode = GetBodyItemValue(domBody, "cwhcode", 0)
            strInvCode = GetBodyItemValue(domBody, "cinvcode", 0)
            strFree1 = GetBodyItemValue(domBody, "cfree1", 0)
            strFree2 = GetBodyItemValue(domBody, "cfree2", 0)
            strFree3 = GetBodyItemValue(domBody, "cfree3", 0)
            strFree4 = GetBodyItemValue(domBody, "cfree4", 0)
            strFree5 = GetBodyItemValue(domBody, "cfree5", 0)
            strFree6 = GetBodyItemValue(domBody, "cfree6", 0)
            strFree7 = GetBodyItemValue(domBody, "cfree7", 0)
            strFree8 = GetBodyItemValue(domBody, "cfree8", 0)
            strFree9 = GetBodyItemValue(domBody, "cfree9", 0)
            strFree10 = GetBodyItemValue(domBody, "cfree10", 0)
            strBatch = GetBodyItemValue(domBody, "cbatch", 0)
            cvmivencode = GetBodyItemValue(domBody, "cvmivencode", 0)
            siSOsID = ""
            If LCase(GetBodyItemValue(domBody, "cSRPolicy", 0)) = "lp" Then
                siSOsID = GetBodyItemValue(domBody, "cdemandid", 0)
'                If Trim(siSOsID) = "" Then siSOsID = "0"
'            Else
'                siSOsID = "0"
            End If
            
            If enmReferType = Batch Then
'                iQuantity,(CASE WHEN igrouptype<>2 THEN NULL ELSE  iNum END) AS iNum,
                
                GetSQLToValid = " SELECT currentstock.cBatch," & _
                         " aa_batchproperty.cbatchproperty1,aa_batchproperty.cbatchproperty2,aa_batchproperty.cbatchproperty3,aa_batchproperty.cbatchproperty4,aa_batchproperty.cbatchproperty5,ISNULL(aa_batchproperty.cbatchproperty6,N'') as cbatchproperty6," & _
                         " ISNULL(aa_batchproperty.cbatchproperty7,N'') as cbatchproperty7,ISNULL(aa_batchproperty.cbatchproperty8,N'') as cbatchproperty8,ISNULL(aa_batchproperty.cbatchproperty9,N'') as cbatchproperty9,aa_batchproperty.cbatchproperty10," & _
                         " ISNULL(currentstock.cFree1,N'') AS cFree1,ISNULL(currentstock.cFree2,N'') as cFree2,ISNULL(currentstock.cFree3,N'') AS cFree3,ISNULL(currentstock.cFree4,N'') as cFree4,ISNULL(currentstock.cFree5,N'') as cFree5,ISNULL(currentstock.cFree6,N'') as cFree6," & _
                         " ISNULL(currentstock.cFree7,N'') as cFree7,ISNULL(currentstock.cFree8,N'') as cFree8,ISNULL(currentstock.cFree9,N'') as cFree9,ISNULL(currentstock.cFree10,N'') as cFree10, dMdate, dVDate,iexpiratdatecalcu,dExpirationdate,cExpirationdate,N'' as ccode,N'' as ibatch " & _
                         IIf(strWhCode = "", ",CurrentStock.cwhcode,cWhName,case when isnull(CurrentStock.cvmivencode,N'')='' then 0 else 1 end as bproxywh ", "") & _
                         ",case  isnull(currentstock.cmassunit,0) when N'1' then N'年' when N'2' then N'月' when N'3' then N'日' else N'" + GetBodyItemValue(domBody, "cmassunit", 0) + "' end as cmassunit,case when isnull(currentstock.imassdate,0)=0 then N'' else CAST(currentstock.iMassDate AS CHAR(5)) end as imassdate,vendor.cvenname as cvmivenname,CurrentStock.cvmivencode" & _
                         " FROM CurrentStock INNER JOIN Inventory ON CurrentStock.cInvCode=Inventory.cInvCode left join vendor on vendor.cvencode=CurrentStock.cvmivencode " & _
                         " inner join WareHouse on CurrentStock.cwhcode=warehouse.cWhcode left join aa_batchproperty on currentstock.cinvcode=aa_batchproperty.cinvcode and isnull(currentstock.cbatch,N'')=isnull(aa_batchproperty.cbatch,N'') " & _
                         " and ISNULL(currentstock.cFree1,N'')= ISNULL(aa_batchproperty.cFree1,N'') and ISNULL(currentstock.cFree2,N'')= ISNULL(aa_batchproperty.cFree2,N'')" & _
                         " and ISNULL(currentstock.cFree3,N'')= ISNULL(aa_batchproperty.cFree3,N'') and ISNULL(currentstock.cFree4,N'')= ISNULL(aa_batchproperty.cFree4,N'')" & _
                         " and ISNULL(currentstock.cFree5,N'')= ISNULL(aa_batchproperty.cFree5,N'') and ISNULL(currentstock.cFree6,N'')= ISNULL(aa_batchproperty.cFree6,N'')" & _
                         " and ISNULL(currentstock.cFree7,N'')= ISNULL(aa_batchproperty.cFree7,N'') and ISNULL(currentstock.cFree8,N'')= ISNULL(aa_batchproperty.cFree8,N'')" & _
                         " and ISNULL(currentstock.cFree9,N'')= ISNULL(aa_batchproperty.cFree9,N'') and ISNULL(currentstock.cFree10,N'')= ISNULL(aa_batchproperty.cFree10,N'')" & _
                         " Where CurrentStock.cInvCode = N'" & strInvCode & "'" & _
                         IIf(strWhCode = "", "", " AND CurrentStock.cWhcode= N'" & strWhCode & "'") & _
                         " AND ISNULL(bStopFlag,0)=0 and isnull(bgspstop,0)=0 " & _
                         IIf(strFree1 = "", "", " AND currentstock.cFree1='" & strFree1 & "'") & IIf(strFree2 = "", "", " AND currentstock.cFree2='" & strFree2 & "'") & _
                         IIf(strFree3 = "", "", " AND currentstock.cFree3='" & strFree3 & "'") & IIf(strFree4 = "", "", " AND currentstock.cFree4='" & strFree4 & "'") & _
                         IIf(strFree5 = "", "", " AND currentstock.cFree5='" & strFree5 & "'") & IIf(strFree6 = "", "", " AND currentstock.cFree6='" & strFree6 & "'") & _
                         IIf(strFree7 = "", "", " AND currentstock.cFree7='" & strFree7 & "'") & IIf(strFree8 = "", "", " AND currentstock.cFree8='" & strFree8 & "'") & _
                         IIf(strFree9 = "", "", " AND currentstock.cFree9='" & strFree9 & "'") & IIf(strFree10 = "", "", " AND currentstock.cFree10='" & strFree10 & "'") & _
                         IIf(strBatch = "", "", " AND currentstock.cBatch=N'" & strBatch & "'") & IIf(cvmivencode = "", "", " AND CurrentStock.cvmivencode like N'%" & cvmivencode & "%'") & _
                         " and isnull(CurrentStock.isodid,N'')='" & siSOsID & "' And ((case when IsNull(isaconmode, 0) = 0 then " + IIf(clsSys.bBatOverStock = False, " ( " & clsSys.strBatchStockMethod & ") ", "1") + " else 1 end)>0 or (case when IsNull(isaconmode, 0) = 0 then " + IIf(clsSys.bBatOverStock = False, "(" & clsSys.strBatchStockNumMethod & ") ", "1") + " else 1 end)>0)"

            Else
                If enmReferType = InvTrack Then
                    strCode = GetBodyItemValue(domBody, "ccode", 0)
                    strTaxRate = GetBodyItemValue(domBody, "itaxrate", 0)
                    GetSQLToValid = " SELECT RdRecord.cCode,Warehouse.cWhName,Convert(char(10),RdRecord.dDate,121) as dDate,RdRecords.cBatch,RdRecords.iUnitCost,  " & _
                                         " Convert(char(10),dVDate,121) as dVDate,Convert(char(10),dMadeDate,121) AS dMDate,Convert(char(10),dExpirationdate,121) as dExpirationdate,cExpirationdate,iexpiratdatecalcu,(CASE WHEN ISNULL(RdRecords.iMassDate,0)=0 THEN N'' ELSE CAST(RdRecords.iMassDate AS CHAR(5)) END) AS iMassDate ," & _
                             " ISNULL(RdRecords.cFree1,N'') AS cFree1,ISNULL(RdRecords.cFree2,N'') as cFree2,ISNULL(RdRecords.cFree3,N'') AS cFree3," & _
                             " ISNULL(RdRecords.cFree4,N'') as cFree4,ISNULL(RdRecords.cFree5,N'') as cFree5,ISNULL(RdRecords.cFree6,N'') as cFree6," & _
                             " ISNULL(RdRecords.cFree7,N'') as cFree7,ISNULL(RdRecords.cFree8,N'') as cFree8,ISNULL(RdRecords.cFree9,N'') as cFree9,ISNULL(RdRecords.cFree10,N'') as cFree10,ISNULL(v.cVenAbbName,N'') AS cVenAbbName ,RdRecord.cWhCode,RdRecords.AutoID AS iBatch, " & _
                         " aa_batchproperty.cbatchproperty1,aa_batchproperty.cbatchproperty2,aa_batchproperty.cbatchproperty3,aa_batchproperty.cbatchproperty4,aa_batchproperty.cbatchproperty5,ISNULL(aa_batchproperty.cbatchproperty6,N'') as cbatchproperty6," & _
                         " ISNULL(aa_batchproperty.cbatchproperty7,N'') as cbatchproperty7,ISNULL(aa_batchproperty.cbatchproperty8,N'') as cbatchproperty8,ISNULL(aa_batchproperty.cbatchproperty9,N'') as cbatchproperty9,aa_batchproperty.cbatchproperty10," & _
                             " cAssComUnitCode,cComUnitCode,cSAComUnitCode,cgroupcode,igrouptype," & IIf(strVouchType = "28", " '0' as iTaxRateOld ", "isnull(Inventory.iTaxRate,'" & strTaxRate & "') as iTaxRate ") & _
                             ",case  isnull(rdrecords.cmassunit,0) when N'1' then N'年' when N'2' then N'月' when N'3' then N'日' else  '" + GetBodyItemValue(domBody, "cmassunit", 0) + "' end as cmassunit,Vendor.cvenname as cvmivenname,RdRecords.cvmivencode,bproxywh" & _
                             " FROM MainBatch INNER JOIN  RdRecords ON MainBatch.RdID = RdRecords.AutoID INNER JOIN RdRecord ON RdRecords.ID = RdRecord.ID INNER JOIN  Inventory ON " & _
                             " RdRecords.cInvCode = Inventory.cInvCode INNER JOIN Warehouse ON RdRecord.cWhCode = Warehouse.cWhCode LEFT OUTER JOIN Vendor ON RdRecords.cvmivencode = Vendor.cVenCode " & _
                             " left join vendor v on v.cvencode=(case when isnull(rdrecords.cbvencode,N'')=N'' then rdrecord.cVenCode else rdrecords.cbvencode end) " & _
                            " left join aa_batchproperty on RdRecords.cinvcode=aa_batchproperty.cinvcode  and isnull(RdRecords.cbatch,N'')=isnull(aa_batchproperty.cbatch,N'') " & _
                         " and ISNULL(RdRecords.cFree1,N'')= ISNULL(aa_batchproperty.cFree1,N'') and ISNULL(RdRecords.cFree2,N'')= ISNULL(aa_batchproperty.cFree2,N'')" & " and ISNULL(RdRecords.cFree3,N'')= ISNULL(aa_batchproperty.cFree3,N'') and ISNULL(RdRecords.cFree4,N'')= ISNULL(aa_batchproperty.cFree4,N'')" & _
                         " and ISNULL(RdRecords.cFree5,N'')= ISNULL(aa_batchproperty.cFree5,N'') and ISNULL(RdRecords.cFree6,N'')= ISNULL(aa_batchproperty.cFree6,N'')" & " and ISNULL(RdRecords.cFree7,N'')= ISNULL(aa_batchproperty.cFree7,N'') and ISNULL(RdRecords.cFree8,N'')= ISNULL(aa_batchproperty.cFree8,N'')" & _
                         " and ISNULL(RdRecords.cFree9,N'')= ISNULL(aa_batchproperty.cFree9,N'') and ISNULL(RdRecords.cFree10,N'')= ISNULL(aa_batchproperty.cFree10,N'')" & _
                         " Where  MainBatch.bIsNull = 0 " & IIf(strWhCode = "", "", " AND RdRecord.cWhcode=N'" & strWhCode & "'") & _
                             "    And    RdRecords.cInvCode =N'" & strInvCode & "'" & _
                             IIf(strFree1 = "", "", " AND RdRecords.cFree1=N'" & strFree1 & "'") & IIf(strFree2 = "", "", " AND RdRecords.cFree2=N'" & strFree2 & "'") & IIf(strFree3 = "", "", " AND RdRecords.cFree3=N'" & strFree3 & "'") & IIf(strFree4 = "", "", " AND RdRecords.cFree4=N'" & strFree4 & "'") & _
                             IIf(strFree5 = "", "", " AND RdRecords.cFree5=N'" & strFree5 & "'") & IIf(strFree6 = "", "", " AND RdRecords.cFree6=N'" & strFree6 & "'") & IIf(strFree7 = "", "", " AND RdRecords.cFree7=N'" & strFree7 & "'") & IIf(strFree8 = "", "", " AND RdRecords.cFree8=N'" & strFree8 & "'") & _
                             IIf(strFree9 = "", "", " AND RdRecords.cFree9=N'" & strFree9 & "'") & IIf(strFree10 = "", "", " AND RdRecords.cFree10=N'" & strFree10 & "'") & IIf(strBatch = "", "", " AND RdRecords.cBatch=N'" & strBatch & "'") & IIf(cvmivencode = "", "", " AND cvmivencode like N'%" & cvmivencode & "%'") & _
                             " and isnull(MainBatch.isodid,N'')='" & siSOsID & "'" & _
                             IIf(Not clsSys.bVerifyOut, "", " AND ISNULL(RdRecord.cHandler,N'')<>N''")

                    GetSQLToValid = GetSQLToValid & IIf(strCode = "", "", " AND RdRecord.cCode=N'" & strCode & "'")
                
                Else
                    strVenCode = GetBodyItemValue(domBody, "cvenabbname", 0)
                    GetSQLToValid = "SELECT  RdRecord.cVenCode,Vendor.cVenAbbName, case when btrack=1 then RdRecords.AutoID else null end as ibatch,case when btrack=1 then RdRecord.cCode else null end as ccode, " & _
                             " ltrim(str((Case When (inum Is Null ) Then 0 Else abs(inum) End)-(Case When(isoutnum Is Null) Then 0 Else isoutnum End),20," & CStr(clsSys.iNumBit) & ")) AS AccountNum," & _
                             " ltrim(str((Case When (iquantity Is Null) Then 0 Else abs(iquantity) End)-(Case When(isoutquantity Is Null) Then 0 Else isoutquantity End),20," & clsSys.SysInformation.cMQBit & ")) AS AccountQuantity, " & _
                             " ISNULL(RdRecords.cFree1,N'') AS cFree1, ISNULL(RdRecords.cFree2,N'') AS cFree2,ISNULL(RdRecords.cFree3,N'') AS cFree3, ISNULL(RdRecords.cFree4,N'') AS cFree4, ISNULL(RdRecords.cFree5,N'') AS cFree5, " & _
                             " ISNULL(RdRecords.cFree6,N'') AS cFree6, ISNULL(RdRecords.cFree7,N'') AS cFree7, ISNULL(RdRecords.cFree8,N'') AS cFree8, ISNULL(RdRecords.cFree9,N'') AS cFree9, ISNULL(RdRecords.cFree10,N'') AS cFree10, " & _
                         " aa_batchproperty.cbatchproperty1,aa_batchproperty.cbatchproperty2,aa_batchproperty.cbatchproperty3,aa_batchproperty.cbatchproperty4,aa_batchproperty.cbatchproperty5,ISNULL(aa_batchproperty.cbatchproperty6,N'') as cbatchproperty6," & _
                         " ISNULL(aa_batchproperty.cbatchproperty7,N'') as cbatchproperty7,ISNULL(aa_batchproperty.cbatchproperty8,N'') as cbatchproperty8,ISNULL(aa_batchproperty.cbatchproperty9,N'') as cbatchproperty9,aa_batchproperty.cbatchproperty10," & _
                             " RdRecord.dDate ,RdRecords.dMadeDate, RdRecords.dVDate,RdRecords.dExpirationdate,RdRecords.cExpirationdate,RdRecords.iexpiratdatecalcu,Vendor1.cvenname as cvmivenname,RdRecords.cvmivencode,case when isnull(RdRecords.cvmivencode,'')='' then 0 else 1 end as bproxywh" & _
                             " FROM RdRecord INNER JOIN RdRecords ON RdRecord.ID = RdRecords.ID LEFT OUTER JOIN  Vendor ON RdRecord.cVenCode = Vendor.cVenCode  LEFT OUTER JOIN  Vendor Vendor1 ON RdRecords.cvmivencode = Vendor.cVenCode left join inventory on rdrecords.cinvcode=inventory.cinvcode " & _
                            " left join aa_batchproperty on RdRecords.cinvcode=aa_batchproperty.cinvcode and isnull(RdRecords.cbatch,N'')=isnull(aa_batchproperty.cbatch,N'') " & _
                         " and ISNULL(RdRecords.cFree1,N'')= ISNULL(aa_batchproperty.cFree1,N'') and ISNULL(RdRecords.cFree2,N'')= ISNULL(aa_batchproperty.cFree2,N'')" & " and ISNULL(RdRecords.cFree3,N'')= ISNULL(aa_batchproperty.cFree3,N'') and ISNULL(RdRecords.cFree4,N'')= ISNULL(aa_batchproperty.cFree4,N'')" & _
                         " and ISNULL(RdRecords.cFree5,N'')= ISNULL(aa_batchproperty.cFree5,N'') and ISNULL(RdRecords.cFree6,N'')= ISNULL(aa_batchproperty.cFree6,N'')" & " and ISNULL(RdRecords.cFree7,N'')= ISNULL(aa_batchproperty.cFree7,N'') and ISNULL(RdRecords.cFree8,N'')= ISNULL(aa_batchproperty.cFree8,N'')" & _
                         " and ISNULL(RdRecords.cFree9,N'')= ISNULL(aa_batchproperty.cFree9,N'') and ISNULL(RdRecords.cFree10,N'')= ISNULL(aa_batchproperty.cFree10,N'')" & _
                             " Where 1=1 " & _
                             IIf(strWhCode = "", "", " AND cWhcode=N'" & strWhCode & "'") & _
                            "    And    rdrecords.cInvCode =N'" & strInvCode & "'" & _
                            IIf(strFree1 = "", "", " AND rdrecords.cFree1=N'" & strFree1 & "'") & IIf(strFree2 = "", "", " AND rdrecords.cFree2=N'" & strFree2 & "'") & _
                            IIf(strFree3 = "", "", " AND rdrecords.cFree3=N'" & strFree3 & "'") & IIf(strFree4 = "", "", " AND rdrecords.cFree4=N'" & strFree4 & "'") & _
                            IIf(strFree5 = "", "", " AND rdrecords.cFree5=N'" & strFree5 & "'") & IIf(strFree6 = "", "", " AND rdrecords.cFree6=N'" & strFree6 & "'") & _
                            IIf(strFree7 = "", "", " AND rdrecords.cFree7=N'" & strFree7 & "'") & IIf(strFree8 = "", "", " AND rdrecords.cFree8=N'" & strFree8 & "'") & _
                            IIf(strFree9 = "", "", " AND rdrecords.cFree9=N'" & strFree9 & "'") & IIf(strFree10 = "", "", " AND rdrecords.cFree10=N'" & strFree10 & "'") & _
                            IIf(strBatch = "", "", " AND rdrecords.cBatch=N'" & strBatch & "'") & IIf(cvmivencode = "", "", " AND cvmivencode like N'%" & cvmivencode & "%'") & _
                            " and isnull(RdRecords.isodid,N'')='" & siSOsID & "'" & _
                            IIf(strVenCode = "", "", " AND (RdRecord.cVenCode=N'" & strVenCode & "' OR Vendor.cVenAbbName=N'" & strVenCode & "')")
                
                End If
            
            
            End If
                     
            
    End Select
'    Set clsAuth = Nothing
End Function
Private Function VerifyHead(sKey As String, domHead As DOMDocument, Optional domBody As DOMDocument) As String
    Dim rstFile As New ADODB.Recordset
    Dim strAuth As String
    Dim strXMl As String
    Dim iCount As Integer
    Dim strSQL As String
    Dim bCus_Dep As Boolean
    Dim bCus_Per As Boolean
    Dim sDate As String, swhere As String
    Dim sKeyValue As String
    Dim ele As IXMLDOMElement, Attr As IXMLDOMAttribute 'zcy解决用户重新录入客户编码时代入相应业务员和部门,不考虑手工录入业务员和部门的情况
    Dim o_CRM As Object, bHavCus As Boolean
    Dim iExchRate As Double
    Dim sGatheringPlanFilter As String
     
     
    bCus_Dep = True
    bCus_Per = True
    
    m_UseMode = CS
    
    VerifyHead = ""
    strXMl = "<?xml version='1.0' encoding='GB2312'?>" & Chr(13) & "<head>" & Chr(13)
    sKeyValue = GetHeadItemValue(domHead, sKey)
    
    If LCase(Left(sKey, 7)) = "cdefine" Or LCase(sKey) = "all" Then     '表头自定义项
        
        If bNeedDefCheck(m_Conn, GetDefineID(sKey)) = False Then
            Exit Function
        End If
        
        Dim iIndex As Integer
        Dim iFromIndex As Integer
        Dim iToIndex As Integer
        If LCase(sKey) = "all" Then
            iFromIndex = 1
            iToIndex = 16
        Else
            iFromIndex = CInt(Right(sKey, 1))
            iToIndex = CInt(Right(sKey, 1))
        End If

        For iIndex = iFromIndex To iToIndex
            rstFile.Source = "select cvalue as " & "cDefine" & iIndex & ",calias from userdefine WHERE cID= N'" _
                            & GetDefineID("cDefine" & iIndex) & "' and (cvalue=N'" & GetHeadItemValue(domHead, "cDefine" & iIndex) & "' or calias=N'" & GetHeadItemValue(domHead, "cDefine" & iIndex) & "')"
            rstFile.Open , m_Conn

            If GetHeadItemValue(domHead, "cDefine" & iIndex) = "" Then  '(1)如果清空表头项
                If LCase(sKey) <> "all" Then
                    GoTo ClearDomHead
                End If
            Else
                If Not (rstFile.BOF And rstFile.EOF) Then                   '(2)如果合法
                    If LCase(sKey) <> "all" Then
                        GoTo SetDomHead
                    End If
                Else                                                        '(3)如不合法
                    If m_UseMode = CS Then
                        VerifyHead = GetStringPara("U8.SA.USSASERVER.clscommcheck.00388", iIndex) 'Para zh-CN：自定义项{0}录入不正确，请检查!
                        Exit Function
                    ElseIf m_UseMode = BS Then
                        If LCase(sKey) <> "all" Then
                            VerifyHead = GetStringPara("U8.SA.USSASERVER.clscommcheck.00388", iIndex) 'Para zh-CN：自定义项{0}录入不正确，请检查!
                            Exit Function
                        Else
                            VerifyHead = VerifyHead & GetStringPara("U8.SA.USSASERVER.clscommcheck.00388", iIndex) & Chr(13) 'Para zh-CN：自定义项{0}录入不正确，请检查!                        End If
                        End If
                    End If
                End If
            End If
            rstFile.Close
        Next iIndex
    End If
    ''商机等信息
    '商机 费用支出单
    
    '商机进程
    If LCase(sKey) = "icourseseq" Then
        If sKeyValue = "" Then
            strXMl = strXMl & "<R K='ioppcourseid' V=''/>" & Chr(13)
            strXMl = strXMl & "<R K='icourseseq' V=''/>" & Chr(13)
            strXMl = strXMl & "<R K='csalphasename' V=''/>" & Chr(13)
            strXMl = strXMl & "</head>"
            domHead.loadXML strXMl
            Exit Function
        End If
        Set o_CRM = CreateObject("UFCRMSRVSALE.clsOppPro")
        If Not o_CRM Is Nothing Then
            If rstFile.State = 1 Then rstFile.Close
            Set rstFile = o_CRM.GetRecordRst(clsSys.objlogin, GetHeadItemValue(domHead, "coppcode"), False, sKeyValue, False, domHead.selectSingleNode("//z:row"))
            If rstFile.EOF Then
                VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00394") 'zh-CN：商机进程不正确，请检查!
                Exit Function
            Else
                strXMl = strXMl & "<R K='ioppcourseid' V='" & rstFile("ioppcourseid") & "'/>" & Chr(13)
                strXMl = strXMl & "<R K='icourseseq' V='" & rstFile("icourseseq") & "'/>" & Chr(13)
                'strXMl = strXMl & "<R K='csalphasename' V='" & rstFile("iCourseSeq") & rstFile("csalphasename") & "'/>" & Chr(13)
                strXMl = strXMl & "<R K='csalphasename' V='" & rstFile("csalphasename") & "'/>" & Chr(13)
                strXMl = strXMl & "</head>"
                domHead.loadXML strXMl
                Exit Function
            End If
        Else
            VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00395") 'zh-CN：无法创建crm组件，请检查环境!
            Exit Function
        End If
        If rstFile.State = 1 Then rstFile.Close
        Set o_CRM = Nothing
    End If
    '活动编号
    If LCase(sKey) = "cactcode" And sKeyValue <> "" Then
        Set o_CRM = CreateObject("UFCRMSRVSALE.clsActivity")
        If Not o_CRM Is Nothing Then
            If rstFile.State = 1 Then rstFile.Close
            Set rstFile = o_CRM.GetRecordRst(clsSys.objlogin, False, sKeyValue, False, domHead.selectSingleNode("//z:row"), "W")
            If rstFile.EOF Then
                VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00396") 'zh-CN：活动编号不正确，请检查!
            Else
                Set ele = domHead.selectSingleNode("//z:row")
                '是否有客户信息
                If GetRstVal(rstFile, "ccuscode", False) <> "" Then
                    If LCase(GetElEAtrVal(ele, "ccuscode")) <> LCase(GetRstVal(rstFile, "ccuscode", False)) Then
                        ele.setAttribute "ccuscode", rstFile("ccuscode")
                        ele.setAttribute "ccusabbname", rstFile("ccusabbname")
                        strXMl = strXMl & "<R K='ccuscode' V='" & GetRstVal(rstFile, "ccuscode") & "'/>" & Chr(13)
                        strXMl = strXMl & "<R K='ccusabbname' V='" & GetRstVal(rstFile, "ccusabbname") & "'/>" & Chr(13)
                    Else
                        bHavCus = True
                    End If
                Else
                    If GetElEAtrVal(ele, "ccuscode") <> "" Then
                        ele.setAttribute "ccuscode", ""
                        ele.setAttribute "ccusabbname", ""
                        strXMl = strXMl & "<R K='ccuscode' V=''/>" & Chr(13)
                        strXMl = strXMl & "<R K='ccusabbname' V=''/>" & Chr(13)
                    Else
                        bHavCus = True
                    End If
                End If
                Dim strOppCode As String
                If GetRstVal(rstFile, "coppcode", False) <> "" Then
                    ele.setAttribute "coppcode", rstFile("coppcode")
                    ele.setAttribute "cactcode", rstFile("cactcode")
                    strOppCode = rstFile("coppcode")
                    strXMl = strXMl & "<R K='csocode' V=''/>" & Chr(13)
                    strXMl = strXMl & "<R K='cactcode' V='" & GetRstVal(rstFile, "cactcode") & "'/>" & Chr(13)
                    strXMl = strXMl & "<R K='coppcode' V='" & GetRstVal(rstFile, "coppcode") & "'/>" & Chr(13)
                    strXMl = strXMl & "<R K='ioppcourseid' V='" & GetRstVal(rstFile, "ioppcourseid") & "'/>" & Chr(13)
                    strXMl = strXMl & "<R K='icourseseq' V='" & GetRstVal(rstFile, "icourseseq") & "'/>" & Chr(13)
                    'strXMl = strXMl & "<R K='csalphasename' V='" & GetRstVal(rstFile, "iCourseSeq") & GetRstVal(rstFile, "csalphasename") & "'/>" & Chr(13)
                    strXMl = strXMl & "<R K='csalphasename' V='" & GetRstVal(rstFile, "csalphasename") & "'/>" & Chr(13)
                    '取得商机上的部门
                    Set o_CRM = CreateObject("UFCRMSRVSALE.clsOpportuntity")
                    If o_CRM Is Nothing Then
                        VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00397") 'zh-CN：无法创建对象UFCRMSRVSALE.clsOpportuntity!
                        Set rstFile = Nothing
                        Exit Function
                    End If
                    Set rstFile = o_CRM.GetRecordRst(clsSys.objlogin, False, strVouchType, strOppCode, False, domHead.selectSingleNode("//z:row"), "W")
                    If Not rstFile.EOF Then
                        strXMl = strXMl & "<R K='cdepcode' V='" & GetRstVal(rstFile, "cdepcode") & "'/>" & Chr(13)
                        strXMl = strXMl & "<R K='cdepname' V='" & GetRstVal(rstFile, "cdepname") & "'/>" & Chr(13)
                        strXMl = strXMl & "<R K='cpersoncode' V='" & GetRstVal(rstFile, "COPPPRINCIPAL") & "'/>" & Chr(13)
                        strXMl = strXMl & "<R K='cpersonname' V='" & GetRstVal(rstFile, "COPPPRINCIPALNAME") & "'/>" & Chr(13)
                    End If
                    
                Else
                    strXMl = strXMl & "<R K='coppcode' V=''/>" & Chr(13)
                    strXMl = strXMl & "<R K='ioppcourseid' V=''/>" & Chr(13)
                    strXMl = strXMl & "<R K='icourseseq' V=''/>" & Chr(13)
                    strXMl = strXMl & "<R K='csalphasename' V=''/>" & Chr(13)
                End If
                Set o_CRM = Nothing
                If Not bHavCus Then
                    sKey = "ccusabbname"
                    sKeyValue = GetHeadItemValue(domHead, sKey)
                Else
                    strXMl = strXMl & "</head>"
                    domHead.loadXML strXMl
                End If
            End If
        Else
            VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00398") 'zh-CN：无法创建crm组件，请检查环境!
        End If
        If rstFile.State = 1 Then rstFile.Close
        Set o_CRM = Nothing
    End If
    If LCase(sKey) = "ccontactname" Or LCase(sKey) = "ccontactcode" Then
        If sKeyValue = "" Then
            strXMl = strXMl & "<R K='ccontactname' V=''/>" & Chr(13)
            strXMl = strXMl & "<R K='ccontactcode' V=''/>" & Chr(13)
            strXMl = strXMl & "</head>"
            domHead.loadXML strXMl
        Else
            strSQL = Trim(GetHeadItemValue(domHead, "ccuscode"))
            If strSQL = "" Then
                VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00399") 'zh-CN：请先填写客户!
            Else
                strSQL = " crm_contact.ccuscode ='" + strSQL + "'"
                Set o_CRM = CreateObject("UFCRMSRVSALE.clsContact")
                If Not o_CRM Is Nothing Then
                    If rstFile.State = 1 Then rstFile.Close
                    Set rstFile = o_CRM.GetRecordRst(clsSys.objlogin, False, sKey, sKeyValue, False, , strSQL)
                    If rstFile.EOF Then
                        VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00400") 'zh-CN：联系人不正确，请检查!
                    Else
                        GoTo SetDomHead
                    End If
                Else
                    VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00401") 'zh-CN：无法创建crm组件，请检查环境!
                End If
            End If
            If rstFile.State = 1 Then rstFile.Close
            Set o_CRM = Nothing
        End If
    End If
    '录入订单号码
    If LCase(sKey) = "csocode" And sKeyValue <> "" And strVouchType = "99" Then
        strSQL = "select a.ccuscode,a.ccusabbname ,csocode,isnull(ccloser,N'') as ccloser, isnull(cverifier,N'') as cverifier from so_somain inner join customer a on so_somain.ccuscode=a.ccuscode where csocode=N'" + sKeyValue + "' "
        rstFile.Open strSQL, m_Conn, adOpenForwardOnly, adLockReadOnly
        If rstFile.EOF Then
            VerifyHead = GetStringPara("U8.SA.USSASERVER.clscommcheck.00402", sKeyValue) '或者未审核或者已经关闭!" 'Para zh-CN：订单{0}不存在,请检查!
        Else
            If rstFile("ccloser") <> "" Then
                VerifyHead = GetStringPara("U8.SA.USSASERVER.clscommcheck.00404", sKeyValue) 'Para zh-CN：订单{0}已经关闭,请检查!
            Else
                If rstFile("cverifier") <> "" Then
                    If GetHeadItemValue(domHead, "ccuscode") <> "" Then
                        If LCase(rstFile!cCusCode) <> LCase(GetHeadItemValue(domHead, "ccuscode")) Then
                            VerifyHead = GetStringPara("U8.SA.USSASERVER.clscommcheck.00406", sKeyValue) 'Para zh-CN：订单{0}的客户不是当前单据上的客户!
                        End If
                    Else
                        Set ele = domHead.selectSingleNode("//z:row")
                        ele.setAttribute "ccuscode", rstFile!cCusCode
                        ele.setAttribute "ccusabbname", rstFile!cCusAbbName
                        sKey = "ccusabbname"  ''获得客户相关信息
                    End If
                Else
                    VerifyHead = GetStringPara("U8.SA.USSASERVER.clscommcheck.00408", sKeyValue) 'Para zh-CN：订单{0}未审核,请检查!
                End If
            End If
        End If
        If rstFile.State = 1 Then rstFile.Close
    End If
    
    If LCase(sKey) = "ccusabbname" Or LCase(sKey) = "all" Then
        sKeyValue = GetHeadItemValue(domHead, sKey)
        Dim strPerAuth As String
        Dim strDepAuth As String
        
        If clsSys.bAuth_Cus Then
            strAuth = clsSys.clsAuth.GetAuthString("CUSTOMER", , "W")
        End If
'        If clsSys.bAuth_Per Then
'            strPerAuth = clsSys.clsAuth.GetAuthStringcode("PERSON", , "W")
'            If strPerAuth = "1=2" Then
'                strPerAuth = " isnull(person.cpersoncode,N'')=N'' "
'            ElseIf strPerAuth <> "" Then
'                strPerAuth = " (isnull(person.cpersoncode,N'')=N'' or isnull(person.cpersoncode,N'') in (" + strPerAuth + "))"
'            End If
'        End If
'        If clsSys.bAuth_Dep Then
'            strDepAuth = clsSys.clsAuth.GetAuthStringcode("DEPARTMENT", , "W")
'            If strDepAuth = "1=2" Then
'                strDepAuth = " isnull(department.cdepcode,N'')=N'' "
'            ElseIf strDepAuth <> "" Then
'                strDepAuth = " (isnull(department.cdepcode,N'')=N'' or isnull(department.cdepcode,N'') in (" + strDepAuth + "))"
'            End If
'        End If
        

        If LCase(sKey) = "ccusabbname" And strVouchType <> "92" Then
            Set ele = domHead.documentElement.selectSingleNode("//z:row")
        End If
        
        sDate = Left(GetHeadItemValue(domHead, "ddate"), 10)
        If Not (strVouchType = "92" Or strVouchType = "00") And sDate <> "" Then
            If Not (GetHeadItemValue(domHead, "breturnflag") = "1" Or LCase(GetHeadItemValue(domHead, "breturnflag")) = "true") Then
                If Not (GetHeadItemValue(domHead, "bfirst") = "1" Or LCase(GetHeadItemValue(domHead, "bfirst")) = "true") Then
                    swhere = " and isnull(Customer.dEndDate,'9999-12-31')>'" & sDate & "'"
                End If
            End If
        End If
        strSQL = " SELECT "
'        strSQL = strSQL & IIf(GetHeadItemValue(domHead, "cpersonName") <> "", "", "isnull(Person.cPersonCode,N'') as chandler,isnull(cPersonName,N'') as cPersonName,isnull(person.cPersonCode,N'') as cpersoncode,")
'        strSQL = strSQL & IIf(GetHeadItemValue(domHead, "cDepName") <> "", "", "isnull(department.cdepcode,N'') as cdepcode,isnull(cDepName,N'') as cDepName,")
'        'strSQL = strSQL & " ccusoaddress as cshipaddress,"
        strSQL = strSQL & "isnull(Person.cPersonCode,N'') as chandler,isnull(cPersonName,N'') as cPersonName,isnull(person.cPersonCode,N'') as cpersoncode,"
        strSQL = strSQL & "isnull(department.cdepcode,N'') as cdepcode,isnull(cDepName,N'') as cDepName,"

        strSQL = strSQL & " cusadd.cdeliveradd as cshipaddress ,cusadd.caddcode,cusadd.cdeliverunit,crm.ccontactname ,crm.cmobilephone,crm.cofficephone,"
        
        strSQL = strSQL & " Customer.cCusCode, Customer.cCusName, Customer.cCusAbbName, Customer.cCCCode, Customer.cDCCode, Customer.cTrade, Customer.cCusAddress,"
        strSQL = strSQL & " Customer.cCusPostCode, Customer.cCusRegCode, Customer.cCusBank, Customer.cCusAccount, Customer.dCusDevDate, Customer.cCusLPerson,"
        strSQL = strSQL & " Customer.cCusEmail,Customer.cCusPerson, Customer.cCusPhone, Customer.cCusFax,Customer.cCusBP, Customer.cCusHand, Customer.cCusPPerson,"
        strSQL = strSQL & " Customer.iCusDisRate, Customer.cCusCreGrade, Customer.iCusCreLine,Customer.iCusCreDate, Customer.cCusPayCond,Customer.cCusOAddress,"
        strSQL = strSQL & " Customer.cCusOType,Customer.cCusHeadCode, Customer.cCusWhCode,Customer.cCusDepart, Customer.iARMoney, Customer.dLastDate,Customer.iLastMoney,"
        strSQL = strSQL & " Customer.dLRDate, Customer.iLRMoney,Customer.dEndDate, Customer.iFrequency, Customer.cCusDefine1,Customer.cCusDefine2, Customer.cCusDefine3,"
        strSQL = strSQL & " Customer.iCostGrade,Customer.cCreatePerson, Customer.cModifyPerson,Customer.dModifyDate, Customer.cRelVendor, Customer.iId,Customer.cPriceGroup,Customer.cOfferGrade, Customer.iOfferRate,Customer.cCusDefine4,"
        strSQL = strSQL & " Customer.cCusDefine5, Customer.cCusDefine6,Customer.cCusDefine7, Customer.cCusDefine8, Customer.cCusDefine9,Customer.cCusDefine10,"
        strSQL = strSQL & " Customer.cCusDefine11,Customer.cCusDefine12, Customer.cCusDefine13,Customer.cCusDefine14, Customer.cCusDefine15,Customer.cCusDefine16 ,"
        strSQL = strSQL & " isnull(cPayCode,N'') as cPayCode, isnull(cPayName,N'') as cPayName, isnull(ShippingChoice.cSCCode,N'') as cSCCode,isnull(ShippingChoice.cSCName,N'') as cSCName,"
        strSQL = strSQL & " customer.cCusCreditCompany as ccreditcuscode,creditcustomer.ccusname as ccreditcusname, "
        '870 收付款协议
        strSQL = strSQL & " isnull(customer.bCreditDate,0) as bCreditDate,isnull(customer.cCusSAProtocol,N'') as cCusSAProtocol ,isnull(customer.ccusotherprotocol,N'') as ccusotherprotocol,"
        '870 收付款协议
        strSQL = strSQL & " isnull(customer.ccusexch_name,'') as ccusexch_name "
        strSQL = strSQL & " from customer  left join person on customer.cCusPPerson=person.cpersoncode left join department on customer.cCusDepart=department.cdepcode"
        strSQL = strSQL & " left join ShippingChoice on customer.cCusOType=ShippingChoice.cSCCode left join  PayCondition on customer.cCusPayCond=PayCondition.cPayCode"
        strSQL = strSQL & " left join cusdeliveradd cusadd on (cusadd.ccuscode = customer.ccuscode and cusadd.bdefault=1)  "
        strSQL = strSQL & " left join crm_contact crm on (crm.ccontactcode=cusadd.clinkperson) left join customer as creditcustomer on customer.cCusCreditCompany=creditcustomer.ccuscode"
        
        rstFile.Source = strSQL & " WHERE (customer.ccusAbbName=N'" & GetHeadItemValue(domHead, "cCusAbbName") & "' OR customer.cCusCode=N'" & GetHeadItemValue(domHead, "cCusAbbName") & "') " _
        & swhere & IIf(strAuth = "", "", " and (" & IIf(strAuth = "1=2", "1=2", " customer.iid in (" & strAuth & ")") & ")")
        '& IIf(strDepAuth = "", "", " and " & strDepAuth) & IIf(strPerAuth = "", "", " and " & strPerAuth)
        rstFile.Open , m_Conn
        rstFile.Filter = "ccusAbbName='" & GetHeadItemValue(domHead, "cCusAbbName") & "' and cCusCode='" & GetHeadItemValue(domHead, "cCusCode") & "'"
        If rstFile.EOF Then rstFile.Filter = ""
        
        If GetHeadItemValue(domHead, "cCusAbbName") = "" Then           '(1)如果清空表头项
            If LCase(sKey) <> "all" Then
                'VerifyHead = "客户不可以为空!"
                GoTo ClearDomHead
            End If
        Else
            If Not (rstFile.BOF And rstFile.EOF) Then                       '(2)如果合法
                If GetHeadItemValue(domHead, "cpersonName") = "" Or GetHeadItemValue(domHead, "cDepName") = "" Then
                    If clsSys.bAuth_Dep Or clsSys.bAuth_Per Then
                        ''初始化权限
'                        If Not clsSys.bAuth_Cus Then clsAuth.Init clsSys.objlogin.UfDbName, clsSys.objlogin.cUserId, False, "SA"
                        If GetHeadItemValue(domHead, "cpersonName") = "" Then
                            If rstFile("cPersonCode") <> "" Then
                                If Not clsSys.clsAuth.IsHoldAuth("Person", rstFile("cPersoncode")) Then
                                    bCus_Per = False
                                End If
                            End If
                        Else
                            bCus_Per = False
                        End If
                        If GetHeadItemValue(domHead, "cDepName") = "" Then
                            If rstFile("cDepCode") <> "" Then
                                If Not clsSys.clsAuth.IsHoldAuth("Department", rstFile("cDepcode")) Then
                                    bCus_Dep = False
                                End If
                            End If
                        Else
                            bCus_Dep = False
                        End If
                    End If
                Else
                    bCus_Per = False
                    bCus_Dep = False
                End If
                '需要取客户默认汇率，此信息在单据控制台中加入
                If GetHeadItemValue(domHead, "bcheckcusexchname") = "1" Then
                    If Not rstFile("ccusexch_name") = "" Then
                        If rstFile("ccusexch_name") <> GetHeadItemValue(domHead, "cexch_name") Then
                            iExchRate = clsSys.GetExchRate(rstFile("ccusexch_name"), IIf(GetHeadItemValue(domHead, "ddate") <> "", ConvDateForDOM(GetHeadItemValue(domHead, "ddate")), clsSys.objlogin.CurDate), clsSys.objlogin)
                            strXMl = strXMl & "<R K='cexch_name' V='" & GetRstVal(rstFile, "ccusexch_name") & "'/>" & Chr(13)
                            strXMl = strXMl & "<R K='bCountExch' V='1'/>" & Chr(13)
                            strXMl = strXMl & "<R K='iexchrate' V='" & iExchRate & "'/>" & Chr(13)
                        Else
                            strXMl = strXMl & "<R K='bCountExch' V='0'/>" & Chr(13)
                        End If
                    End If
                Else
                    strXMl = strXMl & "<R K='bCountExch' V='0'/>" & Chr(13)
                End If
'                If strVouchType = "95" Or strVouchType = "92" Then
'                    strXMl = strXMl & "<R K='ccusinvcode' V=''/>" & Chr(13)
'                    strXMl = strXMl & "<R K='ccusinvname' V=''/>" & Chr(13)
'                End If
                '870 收付款协议
                
                If LCase(sKey) <> "all" Then
                    GoTo SetDomHead
                End If
            Else                                                            '(3)如不合法
                If m_UseMode = CS Then
                    VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00410") 'zh-CN：客户录入不正确或者没有权限或者被停用，请检查!
                    GoTo ClearDomHead
                    Exit Function
                ElseIf m_UseMode = BS Then
                    If LCase(sKey) <> "all" Then
                        VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00411") 'zh-CN：客户录入不正确或者没有权限或者被停用，请检查!
                        Exit Function
                    Else
                        VerifyHead = VerifyHead & GetString("U8.SA.USSASERVER.clscommcheck.00412") & Chr(13) 'zh-CN：客户录入不正确或者没有权限或者被停用，请检查!
                    End If
                End If
            End If
        End If
        rstFile.Close
    End If

    If LCase(sKey) = "cdepname" Or LCase(sKey) = "all" Then                     '部门
        If clsSys.bAuth_Dep Then
'            clsAuth.Init clsSys.objlogin.UfDbName, clsSys.objlogin.cUserId, False, "SA"
            strAuth = clsSys.clsAuth.GetAuthStringCode("DEPARTMENT", , "W")
            If strAuth = "1=2" And GetHeadItemValue(domHead, "cdepname") <> "" Then
                VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00413") 'zh-CN：没有部门录入权限，请检查!
                Exit Function
            End If
        End If
        If strAuth = "1=2" Then
            rstFile.Source = GetSQLToValid(Department) & " WHERE 1=2"
        Else
            swhere = ""
            sDate = Left(GetHeadItemValue(domHead, "ddate"), 10)
            If sDate <> "" Then
                If Not (GetHeadItemValue(domHead, "bfirst") = "1" Or LCase(GetHeadItemValue(domHead, "bfirst")) = "true") Then
                    swhere = " and isnull(ddependdate,'9999-12-31')>'" & sDate & "'"
                End If
            End If
            rstFile.Source = GetSQLToValid(Department) & " WHERE bdepend=1  " & swhere & " and (cDepName=N'" & GetHeadItemValue(domHead, "cDepName") & "' OR cDepCode=N'" & GetHeadItemValue(domHead, "cDepName") & "')" & IIf(strAuth = "", "", " and cDepCode in (" & strAuth & ")")
        End If
        rstFile.Open , m_Conn
        rstFile.Filter = "cdepCode='" & GetHeadItemValue(domHead, "cdepcode") & "'"
        If rstFile.EOF Then rstFile.Filter = ""
        If GetHeadItemValue(domHead, "cDepName") = "" Then              '(1)如果清空表头项
            If LCase(sKey) <> "all" Then
                GoTo ClearDomHead
            End If
        Else
            
            If Not (rstFile.BOF And rstFile.EOF) Then                       '(2)如果合法
                If LCase(sKey) <> "all" Then
                    GoTo SetDomHead
                End If
            Else                                                            '(3)如不合法
                If m_UseMode = CS Then
                    VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00414") 'zh-CN：部门录入不正确或者没有权限，请检查!
                    GoTo ClearDomHead
                    Exit Function
                ElseIf m_UseMode = BS Then
                    If LCase(sKey) <> "all" Then
                        VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00415") 'zh-CN：部门录入不正确或者没有权限，请检查!
                        Exit Function
                    Else
                        VerifyHead = VerifyHead & GetString("U8.SA.USSASERVER.clscommcheck.00416") & Chr(13) 'zh-CN：部门录入不正确或者没有权限，请检查!
                    End If
                End If
            End If
        End If
        rstFile.Close
    End If

    
    
    If LCase(sKey) = "cpersonname" Or LCase(sKey) = "all" Then                  '职员
        If clsSys.bAuth_Per Then
'            clsAuth.Init clsSys.objlogin.UfDbName, clsSys.objlogin.cUserId, False, "SA"
            strAuth = clsSys.clsAuth.GetAuthStringCode("PERSON", , "W")
            If strAuth = "1=2" And GetHeadItemValue(domHead, "cPersonName") <> "" Then
                VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00417") 'zh-CN：没有权限录入业务员，请检查!
                Exit Function
            End If
        End If
        If strAuth = "1=2" Then
            rstFile.Source = GetSQLToValid(Person, , domHead) & " WHERE 1=2"
        Else
            rstFile.Source = GetSQLToValid(Person, , domHead) & " WHERE (Person.cPersonName=N'" _
                            & GetHeadItemValue(domHead, "cPersonName") & "' Or Person.cPersonCode=N'" _
                            & GetHeadItemValue(domHead, "cPersonName") & "')" & IIf(strAuth = "", "", " and cPersonCode in (" & strAuth & ")")
        End If
        rstFile.Open , m_Conn
        rstFile.Filter = "cPersonName='" & GetHeadItemValue(domHead, "cPersonName") & "' and cPersonCode='" & GetHeadItemValue(domHead, "cPersoncode") & "'"
        If rstFile.EOF Then rstFile.Filter = ""
        If GetHeadItemValue(domHead, "cPersonName") = "" Then           '(1)如果清空表头项
            If LCase(sKey) <> "all" Then
                GoTo ClearDomHead
            End If
        Else
            
            If Not (rstFile.BOF And rstFile.EOF) Then                       '(2)如果合法
                If GetHeadItemValue(domHead, "cDepName") = "" Then
                    If clsSys.bAuth_Dep And rstFile("cDepCode") <> "" Then
                        ''初始化权限
'                        If Not clsSys.bAuth_Per Then clsAuth.Init clsSys.objlogin.UfDbName, clsSys.objlogin.cUserId, False, "SA"
                        If rstFile("cDepCode") <> "" Then
                            If Not clsSys.clsAuth.IsHoldAuth("Department", rstFile("cDepcode")) Then
                                bCus_Dep = False
                            End If
                        End If
                    End If
                End If
                
                If LCase(sKey) <> "all" Then
                    GoTo SetDomHead
                End If
            Else                                                            '(3)如不合法
                If m_UseMode = CS Then
                    VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00418") 'zh-CN：业务员录入不正确或不属于相关部门或者没有权限，请检查!
                    GoTo ClearDomHead
                    Exit Function
                ElseIf m_UseMode = BS Then
                    If LCase(sKey) <> "all" Then
                        VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00419") 'zh-CN：业务员录入不正确或不属于相关部门或者没有权限，请检查!
                        Exit Function
                    Else
                        VerifyHead = VerifyHead & GetString("U8.SA.USSASERVER.clscommcheck.00420") & Chr(13) 'zh-CN：业务员录入不正确或不属于相关部门或者没有权限，请检查!
                    End If
                End If
            End If
        End If
        rstFile.Close
    End If

    
    If LCase(sKey) = "cexch_name" Or LCase(sKey) = "all" Then                  '外币
        'rstFile.Source = GetSQLToValid(ForeignCurrency) & " WHERE cexch_name='" & GetHeadItemValue(domHead, "cexch_name") & "' Or cexch_Code='" & GetHeadItemValue(domHead, "cexch_name") & "'"
        'rstFile.Source = "SELECT ForeignCurrency.cexch_name,ForeignCurrency.cexch_code,ISNULL(exch.nflat,1) AS iExchRate FROM ForeignCurrency LEFT JOIN exch ON ForeignCurrency.cexch_name=exch.cexch_name " _
                      & "WHERE ForeignCurrency.cexch_name='" & GetHeadItemValue(DOMHead, "cexch_name") & "' Or ForeignCurrency.cexch_Code='" & GetHeadItemValue(DOMHead, "cexch_name") & "'"
        rstFile.Source = "SELECT ForeignCurrency.cexch_name,ForeignCurrency.cexch_code FROM ForeignCurrency  " _
                      & "WHERE ForeignCurrency.cexch_name=N'" & GetHeadItemValue(domHead, "cexch_name") & "' Or ForeignCurrency.cexch_Code=N'" & GetHeadItemValue(domHead, "cexch_name") & "'"
        rstFile.Open , m_Conn
        
        If GetHeadItemValue(domHead, "cexch_name") = "" Then            '(1)如果清空表头项
            If LCase(sKey) <> "all" Then
                VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00421") 'zh-CN：币种不可以为空，请填写正确数据!
                Exit Function
                'GoTo ClearDomHead
            End If
        Else
            
            If Not (rstFile.BOF And rstFile.EOF) Then                       '(2)如果合法
                If LCase(sKey) <> "all" Then
                    
                    iExchRate = clsSys.GetExchRate(rstFile(0), IIf(GetHeadItemValue(domHead, "ddate") <> "", ConvDateForDOM(GetHeadItemValue(domHead, "ddate")), clsSys.objlogin.CurDate), clsSys.objlogin)
                    GoTo SetDomHead
                End If
            Else                                                            '(3)如不合法
                If m_UseMode = CS Then
                    VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00422") 'zh-CN：外币录入不正确，请检查!
                    GoTo ClearDomHead
                    Exit Function
                ElseIf m_UseMode = BS Then
                    If LCase(sKey) <> "all" Then
                        VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00423") 'zh-CN：外币录入不正确，请检查!
                        Exit Function
                    Else
                        VerifyHead = VerifyHead & GetString("U8.SA.USSASERVER.clscommcheck.00424") & Chr(13) 'zh-CN：外币录入不正确，请检查!
                    End If
                End If
            End If
        End If
        rstFile.Close
    End If
    
    ''付款条件
    If LCase(sKey) = "cpayname" Or LCase(sKey) = "all" Then
        rstFile.Source = "Select cpaycode,cpayname from PayCondition where cPayName=N'" & GetHeadItemValue(domHead, "cPayName") _
                        & "' Or cPayCode=N'" & GetHeadItemValue(domHead, "cPayName") & "'"
        rstFile.Open , m_Conn

        If GetHeadItemValue(domHead, "cPayName") = "" Then            '(1)如果清空表头项
            If LCase(sKey) <> "all" Then
                GoTo ClearDomHead
            End If
        Else
            
            If Not (rstFile.BOF And rstFile.EOF) Then                       '(2)如果合法
                If LCase(sKey) <> "all" Then
                    GoTo SetDomHead
                End If
            Else                                                            '(3)如不合法
                If m_UseMode = CS Then
                    VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00425") 'zh-CN：付款条件录入不正确，请检查!
                    GoTo ClearDomHead
                    Exit Function
                ElseIf m_UseMode = BS Then
                    If LCase(sKey) <> "all" Then
                        VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00426") 'zh-CN：付款条件录入不正确，请检查!
                        Exit Function
                    Else
                        VerifyHead = VerifyHead & GetString("U8.SA.USSASERVER.clscommcheck.00427") & Chr(13) 'zh-CN：付款条件录入不正确，请检查!
                    End If
                End If
            End If
        End If
        rstFile.Close
    End If
    
    If LCase(sKey) = "ccusinvcode" Or LCase(sKey) = "ccusinvname" Then
        If GetHeadItemValue(domHead, sKey) <> "" Then
            If GetcInvFromcCusinv(sKey, domHead, domHead, VerifyHead, False) = False Then
                
                Exit Function
            Else
                sKey = "cinvcode"
                sKeyValue = GetHeadItemValue(domHead, sKey)
                strXMl = strXMl & "<R K='ccusinvcode' V='" + FormatStrForDOM(GetHeadItemValue(domHead, "ccusinvcode")) + "'/>" & Chr(13)
                strXMl = strXMl & "<R K='ccusinvname' V='" + FormatStrForDOM(GetHeadItemValue(domHead, "ccusinvname")) + "'/>" & Chr(13)
            End If
        Else
            strXMl = strXMl & "<R K='ccusinvcode' V=''/>" & Chr(13)
            strXMl = strXMl & "<R K='ccusinvname' V=''/>" & Chr(13)
'            strXMl = strXMl & "<R K='cinvcode' V=''/>" & Chr(13)
'            strXMl = strXMl & "<R K='cinvname' V=''/>" & Chr(13)
            strXMl = strXMl & "</head>"
            Call domHead.loadXML(strXMl)
            Exit Function
        End If
    End If
    If LCase(sKey) = "cinvname" Or LCase(sKey) = "all" Or sKey = "cinvcode" Then                      '部门
        If clsSys.bAuth_Inv And Not clsSys.bAdmin Then
'            clsAuth.Init clsSys.objlogin.UfDbName, clsSys.objlogin.cUserId, False, "SA"
            strAuth = clsSys.clsAuth.GetAuthString("INVENTORY", , "W")
            
            If strAuth = "1=2" And GetHeadItemValue(domHead, sKey) <> "" Then
                VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00428") 'zh-CN：没有存货录入权限，请检查!
                If rstFile.State <> adStateClosed Then
                    rstFile.Close
                End If
                Set rstFile = Nothing
                Exit Function
            End If
        End If
        sDate = Left(GetHeadItemValue(domHead, "ddate"), 10)
        If sDate <> "" And Not (strVouchType = "92" Or strVouchType = "98" Or strVouchType = "99" Or strVouchType = "00") Then
            If Not (GetHeadItemValue(domHead, "breturnflag") = "1" Or LCase(GetHeadItemValue(domHead, "breturnflag")) = "true") Then
                If Not (GetHeadItemValue(domHead, "bfirst") = "1" Or LCase(GetHeadItemValue(domHead, "bfirst")) = "true") Then
                    swhere = " and isnull(dEDate,'9999-12-31')>'" & sDate & "'"
                End If
            End If
        End If
        If strAuth = "1=2" Then
            rstFile.Source = "select cinvname,cinvcode,cinvdefine1,cinvdefine2,cinvdefine3,cinvdefine4,cinvdefine5,cinvdefine6," _
                            & "cinvdefine7,cinvdefine8,cinvdefine9,cinvdefine10,cinvdefine11,cinvdefine12,cinvdefine13,cinvdefine14,cinvdefine15,cinvdefine16 from inventory WHERE 1=2"
        Else
            rstFile.Source = "select cinvname,cinvcode,cinvdefine1,cinvdefine2,cinvdefine3,cinvdefine4,cinvdefine5,cinvdefine6," _
                            & "cinvdefine7,cinvdefine8,cinvdefine9,cinvdefine10,cinvdefine11,cinvdefine12,cinvdefine13,cinvdefine14,cinvdefine15,cinvdefine16 from inventory WHERE ( cInvName=N'" & sKeyValue & "' OR cInvCode=N'" & sKeyValue & "' or Inventory.cBarCode=N'" & sKeyValue & "' )  " + swhere + IIf(strAuth = "", "", " and iid in (" & strAuth & ")")
        End If
        rstFile.Open , m_Conn
        rstFile.Filter = "cInvName='" & GetHeadItemValue(domHead, "cInvName") & "' and cInvCode='" & GetHeadItemValue(domHead, "cinvcode") & "'"
        If rstFile.EOF Then rstFile.Filter = ""
        If GetHeadItemValue(domHead, "cInvName") = "" Then              '(1)如果清空表头项
            If LCase(sKey) <> "all" Then
                strXMl = strXMl & "<R K='ccusinvcode' V=''/>" & Chr(13)
                strXMl = strXMl & "<R K='ccusinvname' V=''/>" & Chr(13)
                GoTo ClearDomHead
            End If
        Else
            
            If Not (rstFile.BOF And rstFile.EOF) Then                       '(2)如果合法
                If LCase(sKey) <> "all" Then
                    Call GetCusInvFromInv(strXMl, GetHeadItemValue(domHead, "ccuscode"), rstFile("cinvcode"))
                    GoTo SetDomHead
                End If
            Else                                                            '(3)如不合法
                If m_UseMode = CS Then
                    VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00429") 'zh-CN：存货录入不正确或者没有权限或者已经停止使用，请检查!
                    GoTo ClearDomHead
                    Exit Function
                ElseIf m_UseMode = BS Then
                    If LCase(sKey) <> "all" Then
                        VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00430") 'zh-CN：存货录入不正确，请检查!
                        Exit Function
                    Else
                        VerifyHead = VerifyHead & GetString("U8.SA.USSASERVER.clscommcheck.00431") & Chr(13) 'zh-CN：部门录入不正确，请检查!
                    End If
                End If
            End If
        End If
        rstFile.Close
    End If
    
    If LCase(sKey) = "cstname" Or LCase(sKey) = "all" Then                      '部门
        rstFile.Source = "Select cSTCode,cSTName from Saletype WHERE cSTName=N'" & GetHeadItemValue(domHead, "cSTName") & "' OR cSTCode=N'" & GetHeadItemValue(domHead, "cSTName") & "'"
        rstFile.Open , m_Conn
        
        If GetHeadItemValue(domHead, "cstname") = "" Then              '(1)如果清空表头项
            If LCase(sKey) <> "all" Then
                GoTo ClearDomHead
            End If
        Else
            
            If Not (rstFile.BOF And rstFile.EOF) Then                       '(2)如果合法
                If LCase(sKey) <> "all" Then
                    GoTo SetDomHead
                End If
            Else                                                            '(3)如不合法
                If m_UseMode = CS Then
                    VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00432") 'zh-CN：销售类型录入不正确，请检查!
                    GoTo ClearDomHead
                    Exit Function
                ElseIf m_UseMode = BS Then
                    If LCase(sKey) <> "all" Then
                        VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00433") 'zh-CN：销售类型录入不正确，请检查!
                        Exit Function
                    Else
                        VerifyHead = VerifyHead & GetString("U8.SA.USSASERVER.clscommcheck.00434") & Chr(13) 'zh-CN：销售类型录入不正确，请检查!
                    End If
                End If
            End If
        End If
        rstFile.Close
    End If
    
    If LCase(sKey) = "cscname" Or LCase(sKey) = "all" Then                      '部门
        rstFile.Source = "select cSCCode,cSCname from shippingchoice WHERE cScName=N'" & GetHeadItemValue(domHead, "cScName") _
                & "' OR cScCode=N'" & GetHeadItemValue(domHead, "cScName") & "'"
        rstFile.Open , m_Conn
        
        If GetHeadItemValue(domHead, "cscname") = "" Then              '(1)如果清空表头项
            If LCase(sKey) <> "all" Then
                GoTo ClearDomHead
            End If
        Else
            
            If Not (rstFile.BOF And rstFile.EOF) Then                       '(2)如果合法
                If LCase(sKey) <> "all" Then
                    GoTo SetDomHead
                End If
            Else                                                            '(3)如不合法
                If m_UseMode = CS Then
                    VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00435") 'zh-CN：发运方式录入不正确，请检查!
                    GoTo ClearDomHead
                    Exit Function
                ElseIf m_UseMode = BS Then
                    If LCase(sKey) <> "all" Then
                        VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00436") 'zh-CN：发运方式录入不正确，请检查!
                        Exit Function
                    Else
                        VerifyHead = VerifyHead & GetString("U8.SA.USSASERVER.clscommcheck.00437") & Chr(13) 'zh-CN：发运方式录入不正确，请检查!
                    End If
                End If
            End If
        End If
        rstFile.Close
    End If
    ''本单位开户银行
    If LCase(sKey) = "cbname" Or LCase(sKey) = "all" Then                      '部门
        rstFile.Source = "select cBCode, cBName, cBAccount, bBFlag from bank where bbflag=0 and( cBCode=N'" & GetHeadItemValue(domHead, "cbname") _
                    & "' OR cBName=N'" & GetHeadItemValue(domHead, "cBName") & "' or cBAccount = N'" & GetHeadItemValue(domHead, "cBName") & "')"
        rstFile.Open , m_Conn
        
        If GetHeadItemValue(domHead, "cbname") = "" Then              '(1)如果清空表头项
            If LCase(sKey) <> "all" Then
                GoTo ClearDomHead
            End If
        Else
            
            If Not (rstFile.BOF And rstFile.EOF) Then                       '(2)如果合法
                If LCase(sKey) <> "all" Then
                    GoTo SetDomHead
                End If
            Else                                                            '(3)如不合法
                If m_UseMode = CS Then
                    VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00438") 'zh-CN：开户银行录入不正确，请检查!
                    GoTo ClearDomHead
                    Exit Function
                ElseIf m_UseMode = BS Then
                    If LCase(sKey) <> "all" Then
                        VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00439") 'zh-CN：开户银行录入不正确，请检查!
                        Exit Function
                    Else
                        VerifyHead = VerifyHead & GetString("U8.SA.USSASERVER.clscommcheck.00440") & Chr(13) 'zh-CN：开户银行录入不正确，请检查!
                    End If
                End If
            End If
        End If
        rstFile.Close
    End If
    If LCase(sKey) = "ccusaccount" Then
        rstFile.Open "Select cAccountNum as ccusaccount,cBranch as ccusbank from CustomerBank where cAccountNum='" & GetHeadItemValue(domHead, "ccusaccount") & "' and ccuscode='" _
        & GetHeadItemValue(domHead, "ccuscode") & "'", m_Conn
        If Not (rstFile.BOF And rstFile.EOF) Then                       '(2)如果合法
            If LCase(sKey) <> "all" Then
                GoTo SetDomHead
            End If
        Else                                                            '(3)如不合法
            If m_UseMode = CS Then
                If GetHeadItemValue(domHead, "ccusaccount") <> "" Then
                    VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00438") 'zh-CN：开户银行录入不正确，请检查!
                End If
                
                
                GoTo ClearDomHead
                Exit Function
            ElseIf m_UseMode = BS Then
                If LCase(sKey) <> "all" Then
                    VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00439") 'zh-CN：开户银行录入不正确，请检查!
                    Exit Function
                Else
                    VerifyHead = VerifyHead & GetString("U8.SA.USSASERVER.clscommcheck.00440") & Chr(13) 'zh-CN：开户银行录入不正确，请检查!
                End If
            End If
        End If
        
    End If
    
    '客户收货地址 support 74783
    If LCase(sKey) = "cshipaddress" Or LCase(sKey) = "ccusoaddress" Or LCase(sKey) = "caddcode" Then
        Dim sTmpShipAddress As String
        Dim sTmpCuscode As String
        Dim strAddCode As String
        
        sKey = LCase(sKey)
        'sTmpShipAddress = Trim(GetHeadItemValue(DomHead, "cshipaddress"))
        sTmpShipAddress = Trim(GetHeadItemValue(domHead, sKey))
        sTmpCuscode = Trim(GetHeadItemValue(domHead, "ccuscode"))
        
        strSQL = "select cusadd.cdeliveradd as cshipaddress,cusadd.cdeliveradd as ccusoaddress,cusadd.caddcode,cusadd.cdeliverunit, " _
                    & " crm.ccontactname as ccontactname , crm.cmobilephone, crm.cofficephone " _
                    & " from cusdeliveradd cusadd left join crm_contact crm on (crm.ccontactcode=cusadd.clinkperson) " _
                    & " where cusadd.ccuscode = N'" & sTmpCuscode & "' and (cusadd.cdeliveradd=N'" & sTmpShipAddress & "' or caddcode='" & sTmpShipAddress & "')"
        rstFile.Open strSQL, m_Conn, adOpenStatic, adLockReadOnly, adCmdText
        If rstFile.RecordCount > 1 Then
            rstFile.Filter = "caddcode='" + Trim(GetHeadItemValue(domHead, "caddcode")) + "'"
            If rstFile.RecordCount = 0 Then
                rstFile.Filter = "caddcode='" + sTmpShipAddress + "'"
                If rstFile.RecordCount = 0 Then
                    rstFile.Filter = ""
                End If
            End If
        End If
        If Not (rstFile.EOF And rstFile.BOF) Then
            GoTo SetDomHead
        End If
        
    End If
    
'    Select Case LCase(sKey)
'        Case "cgatheringplan", "cgatheringplanname", "dcreditstart", "icreditdays", "dgatheringdate"
'            Select Case GetCardNumber(strVouchType, domHead)
'                Case "07", "13", "14", "15", "02"
'                    If GetHeadItemValue(domHead, "idisp") = 1 Then
'                        VerifyHead = "当前单据是参照发货单生成，不允许修改收付款协议相关信息。"
'                        Exit Function
'                    End If
'            End Select
'    End Select
    
    
    If LCase(sKey) = "cnextsystem" Then
        strSQL = "select * from  v_aa_enum where enumtype=N'SA.cnextsystem' and enumcode=N'" & sKeyValue & "'"
        rstFile.Open strSQL, m_Conn, adOpenStatic, adLockReadOnly, adCmdText
        If GetHeadItemValue(domHead, "cnextsystem") = "" Then
        Else
            If (rstFile.EOF And rstFile.BOF) Then
                VerifyHead = GetString("U8.SA.xsglsql_2.errmsg.VouchFlow") '"单据流向录入不正确，请检查!"
                Exit Function
            End If
        End If
    End If
    
    
'    Set clsAuth = Nothing
    Set o_CRM = Nothing
    Set rstFile = Nothing
    Exit Function
    
SetDomHead:
    If rstFile.State = 1 Then
    For iCount = 0 To rstFile.Fields.Count - 1
        If LCase(rstFile.Fields(iCount).Name) <> "ufts" Then
            'strXMl = strXMl & Replace("<R K='" & LCase(rstFile.Fields(iCount).Name) & "' V='" & rstFile.Fields(iCount) & "'/>" & Chr(13), "&", "&amp;")
            Select Case LCase(rstFile.Fields(iCount).Name)
                Case "cdepname", "cdepcode"
                    If bCus_Dep Then
                        If LCase(sKey) = "ccusabbname" And FormatStrForDOM(rstFile.Fields(iCount)) = "" And GetHeadItemValue(domHead, "cdepname") <> "" Then
                            'strXMl = strXMl & "<R K='" & LCase(rstFile.Fields(iCount).Name) & "' V='" & FormatStrForDOM(rstFile.Fields(iCount)) & "'/>" & Chr(13)
                        Else
                            strXMl = strXMl & "<R K='" & LCase(rstFile.Fields(iCount).Name) & "' V='" & FormatStrForDOM(rstFile.Fields(iCount)) & "'/>" & Chr(13)
                        End If
                    End If
                Case "cpersonname", "cpersoncode", "chandler"
                    If bCus_Per Then
                        If LCase(sKey) = "ccusabbname" And FormatStrForDOM(rstFile.Fields(iCount)) = "" And GetHeadItemValue(domHead, "cpersonname") <> "" Then
                        Else
                            strXMl = strXMl & "<R K='" & LCase(rstFile.Fields(iCount).Name) & "' V='" & FormatStrForDOM(rstFile.Fields(iCount)) & "'/>" & Chr(13)
                        End If
                    End If
                Case Else
                    If InStr(1, strXMl, "K='" & LCase(rstFile.Fields(iCount).Name) & "'") <= 0 Then
                        strXMl = strXMl & "<R K='" & LCase(rstFile.Fields(iCount).Name) & "' V='" & FormatStrForDOM(rstFile.Fields(iCount)) & "'/>" & Chr(13)
                    End If
            End Select
        End If
    Next iCount
    End If
    If LCase(sKey) = "cexch_name" Then
        strXMl = strXMl & "<R K='iexchrate' V='" & iExchRate & "'/>" & Chr(13)
    End If
    strXMl = strXMl & "</head>"
    If domHead.loadXML(strXMl) = False Then
        VerifyHead = GetString("U8.SA.USSASERVER.clscommcheck.00441") & domHead.parseError.reason 'zh-CN：生成dom出现错误:
    End If
    If LCase(sKey) = "ccusabbname" And Not IsMissing(domBody) And Not domBody Is Nothing Then
        Dim strCusInvCode As String
        Dim strCusInvName As String
        Dim strCusCode As String
        strCusCode = domHead.selectSingleNode("//R[@K='ccuscode']").Attributes.getNamedItem("V").nodeValue
        For Each ele In domBody.selectNodes("//z:row")
            If Not ele.Attributes.getNamedItem("cinvcode") Is Nothing Then
                GetCusInvCodeName strCusCode, ele.Attributes.getNamedItem("cinvcode").nodeValue, strCusInvCode, strCusInvName
                ele.setAttribute "ccusinvcode", strCusInvCode
                ele.setAttribute "ccusinvname", strCusInvName
            Else
                ele.setAttribute "ccusinvcode", ""
                ele.setAttribute "ccusinvname", ""
            End If
        Next ele
    End If
    If rstFile.State = 1 Then rstFile.Close
'    Set clsAuth = Nothing
    Set o_CRM = Nothing
    Set rstFile = Nothing
    Exit Function

ClearDomHead:
    
    For iCount = 0 To rstFile.Fields.Count - 1
        strXMl = strXMl & "<R K='" & LCase(rstFile.Fields(iCount).Name) & "' V=''/>" & Chr(13)
    Next iCount
    strXMl = strXMl & "</head>"
    domHead.loadXML strXMl
    rstFile.Close
'    Set clsAuth = Nothing
    Set o_CRM = Nothing
    Set rstFile = Nothing
    Exit Function
End Function

Public Function HeadCheck(sKey As String, domHead As DOMDocument, Optional domBody As DOMDocument = Nothing) As String
    On Error GoTo DoErr
    
    Dim ErrMsg As String
    
    
    HeadCheck = ""
    Select Case LCase(sKey)
        Case "all"
               HeadCheck = VerifyHead(sKey, domHead)
             If HeadCheck <> "" Then
                Exit Function
             End If

        'Case "cexch_name"
           
        'Case "iTaxRate"
        Case "coppcode"             '商机
'            If strVouchType = "99" Then '费用支出单
'                HeadCheck = VerifyHead(sKey, DomHead, domBody)
'            Else
            
            If GetVouchFromCRM(domHead, domBody, sKey, ErrMsg) = False Then
                HeadCheck = ErrMsg
            End If
'            End If
        
        Case Else
            HeadCheck = VerifyHead(sKey, domHead, domBody)
    End Select
    Exit Function
    
DoErr:
    HeadCheck = err.Description
    
End Function

''2002/06/18/gyp 修改
Public Function BodyCheck(sKey As String, domBody As DOMDocument, Optional domHead As DOMDocument = Nothing, Optional R As Long) As String
    
    Dim strXMLbody As String
    Dim rstFile As New Recordset, rstPrice As New ADODB.Recordset
    Dim iCount As Integer
    Dim keyCode As String
    Dim strHeadTable As String, strBodyTable As String, strSQL As String
    Dim iQuotedPrice As Double, iExchRate As Double, iTaxRate As Double, iQuantity As Double, fCusMinPrice As Double
    Dim fSalecost As Double         ''零售单价
    Dim SettleCost As Double
    Dim kl As Double
    Dim i As Integer
    Dim cFree(1 To 10) As String
    Dim strAuth As String
    Dim strUnitName As String, strUnitCode As String, strGroupCode As String, iChangRate As Double, strErrMsg As String
    Dim ele As IXMLDOMElement
    Dim sDate As String, swhere As String
    Dim domLine As New DOMDocument
    
    
    On Error GoTo DoErr
    
    rstFile.CursorLocation = adUseClient
    BodyCheck = ""
    m_UseMode = CS
    strXMLbody = "<?xml version='1.0' encoding='GB2312'?>" & Chr(13) & "<body sKey='" + sKey + "' >" & Chr(13)
    '存货参照
    sKey = LCase(sKey)
    sKeyOld = LCase(sKey)
    If LCase(sKey) <> "cinvname2" Then
        If LCase(sKey) <> "cinvcode2" Then
            keyCode = GetBodyItemValue(domBody, sKey, 0)
        Else
            keyCode = GetBodyItemValue(domBody, "cinvcode", 0)
        End If
    Else
        keyCode = GetBodyItemValue(domBody, "cinvname", 0)
    End If
    
    If LCase(Left(sKey, 7)) = "cdefine" Or LCase(Left(sKey, 5)) = "cfree" Then
        rstFile.Source = "select cvalue as " & sKey & ",calias,N'' as ibatch ,N'' as cbatch,N'' as ccode from userdefine WHERE cID= N'" _
                            & GetDefineID(sKey) & "' and cvalue=N'" & GetBodyItemValue(domBody, sKey, 0) & "'"
        rstFile.Open , m_Conn
        If GetBodyItemValue(domBody, sKey, 0) = "" Then
            Call ClearDombody(strXMLbody, rstFile, domBody)
            If rstFile.State <> adStateClosed Then
                rstFile.Close
            End If
            Set rstFile = Nothing
            Exit Function
        End If
        
        
        If rstFile.EOF Then
            BodyCheck = GetStringPara("U8.SA.USSASERVER.clscommcheck.00446", clsSys.getDefName(m_Conn, sKey)) 'Para zh-CN：{0} 录入不合法
        '    strXMLbody = strXMLbody & "<R K='" & rstFile.Fields(0).Name & "' V=''/>" & Chr(13)
            Call ClearDombody(strXMLbody, rstFile, domBody)
        Else
        
            If Left(LCase(sKey), 5) = "cfree" Then
                Set ele = domBody.selectNodes("//z:row").Item(0)
                If (clsSys.CostGetType = 0 Or clsSys.CostGetType = 2 Or clsSys.CostGetType = 7) And strVouchType <> "00" And strVouchType <> "98" And strVouchType <> "99" Then
                    If GetBodyItemValue(domBody, "cinvcode", 0) <> "" Then
                        If CBool(GetBodyItemValue(domBody, "bsalepricefree1", 0)) Or CBool(GetBodyItemValue(domBody, "bsalepricefree2", 0)) Or CBool(GetBodyItemValue(domBody, "bsalepricefree3", 0)) _
                            Or CBool(GetBodyItemValue(domBody, "bsalepricefree4", 0)) Or CBool(GetBodyItemValue(domBody, "bsalepricefree5", 0)) Or CBool(GetBodyItemValue(domBody, "bsalepricefree6", 0)) _
                            Or CBool(GetBodyItemValue(domBody, "bsalepricefree7", 0)) Or CBool(GetBodyItemValue(domBody, "bsalepricefree8", 0)) Or CBool(GetBodyItemValue(domBody, "bsalepricefree9", 0)) _
                            Or CBool(GetBodyItemValue(domBody, "bsalepricefree10", 0)) Then
                        
                            ele.setAttribute sKey, rstFile(sKey)
                            BodyCheck = GetPrice(domHead, domBody)
                            If BodyCheck <> "" Then GoTo DOExit
                        End If
                    End If
                End If
                If GetBodyItemValue(domBody, "cbatch", 0) <> "" Then
                    GetBatchProperty m_Conn, ele
                End If
            End If
            
'            Call SetDombody(sKey, strXMLbody, rstFile, domBody, domHead, R)
        End If
        

        
        GoTo DOExit
    End If
    
    Select Case LCase(sKey)
        Case "cdemandcode"
            swhere = ""
            sDate = Left(GetHeadItemValue(domHead, "ddate"), 10)
            If Not (strVouchType = "92" Or strVouchType = "00") And sDate <> "" Then
                If Not (GetHeadItemValue(domHead, "breturnflag") = "1" Or LCase(GetHeadItemValue(domHead, "breturnflag")) = "true") Then
                    If Not (GetHeadItemValue(domHead, "bfirst") = "1" Or LCase(GetHeadItemValue(domHead, "bfirst")) = "true") Then
                        swhere = " and isnull(dstopdate,'9999-12-31')>'" & sDate & "'"
                    End If
                End If
            End If
            keyCode = GetBodyItemValue(domBody, sKey, 0)
            If rstFile.State = 1 Then rstFile.Close
            rstFile.Source = ConvertSQLString(GetSQLToValid(demandclass, domBody, domHead) & IIf(keyCode = "", " where 1=2", " where crclasscode='" + keyCode + "'" & swhere))
            rstFile.Open , m_Conn
            If rstFile.EOF Then           '(1)如果清空表头项
                If keyCode = "" Then
                    If LCase(sKey) <> "all" Then
                        Call ClearDombody(strXMLbody, rstFile, domBody)
                    End If
                Else
                    BodyCheck = GetString("U8.SA.USSASERVER.clssofunc.demandcodevalid") 'zh-CN：费用项目录入不正确，请检查!
'                    Call ClearDombody(strXMLbody, rstFile, domBody)
                End If
            Else
                Call SetDombody(sKey, strXMLbody, rstFile, domBody, domHead, R)
            End If
        Case "bqaneedcheck"
            If GetBodyItemValue(domBody, "cinvcode", 0) <> "" Then
                If keyCode = "是" Then
                    strSQL = "select cinvcode from inventory where cinvcode=N'" + GetBodyItemValue(domBody, "cinvcode", 0) + "' and isnull(bPropertyCheck,0)=1"
                    rstFile.Open strSQL, m_Conn
                    If rstFile.EOF Then
                        BodyCheck = GetString("U8.SA.USSASERVER.clscommcheck.00448") 'zh-CN：存货不存在或者非质检存货!
                        GoTo DOExit
                    End If
                End If
            Else
                BodyCheck = GetString("U8.SA.USSASERVER.clscommcheck.00449") 'zh-CN：请先录入存货!
                GoTo DOExit
            End If
        Case "citemcode"
            If GetBodyItemValue(domBody, "citem_class", 0) = "" And keyCode <> "" Then
                BodyCheck = GetString("U8.SA.USSASERVER.clscommcheck.00450") 'zh-CN：请先填写项目大类
                GoTo DOExit
            End If
            strSQL = "select ctable from fitem where citem_class=N'" & GetBodyItemValue(domBody, "citem_class", 0) & "'"
            rstFile.Open strSQL, m_Conn
            If Not rstFile.EOF Then
                If rstFile(0) = "inventory" Then
                    strSQL = "select cInvCode as citemcode,cInvName as citemname from inventory where (cInvCode=N'" & keyCode & "' or cInvName=N'" & keyCode & "' ) "
                Else
                    strSQL = "select citemcode,citemname,bclose from " & rstFile(0) & " where (citemcode=N'" & keyCode & "' or citemname=N'" & keyCode & "' ) and bclose =0"
                End If
                rstFile.Close
                rstFile.Open strSQL, m_Conn
            Else
                BodyCheck = GetString("U8.SA.USSASERVER.clscommcheck.00451") 'zh-CN：项目大类有错误,请检查!
                Call ClearDombody(strXMLbody, rstFile, domBody)
                GoTo DOExit
            End If
            If keyCode = "" Then           '(1)如果清空表头项
                    If LCase(sKey) <> "all" Then
                        Call ClearDombody(strXMLbody, rstFile, domBody)
                    End If
                Else
                    If Not (rstFile.BOF And rstFile.EOF) Then                       '(2)如果合法
                        If LCase(sKey) <> "all" Then
                            Call SetDombody(sKey, strXMLbody, rstFile, domBody, domHead, R)
                        End If
                    Else                                                            '(3)如不合法
                        If m_UseMode = CS Then
                            BodyCheck = GetString("U8.SA.USSASERVER.clscommcheck.00452") 'zh-CN：项目录入不正确，请检查!
                            Call ClearDombody(strXMLbody, rstFile, domBody)
                            GoTo DOExit
                        ElseIf m_UseMode = BS Then
                            If LCase(sKey) <> "all" Then
                                BodyCheck = GetString("U8.SA.USSASERVER.clscommcheck.00453") 'zh-CN：项目录入不正确，请检查!
                                Call ClearDombody(strXMLbody, rstFile, domBody)
                                GoTo DOExit
                            Else
                                BodyCheck = BodyCheck & GetString("U8.SA.USSASERVER.clscommcheck.00454") & Chr(13) 'zh-CN：项目录入不正确，请检查!
                                Call ClearDombody(strXMLbody, rstFile, domBody)
                            End If
                        End If
                    End If
                End If
                If rstFile.State <> adStateClosed Then
                    rstFile.Close
                End If
        Case "citem_class"
            strSQL = "select citem_class, citem_name as citem_cname,N'' as citemcode,N''as citemname from fitem where citem_class=N'" & keyCode & "' or citem_name=N'" & keyCode & "'"
            rstFile.Open strSQL, m_Conn
            If keyCode = "" Then           '(1)如果清空表头项
                    If LCase(sKey) <> "all" Then
                    'citemcode,citemname
                        'strXMLbody = strXMLbody & "<R K='citemcode' V=''/>" & Chr(13)
                        'strXMLbody = strXMLbody & "<R K='citemname' V=''/>" & Chr(13)
                        Call ClearDombody(strXMLbody, rstFile, domBody)
                    End If
                Else
                    If Not (rstFile.BOF And rstFile.EOF) Then                       '(2)如果合法
                        If LCase(sKey) <> "all" Then
                            Call SetDombody(sKey, strXMLbody, rstFile, domBody, domHead, R)
                        End If
                    Else                                                            '(3)如不合法
                        If m_UseMode = CS Then
                            BodyCheck = GetString("U8.SA.USSASERVER.clscommcheck.00455") 'zh-CN：项目大类录入不正确，请检查!
                            Call ClearDombody(strXMLbody, rstFile, domBody)
                            GoTo DOExit
                        ElseIf m_UseMode = BS Then
                            If LCase(sKey) <> "all" Then
                                BodyCheck = GetString("U8.SA.USSASERVER.clscommcheck.00456") 'zh-CN：项目大类录入不正确，请检查!
                                GoTo DOExit
                            Else
                                BodyCheck = BodyCheck & GetString("U8.SA.USSASERVER.clscommcheck.00457") & Chr(13) 'zh-CN：项目大类录入不正确，请检查!
                            End If
                        End If
                    End If
                End If
                If rstFile.State <> adStateClosed Then
                    rstFile.Close
                End If
        Case "cbatch", "cvenabbname", "ccode"
                'If LCase(sKey) = "cbatch" And Val(GetBodyItemValue(domBody, "iquantity", 0)) < 0 Then Exit Function
                
                If LCase(sKey) = "ccode" And keyCode <> "" Then
                    If Val(GetBodyItemValue(domBody, "iQuantity", 0)) < 0 Then
                        BodyCheck = GetString("U8.SA.USSASERVER.clscommcheck.00458") 'zh-CN：数量小于0， 不可以输入入库单据号!
                        Set rstFile = Nothing
                        Exit Function
                    End If
                End If
                If LCase(sKey) = "ccode" Or LCase(sKey) = "cbatch" Then
                    If GetBodyItemValue(domBody, "itb", 0) = "退补" Or GetBodyItemValue(domBody, "itb", 0) = "1" Then
                        Set rstFile = Nothing
                        Exit Function
                    End If
                End If
                If rstFile.State = 1 Then rstFile.Close
                rstFile.Source = ConvertSQLString((GetSQLToValid(IIf(LCase(sKey) = "cbatch", Batch, IIf(LCase(sKey) = "ccode", InvTrack, VendorTrack)), domBody, domHead)))
                rstFile.Open , m_Conn
                If sKey = "ccode" Then
                    rstFile.Filter = "ibatch =" & Val(GetBodyItemValue(domBody, "ibatch", 0))
                    If rstFile.EOF Then
                        rstFile.Filter = ""
                    End If
                    
                End If
                
                If keyCode = "" Then           '(1)如果清空表头项
                    If LCase(sKey) <> "all" Then
                        Call ClearDombody(strXMLbody, rstFile, domBody)
                    End If
                Else
                    If Not (rstFile.BOF And rstFile.EOF) Then                       '(2)如果合法
                        If LCase(sKey) <> "all" Then
                            If LCase(sKey) = "" Or LCase(sKey) = "cbatch" Then
                                keyCode = rstFile("cBatch")
                            Else
                                If LCase(sKey) = "ccode" Then
                                    keyCode = rstFile("ccode")
                                Else
                                    keyCode = rstFile("cvenabbname")
                                End If
                                    
                            End If
                                
                                
                                ''计算价格zcy跟踪存货取得入库价
                            'If LCase(GetBodyItemValue(domBody, "cfree1", 0)) <> LCase(rstFile("cfree1")) Or LCase(GetBodyItemValue(domBody, "cfree2", 0)) <> LCase(rstFile("cfree2"))
                            '设置记录体
                            
                            If LCase(GetBodyItemValue(domBody, "cfree1", 0)) <> LCase(rstFile("cfree1")) _
                                Or LCase(GetBodyItemValue(domBody, "cfree2", 0)) <> LCase(rstFile("cfree2")) _
                                Or LCase(GetBodyItemValue(domBody, "cfree3", 0)) <> LCase(rstFile("cfree3")) _
                                Or LCase(GetBodyItemValue(domBody, "cfree4", 0)) <> LCase(rstFile("cfree4")) _
                                Or LCase(GetBodyItemValue(domBody, "cfree5", 0)) <> LCase(rstFile("cfree5")) _
                                Or LCase(GetBodyItemValue(domBody, "cfree6", 0)) <> LCase(rstFile("cfree6")) _
                                Or LCase(GetBodyItemValue(domBody, "cfree7", 0)) <> LCase(rstFile("cfree7")) _
                                Or LCase(GetBodyItemValue(domBody, "cfree8", 0)) <> LCase(rstFile("cfree8")) _
                                Or LCase(GetBodyItemValue(domBody, "cfree9", 0)) <> LCase(rstFile("cfree9")) _
                                Or LCase(GetBodyItemValue(domBody, "cfree10", 0)) <> LCase(rstFile("cfree10")) _
                                Or (sKey = "ccode" And clsSys.CostGetType = 1) Then
                                Set ele = domBody.selectSingleNode("//z:row")
                                ele.setAttribute "ibatch", rstFile("ibatch")
                                ele.setAttribute "cfree1", rstFile("cfree1")
                                ele.setAttribute "cfree2", rstFile("cfree2")
                                ele.setAttribute "cfree3", rstFile("cfree3")
                                ele.setAttribute "cfree4", rstFile("cfree4")
                                ele.setAttribute "cfree5", rstFile("cfree5")
                                ele.setAttribute "cfree6", rstFile("cfree6")
                                ele.setAttribute "cfree7", rstFile("cfree7")
                                ele.setAttribute "cfree8", rstFile("cfree8")
                                ele.setAttribute "cfree9", rstFile("cfree9")
                                ele.setAttribute "cfree10", rstFile("cfree10")
                                Set domLine = domBody.cloneNode(True)
                                Call SetDombody(sKey, strXMLbody, rstFile, domBody, domHead, R)
                                Set domBody = domLine.cloneNode(True)
                                SetDomBodyItemValues domBody, strXMLbody
                                BodyCheck = GetPrice(domHead, domBody)
                                If BodyCheck <> "" Then
                                    GoTo DOExit
                                End If
                            Else
                                Call SetDombody(sKey, strXMLbody, rstFile, domBody, domHead, R)
                            End If
                        End If
                    ElseIf LCase(sKey) = "cbatch" And Val(GetBodyItemValue(domBody, "iquantity", 0)) < 0 Then
                        If rstFile.State <> adStateClosed Then
                            rstFile.Close
                        End If
                        strXMLbody = strXMLbody & "<R K='cbatch' V='" & Trim(GetBodyItemValue(domBody, "cbatch", 0)) & "'/>" & Chr(13)
                        strXMLbody = strXMLbody & "</body>"
                        domBody.loadXML strXMLbody
                        Set rstFile = Nothing
                        Exit Function
                    Else
                        BodyCheck = GetStringPara("U8.SA.USSASERVER.clscommcheck.00463", IIf(LCase(sKey) = "cbatch", GetString("U8.SA.USSASERVER.clscommcheck.00460"), IIf(LCase(sKey) = "ccode", GetString("U8.SA.USSASERVER.clscommcheck.00461"), GetString("U8.SA.USSASERVER.clscommcheck.00462"))))  'zh-CN：批次 'zh-CN：入库单号 'zh-CN：供应商 'Para zh-CN：{0}录入不正确，请检查!
                        If rstFile.State <> adStateClosed Then
                            rstFile.Close
                        End If
                        Set rstFile = Nothing
                        Exit Function
                    End If
                End If
                
                If rstFile.State <> adStateClosed Then
                    rstFile.Close
                End If
    
    
        Case "cinvname", "cinvcode", "cinvname2", "currentstock", "cinvcode2", "ccusinvname", "ccusinvcode"
'                keyCode = GetBodyItemValue(DOMBody, sKey, 0)
                Dim strWhCode As String 'zcy仓库编号
                If clsSys.bAuth_Inv Then
'                    clsAuth.Init clsSys.objlogin.UfDbName, clsSys.objlogin.cUserId, False, "SA"
                    strAuth = clsSys.clsAuth.GetAuthString("INVENTORY", , "W")
                    
                    If strAuth = "1=2" Then
                        BodyCheck = GetString("U8.SA.USSASERVER.clscommcheck.00464") 'zh-CN：没有存货录入权限，请检查!
                        If rstFile.State <> adStateClosed Then
                            rstFile.Close
                        End If
                        Set rstFile = Nothing
                        Exit Function
                    End If
                End If
                If rstFile.State = 1 Then rstFile.Close
                If sKey = "ccusinvname" Or sKey = "ccusinvcode" Then
                    If GetcInvFromcCusinv(sKey, domHead, domBody, strErrMsg) = False Then
                        BodyCheck = strErrMsg
                        If rstFile.State <> adStateClosed Then
                            rstFile.Close
                        End If
                        domBody.loadXML strXMLbody + "<R K='ccusinvcode' V=''/>" & Chr(13) + "<R K='ccusinvname' V=''/>" & Chr(13) + "</body>"
                        Set rstFile = Nothing
                        Exit Function
                    Else
                        If LCase(sKey) <> "cinvname2" Then
                            If LCase(sKey) <> "cinvcode2" Then
                                keyCode = GetBodyItemValue(domBody, sKey, 0)
                            Else
                                keyCode = GetBodyItemValue(domBody, "cinvcode", 0)
                            End If
                        Else
                            keyCode = GetBodyItemValue(domBody, "cinvname", 0)
                        End If
                    End If
                End If

                
                ''存货停用 与日期 单据红兰相关
                sDate = Left(GetHeadItemValue(domHead, "ddate"), 10)
                If sDate <> "" And Not (strVouchType = "92" Or strVouchType = "98" Or strVouchType = "99" Or strVouchType = "00") Then
                    If Not (GetHeadItemValue(domHead, "breturnflag") = "1" Or LCase(GetHeadItemValue(domHead, "breturnflag")) = "true") Then
                        If Not (GetHeadItemValue(domHead, "bfirst") = "1" Or LCase(GetHeadItemValue(domHead, "bfirst")) = "true") Then
                            swhere = " and isnull(dEDate,'9999-12-31')>'" & sDate & "'"
                        End If
                    End If
                End If
                
                If LCase(sKey) = "currentstock" Then
                    keyCode = GetBodyItemValue(domBody, "cinvname", 0)
                    strSQL = GetSQLToValid(CurrentStock, domBody, domHead) & " WHERE cinvcode=N'" & GetBodyItemValue(domBody, "cinvcode", 0) & "' and cinvname=N'" & GetBodyItemValue(domBody, "cinvname", 0) & "'" _
                    & swhere & IIf(strAuth = "", "", " and inventory.iid in (" & strAuth & ")")
                    sKey = "cinvname"
                    
                Else
'                    strSQL = GetSQLToValid(inventory, domBody, DOMHead, 0) & " WHERE cinvcode='" & GetBodyItemValue(domBody, "cinvcode", 0) & "' and cinvname='" & keyCode & "' and isnull(dEDate,'')='' and isnull(bSale,0)=1 " & IIf(strAuth = "", "", " and inventory.iid in (" & strAuth & ")")
'                    If GetHeadItemValue(DOMHead, "cbustype") = "直运销售" Then
'                        strSQL = strSQL + " and  isnull(bInvEntrust,0)<>1"
'                    End If
'                    strSQL = strSQL & " union " & GetSQLToValid(inventory, domBody, DOMHead, 0) & " WHERE (cinvcode='" & keyCode & "' OR cinvname='" & keyCode & "' OR    cInvAddCode='" & keyCode & "') and isnull(dEDate,'')='' and isnull(bSale,0)=1" & IIf(strAuth = "", "", " and inventory.iid in (" & strAuth & ")")
                    strSQL = GetSQLToValid(Inventory, domBody, domHead, 0) & " WHERE (cinvcode=N'" & keyCode & "' OR cinvname=N'" & keyCode & "' OR    cInvAddCode=N'" & keyCode & "' OR Inventory.cBarCode=N'" & keyCode & "') " _
                    & swhere & " and isnull(bSale,0)=1" & IIf(strAuth = "", "", " and inventory.iid in (" & strAuth & ")")
                    If sKey = "cinvname2" Or sKey = "cinvcode2" And (strVouchType = "98" Or strVouchType = "99") Then
                        If Val(GetHeadItemValue(domHead, "sbvid")) <> 0 Then
                            strSQL = strSQL & " and cinvcode in (select cinvcode from salebillvouchs where sbvid = " & Val(GetHeadItemValue(domHead, "sbvid")) & " )"
                        End If
                    End If
                    If GetHeadItemValue(domHead, "cbustype") = "直运销售" And sKey <> "cinvname2" Then
                        strSQL = strSQL + " and  isnull(bInvEntrust,0)<>1 and isnull(bPurchase,0)=1 "
                    End If
                End If
                rstFile.Source = strSQL
                
                rstFile.Open , m_Conn
                If rstFile.RecordCount > 1 Then
                'If LCase(sKey) <> "cinvcode" Then
                    'rstFile.Filter = "cinvcode='" & GetBodyItemValue(domBody, "cinvcode", 0) & "' and cinvname='" & keyCode & "'"
                    rstFile.Filter = "cinvcode='" & GetBodyItemValue(domBody, "cinvcode", 0) & "' and cinvname='" & GetBodyItemValue(domBody, "cinvname", 0) & "'"
                End If
                'ElseIf LCase(sKey) = "cinvcode" Then
                '    rstFile.Filter = "cinvcode='" & GetBodyItemValue(domBody, "cinvcode", 0) & "' or cbarcode='" & GetBodyItemValue(domBody, "cinvcode", 0) & "' "
                    'sKey = "cinvname"
                'End If
                If rstFile.EOF Then rstFile.Filter = ""
                
                
                
                
                If keyCode = "" Then           '(1)如果清空表头项
                    If LCase(sKey) <> "all" Then
                        Call ClearDombody(strXMLbody, rstFile, domBody)
                    End If
                Else
                    If Not (rstFile.BOF And rstFile.EOF) Then                       '(2)如果合法
                        If LCase(sKey) <> "all" Then
                            ''需要计算单价
                            If (LCase(sKey) = "cinvname" Or LCase(sKey) = "cinvcode") And rstFile("binvtype") <> 1 And LCase(rstFile("binvtype")) <> "true" Then    ' And Val(GetBodyItemValue(domBody, "iunitprice", 0)) = 0
'                                Dim domLine As New DOMDocument
                                Set domLine = domBody.cloneNode(True)
                                Call SetDombody(sKey, strXMLbody, rstFile, domBody, domHead, R)
                                Set domBody = domLine.cloneNode(True)
                                SetDomBodyItemValues domBody, strXMLbody
                                BodyCheck = GetPrice(domHead, domBody, True)
                                If BodyCheck <> "" Then
                                    If rstFile.State <> adStateClosed Then
                                        rstFile.Close
                                    End If
                                    Set rstFile = Nothing
                                    Exit Function
                                End If
                                strXMLbody = ""
                            Else
                                Call SetDombody(sKey, strXMLbody, rstFile, domBody, domHead, R)
                            End If
                        End If
                    Else                                                            '(3)如不合法
                        If m_UseMode = CS Then
                            BodyCheck = GetString("U8.SA.USSASERVER.clscommcheck.00467") 'zh-CN：存货录入不正确或者没有销售属性或者没有权限或者被停用，请检查!
                            Call ClearDombody(strXMLbody, rstFile, domBody)
                            If rstFile.State <> adStateClosed Then
                                rstFile.Close
                            End If
                            Set rstFile = Nothing
                            Exit Function
                        ElseIf m_UseMode = BS Then
                            If LCase(sKey) <> "all" Then
                                BodyCheck = GetString("U8.SA.USSASERVER.clscommcheck.00468") 'zh-CN：存货录入不正确或者没有权限或者被停用，请检查!
                                If rstFile.State <> adStateClosed Then
                                    rstFile.Close
                                End If
                                Set rstFile = Nothing
                                Exit Function
                            Else
                                BodyCheck = BodyCheck & GetString("U8.SA.USSASERVER.clscommcheck.00469") & Chr(13) 'zh-CN：存货录入不正确或者没有权限或者被停用，请检查!
                            End If
                        End If
                    End If
                End If
'                Set clsAuth = Nothing
                If rstFile.State <> adStateClosed Then
                    rstFile.Close
                End If
        Case "cexpname", "all"        '销售费用

'                keyCode = GetBodyItemValue(DOMBody, sKey, 0)
                rstFile.Source = GetSQLToValid(ExpenseItem) & " WHERE cExpName=N'" & keyCode & "' OR cExpCode=N'" & keyCode & "'"
                rstFile.Open , m_Conn
                If keyCode = "" Then           '(1)如果清空表头项
                    If LCase(sKey) <> "all" Then
                        Call ClearDombody(strXMLbody, rstFile, domBody)
                    End If
                Else
                    If Not (rstFile.BOF And rstFile.EOF) Then                       '(2)如果合法
                        If LCase(sKey) <> "all" Then
                            Call SetDombody(sKey, strXMLbody, rstFile, domBody, domHead, R)
                        End If
                    Else                                                            '(3)如不合法
                        If m_UseMode = CS Then
                            BodyCheck = GetString("U8.SA.USSASERVER.clscommcheck.00470") 'zh-CN：费用项目录入不正确，请检查!
                            Call ClearDombody(strXMLbody, rstFile, domBody)
                            If rstFile.State <> adStateClosed Then
                                rstFile.Close
                            End If
                            Set rstFile = Nothing
                            Exit Function
                        ElseIf m_UseMode = BS Then
                            If LCase(sKey) <> "all" Then
                                BodyCheck = GetString("U8.SA.USSASERVER.clscommcheck.00471") 'zh-CN：费用项目录入不正确，请检查!
                                If rstFile.State <> adStateClosed Then
                                    rstFile.Close
                                End If
                                Set rstFile = Nothing
                                Exit Function
                            Else
                                BodyCheck = BodyCheck & GetString("U8.SA.USSASERVER.clscommcheck.00472") & Chr(13) 'zh-CN：费用项目录入不正确，请检查!
                            End If
                        End If
                    End If
                End If
                If rstFile.State <> adStateClosed Then
                    rstFile.Close
                End If
        
        Case "cwhname", "cwhcode", "all"   '仓库
                If clsSys.bAuth_Wh Then
                    strAuth = clsSys.clsAuth.GetAuthStringCode("WAREHOUSE", , "W")
                    If strAuth = "1=2" And keyCode <> "" Then
                        BodyCheck = GetString("没有仓库录入权限，请检查!") 'zh-CN：没有存货录入权限，请检查!
                        If rstFile.State <> adStateClosed Then
                            rstFile.Close
                        End If
                        Set rstFile = Nothing
                        Exit Function
                    End If
                End If
                
                If keyCode = "" Or strAuth = "1=2" Then
                    rstFile.Source = GetSQLToValid(warehouse, domBody, domHead) & " WHERE 1=2"
                Else
                    rstFile.Source = GetSQLToValid(warehouse, domBody, domHead) & " WHERE (cWhName=N'" & keyCode & "' OR cWhCode=N'" & keyCode & "')" _
                        & IIf(strAuth = "", "", " and cwhcode in(" & strAuth & ")")
                End If
                rstFile.Open , m_Conn
        
                If keyCode = "" Then           '(1)如果清空表头项
                    If LCase(sKey) <> "all" Then
                        Call ClearDombody(strXMLbody, rstFile, domBody)
                    End If
                Else
                    If Not (rstFile.BOF And rstFile.EOF) Then                       '(2)如果合法
                        If LCase(sKey) <> "all" Then
                            Call SetDombody(sKey, strXMLbody, rstFile, domBody, domHead, R)
                        End If
                    Else                                                            '(3)如不合法
                        If m_UseMode = CS Then
                            BodyCheck = GetString("U8.SA.USSASERVER.clscommcheck.00473") 'zh-CN：仓库录入不正确，请检查!
                            Call ClearDombody(strXMLbody, rstFile, domBody)
                            If rstFile.State <> adStateClosed Then
                                rstFile.Close
                            End If
                            Set rstFile = Nothing
                            Exit Function
                        ElseIf m_UseMode = BS Then
                            If LCase(sKey) <> "all" Then
                                BodyCheck = GetString("U8.SA.USSASERVER.clscommcheck.00474") 'zh-CN：仓库录入不正确，请检查!
                                If rstFile.State <> adStateClosed Then
                                    rstFile.Close
                                End If
                                Set rstFile = Nothing
                                Exit Function
                            Else
                                BodyCheck = BodyCheck & GetString("U8.SA.USSASERVER.clscommcheck.00475") & Chr(13) 'zh-CN：仓库录入不正确，请检查!
                            End If
                        End If
                    End If
                End If
                If rstFile.State <> adStateClosed Then
                    rstFile.Close
                End If
        Case "cvmivencode", "cvmivenname" '代管仓
            rstFile.Source = GetSQLToValid(vendor, domBody, domHead) & " WHERE cvencode = N'" & keyCode & "' OR cvenname = N'" & keyCode & "'"
            rstFile.Open , m_Conn
            If keyCode = "" Then
                 Call ClearDombody(strXMLbody, rstFile, domBody)
            Else
                If Not (rstFile.BOF And rstFile.EOF) Then                       '(2)如果合法
                    Call SetDombody(sKey, strXMLbody, rstFile, domBody, domHead, R)
                Else
                    BodyCheck = GetStringPara("U8.SA.USSASERVER.clscommcheck.00463", GetString("U8.SA.USSASERVER.clscommcheck.00462"))   'zh-CN：批次 'zh-CN：入库单号 'zh-CN：供应商 'Para zh-CN：{0}录入不正确，请检查!
'                    BodyCheck = GetString("U8.SA.xsglsql_2.errmsg.VenErr")
                End If
            End If
    End Select
    Exit Function
DOExit:
    If Not rstFile Is Nothing Then
        If rstFile.State <> adStateClosed Then
            rstFile.Close
        End If
        Set rstFile = Nothing
    End If
    Exit Function
DoErr:
    BodyCheck = err.Description
    If rstFile.State <> adStateClosed Then
        rstFile.Close
    End If
    Set rstFile = Nothing
End Function
Public Function BodyAllFetchPrice(cnn As ADODB.Connection, domHead As DOMDocument, domBody As DOMDocument, clsSASystem As Object) As String
    Dim domBodyLine As New DOMDocument
    Dim ele As IXMLDOMElement
    Dim lst As IXMLDOMNodeList
    Dim lstTmp As IXMLDOMNodeList
    Dim eleBody As IXMLDOMElement
    Dim Attr As IXMLDOMAttribute
    
    Set clsSys = clsSASystem
    Set m_Conn = cnn
    BodyAllFetchPrice = ""
    Set lst = domBody.selectNodes("//z:row")
    For Each ele In lst
        Set domBodyLine = domBody.cloneNode(True)
        Set lstTmp = domBodyLine.selectNodes("//z:row")
        For Each eleBody In lstTmp
            domBodyLine.selectSingleNode("//rs:data").removeChild eleBody
        Next
        Set eleBody = ele.cloneNode(True)
        domBodyLine.selectSingleNode("//rs:data").appendChild eleBody
        GetPrice domHead, domBodyLine, True
        Set eleBody = domBodyLine.selectSingleNode("//z:row")
        For Each Attr In eleBody.Attributes
            If LCase(Attr.Name) <> "xmlns:z" Then
                ele.setAttribute LCase(Attr.Name), Attr.value
            End If
        Next
        If ele.Attributes.getNamedItem("editprop") Is Nothing Then
            ele.setAttribute "editprop", "M"
        Else
            If ele.getAttribute("editprop") = "" Then
                ele.setAttribute "editprop", "M"
            End If
        End If
    Next
    Set domBodyLine = Nothing
End Function
Private Function GetInvBGsp(strInvCode As String) As String
    Dim rst As New ADODB.Recordset

    rst.CursorLocation = adUseClient
    rst.Open "select case when (isnull(bPropertyCheck,0)=1 and iGroupType <>2 )  then N'是' else N'否' end as bgsp from inventory where cinvcode=N'" + strInvCode + "'", m_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not rst.EOF Then
        GetInvBGsp = rst(0)
    End If
    rst.Close
    Set rst = Nothing
End Function
Private Function GetInvFieldValue(strFieleName As String, strInvCode As String) As String
    Dim rst As New ADODB.Recordset
    rst.CursorLocation = adUseClient
    
    rst.Open "select isnull(" + strFieleName + ",N'') from inventory where cinvcode=N'" + strInvCode + "'", m_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not rst.EOF Then
        GetInvFieldValue = rst(0)
    End If
    rst.Close
    Set rst = Nothing
End Function
Private Function IsVoucherNeedGetPrice(domBody As DOMDocument) As Boolean
            Dim strContractId As String
            Dim strBRefPa As String
            Dim strSosid As String
            Dim strIdlsid As String
            Dim blnGetPrice As Boolean
            Dim strQuoCode As String
            Dim strCorID As String
    
    IsVoucherNeedGetPrice = True
    
    strContractId = GetBodyItemValue(domBody, "ccontractid", 0)
    strBRefPa = GetBodyItemValue(domBody, "brefpa", 0)
    If strBRefPa = "" Then strBRefPa = "False"
    If strContractId <> "" Or CBool(strBRefPa) Then
        IsVoucherNeedGetPrice = False
    End If
    strQuoCode = GetBodyItemValue(domBody, "cquocode", 0)
    If strVouchType <> "16" And strQuoCode <> "" Then
        IsVoucherNeedGetPrice = False
    End If
    strSosid = GetBodyItemValue(domBody, "isosid", 0)
    If strVouchType <> "97" And strSosid <> "" Then
        IsVoucherNeedGetPrice = False
    End If
    strIdlsid = GetBodyItemValue(domBody, "idlsid", 0)
    If strVouchType <> "05" And strVouchType <> "06" And strVouchType <> "00" And strIdlsid <> "" Then
        IsVoucherNeedGetPrice = False
    End If
    strCorID = GetBodyItemValue(domBody, "icorid", 0)
    If (strVouchType = "05" Or strVouchType = "06" Or strVouchType = "07") And strCorID <> "" Then
        IsVoucherNeedGetPrice = False
    End If
End Function
Public Function GetGSPSignal(domHead As DOMDocument, domBody As DOMDocument) As String
    Dim strVouchType As String
    Dim strQuantity As String
    Dim ele As IXMLDOMElement
    Dim strInvType As String
    Dim strInvCode As String
    
    Set ele = domBody.selectSingleNode("//z:row")
    strQuantity = GetElEAtrVal(ele, "iquantity")
    strVouchType = GetHeadItemValue(domHead, "cvouchtype")
    strInvType = GetBodyItemValue(domBody, "binvtype", 0)
    If strInvType = "" Then strInvType = "0"
    strInvCode = GetBodyItemValue(domBody, "cinvcode", 0)
    If (strVouchType = "05" Or strVouchType = "06" Or strVouchType = "26" Or strVouchType = "27" Or strVouchType = "28" Or strVouchType = "29") And Not CBool(strInvType) Then
        If Val(strQuantity) < 0 And clsSys.bGspStart Then
            Dim strGsp As String
            strGsp = GetInvBGsp(strInvCode)
            GetGSPSignal = strGsp
        Else
            GetGSPSignal = "否"
        End If
    End If
End Function

Public Function GetPrice(domHead As DOMDocument, domBody As DOMDocument, Optional blnReget As Boolean = False) As String
    Dim iQuotedPrice As Double
    Dim fSalecost As Double
    Dim SettleCost As Double
    Dim kl As Double
    Dim fCusMinPrice As Double
    Dim iExchRate As Double
    Dim bIscal As Boolean
    Dim ele As IXMLDOMElement
    Dim iTaxRate As Double
    Dim iSum As Double
    Dim iTax As Double
    Dim iMoney As Double
    Dim iDisCount As Double
    Dim strQuantity As String
    Dim iNatMoney As Double
    Dim iTaxUnitPrice As Double
    Dim iUnitPrice As Double
    Dim kl2 As Double
    Dim iNatTax As Double
    Dim iNatSum As Double
    Dim iNatDisCount As Double
    Dim strInvCode As String
    Dim dblOldPrice As Double
    Dim strQtyPrice As String
    Dim strVouType As String
    
    strInvCode = GetBodyItemValue(domBody, "cinvcode", 0)
    strQtyPrice = clsSys.GetSysDicOption("SA", "bquantityprice")  ' getAccinformation(m_Conn, "SA", "bquantityprice")
    If strQtyPrice = "" Then strQtyPrice = "True"
    kl = 100
    Set ele = domBody.selectSingleNode("//z:row")
    strQuantity = GetElEAtrVal(ele, "iquantity")
    Dim strInvType As String
    strInvType = GetBodyItemValue(domBody, "binvtype", 0)
    If strInvType = "" Then strInvType = "0"
    strVouType = GetHeadItemValue(domHead, "cvouchtype")
    If (strVouType = "05" Or strVouType = "06" Or strVouType = "26" Or strVouType = "27" Or strVouType = "28" Or strVouType = "29") And Not CBool(strInvType) Then
        If Val(strQuantity) < 0 And clsSys.bGspStart Then
            Dim strGsp As String
            strGsp = GetInvBGsp(strInvCode)
            ele.setAttribute "bgsp", strGsp
        Else
            ele.setAttribute "bgsp", "否"
        End If
    End If
    If Not blnReget Then
        If Not IsVoucherNeedGetPrice(domBody) Then Exit Function
    End If
    
    If clsSys.CostGetType = 2 And CBool(strQtyPrice) And strQuantity = "" Then Exit Function
    
    GetPrice = CountPrice(domHead, domBody, iQuotedPrice, fSalecost, SettleCost, kl, fCusMinPrice, CBool(strQtyPrice))
    If GetPrice <> "" Then Exit Function
    If iQuotedPrice = 0 And Not blnReget Then
        dblOldPrice = Val(GetBodyItemValue(domBody, "iquotedprice", 0))
        If dblOldPrice <> 0 Then
            iQuotedPrice = dblOldPrice
            kl = Val(GetBodyItemValue(domBody, "kl", 0))
            If kl = 0 Then kl = 100
            SettleCost = iQuotedPrice * kl / 100
            fSalecost = Val(GetBodyItemValue(domBody, "fsalecost", 0))
            fCusMinPrice = Val(GetBodyItemValue(domBody, "fcusminprice", 0))
        End If
    End If
    If iQuotedPrice = 0 And Val(GetBodyItemValue(domBody, "iunitprice", 0)) <> 0 And Not blnReget Then Exit Function
    iExchRate = CDbl(Val(GetHeadItemValue(domHead, "iexchrate")))
    bIscal = clsSys.IsCal(GetHeadItemValue(domHead, "cexch_name"))
    kl2 = 100
    Call getKL2(kl2, domHead, domBody, m_Conn, GetPrice, "iquantity")
    SettleCost = SettleCost * kl2 / 100
    ele.setAttribute "kl2", kl2
        
    If iExchRate = 0 Then iExchRate = 1
    If bIscal = False Then
        iExchRate = 1 / iExchRate
    End If
    
    iTaxRate = Val(GetBodyItemValue(domBody, "itaxrate", 0))
    If strVouType = "28" And iQuotedPrice <> dblOldPrice Then
        ele.setAttribute "itaxrate", "0"
        If iTaxRate <> -100 Then
            If clsSys.bSalePrice Then
                Dim dblInvTaxRate As Double
                dblInvTaxRate = Val(GetInvFieldValue("itaxrate", strInvCode))
                iQuotedPrice = iQuotedPrice * 100 / (100 + dblInvTaxRate)
                SettleCost = SettleCost * 100 / (100 + dblInvTaxRate)
            End If
        Else
            iQuotedPrice = 0
            SettleCost = 0
        End If
        iTaxRate = 0
    End If
    ele.setAttribute "iquotedprice", CStr(iQuotedPrice)
        
    If Not clsSys.bSalePrice Then
        iQuotedPrice = iQuotedPrice * (1 + iTaxRate / 100)
        SettleCost = SettleCost * (1 + iTaxRate / 100)
    End If
    If clsSys.bSalePrice Then
        iSum = FormatNum(Val(strQuantity) * FormatNum(SettleCost, clsSys.iBillPriceBit), 2)
        iTax = FormatNum(iSum * (iTaxRate / (100 + iTaxRate)), 2)
        iMoney = iSum - iTax
    Else
        iMoney = FormatNum(Val(strQuantity) * FormatNum(SettleCost * 100 / (100 + iTaxRate), clsSys.iBillPriceBit), 2)
        iTax = FormatNum(iMoney * iTaxRate / 100, 2)
        iSum = iMoney + iTax
    End If
    If kl = 100 And kl2 = 100 Then
        iDisCount = 0
    Else
        iDisCount = FormatNum(Val(strQuantity) * iQuotedPrice, 2) - iSum
    End If
    iNatMoney = Format(iMoney, "#0.00") * iExchRate
    iNatTax = Format(iTax, "#0.00") * iExchRate
    iNatSum = CDbl(Format(iNatMoney, "#0.00")) + CDbl(Format(iNatTax, "#0.00"))
    iNatDisCount = Format(iDisCount, "#0.00") * iExchRate
    ele.setAttribute "isum", iSum
    ele.setAttribute "imoney", iMoney
    ele.setAttribute "itax", iTax
    ele.setAttribute "idiscount", iDisCount
    ele.setAttribute "inatsum", iNatSum
    ele.setAttribute "inatmoney", iNatMoney
    ele.setAttribute "inattax", iNatTax
    ele.setAttribute "inatdiscount", iNatDisCount
         ''U860 修改算法
'    If Val(strQuantity) <> 0 Then
        If clsSys.bSalePrice Then
            iTaxUnitPrice = FormatNum(SettleCost, clsSys.iBillPriceBit)
            '支持问题66411
            If iTaxRate <> 0 Then
                If Val(strQuantity) <> 0 Then
                    iUnitPrice = iMoney / Val(strQuantity)
                Else
                    iUnitPrice = FormatNum(SettleCost / (1 + iTaxRate / 100), clsSys.iBillPriceBit)
                End If
            Else
                iUnitPrice = iTaxUnitPrice
            End If
        Else
            iUnitPrice = FormatNum(SettleCost / (1 + iTaxRate / 100), clsSys.iBillPriceBit)
            '支持问题66411
            If iTaxRate <> 0 Then
                If Val(strQuantity) <> 0 Then
                    iTaxUnitPrice = iSum / Val(strQuantity)
                Else
                    iTaxUnitPrice = SettleCost
                End If
            Else
                iTaxUnitPrice = iUnitPrice
            End If
        End If
'    End If
    ele.setAttribute "itaxunitprice", iTaxUnitPrice
    ele.setAttribute "iunitprice", iUnitPrice
    ele.setAttribute "kl", kl
    ele.setAttribute "dkl1", 100 - kl
    ele.setAttribute "dkl2", 100 - kl2
    ele.setAttribute "inattaxunitprice", iExchRate * iTaxUnitPrice
    ele.setAttribute "inatunitprice", iExchRate * iUnitPrice
    ele.setAttribute "fcusminprice", fCusMinPrice
    ele.setAttribute "fsalecost", fSalecost
    ele.setAttribute "fsaleprice", FormatNum(Val(strQuantity) * fSalecost, 2)
    
'    Dim strGroupType As String
'    Dim iChangRate As Double
'    strGroupType = GetBodyItemValue(domBody, "igrouptype", 0)
'    If IIf(strGroupType = "", 0, Val(strGroupType)) <> 0 Then
'        iChangRate = Val(GetBodyItemValue(domBody, "iinvexchrate", 0))
'        If clsSys.iFloatRateRule = 0 And Val(strGroupType) = 2 Then
'            If Val(GetBodyItemValue(domBody, "inum", 0)) = 0 Then
'                'ele.setAttribute "iinvexchrate", ""
'                If Val(GetBodyItemValue(domBody, "iinvexchrate", 0)) <> 0 Then
'                    ele.setAttribute "inum", IIf(strQuantity = "", "", IIf(Val(strQuantity) = 0, "0", Val(strQuantity) / iChangRate))
'                End If
'            Else
'                If Val(strQuantity) <> 0 Then
'                    ele.setAttribute "iinvexchrate", CStr(Abs(Val(strQuantity) / Val(GetBodyItemValue(domBody, "inum", 0))))
'                    ele.setAttribute "inum", CStr(Sgn(Val(strQuantity)) * Abs(Val(GetBodyItemValue(domBody, "inum", 0))))
'                Else
'                    ele.setAttribute "iinvexchrate", "0"
''                    ele.setAttribute "inum", CStr(Sgn(Val(strQuantity)) * Abs(Val(GetBodyItemValue(domBody, "inum", 0))))
'                End If
'            End If
'        Else
'            If iChangRate = 0 Then
'                ele.setAttribute "inum", "0"
'            Else
'                ele.setAttribute "inum", IIf(strQuantity = "", "", IIf(Val(strQuantity) = 0, "0", Val(strQuantity) / iChangRate))
'            End If
'        End If
'    End If
End Function

''根据存货代码,存货编码等计算报价,零售单价,结算价格,扣率 zzg
'Public Function CountPrice(cCusCode As String, cInvCode As String, cFree() As String, iQuotedPrice As Double, fSalecost As Double, SettleCost As Double, kl As Double, ddate As String, strExchName As String, Optional domBody As Variant, Optional domHead As DOMDocument) As String
Public Function CountPrice(domHead As DOMDocument, domBody As DOMDocument, iQuotedPrice As Double, fSalecost As Double, SettleCost As Double, kl As Double, fCusMinPrice As Double, Optional blnQtyPrice As Boolean) As String
    Dim rstPrice As New ADODB.Recordset
    Dim iMode As Integer, i As Integer
    Dim strSQL As String, strErr As String
    Dim m_DOM As DOMDocument
    Dim m_ele As IXMLDOMElement
    Dim m_Obj As Object
    Dim m_Node As IXMLDOMNode
    Dim strPass As String
    Dim m_DomR As DOMDocument
    Dim strwhere As String
    Dim cCusCode As String
    Dim cInvCode As String
    Dim cFree(10) As String
    Dim blnSalePriceFree(10) As Boolean
    Dim ErrMsg As String
    Dim ddate As String, strExchName As String
    Dim iExchRate As Double
    Dim bIscal As Boolean
    Dim strFree1 As String
    Dim strFree2 As String
    
    CountPrice = ""
    cCusCode = GetHeadItemValue(domHead, "ccuscode")
    strExchName = GetHeadItemValue(domHead, "cexch_name")
    ddate = ConvDateForDOM(GetHeadItemValue(domHead, "ddate"))
    For i = 1 To 10
        cFree(i) = GetBodyItemValue(domBody, "cfree" & CStr(i), 0)
    Next
    cInvCode = GetBodyItemValue(domBody, "cinvcode", 0)
    If cInvCode = "" Then Exit Function
    Set m_ele = domBody.selectSingleNode("//z:row")
    If m_ele.Attributes.getNamedItem("bsalepricefree1") Is Nothing Then
        rstPrice.Open "select isnull(bsalepricefree1,0) as bsalepricefree1,isnull(bsalepricefree2,0) as bsalepricefree2,isnull(bsalepricefree3,0) as bsalepricefree3,isnull(bsalepricefree4,0) as bsalepricefree4," _
            & "isnull(bsalepricefree5,0) as bsalepricefree5,isnull(bsalepricefree6,0) as bsalepricefree6,isnull(bsalepricefree7,0) as bsalepricefree7,isnull(bsalepricefree8,0) as bsalepricefree8," _
            & "isnull(bsalepricefree9,0) as bsalepricefree9,isnull(bsalepricefree10,0) as bsalepricefree10 from inventory_sub where cinvsubcode=N'" + cInvCode + "'", m_Conn, 3, 1
        If rstPrice.EOF Then
            rstPrice.Close
            Set rstPrice = Nothing
            Exit Function
        End If
        Dim fld As ADODB.Field
        For Each fld In rstPrice.Fields
            m_ele.setAttribute fld.Name, fld.value
        Next
        rstPrice.Close
    End If
    Dim blnFree As Boolean
    blnFree = True
    For i = 1 To 10
        strFree1 = GetAttrValue(m_ele, "bsalepricefree" & CStr(i))
        If strFree1 = "" Then strFree1 = "0"
        blnSalePriceFree(i) = CBool(strFree1)
        strFree2 = cFree(i)
        If CBool(strFree1) And strFree2 = "" Then
            blnFree = False
        End If
'        If blnFree Then Exit For
    Next
'    If strFree2 = "" Then strFree2 = "0"
'    If (clsSys.CostGetType = 0 Or clsSys.CostGetType = 2) And Not blnFree Then
''        And ((CBool(strFree1) And cFree(1) = "" And mId(clsSys.FreePriceType, Val(Right("cfree1", 1)), 1)) Or (CBool(strFree2) And cFree(2) = "" And mId(clsSys.FreePriceType, Val(Right("cfree2", 1)), 1) = "1")) Then
'        Set rstPrice = Nothing
'        Exit Function
'    End If
    Select Case clsSys.CostGetType
        Case 0   ''最新售价
            For i = 1 To 10
'                strFree1 = m_ele.Attributes.getNamedItem("bsalepricefree" + CStr(i)).nodeValue
'                If strFree1 = "" Then strFree1 = "0"
                If blnSalePriceFree(i) Then
                    strwhere = strwhere & " and isnull(cfree" + CStr(i) + ",N'')=N'" & cFree(i) & "'"
                End If
            Next
'            If mId(clsSys.FreePriceType, Val(Right("cfree1", 1)), 1) = "1" Then
'                strwhere = " and isnull(cfree1,N'')=N'" & cFree(1) & "'"
'            Else
'                strwhere = ""
'            End If
'            If mId(clsSys.FreePriceType, Val(Right("cfree2", 1)), 1) = "1" Then
'                strwhere = strwhere & " and " & "isnull(cfree2,N'')=N'" & cFree(2) & "'"
'                'strwhere = strwhere & IIf(strwhere = "", "", " and ") & "isnull(cfree2,N'')=N'" & cFree(2) & "'"
'            End If
            
            Select Case clsSys.CostReferVouch
                Case 0
                    strSQL = "select top 1 isnull(iQuotedPrice,0) as iQuotedPrice,isnull(fSaleCost,0) as fSaleCost  from so_somain a inner join so_sodetails b on a.id=b.id where cinvcode=N'" & cInvCode _
                            & "' " & strwhere & IIf(clsSys.CostRefCustomer, "and ccuscode=N'" & cCusCode & "'", "") & "and cexch_name=N'" & strExchName & "' order by ddate desc,a.id desc"
                Case 1
                    strSQL = "select top 1 isnull(iQuotedPrice,0) as iQuotedPrice,isnull(fSaleCost,0) as fSaleCost from dispatchlist a inner join dispatchlists b on a.dlid=b.dlid where cinvcode=N'" & cInvCode _
                            & "' " & strwhere & IIf(clsSys.CostRefCustomer, "and ccuscode=N'" & cCusCode & "'", "") & "and cexch_name=N'" & strExchName & "' order by ddate desc,a.dlid desc"
                Case 2
                    strSQL = "select top 1 isnull(iQuotedPrice,0) as iQuotedPrice,isnull(fSaleCost,0) as fSaleCost from salebillvouch a inner join salebillvouchs b on a.sbvid=b.sbvid where cinvcode=N'" & cInvCode _
                            & "' " & strwhere & IIf(clsSys.CostRefCustomer, "and ccuscode=N'" & cCusCode & "'", "") & "and cexch_name=N'" & strExchName & "' order by ddate desc,a.sbvid desc"
                Case 3
                    strSQL = "select top 1 isnull(iQuotedPrice,0) as iQuotedPrice,isnull(fSaleCost,0) as fSaleCost  from sa_quomain a inner join sa_quodetails b on a.id=b.id where cinvcode=N'" & cInvCode _
                            & "' " & strwhere & IIf(clsSys.CostRefCustomer, "and ccuscode=N'" & cCusCode & "'", "") & "and cexch_name=N'" & strExchName & "' order by ddate desc,a.id desc"
            End Select
            rstPrice.Open strSQL, m_Conn, 3, 1
            If Not rstPrice.EOF Then
                iQuotedPrice = rstPrice("iQuotedPrice")
                fSalecost = rstPrice("fSaleCost")
                SettleCost = iQuotedPrice
                kl = 100
            Else
                iQuotedPrice = 0
                fSalecost = 0
                SettleCost = iQuotedPrice
                kl = 100
            End If
'            fCusMinPrice = Val(GetInvFieldValue("iinvlscost", cInvCode))
           
        Case 1   ''最新成本加成
                'zcy
                Dim iBatch As Long '入库单子表id
                Dim rs As New ADODB.Recordset
                Dim iTrack As Long '跟踪存货标记
                Dim iCostPrice As Double '入库成本
                If Not IsMissing(domBody) Then
                    iBatch = Val(GetBodyItemValue(domBody, "ibatch", 0))
                    If LCase(GetBodyItemValue(domBody, "btrack", 0)) = "true" Then
                        iTrack = 1
                    Else
                        iTrack = Val(GetBodyItemValue(domBody, "btrack", 0))
                    End If
                End If
                
                strSQL = "select isnull(iInvNCost,0) as iInvNCost ,isnull(iInvSaleCost,0) as iInvSaleCost,isnull(iExpSaleRate,0) as iExpSaleRate from inventory where cinvcode=N'" & cInvCode & "'"
                rstPrice.Open strSQL, m_Conn, 3, 1
                  
                If iBatch <> 0 And (iTrack = 1 Or iTrack = -1) Then '有子表id并且是跟踪存货
                    rs.Open "Select isnull(iUnitCost,0) as iUnitCost from rdrecords Where AutoID=" & iBatch, m_Conn, adOpenForwardOnly, adLockReadOnly
                    If Not (rs.BOF And rs.EOF) Then
                        iCostPrice = rs("iUnitCost")
                    End If
                    iQuotedPrice = iCostPrice * (1 + rstPrice("iExpSaleRate") / 100)
                    rs.Close
                Else
                  iQuotedPrice = rstPrice("iInvNCost") * (1 + rstPrice("iExpSaleRate") / 100)
                End If
                iExchRate = CDbl(Val(GetHeadItemValue(domHead, "iexchrate")))
                bIscal = clsSys.IsCal(GetHeadItemValue(domHead, "cexch_name"))
                If iExchRate = 0 Then iExchRate = 1
                If bIscal = False Then
                    iExchRate = 1 / iExchRate
                End If
                iQuotedPrice = iQuotedPrice / iExchRate
                SettleCost = iQuotedPrice
                fSalecost = rstPrice("iInvSaleCost")
                kl = 100
'                fCusMinPrice = Val(GetInvFieldValue("iinvlscost", cInvCode))
                
                rstPrice.Close
        Case 2
            Dim DOMPara As New DOMDocument
            Dim strPara As String
            Dim strQuantity As String
            Dim ele As IXMLDOMElement
            
            If blnQtyPrice Then
                If Not IsMissing(domBody) Then
                    strQuantity = GetBodyItemValue(domBody, "iquantity", 0)
                End If
            Else
                strQuantity = "0"
            End If
            
            strPara = "<Data> "
            strPara = strPara & "<row name='ccuscode' datatype='202' length='20' type='1' value ='" + cCusCode + "' />"
            strPara = strPara & "<row name='cinvcode' datatype='202' length='20' type='1' value ='" + cInvCode + "' />"
            For i = 1 To 10
                If blnSalePriceFree(i) Then
                    strPara = strPara & "<row name='sfree" + CStr(i) + "' datatype='202' length='20' type='1' value ='" + cFree(i) + "' />"
                Else
                    strPara = strPara & "<row name='sfree" + CStr(i) + "' datatype='202' length='20' type='1' value ='' />"
                End If
            Next
'            If mId(clsSys.FreePriceType, Val(Right("cfree2", 1)), 1) = "1" Then
'                strPara = strPara & "<row name='sfree2' datatype='202' length='20' type='1' value ='" + cFree(2) + "' />"
'            Else
'                strPara = strPara & "<row name='sfree2' datatype='202' length='20' type='1' value ='' />"
'            End If
            strPara = strPara & "<row name='sdate' datatype='202' length='10' type='1' value ='" + ddate + "' />"
            strPara = strPara & "<row name='exchname' datatype='202' length='8' type='1' value ='" + strExchName + "' />"
            strPara = strPara & "<row name='quantity' datatype='5' length='10' type='1' value ='" + CStr(Abs(Val(strQuantity))) + "' />"
            strPara = strPara & "<row name='bsales' datatype='11' length='1' type='1' value ='1' />"
            strPara = strPara & "<row name='nInvPrice' datatype='5' length='10' type='2' value ='' />"
            strPara = strPara & "<row name='nSalePrice' datatype='5' length='10' type='2' value ='' />"
            strPara = strPara & "<row name='nInvNowPrice' datatype='5' length='10' type='2' value ='' />"
            strPara = strPara & "<row name='nDisRate' datatype='5' length='10' type='2' value ='' />"
            strPara = strPara & "<row name='minPrice' datatype='5' length='10' type='2' value ='' />"
            strPara = strPara & "<row name='sErr' datatype='202' length='200' type='2' value ='' />"
            strPara = strPara & " </Data>"
            DOMPara.loadXML strPara
            strErr = ""
            Dim fCusMinPriceTmp As Double
            strErr = DOMPara.documentElement.selectSingleNode("row[@name='sErr']").Attributes.getNamedItem("value").Text
            CountPrice = strErr
            iQuotedPrice = Val(DOMPara.documentElement.selectSingleNode("row[@name='nInvPrice']").Attributes.getNamedItem("value").Text)
            fSalecost = Val(DOMPara.documentElement.selectSingleNode("row[@name='nSalePrice']").Attributes.getNamedItem("value").Text)
            SettleCost = Val(DOMPara.documentElement.selectSingleNode("row[@name='nInvNowPrice']").Attributes.getNamedItem("value").Text)
            kl = Val(DOMPara.documentElement.selectSingleNode("row[@name='nDisRate']").Attributes.getNamedItem("value").Text)
            If SettleCost = 0 Then
                SettleCost = iQuotedPrice * kl / 100
            End If
'            If fCusMinPrice = 0 Then
'                fCusMinPrice = GetInvFieldValue("iinvlscost", cInvCode)
'            End If
            Set DOMPara = Nothing
        Case 7  ''增加二次开发接口
            Set m_DOM = New DOMDocument
            If m_DOM.Load(App.Path + IIf(Right(App.Path, 1) <> "\", "\", "") + "DefPrice.xml") = False Then
                CountPrice = GetString("U8.SA.USSASERVER.clscommcheck.00476") 'zh-CN：加载自定义格式错误，请检查
            End If
            On Error Resume Next
            Set m_ele = m_DOM.selectNodes("//interface/zt" + clsSys.CurrentAccID).Item(0)
            strPass = m_ele.selectSingleNode("PassWord").Text
            If CheckPass(strPass) = False Then
                CountPrice = GetString("U8.SA.USSASERVER.clscommcheck.00477") 'zh-CN：密码错误，请检查配置文件
                Exit Function
            End If
            Set m_Node = m_ele.selectSingleNode("ClassName")
            
            Set m_Obj = CreateObject(m_Node.Text)
            If Not m_Obj Is Nothing Then
                Set m_DomR = New DOMDocument
                Call m_Obj.CountPrice(m_Conn, domHead, domBody, m_DomR, ErrMsg)
                If ErrMsg <> "" Then
                    CountPrice = ErrMsg
                Else
                    Set m_ele = m_DomR.selectSingleNode("//price")
                    If Not m_ele Is Nothing Then
                        iQuotedPrice = CDbl(Val(GetElEAtrVal(m_ele, "iquotedprice")))
                        SettleCost = CDbl(Val(GetElEAtrVal(m_ele, "settlecost")))
                        fSalecost = CDbl(Val(GetElEAtrVal(m_ele, "fsalecost")))
                        kl = CDbl(Val(GetElEAtrVal(m_ele, "kl")))
                    End If
                End If
            Else
                CountPrice = GetStringPara("U8.SA.USSASERVER.clscommcheck.00478", m_Node.Text) 'Para zh-CN：无法创建对象{0}
            End If
            Set m_ele = Nothing
            Set m_Node = Nothing
            Set m_DOM = Nothing
     End Select
     
End Function

'Public Sub InitLogin(ioLogin As U8Login.clsLogin)
' Set objlogin = ioLogin
'End Sub
'获得换算率，并检查辅计量单位的正确性
Public Function GetInvExchRate(strInvCode As String, strGroupCode As String, strInvA_Unit As String, strMessage As String, strComunitCode As String) As Double
Dim RecTemp As ADODB.Recordset
Dim strSQL As String
Set RecTemp = New ADODB.Recordset
On Error GoTo err_Handle
strSQL = "SELECT CComunitCode,CComUnitName,ISNULL(iChangRate,0) AS iChangRate,iProportion FROM  ComputationUnit WHERE cGroupCode=N'" & strGroupCode & "'" & _
         " AND (CComunitCode=N'" & strInvA_Unit & "' OR CComUnitName=N'" & strInvA_Unit & "')"

RecTemp.CursorLocation = adUseClient
RecTemp.Open strSQL, m_Conn, adOpenForwardOnly, adLockReadOnly
If RecTemp.RecordCount > 0 Then
    GetInvExchRate = RecTemp!iChangRate
    strComunitCode = RecTemp!cComunitCode
    strInvA_Unit = RecTemp!cComUnitName
Else
    strMessage = GetString("U8.SA.USSASERVER.clscommcheck.00479") 'zh-CN：该计量单位不存在!
    strComunitCode = ""
    strInvA_Unit = ""

    RecTemp.Close
End If
Exit Function

err_Handle:
    strMessage = err.Description
End Function
Private Sub SetDomBodyItemValues(domBody As DOMDocument, xmlValues As String)
    Dim tmpDomBody As New DOMDocument
    tmpDomBody.loadXML xmlValues
    
    Dim ele As IXMLDOMElement
    Dim lst As IXMLDOMNodeList
    Dim eleBody As IXMLDOMElement
    Set eleBody = domBody.selectSingleNode("//z:row")
    Set lst = tmpDomBody.selectNodes("//R")
    For Each ele In lst
        eleBody.setAttribute LCase(ele.Attributes.getNamedItem("K").Text), ele.Attributes.getNamedItem("V").Text
    Next
    Set tmpDomBody = Nothing
End Sub
Public Sub SetDombody(sKey As String, strXMLbody As String, rstFile As ADODB.Recordset, domBody As DOMDocument, Optional domHead As DOMDocument = Nothing, Optional R As Long)
    
    Dim iCount As Integer
    Dim iTaxRate As Double
    Dim iUnitPrice  As Double, iTaxUnitPrice  As Double, iNum As Double
    Dim cMUnit As String     '主计量单位
    Dim cAumit As String     '辅计量单位
    Dim cMUnitCode As String     '主计量单位编码
    Dim cAumitCode As String     '辅计量单位编码
    Dim strGroupCode As String
    Dim iChangRate As Double
    Dim strErrMsg As String
    Dim strSQL As String
    Dim strBGSP As String
    Dim iSum As Double, iMoney As Double, iTax As Double, iDisCount As Double
    Dim iNatSum As Double, iNatMoney As Double, iNatTax As Double, iNatDisCount As Double
    Dim kl2 As Double  ''增加扣率2
    Dim ele As IXMLDOMElement
    Dim bQANeedCheck As Boolean
    Dim strQuantity As String
    
'    If kl = 0 Then kl = 100
    For iCount = 0 To rstFile.Fields.Count - 1
        If Not (LCase(rstFile.Fields(iCount).Name) = "iquantity" Or LCase(rstFile.Fields(iCount).Name) = "inum") Then
            If Not LCase(rstFile.Fields(iCount).Name) = "bgsp" And Not LCase(rstFile.Fields(iCount).Name) = "ufts" Then
                If rstFile.Fields(iCount).Type = adBoolean Then
                        strXMLbody = strXMLbody & "<R K='" & LCase(rstFile.Fields(iCount).Name) & "' V='" & IIf(rstFile.Fields(iCount) = True, 1, 0) & "'/>" & Chr(13)
                Else
                    If LCase(sKey) = "cinvname" Then
                        If LCase(rstFile.Fields(iCount).Name) = "itaxrate" Then
                            If IsNull(rstFile("itaxrate")) Then
                                strXMLbody = strXMLbody & "<R K='" & LCase(rstFile.Fields(iCount).Name) & "' V='" & GetHeadItemValue(domHead, "itaxrate") & "'/>" & Chr(13)
                            Else
                                strXMLbody = strXMLbody & "<R K='" & LCase(rstFile.Fields(iCount).Name) & "' V='" & rstFile.Fields(iCount) & "'/>" & Chr(13)
                            End If
                        Else
                            strXMLbody = strXMLbody & "<R K='" & LCase(rstFile.Fields(iCount).Name) & "' V='" & FormatStrForDOM(rstFile.Fields(iCount)) & "'/>" & Chr(13)
                        End If
                    ElseIf (LCase(sKey) = "cinvname2" Or LCase(sKey) = "cinvcode2") And (strVouchType = "97" Or strVouchType = "16") And Left(LCase(rstFile.Fields(iCount).Name), 5) = "cfree" Then
                        '不回写
                    Else
                        strXMLbody = strXMLbody & "<R K='" & LCase(rstFile.Fields(iCount).Name) & "' V='" & FormatStrForDOM(rstFile.Fields(iCount)) & "'/>" & Chr(13)
                    End If
                End If
            End If
        End If
    Next iCount
    'zcy
    If LCase(sKey) = "cinvcode" Or LCase(sKey) = "cinvname" Or sKey = "cinvname2" Or sKey = "cinvcode2" Then
        strXMLbody = strXMLbody & "<R K='ccusinvcode' V='" & GetBodyItemValue(domBody, "ccusinvcode", 0) & "'/>" & Chr(13)
        Call GetCusInvFromInv(strXMLbody, GetHeadItemValue(domHead, "ccuscode"), rstFile("cinvcode"))
    End If
    
    If LCase(sKey) = "cwhcode" Or LCase(sKey) = "cwhname" Then 'wxy added for 晴空供应商
        strXMLbody = strXMLbody & "<R K='cvmivencode' V=''/>" & Chr(13)
        strXMLbody = strXMLbody & "<R K='cvmivenname' V=''/>" & Chr(13)
    End If
    
    'If sKey = "ccode" Then GoTo CountPrice
    If LCase(sKey) = "cinvcode" Or LCase(sKey) = "cinvname" Or LCase(Left(sKey, 5)) = "cfree" Or sKey = "ccode" Or (strVouchType = "97" And (sKey = "cinvname2" Or sKey = "cinvcode2")) Then
        ''取得客户存货对照表
        If sKey = "ccode" And clsSys.CostGetType <> 1 Then GoTo DOExit  'zcy modify 2003-8-5
        strQuantity = GetBodyItemValue(domBody, "iquantity", 0)
        If LCase(sKey) = "cinvcode" Or LCase(sKey) = "cinvname" Or sKey = "cinvname2" Or sKey = "cinvcode2" Then
            If strVouchType = "05" Or strVouchType = "06" Or strVouchType = "26" Or strVouchType = "27" Or strVouchType = "28" Or strVouchType = "29" Then
                If rstFile("binvtype") = 1 Or LCase(rstFile("binvtype")) = "true" Then
                    strQuantity = 0
                    strXMLbody = strXMLbody & "<R K='bgsp' V='否'/>" & Chr(13)
                    ''质量检验取得存货默认值
                    strXMLbody = strXMLbody & "<R K='bQANeedCheck' V='否'/>" & Chr(13)
    
                Else
                    
                    bQANeedCheck = clsSys.bInvNeedQMCheck(rstFile("cinvcode"), GetHeadItemValue(domHead, "ccuscode"), IIf(LCase(GetBodyItemValue(domHead, "breturnflag", 0)) = "false" Or LCase(GetBodyItemValue(domHead, "breturnflag", 0)) = "0", True, False))
                    ''质量检验取得存货默认值
                    strXMLbody = strXMLbody & "<R K='bQANeedCheck' V='" + IIf(bQANeedCheck, "是", "否") + "'/>" & Chr(13)
                End If
                strXMLbody = strXMLbody & "<R K='bQAUrgency' V='否'/>" & Chr(13)
                strXMLbody = strXMLbody & "<R K='iqaquantity' V='0'/>" & Chr(13)
                strXMLbody = strXMLbody & "<R K='iqanum' V='0'/>" & Chr(13)
                strXMLbody = strXMLbody & "<R K='bQAChecking' V='否'/>" & Chr(13)
                strXMLbody = strXMLbody & "<R K='bQAChecked' V='否'/>" & Chr(13)
            Else
                If rstFile("binvtype") = 1 Or LCase(rstFile("binvtype")) = "true" Then
                    strQuantity = 0
                End If
            End If
        End If
'        If LCase(Left(sKey, 5)) = "cfree" Then
'            'zcy If clsSys.CostGetType = 2 And clsSys.CurrentPriceType = "存货价格" Then  客户价格也要考虑自由项
'            If clsSys.CostGetType = 0 Or clsSys.CostGetType = 2 Or clsSys.CostGetType = 7 Then
'                If Mid(clsSys.FreePriceType, Val(Mid(sKey, 6)), 1) = "0" Then
'                    GoTo DOExit
'                Else
'                    If strVouchType <> "28" Then
'                        iTaxRate = Val(GetBodyItemValue(domBody, "itaxrate", 0))
'                        GoTo CountPrice
'                    Else
'                        strSQL = GetSQLToValid(inventory, domBody, domHead, 0) & " WHERE cinvcode='" & GetBodyItemValue(domBody, "cinvcode", 0) & "'"
'                        'strSql = strSql & " union " & GetSQLToValid(inventory) & " WHERE cinvcode='" & keyCode & "' OR cinvname='" & keyCode & "'"
'                        If rstFile.State = 1 Then rstFile.Close
'                        rstFile.Source = strSQL
'                        rstFile.Open , m_Conn
'                        iTaxRate = IIf(IsNull(rstFile("itaxrateold")), 0, rstFile("itaxrateold"))
'                    End If
'                End If
'            Else
'                GoTo DOExit
'            End If
'        End If
        If Val(GetBodyItemValue(domBody, "kl", 0)) = 0 Then
            strXMLbody = strXMLbody & "<R K='kl' V='100'/>" & Chr(13)
        End If
        If Val(GetBodyItemValue(domBody, "kl2", 0)) = 0 Then
            strXMLbody = strXMLbody & "<R K='kl2' V='100'/>" & Chr(13)
        End If
        strXMLbody = strXMLbody & "<R K='igrouptype' V='" & rstFile!iGroupType & "'/>" & Chr(13)
        If IIf(IsNull(rstFile!cgroupcode), "", rstFile!cgroupcode) <> "" Then
            strGroupCode = rstFile!cgroupcode
            strXMLbody = strXMLbody & "<R K='igroupcode' V='" & FormatStrForDOM(strGroupCode) & "'/>" & Chr(13)
        End If
        
        If IIf(IsNull(rstFile!cSAComUnitCode), "", rstFile!cSAComUnitCode) = "" Then
            cAumitCode = IIf(IsNull(rstFile!cAssComUnitCode), "", FormatStrForDOM(rstFile!cAssComUnitCode))
        Else
            cAumitCode = IIf(IsNull(rstFile!cSAComUnitCode), "", FormatStrForDOM(rstFile!cSAComUnitCode))
        End If
        'strXMLbody = strXMLbody & "<R K='igroupcode' V='" & strGroupCode & "'/>" & Chr(13)
      
        cMUnitCode = IIf(IsNull(rstFile!cComunitCode), "", FormatStrForDOM(rstFile!cComunitCode))
        strXMLbody = strXMLbody & "<R K='ccomunitcode' V='" & cMUnitCode & "'/>" & Chr(13)
        'strXMLbody = strXMLbody & "<R K='casscomunitcode' V='" & cAumitCode & " '/>" & Chr(13)
        
        strXMLbody = strXMLbody & "<R K='cunitid' V='" & cAumitCode & "'/>" & Chr(13)
        
        
        cMUnit = GetUnitName(cMUnitCode, strGroupCode, iChangRate, strErrMsg)
        If strErrMsg = "" Then
            strXMLbody = strXMLbody & "<R K='cinvm_unit' V='" & cMUnit & "'/>" & Chr(13)
        End If
        If IIf(IsNull(rstFile!iGroupType), 0, rstFile!iGroupType) <> 0 Then
            cAumit = GetUnitName(cAumitCode, strGroupCode, iChangRate, strErrMsg)
            If strErrMsg = "" Then
                strXMLbody = strXMLbody & "<R K='cinva_unit' V='" & cAumit & "'/>" & Chr(13)
                strXMLbody = strXMLbody & "<R K='iinvexchrate' V='" & iChangRate & "'/>" & Chr(13)
                If iChangRate = 0 Then
                    strXMLbody = strXMLbody & "<R K='inum' V='0'/>" & Chr(13)
                Else
                    strXMLbody = strXMLbody & "<R K='inum' V='" & IIf(strQuantity = "", "", IIf(Val(strQuantity) = 0, "0", Val(strQuantity) / iChangRate)) & " '/>" & Chr(13)
                End If
            End If
        Else
            strXMLbody = strXMLbody & "<R K='cinva_unit' V=''/>" & Chr(13)
            strXMLbody = strXMLbody & "<R K='iinvexchrate' V=''/>" & Chr(13)
            strXMLbody = strXMLbody & "<R K='inum' V=''/>" & Chr(13)
        End If
        strXMLbody = strXMLbody & "<R K='cinvm_unit' V='" & cMUnit & "'/>" & Chr(13)
        strXMLbody = strXMLbody & "<R K='cinva_unit' V='" & cAumit & "'/>" & Chr(13)
        strXMLbody = strXMLbody & "<R K='iquantity' V='" & strQuantity & "'/>" & Chr(13)
        If (strVouchType = "97" Or strVouchType = "16") Then        'And (sKey = "cinvname2" Or sKey = "cinvcode2")
            'support 73588 预完工日期修改 2005.10.17 by jzq
            If strVouchType = "97" Then
                Dim ddate As Date
                Dim dPreMoDate As Date
                Dim dPreDate As Date
                Dim dPreMoDateBT As String
                Dim dPreDateBT As String
                
                ddate = ConvDateForDOM(GetHeadItemValue(domHead, "ddate"))
                dPreMoDateBT = Trim(GetHeadItemValue(domHead, "dpremodatebt"))
                If dPreMoDateBT <> "" Then dPreMoDateBT = ConvDateForDOM(dPreMoDateBT)
                dPreDateBT = Trim(GetHeadItemValue(domHead, "dpredatebt"))
                If dPreDateBT <> "" Then dPreDateBT = ConvDateForDOM(dPreDateBT)
                
                
                '预完工日期
                If Trim(dPreMoDateBT) = "" Then
                    dPreMoDate = ddate  '+ IIf(IsNull(rstFile.Fields("iAdvanceDate")), 0, rstFile.Fields("iAdvanceDate"))
                ElseIf IsDate(dPreMoDateBT) Then
                    dPreMoDate = CDate(dPreMoDateBT)
                Else
                    dPreMoDate = ddate
                End If
                
                '预发货日期
                If Trim(dPreDateBT) <> "" Then
                    dPreDate = dPreDateBT
                Else
                    dPreDate = dPreMoDate
                End If
                If Trim(GetBodyItemValue(domBody, "dpremodate", 0)) = "" Then
                    strXMLbody = strXMLbody & "<R K='dpremodate' V='" & dPreMoDate & "'/>" & Chr(13)
                    strXMLbody = strXMLbody & "<R K='dpredate' V='" & dPreDate & "'/>" & Chr(13)
                    strXMLbody = strXMLbody & "<R K='iadvancedate' V='" & rstFile.Fields("iadvancedate") & "'/>" & Chr(13)
                End If
            End If
            GoTo DOExit
        End If
        If strVouchType <> "28" Then
            iTaxRate = IIf(IsNull(rstFile("itaxrate")), 0, rstFile("itaxrate"))
        Else
            iTaxRate = IIf(IsNull(rstFile("itaxrateold")), 0, rstFile("itaxrateold"))
        End If
        strXMLbody = strXMLbody & "<R K='itaxrate' V='" & iTaxRate & "'/>" & Chr(13)
        
        
        
    End If
DOExit:
    strXMLbody = strXMLbody & "</body>"
    domBody.loadXML strXMLbody
    If rstFile.State = 1 Then
        rstFile.Close
        Set rstFile = Nothing
    End If


End Sub

Public Sub ClearDombody(strXMLbody As String, rstFile As ADODB.Recordset, domBody As DOMDocument)
    Dim iCount As Integer
    Dim tmpDomBody As New DOMDocument '同步850不清除自由项
    Dim sKey As String
    Dim m_ele As IXMLDOMElement
    tmpDomBody.loadXML strXMLbody & "</body>"
    Set m_ele = tmpDomBody.selectSingleNode("body")
    sKey = m_ele.getAttribute("sKey")
    Set m_ele = Nothing
    Set tmpDomBody = Nothing
    
    If LCase(rstFile(iCount).Name) = "ccode" Then
        strXMLbody = strXMLbody & "<R K='ccode' V=''/>" & Chr(13)
        strXMLbody = strXMLbody & "<R K='ibatch' V=''/>" & Chr(13)
        strXMLbody = strXMLbody & "<R K='inufts' V=''/>" & Chr(13)
        strXMLbody = strXMLbody & "<R K='cvmivencode' V=''/>" & Chr(13)
        strXMLbody = strXMLbody & "<R K='cvmivenname' V=''/>" & Chr(13)
        strXMLbody = strXMLbody & "<R K='cvenabbname' V=''/>" & Chr(13)
    Else
        For iCount = 0 To rstFile.Fields.Count - 1
            ''跟踪型存货会删除掉税率 ,保质期
            If LCase(rstFile(iCount).Name) <> "iexpiratdatecalcu" And LCase(rstFile(iCount).Name) <> "cmassunit" And LCase(rstFile(iCount).Name) <> "imassdate" And LCase(rstFile(iCount).Name) <> "itaxrate" And LCase(rstFile(iCount).Name) <> "cgroupcode" And LCase(rstFile(iCount).Name) <> "igrouptype" Then
                If Not ((sKey = "cbatch" Or sKey = "ccode") And Left(LCase(rstFile(iCount).Name), 5) = "cfree") Then
                    strXMLbody = strXMLbody & "<R K='" & rstFile.Fields(iCount).Name & "' V=''/>" & Chr(13)
                End If
            End If
        Next iCount
    End If
    strXMLbody = strXMLbody & "</body>"
    domBody.loadXML strXMLbody
    If rstFile.State <> adStateClosed Then
        rstFile.Close
    End If
End Sub

'获得计量单位名称
Public Function GetUnitName(strUnitCode As String, strGroupCode As String, iChangRate As Double, strErrMsg As String) As String
    Dim RecTemp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo err_Handle
    strSQL = " SELECT cComunitCode, cComUnitName, cGroupCode, ISNULL(iChangRate,0) AS iChangRate , iProportion,  iNumber" & _
             " From ComputationUnit " & _
             " WHERE cComunitCode =N'" & strUnitCode & "'" '& _
             " AND cGroupCode=N'" & strGroupCode & "'"
             
    RecTemp.CursorLocation = adUseClient
    RecTemp.Open strSQL, m_Conn
    If RecTemp.RecordCount > 0 Then
        strUnitCode = RecTemp!cComunitCode
        iChangRate = RecTemp!iChangRate
        GetUnitName = RecTemp!cComUnitName
    End If
    Exit Function
err_Handle:
    strErrMsg = err.Description
End Function
Public Function getKL2(kl2 As Double, domHead As DOMDocument, domBody As DOMDocument, CN As ADODB.Connection, ErrMsg As String, Optional sKey As String) As Boolean
'Public Function getKL2(kl2 As Double, domBody As DOMDocument, CN As ADODB.Connection, errMsg As String, Optional sKey As String) As Boolean
    kl2 = CDbl(Val(GetBodyItemValue(domBody, "kl2", 0)))
    If Not clsSys.bQuantityDisRate Then getKL2 = True: Exit Function
    Dim cmdGetKL As New ADODB.Command
    Dim parGetKl As New ADODB.Parameter
    Dim strInvCode As String, iQuantity As Double, strCusCode As String
    
        
    On Error GoTo DoErr
    
    strInvCode = GetBodyItemValue(domBody, "cinvcode", 0)
    strCusCode = GetHeadItemValue(domHead, "ccuscode")
    If IsMissing(sKey) Or LCase(sKey) = "iquantity" Then
        iQuantity = Val(GetBodyItemValue(domBody, "iquantity", 0))
    Else
        iQuantity = Val(GetBodyItemValue(domBody, "iinvexchrate", 0)) * Val(GetBodyItemValue(domBody, "inum", 0))
    End If
    cmdGetKL.CommandText = "SA_FetchQuantityDisRate"
    cmdGetKL.CommandType = adCmdStoredProc
    Set parGetKl = cmdGetKL.CreateParameter("sCusCode", adVarChar, adParamInput, 20, strCusCode)
    cmdGetKL.Parameters.Append parGetKl
    Set parGetKl = cmdGetKL.CreateParameter("sInvCode", adVarChar, adParamInput, 20, strInvCode)
    cmdGetKL.Parameters.Append parGetKl
    Set parGetKl = cmdGetKL.CreateParameter("nQuantity", adDouble, adParamInput, , str(iQuantity))
    cmdGetKL.Parameters.Append parGetKl
    Set parGetKl = cmdGetKL.CreateParameter("nDisRate", adDouble, adParamOutput)
    cmdGetKL.Parameters.Append parGetKl
    cmdGetKL.ActiveConnection = CN
    cmdGetKL.Execute
    kl2 = IIf(IsNull(cmdGetKL.Parameters("nDisRate").value), 0, cmdGetKL.Parameters("nDisRate").value)
    kl2 = CDbl(FormatNum(kl2, clsSys.cKLBit))
    Set cmdGetKL = Nothing
    getKL2 = True
    Exit Function
DoErr:
    ErrMsg = err.Description
    getKL2 = False
    Set cmdGetKL = Nothing
End Function
Public Function CheckMulitUnitInf(strUnitCode As String, strGroupCode As String, iChangRate As Double, strErrMsg As String, bCheckUnitCode As Boolean) As String
    Dim RecTemp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo err_Handle
    strSQL = " SELECT cComunitCode, cComUnitName, cGroupCode, ISNULL(iChangRate,0) AS iChangRate , iProportion,  iNumber" & _
             " From ComputationUnit " & _
             " WHERE 1=1 " & _
             IIf(bCheckUnitCode = True, " AND (cComunitCode =N'" & strUnitCode & "' OR cComUnitName =N'" & strUnitCode & "')", "") & _
             IIf(strGroupCode = "", "", " AND cGroupCode=N'" & strGroupCode & "'") & _
             IIf(bCheckUnitCode = True, "", " AND iChangRate=" & iChangRate)  '& " AND bMainUnit=0"
             
    RecTemp.CursorLocation = adUseClient
    RecTemp.Open strSQL, m_Conn
    If RecTemp.RecordCount > 0 Then
        strUnitCode = RecTemp!cComunitCode
        iChangRate = RecTemp!iChangRate
        CheckMulitUnitInf = RecTemp!cComUnitName
    Else
        If bCheckUnitCode Then
           strErrMsg = GetString("U8.SA.USSASERVER.clscommcheck.00488") 'zh-CN：输入的计量单位编码在计量单位组中不存在!
        Else
           strErrMsg = GetString("U8.SA.USSASERVER.clscommcheck.00489") 'zh-CN：输入的换算率在计量单位组中不存在!
        End If
           
    End If
    Exit Function
err_Handle:
    strErrMsg = err.Description
End Function

''格式化数字
Private Function FormatNum(NumValue, Dec As Integer) As Variant
    Dim tmpStr As String, tmpFString As String
    
    If Dec < 0 Then Dec = 0
    tmpFString = "####0" & IIf(Dec = 0, "", ".") & String(Val(Dec), "0")
    FormatNum = (Format(Val(NumValue), tmpFString))
    'FormatNum = tmpStr
End Function
''由客户存货对照转化为存货编码
Private Function GetcInvFromcCusinv(sKey As String, domHead As DOMDocument, domBody As DOMDocument, ErrMsg As String, Optional bClear As Boolean = False) As Boolean
    Dim strCusCode As String
    Dim tmpRst As New ADODB.Recordset
    Dim ele As IXMLDOMElement, strSQL As String
    Dim strKeyValue As String
    Dim strCusInvCode As String
    
    strCusCode = GetHeadItemValue(domHead, "ccuscode")
    strKeyValue = GetBodyItemValue(domBody, sKey, 0)
    strCusInvCode = GetBodyItemValue(domBody, "ccusinvcode", 0)
    If strCusCode = "" Then
        Set tmpRst = Nothing
        ErrMsg = GetString("U8.SA.USSASERVER.clscommcheck.00490") 'zh-CN：请先填写客户!
        Exit Function
    End If
    strSQL = "select inventory.cinvcode,cinvname,isnull(ccusinvcode,N'') as ccusinvcode,isnull(ccusinvname,N'') as ccusinvname from CusInvContrapose inner join inventory on " _
        & " CusInvContrapose.cinvcode=inventory.cinvcode  where ccuscode=N'" + strCusCode + "' and (ccusinvcode=N'" + strKeyValue + "' or ccusinvname=N'" + strKeyValue + "')"
    tmpRst.Open strSQL, m_Conn, adOpenKeyset, adLockReadOnly
    If Not tmpRst.EOF Then
        tmpRst.Filter = sKey & "='" + strKeyValue + "'"
        If tmpRst.EOF Then tmpRst.Filter = ""
        If tmpRst.RecordCount > 1 Then
            tmpRst.Filter = "ccusinvcode='" + strCusInvCode + "'"
        End If
        If tmpRst.EOF Then tmpRst.Filter = ""
    Else
        tmpRst.Close
        Set tmpRst = Nothing
        If bClear = False Then
            If strKeyValue <> "" Then
                ErrMsg = GetString("U8.SA.USSASERVER.clscommcheck.00491") 'zh-CN：填写错误，客户存货对照表中不存在匹配的记录!
            End If
        Else
            Set ele = domBody.selectSingleNode("//z:row")
            ele.setAttribute "cinvcode", ""
            ele.setAttribute "cinvname", ""
        End If
        Exit Function
    End If
    Set ele = domBody.selectSingleNode("//z:row")
    If strVouchType = "98" Or strVouchType = "99" Or strVouchType = "92" Or strVouchType = "95" Or GetBodyItemValue(domBody, "cinvcode", 0) = tmpRst("cinvcode") Then
        sKey = "cinvcode2"
    Else
        sKey = "cinvcode"
    End If
    ele.setAttribute "cinvcode", tmpRst("cinvcode")
    ele.setAttribute "cinvname", tmpRst("cinvname")
    ele.setAttribute "ccusinvcode", tmpRst("ccusinvcode")
    ele.setAttribute "ccusinvname", tmpRst("ccusinvname")
    GetcInvFromcCusinv = True
    tmpRst.Close
    Set tmpRst = Nothing
End Function
''由客户存货对照转化为存货编码
Private Function GetCusInvFromInv(strXMl As String, strCusCode As String, strcInvcode As String) As Boolean
    Dim domXml As New DOMDocument
    Dim tmpRst As New ADODB.Recordset
    Dim ele As IXMLDOMElement, strSQL As String
    Dim strKeyValue As String
    Dim nod As IXMLDOMNode
    Dim lngPos1 As Long
    Dim lngpos2 As String
    
    Dim strValue As String
    tmpRst.CursorLocation = adUseClient
    If strCusCode <> "" Then
        lngPos1 = InStr(3, strXMl, "<")
'        lngpos2 = InStr(lngPos1, strXMl, " ")
'        If lngpos2 = 0 Then lngpos2 = InStr(lngPos1, strXMl, ">") - 1
        strValue = Trim(Mid(strXMl, lngPos1 + 1, 4))
        Dim strCusInvCode As String
        domXml.loadXML strXMl & " </" & strValue & ">"
        Set nod = domXml.documentElement.selectSingleNode("R[@K='ccusinvcode']")
        If Not nod Is Nothing Then
            strCusInvCode = nod.Attributes.getNamedItem("V").Text
        End If
        strSQL = "select isnull(ccusinvcode,N'') as ccusinvcode,isnull(ccusinvname,N'') as ccusinvname from CusInvContrapose  where ccuscode=N'" + strCusCode + "' and cinvcode=N'" + strcInvcode + "'" + IIf(strCusInvCode = "", "", " and ccusinvcode='" + strCusInvCode + "'")
        tmpRst.Open strSQL, m_Conn, adOpenKeyset, adLockReadOnly
        If Not tmpRst.EOF Then
            strXMl = strXMl & "<R K='ccusinvcode' V='" & tmpRst("ccusinvcode") & "'/>" & Chr(13)
            strXMl = strXMl & "<R K='ccusinvname' V='" & tmpRst("ccusinvname") & "'/>" & Chr(13)
        Else
            strXMl = strXMl & "<R K='ccusinvcode' V=''/>" & Chr(13)
            strXMl = strXMl & "<R K='ccusinvname' V=''/>" & Chr(13)
        End If
        tmpRst.Close
    End If
    Set tmpRst = Nothing
    Set domXml = Nothing
End Function
''由客户存货对照转化为存货编码
Private Function GetCusInvCodeName(strCusCode As String, strInvCode As String, strCusInvCode As String, strCusInvName As String) As Boolean
    Dim tmpRst As New ADODB.Recordset
    tmpRst.CursorLocation = adUseClient
    Dim strSQL As String
   
    If strCusCode <> "" And strInvCode <> "" Then
        strSQL = "select isnull(ccusinvcode,N'') as ccusinvcode,isnull(ccusinvname,N'') as ccusinvname from CusInvContrapose  where ccuscode=N'" + strCusCode + "' and cinvcode=N'" + strInvCode + "'"
        tmpRst.Open strSQL, m_Conn, adOpenKeyset, adLockReadOnly
        If Not tmpRst.EOF Then
            strCusInvCode = IIf(IsNull(tmpRst.Fields("ccusinvcode")), "", tmpRst.Fields("ccusinvcode"))
            strCusInvName = IIf(IsNull(tmpRst.Fields("ccusinvname")), "", tmpRst.Fields("ccusinvname"))
        Else
            strCusInvCode = ""
            strCusInvName = ""
        End If
        tmpRst.Close
    Else
        strCusInvCode = ""
        strCusInvName = ""
    End If
    Set tmpRst = Nothing
End Function

'获得CRM数据，拼成销售数据, 根据输入的商机编号获得指定的单据视图。
Private Function GetVouchFromCRM(oDomHead As DOMDocument, oDomBody As DOMDocument, sKey As String, ErrMsg As String) As Boolean
    Dim oCRMServer As Object
    Dim rst As New ADODB.Recordset
    
    Dim sKeyValue As String
    Dim oDOMCRM As New DOMDocument, oDOMTmp As DOMDocument
    Dim oDOMHNew As New DOMDocument, oDOMBNew As New DOMDocument  '新单据对象
    Dim oDOMHTmp As New DOMDocument, oDOMBTmp As New DOMDocument  '新单据临时对象
    Dim eleSA As IXMLDOMElement, eleCRM As IXMLDOMElement, eleH As IXMLDOMElement
    Dim oNDList As IXMLDOMNodeList
    
    Dim NdList As IXMLDOMNodeList
    Dim ndrs As IXMLDOMNode
    Dim nd As IXMLDOMNode, ndTmp As IXMLDOMNode
    Dim eleTmp As IXMLDOMElement, eleR As IXMLDOMElement
    Dim i As Long
    Dim bClearBody As Boolean
    
    sKeyValue = GetHeadItemValue(oDomHead, sKey)
    If sKeyValue = "" Then
        Set eleH = oDomHead.selectSingleNode("//z:row")
        eleH.setAttribute "ioppcourseid", ""
        eleH.setAttribute "icourseseq", ""
        eleH.setAttribute "csalphasename", ""
        eleH.setAttribute "coppcode", ""
        Exit Function
    End If
    
    Set oCRMServer = CreateObject("UFCRMSRVSALE.clsOpportuntity")
    
    
    
    If oCRMServer Is Nothing Then
        ErrMsg = GetString("U8.SA.USSASERVER.clscommcheck.00492") 'zh-CN：无法创建对象UFCRMSRVSALE.clsOpportuntity!
        Set rst = Nothing
        Exit Function
    End If
    '861测试问题139685
    Set rst = oCRMServer.GetRecordRst(clsSys.objlogin, False, strVouchType, sKeyValue, False, oDomHead.selectSingleNode("//z:row"), "W")
    'Set Rst = oCRMServer.GetRecordRst(clsSys.objlogin, False, strVouchType, sKeyValue, False, oDomHead.selectSingleNode("//z:row"), "T")
    If rst.EOF Then
        ErrMsg = GetString("U8.SA.USSASERVER.clscommcheck.00493") 'zh-CN：商机录入不正确,请检查!
        Set eleH = oDomHead.selectSingleNode("//z:row")
        eleH.setAttribute "ioppcourseid", ""
        eleH.setAttribute "icourseseq", ""
        eleH.setAttribute "csalphasename", ""
        eleH.setAttribute "coppcode", ""
        If rst.State = 1 Then rst.Close
        Set rst = Nothing
        Set oCRMServer = Nothing
        Exit Function
    Else
        If LCase(rst("ccuscode")) <> LCase(GetHeadItemValue(oDomHead, "ccuscode")) And LCase(GetHeadItemValue(oDomHead, "ccuscode")) <> "" Then
            ErrMsg = GetString("U8.SA.USSASERVER.clscommcheck.00494") 'zh-CN：商机上的客户不是当前客户,请检查!
            Set eleH = oDomHead.selectSingleNode("//z:row")
            eleH.setAttribute "ioppcourseid", ""
            eleH.setAttribute "icourseseq", ""
            eleH.setAttribute "csalphasename", ""
            If rst.State = 1 Then rst.Close
            Set rst = Nothing
            Set oCRMServer = Nothing
            Exit Function
        End If
    End If
'    If strVouchType = "99" Then
'        GetVouchFromCRM = True
'        GoTo DOExit
'    End If
    If strVouchType = "99" Then
        bClearBody = False
    Else
        If GetHeadItemValue(oDomHead, "bclearbody") = "1" Then
            bClearBody = True
        Else
            bClearBody = False
        End If
    End If
    rst.Save oDOMCRM, adPersistXML
    oDOMHNew.loadXML oDomHead.xml: oDOMBNew.loadXML oDomBody.xml   '初始化新单据数据
    
    
    
    '获取单据头数据
    Set eleSA = oDOMHNew.selectSingleNode("//z:row")
    Set eleCRM = oDOMCRM.selectNodes("//z:row").Item(0)
    oDomHead.loadXML oDOMHNew.xml
    '添加校验客户
    Set eleSA = oDOMHNew.selectSingleNode("//z:row")
    eleSA.setAttribute "ddate", "1900-12-31"
    Call VerifyHead("ccusabbname", oDOMHNew, oDomBody)
    
    Set eleTmp = oDomHead.selectNodes("//z:row").Item(0)
    For Each eleR In oDOMHNew.selectNodes("//R")
        '如果原始数据没有对应项目则采用校验返回值 ，主要避免更改部门、业务员
        If GetElEAtrVal(eleTmp, LCase(eleR.getAttribute("K"))) = "" Then
            eleTmp.setAttribute LCase(eleR.getAttribute("K")), eleR.getAttribute("V")
        End If
    Next eleR
          
    If strVouchType = "99" Then
        Set eleH = oDomHead.selectSingleNode("//z:row")
        eleH.setAttribute "cexch_name", clsSys.sCurrencyName
        eleH.setAttribute "iexchrate", 1
        Set oCRMServer = Nothing
        Set oCRMServer = CreateObject("UFCRMSRVSALE.clsOppPro")
        Set rst = oCRMServer.GetOppLastPro(clsSys.objlogin, GetHeadItemValue(oDomHead, "coppcode"))
        If Not rst Is Nothing Then
            If Not rst.EOF Then
                eleH.setAttribute "icourseseq", GetRstVal(rst, "icourseseq", False)
                eleH.setAttribute "ioppcourseid", GetRstVal(rst, "ioppcourseid", False)
                'eleH.setAttribute "csalphasename", GetRstVal(rst, "iCourseSeq") & GetRstVal(rst, "csalphasename")
                eleH.setAttribute "csalphasename", GetRstVal(rst, "csalphasename", False)
                ''strXMl = strXMl & "<R K='csalphasename' V='" & rstFile("iCourseSeq") & rstFile("csalphasename") & "'/>" & Chr(13)
            Else
                eleH.setAttribute "icourseseq", ""
                eleH.setAttribute "ioppcourseid", ""
                eleH.setAttribute "csalphasename", ""
            End If
        End If
        GoTo DOExit
    End If
    If IsNull(rst("cinvcode")) Then
        GoTo DOExit
    End If
    '    '根据bClearBody 处理员表体 删除临时表体记录
    Set ndrs = oDOMBNew.selectSingleNode("//rs:data")
    Set NdList = oDOMBNew.selectNodes("//z:row")
    If bClearBody = True Then
        If (Not ndrs Is Nothing) And NdList.length <> 0 Then
            For Each nd In NdList
                Select Case Trim(UCase(GetNodeAtrVal(nd, "editprop")))
                    Case "A"  '新增加的
                        ndrs.removeChild nd
                    Case "M", ""   '数据库中已经存在的部分
                        Set eleTmp = nd
                        eleTmp.setAttribute "editprop", "D"
                    Case "D"
                        '不处理
                End Select
            Next
        End If
    End If
    
    '获取单据体数据
    Set nd = oDOMBNew.selectSingleNode("//rs:data")
    If nd Is Nothing Then
        Set eleTmp = oDOMBNew.createElement("rs:data")
        oDOMBNew.documentElement.appendChild eleTmp
    End If
    
    Set oDOMTmp = New DOMDocument
    oDOMTmp.loadXML oDOMBNew.xml
    Set ndrs = oDOMTmp.selectSingleNode("//rs:data")
    For Each nd In oDOMTmp.selectNodes("//z:row")
        ndrs.removeChild nd
    Next nd
'    Dim domBodyLine As New DOMDocument
    
    Dim dblQuo As Double
    Dim dblSale As Double
    Dim dblSett As Double
    Dim dblKL As Double
    Dim dblCusMinPrice As Double
    Dim strQtyPrice As String
    
    strQtyPrice = clsSys.GetSysDicOption("SA", "bquantityprice")  ' getAccinformation(m_Conn, "SA", "bquantityprice")
    If strQtyPrice = "" Then strQtyPrice = "True"
    
    Set NdList = oDOMCRM.selectNodes("//z:row")
    i = oDOMBNew.selectNodes("//z:row").length
    ''组织
    oDOMHTmp.loadXML oDomHead.xml
    Set eleH = oDomHead.selectSingleNode("//z:row")
    For Each eleCRM In NdList
        
            Set nd = oDOMBNew.selectSingleNode("//rs:data")
            Set eleSA = oDOMBNew.createElement("z:row")
            eleSA.setAttribute "xmlns:z", "#RowsetSchema"
        If GetElEAtrVal(eleSA, "cinvcode") <> "" Then
            
            oDOMBTmp.loadXML oDOMTmp.xml
            Set ndTmp = oDOMBTmp.selectSingleNode("//rs:data")
            ndTmp.appendChild eleSA.cloneNode(True)
            ErrMsg = ""
            ErrMsg = Me.BodyCheck("cinvname2", oDOMBTmp, oDOMHTmp, 1)
            If ErrMsg = "" Then
                For Each eleR In oDOMBTmp.selectNodes("//R")
                    If Left(LCase(eleR.getAttribute("K")), 5) <> "cfree" Then
                        eleSA.setAttribute LCase(eleR.getAttribute("K")), eleR.getAttribute("V")
                    End If
                Next eleR
                
                eleSA.setAttribute "autoid", ""
                eleSA.setAttribute "id", ""
                eleSA.setAttribute "editprop", "A"
                eleSA.setAttribute "dpremodate", Format(GetElEAtrVal(eleH, "ddate"), "yyyy-mm-dd")
                eleSA.setAttribute "dpredate", Format(GetElEAtrVal(eleH, "ddate"), "yyyy-mm-dd")
                eleSA.setAttribute "kl", "100"
                eleSA.setAttribute "kl2", "100"
                eleSA.setAttribute "dkl1", "0"
                eleSA.setAttribute "dkl2", "0"
                nd.appendChild eleSA  ''在新数据中增加行
                
                oDOMBTmp.loadXML oDOMTmp.xml
                Set ndTmp = oDOMBTmp.selectSingleNode("//rs:data")
                ndTmp.appendChild eleSA.cloneNode(True)
                dblCusMinPrice = 0
                dblQuo = 0
                ErrMsg = CountPrice(oDOMHTmp, oDOMBTmp, dblQuo, dblSale, dblSett, dblKL, dblCusMinPrice, CBool(strQtyPrice))
                eleSA.setAttribute "fcusminprice", CStr(dblCusMinPrice)
                ErrMsg = o_cac.BodyCheck("iquotedprice", oDOMBTmp, oDOMHTmp, 1, strVouchType)
                If ErrMsg <> "" Then
                    GoTo DOExit
                End If
                Set eleTmp = oDOMBNew.selectNodes("//z:row").Item(i)
                For Each eleR In oDOMBTmp.selectNodes("//R")
                    eleTmp.setAttribute LCase(eleR.getAttribute("K")), eleR.getAttribute("V")
                Next eleR
                If Val(GetElEAtrVal(eleH, "itaxrate")) = 0 Then
                    
                    eleH.setAttribute "itaxrate", GetElEAtrVal(eleTmp, "itaxrate")
                End If
                ''cellcheck 存货已取得必要信息。
'                Set domBodyLine = Voucher.GetLineDom(nRow)
                i = i + 1
            End If
        End If
    Next eleCRM
    'HeadCheck(vsKey, vdomHead, vstrVouchType, Optional vDomBody) As String
   
    Set oDomBody = oDOMBNew
    GetVouchFromCRM = True
DOExit:
    o_cac.vbExchType = clsSys.IsCal(GetHeadItemValue(oDomHead, "cexch_name"))
    Call o_cac.HeadCheck("iexchrate", oDomHead, strVouchType, oDomBody)
    If rst.State = 1 Then rst.Close
    Set rst = Nothing
    Set oCRMServer = Nothing
    Exit Function
DoErr:
    ErrMsg = GetStringPara("U8.SA.USSASERVER.clscommcheck.00497", err.Description) 'Para zh-CN：取crm数据发生错误,{0}
    If rst.State = 1 Then rst.Close
    Set rst = Nothing
    Set oCRMServer = Nothing
End Function
Private Function SeekStock(oVouchPreDom As DOMDocument, ele As IXMLDOMElement) As Boolean
    Dim blnLP As Boolean
    Dim strSosid As String
    Dim strIBatch As String
    Dim strWhCode As String
    Dim strInvCode As String
    Dim strFree1 As String, strFree2 As String, strFree3 As String, strFree4 As String, strFree5 As String
    Dim strFree6 As String, strFree7 As String, strFree8 As String, strFree9 As String, strFree10 As String
    Dim strBatch As String
    Dim dVouQuantity As Double, dVouNum As Double
    Dim eleBatch As IXMLDOMElement
    
    If LCase(GetElEAtrVal(ele, "cSRPolicy")) = "lp" Then
        blnLP = True
    Else
        blnLP = False
    End If
    
    strIBatch = GetElEAtrVal(ele, "ibatch")
    strWhCode = GetElEAtrVal(ele, "cwhcode")
    strInvCode = GetElEAtrVal(ele, "cinvcode")
    strBatch = GetElEAtrVal(ele, "cbatch")
    strFree1 = GetElEAtrVal(ele, "cfree1")
    strFree2 = GetElEAtrVal(ele, "cfree2")
    strFree3 = GetElEAtrVal(ele, "cfree3")
    strFree4 = GetElEAtrVal(ele, "cfree4")
    strFree5 = GetElEAtrVal(ele, "cfree5")
    strFree6 = GetElEAtrVal(ele, "cfree6")
    strFree7 = GetElEAtrVal(ele, "cfree7")
    strFree8 = GetElEAtrVal(ele, "cfree8")
    strFree9 = GetElEAtrVal(ele, "cfree9")
    strFree10 = GetElEAtrVal(ele, "cfree10")
    strSosid = Val(GetElEAtrVal(ele, "isosid"))
    dVouQuantity = GetElEAtrVal(ele, "iquantity")
    dVouNum = Val(GetElEAtrVal(ele, "inum"))
    Set eleBatch = oVouchPreDom.selectSingleNode("//row[@cwhcode='" + strWhCode + "' and @cinvcode='" + strInvCode + "' and @cfree1='" + strFree1 _
        + "' and @cfree2='" + strFree2 + "' and @cfree3='" + strFree3 + "' and @cfree4='" + strFree4 + "' and @cfree5='" + strFree5 _
        + "' and @cfree6='" + strFree6 + "' and @cfree7='" + strFree7 + "' and @cfree8='" + strFree8 + "' and @cfree9='" + strFree9 _
        + "' and @cfree10='" + strFree10 + "'" & IIf(strIBatch = "", " and @cbatch='" + strBatch + "'" + IIf(blnLP, " and @isosid='" + strSosid + "'", ""), " and @ibatch='" + strIBatch + "'") + "]")
    If eleBatch Is Nothing Then
        Set eleBatch = oVouchPreDom.createElement("row")
        eleBatch.setAttribute "cwhcode", strWhCode
        eleBatch.setAttribute "cinvcode", strInvCode
        eleBatch.setAttribute "cfree1", strFree1
        eleBatch.setAttribute "cfree2", strFree2
        eleBatch.setAttribute "cfree3", strFree3
        eleBatch.setAttribute "cfree4", strFree4
        eleBatch.setAttribute "cfree5", strFree5
        eleBatch.setAttribute "cfree6", strFree6
        eleBatch.setAttribute "cfree7", strFree7
        eleBatch.setAttribute "cfree8", strFree8
        eleBatch.setAttribute "cfree9", strFree9
        eleBatch.setAttribute "cfree10", strFree10
        eleBatch.setAttribute "cbatch", strBatch
        eleBatch.setAttribute "ibatch", strIBatch
        eleBatch.setAttribute "isosid", strSosid
        eleBatch.setAttribute "iquantity", -1 * dVouQuantity
        eleBatch.setAttribute "inum", -1 * dVouNum
        oVouchPreDom.documentElement.appendChild eleBatch
    Else
        eleBatch.setAttribute "iquantity", Val(GetElEAtrVal(eleBatch, "iquantity")) - dVouQuantity
        eleBatch.setAttribute "inum", Val(eleBatch.getAttribute("inum")) - dVouNum
    End If
End Function
Private Function GetPreiBatchInf(strVouchType As String, CN As ADODB.Connection, ele As IXMLDOMElement, _
   sPreBatch As String, iPreQuan As Double, iPreNum As Double, ErrMsg As String, itype As Integer) As Boolean
    Dim rst As New ADODB.Recordset
    Dim strSQL As String, sName As String
    If itype = 0 Then
        sName = "isnull(ibatch,0) as batch"
    ElseIf itype = 1 Then
        sName = "isnull(cbatch,'') as batch"
    End If
    strSQL = "select " + sName + ",isnull(iquantity,0) as iPreQuan,isnull(inum,0) as iPreNum from "
    Select Case strVouchType
        Case "05", "06", "00"
            strSQL = strSQL + " dispatchlists where idlsid =" & GetElEAtrVal(ele, "idlsid")
        Case "26", "27", "28", "29"
            strSQL = strSQL + " salebillvouchs where autoid =" & GetElEAtrVal(ele, "autoid")
    End Select
    rst.Open strSQL, CN, adOpenForwardOnly, adLockReadOnly
    If Not rst.EOF Then
        sPreBatch = rst("batch")
        iPreQuan = rst("iPreQuan")
        iPreNum = rst("iPreNum")
        GetPreiBatchInf = True
    Else
        ErrMsg = GetString("U8.SA.ussaserver.errmsg.Vouchchange") '"单据已经发生改变，请刷新! "
    End If
    rst.Close
    Set rst = Nothing
End Function

