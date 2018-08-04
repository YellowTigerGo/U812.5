/*=========================== View EF_PrintPolicy_VCH =============================*/
print 'EF_PrintPolicy_VCH' 
if exists (select * from sysobjects where id = object_id(N'[dbo].[EF_PrintPolicy_VCH]') and sysstat & 0xf = 2)
     drop view [dbo].[EF_PrintPolicy_VCH]
GO

create view EF_PrintPolicy_VCH
AS
select LEFT(PolicyID,6) as vouchertype,VchID as vouchercode,SUM(total) as total from PrintPolicy_VCH group by LEFT(PolicyID,6),
VchID

GO

/*=========================== View EF_Inventory =============================*/
print 'EF_Inventory' 
if exists (select * from sysobjects where id = object_id(N'[dbo].[EF_Inventory]') and sysstat & 0xf = 2)
     drop view [dbo].[EF_Inventory]
GO

create view EF_Inventory
as
select 
inventory.cinvcode,cinvname,cinvstd,cinvaddcode,cinvmnemcode,(inventory.cinvccode) as cinvccode,(inventoryclass.cinvcname) as cinvcname,creplaceitem,(inventory.cposition) as cposition,(position.cposname) as cposname,(inventory.bsale) as bsale,(inventory.bexpsale) as bexpsale,bself,bcomsume,bproducing,bservice,(inventory.bproxyforeign) as bproxyforeign,(inventory.bptomodel) as bptomodel,baccessary,bbondedinv,(inventory.batomodel) as batomodel,(inventory.bcheckitem) as bcheckitem,(inventory.bequipment) as bequipment,(inventory.bpiece) as bpiece,(inventory.bplaninv) as bplaninv,(inventory.bsrvfittings) as bsrvfittings,(inventory.bsrvitem) as bsrvitem,bsrvproduct,bprjmat,binvasset,iimptaxrate,(ex_ciqarchive.cciqcode) as cciqcode,cciqname,finvciqexch,itaxrate,iinvweight,ivolume,iinvrcost,iinvsprice,bpurchase,bsuitretail,bconsiderfreestock,froundfactor,Inventory_Sub.bchecksubitemcost,inearrejectdays,bimport,fmaterialcost,iinvscost,iinvlscost,iinvncost,(inventory.cvencode) as cvencode,(vendor.cvenname) as cvenname,cpurpersoncode,(purperson.cpersonname) as cpurpersonname,iinvadvance,iinvbatch,(inventory.isafenum) as isafenum,(inventory.itopsum) as itopsum,(inventory.ilowsum) as ilowsum,ioverstock,cinvabc,binvrohs,binvbatch,binvquality,bserial,binventrust,binvoverstock,(inventory.dsdate) as dsdate,cwunit,(wunit.ccomunitname) as cwunitname,(inventory.dedate) as dedate,fgrossw,fheight,flength,cvgroupcode,(vgroup.cgroupname) as cvgroupname,cvunit,(vunit.ccomunitname) as cvunitname,cwgroupcode,(wgroup.cgroupname) as cwgroupname,fwidth,binvtype,fcurllaborcost,fcurlvarmanucost,fcurlfixmanucost,fcurlomcost,fnextllaborcost,fnextlvarmanucost,fnextlfixmanucost,fnextlomcost,iinvmpcost,iwarrantyperiod,iwarrantyunit,(inventory.cquality) as cquality,(inventory.ccreateperson) as ccreateperson,dinvcreatedatetime,(inventory.cmodifyperson) as cmodifyperson,cinvappdocno,(inventory.dmodifydate) as dmodifydate,cvaluetype,foutexcess,finexcess,imassdate,iexpiratdatecalcu,iwarndays,fexpensesexch,btrack,bbarcode,(aa_authclass.caccode) as caccode,(aa_authclass.cacname) as cacname,(inventory.cbarcode) as cbarcode,cinvdefine1,cinvdefine2,cinvdefine3,cinvdefine4,cinvdefine5,cinvdefine6,cinvdefine7,cinvdefine8,cinvdefine9,cinvdefine10,cinvdefine11,cinvdefine12,cinvdefine13,cinvdefine14,cinvdefine15,cinvdefine16,bfree1,bfree2,bfree3,bfree4,bfree5,bfree6,bfree7,bfree8,bfree9,bfree10,bcontrolfreerange1,bcontrolfreerange2,bcontrolfreerange3,bcontrolfreerange4,bcontrolfreerange5,bcontrolfreerange6,bcontrolfreerange7,bcontrolfreerange8,bcontrolfreerange9,bcontrolfreerange10,bconfigfree1,bconfigfree2,bconfigfree3,bconfigfree4,bconfigfree5,bconfigfree6,bconfigfree7,bconfigfree8,bconfigfree9,bconfigfree10,bcheckfree1,bcheckfree2,bcheckfree3,bcheckfree4,bcheckfree5,bcheckfree6,bcheckfree7,bcheckfree8,bcheckfree9,bcheckfree10,bpurpricefree1,bpurpricefree2,bpurpricefree3,bpurpricefree4,bpurpricefree5,bpurpricefree6,bpurpricefree7,bpurpricefree8,bpurpricefree9,bpurpricefree10,bsalepricefree1,bsalepricefree2,bsalepricefree3,bsalepricefree4,bsalepricefree5,bsalepricefree6,bsalepricefree7,bsalepricefree8,bsalepricefree9,bsalepricefree10,bompricefree1,bompricefree2,bompricefree3,bompricefree4,bompricefree5,bompricefree6,bompricefree7,bompricefree8,bompricefree9,bompricefree10,(inventory.cgroupcode) as cgroupcode,(computationgroup.cgroupname) as cgroupname,(inventory.igrouptype) as igrouptype,(inventory.ccomunitcode) as ccomunitcode,(computationunit.ccomunitname) as ccomunitname,csacomunitcode,(saunit.ccomunitname) as csacomunitname,(puunit.ccomunitname) as cpucomunitname,cpucomunitcode,cstcomunitcode,(stunit.ccomunitname) as cstcomunitname,ccacomunitcode,(caunit.ccomunitname) as ccacomunitname,cproductunit,(productunit.ccomunitname) as cproductunitname,(inventory.cfrequency) as cfrequency,(inventory.ifrequency) as ifrequency,(inventory.idays) as idays,cdtperiod,bdtwarninv,bpropertycheck,(inventory.cmassunit) as cmassunit,(inventory.dlastdate) as dlastdate,forderuplimit,finvoutuplimit,idrawbatch,fminsplit,iwastage,bsolitude,centerprise,caddress,cfile,clabel,ccheckout,clicence,bspecialties,(inventory.cdefwarehouse) as cdefwarehouse,(warehouse.cwhname) as cwhname,fretailprice,iexpsalerate,iadvancedate,ccurrencyname,cproduceaddress,cproducenation,cregisterno,centerno,cpackingtype,cenglishname,bperioddt,cpreparationtype,ccommodity,cnotpatentname,(atp_projectmain.cprojectexplain) as cprojectexplain,(irecipebatch) as irecipebatch,bpuquota,cinvprojectcode,bcheckbsatp,(inventory.iropmethod) as iropmethod,(inventory.ibatchrule) as ibatchrule,(inventory.fsubscribepoint) as fsubscribepoint,(inventory.iassureprovidedays) as iassureprovidedays,fsupplymulti,(inventory.fvagquantity) as fvagquantity,ialteradvance,isupplyday,isupplytype,ibomexpandunittype,csrpolicy,falterbasenum,fminsupply,fmaxsupply,irequiretrackstyle,cplanmethod,(mps_timefence.tfcode) as tfcode,(mps_timefence.description) as description,(v_mps_atp.description) as v_mps_atp_description,cinvpersoncode,(v_mps_atp.atpcode) as atpcode,(person.cpersonname) as cpersonname,cinvdepcode,cdepname,breplan,(inventory.brop) as brop,cengineerfigno,ioverlapday,iplantfday,iacceptdelaydays,iacceptearlydays,icheckatp,(inventory.bbommain) as bbommain,(inventory.bbomsub) as bbomsub,(inventory.bproductbill) as bproductbill,(inventory.bmps) as bmps,bintotalcost,bbillunite,binvkeypart,bcutmantissa,(inventory.iteststyle) as iteststyle,(inventory.idtmethod) as idtmethod,(inventory.fdtrate) as fdtrate,(inventory.fdtnum) as fdtnum,(inventory.cdtunit) as cdtunit,(dtunitaliastable.ccomunitname) as cdtunitname,(inventory.idtstyle) as idtstyle,(inventory.iqtmethod) as iqtmethod,(qmcheckproject.cprojectcode) as cprojectcode,(qmcheckproject.cprojectname) as cprojectname,cdtaql,breceiptbydt,binbyprocheck,(qmrandorcheck.crulename) as crulename,(inventory.crulecode) as crulecode,itestrule,idtlevel,cshopunit,(shopunit.ccomunitname) as cshopunitname,bimportmedicine,bfirstbusimedicine,bforeexpland,cinvplancode,fconvertrate,dreplacedate,binvmodel,bkccutmantissa,fbuyexcess,fprjmatlimit,isurenesstype,idatetype,idatesum,idynamicsurenesstype,ibestrowsum,ipercentumsum,pictureguid,bisattachfile,bbatchcreate,bbatchproperty1,bbatchproperty2,bbatchproperty3,bbatchproperty4,bbatchproperty5,bbatchproperty6,bbatchproperty7,bbatchproperty8,bbatchproperty9,bbatchproperty10,imaterialscycle,iplancheckday,btracksalebill,idrawtype,bsckeyprojections,(inventory_sub.isupplyperiodtype) as isupplyperiodtype,(inventory_sub.itimebucketid) as itimebucketid,(inventory_sub.iavailabilitydate) as iavailabilitydate,bcheckbatch,ipfbatchqty,iallocateprintdgt,iplandefault,(ibigday) as ibigday,(ibigmonth) as ibigmonth,(ismallday) as ismallday,(ismallmonth) as ismallmonth,bmngoldpart,ioldpartmngrule,(bfeaturematch) as bfeaturematch,(bproducebyfeatureallocate) as bproducebyfeatureallocate,(bmaintenance) as bmaintenance,(imaintenancecycle) as imaintenancecycle,(imaintenancecycleunit) as imaintenancecycleunit,bcoupon,bstorecard,bprocessproduct,bprocessmaterial   
from Inventory  
Left Join InventoryClass on Inventory.cInvCCode=InventoryClass.cInvCCode  
Left Join ex_ciqarchive  on Inventory.cciqcode=ex_ciqarchive.cciqcode  
Left Join Vendor on Inventory.cVenCode=Vendor.cVenCode  
Left Join Position on Inventory.cPosition=Position.cPosCode  
Left Join computationGroup on Inventory.cGroupcode=computationGroup.cGroupCode  
Left Join  ComputationUnit on Inventory.cComUnitCode=ComputationUnit.cComUnitCode  
Left Join ComputationUnit as PUUnit on Inventory.cPUComUnitcode=PUUnit.cComUnitCode  
Left Join ComputationUnit  as SAUnit on Inventory.cSAComUnitCode=SAUnit.cComUnitCode  
Left Join ComputationUnit as STUnit on Inventory.cSTComUnitCode=STUnit.cComUnitCode  
Left Join ComputationUnit as CAUnit on Inventory.cCAComUnitCode=CAUnit.cComUnitCode  
Left Join Warehouse on Inventory.cDefWareHouse=Warehouse.cWhCode  
Left join ComputationUnit as DTUnitAliasTable on  Inventory.cDTUnit=DTUnitAliasTable.cComUnitCode  
Left join ComputationGroup as WGroup on  Inventory.cWGroupCode=WGroup.cGroupCode  
Left join ComputationGroup as VGroup on  Inventory.cVGroupCode=VGroup.cGroupCode  
Left Join ComputationUnit as VUnit on Inventory.cVUnit=VUnit.cComUnitCode  
Left Join ComputationUnit as WUnit on Inventory.cWUnit=WUnit.cComUnitCode  
Left Join ComputationUnit as ProductUnit on Inventory.cProductUnit=ProductUnit.cComUnitCode  
Left Join ComputationUnit as ShopUnit on Inventory.cShopUnit=ShopUnit.cComUnitCode  
Left Join Person   on Inventory.cInvPersonCode=Person.cPersonCode  
Left Join Person  as PurPerson on Inventory.cPurPersonCode=PurPerson.cPersonCode  
Left Join department   on Inventory.cInvDepCode=department.cDepCode  
Left Join mps_timefence  on Inventory.iInvTfId=mps_timefence.TfId   
Left join qmCheckProject on  Inventory.iQTMethod=qmCheckProject.Id  
left join AA_AuthClass on Inventory.iid=AA_AuthClass.id  
Left Join ATP_ProjectMain on Inventory.cInvProjectCode=ATP_ProjectMain.cProjectCode  
Left Join v_mps_atp  on Inventory.iInvATPId=v_mps_atp.ATPId   
Left Join QMRANDORCHECK  on Inventory.cRuleCode=QMRANDORCHECK.cRuleCode  
Left Join Inventory_Sub on Inventory.cInvCode=Inventory_Sub.cInvSubCode  
Left Join Inventory_extradefine on Inventory.cInvCode=Inventory_extradefine.cInvCode

GO
/*=========================== View EF_V_fitemss97 =============================*/
print 'EF_V_fitemss97' 
if exists (select * from sysobjects where id = object_id(N'[dbo].[EF_V_fitemss97]') and sysstat & 0xf = 2)
     drop view [dbo].[EF_V_fitemss97]
GO

create view EF_V_fitemss97
AS
SELECT f.citemcode,f.citemname,f.citemccode,c.citemcname,f.客户 as ccusname,f.规格 as citemgg,f.总图号 as citemth 

FROM fitemss97 f
left outer join fitemss97class c on f.citemccode=c.cItemCcode

GO
--
--drop view [dbo].[V_EF_ProjectMRP]
--drop view [dbo].[V_EF_ProjectMRPs]
--drop view [dbo].[V_List_EF_ProjectMRP]
--
--select * from  [dbo].[V_EF_ProjectMRP]
--select * from [dbo].[V_EF_ProjectMRPs]
--select * from [dbo].[V_List_EF_ProjectMRP]

 
print '1 项目需求计划表头视图 dbo.V_EF_ProjectMRP... '
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_EF_ProjectMRP]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_EF_ProjectMRP]
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
CREATE VIEW [dbo].[V_EF_ProjectMRP]
AS
SELECT
a.id,
a.ccode,
a.ddate,
a.cmaker,
a.cmakerddate,
a.cmodifer,
a.cmodiferDate,
a.cmodifier,
a.dmoddate,
convert(nvarchar,a.dmodifysystime,120) as dmodifysystime,
a.cverifier,
a.dverifydate,
a.ccloser,
a.dcloserdate,
a.vt_id,
CONVERT(char, CONVERT(money, a.ufts), 2) AS ufts ,
a.cvouchtype,
a.t_cdepcode,					--部门编码
e.cDepName as t_cdepname,		--部门名称 
a.t_cpersoncode,				--人员编码
i.cPersonName  as t_cpersonname,  --人员名称 
a.t_ccuscode,					--客户编码
f.cCusName as t_ccusname,		--客户名称
a.t_cvencode,					--供应商编码
g.cVenName  as t_cvenname,		--供应商名称
a.t_cwhcode,					--仓库编码
h.cWhName  as t_cwhname,		--仓库名称
a.t_cinvcode,					----物料号（存货编码）
a.t_cfree1,						--存货自由项1
a.t_cfree2,						--存货自由项2
a.t_cfree3,						--存货自由项3
a.t_cfree4,						--存货自由项4
a.t_cfree5,						--存货自由项5
a.t_cfree6,						--存货自由项6
a.t_cfree7,						--存货自由项7
a.t_cfree8,						--存货自由项8
a.t_cfree9,						--存货自由项9
a.t_cfree10,					--存货自由项10
b.cInvName as t_cinvname, 		--名称（存货名称）
a.ireturncount,
a.iswfcontrolled,
a.iverifystate,
a.VoucherId,
a.VoucherCode,
a.VoucherType,
a.define1,
a.define2,
a.define3,
a.define4,
a.define5,
a.define6,
a.define7,
a.define8,
a.define9,
a.define10,
a.define11,
a.define12,
a.define13,
a.define14,
a.define15,
a.define16
,a.citem_class
,fitem.citem_name as citem_cname
,a.citemcode
,im.citemname
,im.ccusname
,im.citemth
,im.citemgg
,a.ipqty
,a.cmemo
,a.cppcode
,case when isnull(a.cppcode,'')='' then '未执行' else '已执行' end as bzx
,P.Total AS iprintcount
,case when ISNULL(a.dnverifytime,'')='' then a.dverifydate else convert(nvarchar,a.dnverifytime,120) end as dnverifytime
from  EF_ProjectMRP  a
LEFT OUTER JOIN inventory b on  a.t_cinvcode=b.cinvcode					-- inventory存货档案表
LEFT OUTER JOIN Department e on a.t_cdepcode=e.cDepCode					--部门表关联
LEFT OUTER JOIN Person  i on a.t_cpersoncode=i.cPersonCode				--人员表关联
LEFT OUTER JOIN Customer f on a.t_ccuscode=f.cCusCode					--客户表关联
LEFT OUTER JOIN Vendor g on a.t_cvencode=g.cVenCode						--供应商关联
LEFT OUTER JOIN Warehouse  h on a.t_cwhcode=h.cWhCode					--仓库表关联
LEFT OUTER JOIN EF_V_fitemss97 im on a.citemcode=im.citemcode
LEFT OUTER JOIN fitem ON A.cItem_class=fitem.citem_class
LEFT OUTER JOIN EF_PrintPolicy_VCH P ON A.ccode=p.vouchercode AND a.cVouchType=P.vouchertype
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


print '2 项目需求计划表体视图 dbo.V_EF_ProjectMRPs... '
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_EF_ProjectMRPs]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_EF_ProjectMRPs]
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
CREATE VIEW [dbo].[V_EF_ProjectMRPs]
AS
SELECT
a.autoid,
a.id,
a.b_cdepcode,					--部门编码
e.cDepName as b_cdepname,		--部门名称 
a.b_cpersoncode,				--人员编码
i.cPersonName  as b_cpersonname,  --人员名称 
a.b_ccuscode,					--客户编码
f.cCusName as b_ccusname,		--客户名称
a.b_cvencode,					--供应商编码
g.cVenName  as b_cvenname,		--供应商名称
a.b_cwhcode,					--仓库编码
h.cWhName  as b_cwhname,		--仓库名称
a.b_cinvcode,					----物料号（存货编码）
b.cInvName as b_cinvname, 		--名称（存货名称）
b.cInvStd as b_cinvstd,          --规格
U.ccomunitname as b_ccomunitname,  --主计量单位
a.b_cfree1,						--存货自由项1
a.b_cfree2,						--存货自由项2
a.b_cfree3,						--存货自由项3
a.b_cfree4,						--存货自由项4
a.b_cfree5,						--存货自由项5
a.b_cfree6,						--存货自由项6
a.b_cfree7,						--存货自由项7
a.b_cfree8,						--存货自由项8
a.b_cfree9,						--存货自由项9
a.b_cfree10,					--存货自由项10
b.cInvDefine1
,b.cInvDefine2
,b.cInvDefine3
,b.cInvDefine4
,b.cInvDefine5
,b.cInvDefine6
,b.cInvDefine7
,b.cInvDefine8
,b.cInvDefine9
,b.cInvDefine10
,b.cInvDefine11
,b.cInvDefine12
,b.cInvDefine13
,b.cInvDefine14
,b.cInvDefine15
,b.cInvDefine16
,a.define22,
a.define23,
a.define24,
a.define25,
a.define26,
a.define27,
a.define28,
a.define29,
a.define30,
a.define31,
a.define32,
a.define33,
a.define34,
a.define35,
a.define36,
a.define37

,a.iinvexchrate
,a.auxunitcode
,uf.cComUnitName as auxunitname
,a.cpart
,a.iunitqty
,a.cperform
,a.iqty
,a.cbmemo
,a.coutsourced
,a.cbcloser
,a.cbclosedate
,a.illqty
,isnull(a.iqty,0)-isnull(a.illqty,0) as iwlyqty
,a.cinvname_cad
,a.irowno
from  EF_ProjectMRPs   a
LEFT OUTER JOIN inventory b on  a.b_cinvcode=b.cinvcode					-- inventory存货档案表
LEFT OUTER JOIN ComputationUnit U ON U.cComunitCode=B.cComUnitCode
LEFT OUTER JOIN ComputationUnit UF on UF.cComunitCode=A.AuxUnitCode
LEFT OUTER JOIN Department e on a.b_cdepcode=e.cDepCode					--部门表关联
LEFT OUTER JOIN Person  i on a.b_cpersoncode=i.cPersonCode				--人员表关联
LEFT OUTER JOIN Customer f on a.b_ccuscode=f.cCusCode					--客户表关联
LEFT OUTER JOIN Vendor g on a.b_cvencode=g.cVenCode						--供应商关联
LEFT OUTER JOIN Warehouse  h on a.b_cwhcode=h.cWhCode					--仓库表关联
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

print '3 项目需求计划列表视图 dbo.V_List_EF_ProjectMRP... '
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_List_EF_ProjectMRP]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_List_EF_ProjectMRP]
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
CREATE VIEW [dbo].[V_List_EF_ProjectMRP]
AS
SELECT
a.id,--id
a.ccode,--单据编号
a.ddate,--单据日期
a.cmaker,--制单人
a.cmakerddate,--制单日期
a.cmodifer,--变更人
a.cmodiferDate,--变更日期
a.cmodifier,
a.dmoddate,
a.dmodifysystime,
a.cverifier,--审核人
a.dverifydate,--审核日期
a.ccloser,
a.dcloserdate,
a.vt_id,--vt_id
a.ufts,--ufts
a.cvouchtype,--cvouchtype
a.t_cdepcode,--部门编码
a.t_cdepname,--部门名称
a.t_cpersoncode,--人员编码
a.t_cpersonname,--人员名称
a.t_ccuscode,--客户编码
a.t_ccusname,--客户名称
a.t_cvencode,--供应商编码
a.t_cvenname,--供应商名称
a.t_cwhcode,--仓库编码
a.t_cwhname,--仓库名称
a.t_cinvcode,--存货编码
a.t_cfree1,						--存货自由项1
a.t_cfree2,						--存货自由项2
a.t_cfree3,						--存货自由项3
a.t_cfree4,						--存货自由项4
a.t_cfree5,						--存货自由项5
a.t_cfree6,						--存货自由项6
a.t_cfree7,						--存货自由项7
a.t_cfree8,						--存货自由项8
a.t_cfree9,						--存货自由项9
a.t_cfree10,					--存货自由项10
a.t_cinvname,--存货名称
a.ireturncount,--ireturncount
a.iswfcontrolled,--iswfcontrolled
a.iverifystate,--iverifystate
a.VoucherId,--VoucherId
a.VoucherCode,--VoucherCode
a.VoucherType,--VoucherType
a.define1,--define1
a.define2,--define2
a.define3,--define3
a.define4,--define4
a.define5,--define5
a.define6,--define6
a.define7,--define7
a.define8,--define8
a.define9,--define9
a.define10,--define10
a.define11,--define11
a.define12,--define12
a.define13,--define13
a.define14,--define14
a.define15,--define15
a.define16,--define16
a.citem_class
,a.citem_cname
,a.citemcode
,a.citemname
,a.ccusname
,a.citemth
,a.citemgg
,a.ipqty
,a.cmemo
,a.cppcode
,a.bzx
,a.iprintcount
,a.dnverifytime
,
b.autoid,--autoid
--b.id,--id
b.b_cdepcode,--部门编码
b.b_cdepname,--部门名称
b.b_cpersoncode,--人员编码
b.b_cpersonname,--人员名称
b.b_ccuscode,--客户编码
b.b_ccusname,--客户名称
b.b_cvencode,--供应商编码
b.b_cvenname,--供应商名称
b.b_cwhcode,--仓库编码
b.b_cwhname,--仓库名称
b.b_cinvcode,--存货编码
b.b_cfree1,						--存货自由项1
b.b_cfree2,						--存货自由项2
b.b_cfree3,						--存货自由项3
b.b_cfree4,						--存货自由项4
b.b_cfree5,						--存货自由项5
b.b_cfree6,						--存货自由项6
b.b_cfree7,						--存货自由项7
b.b_cfree8,						--存货自由项8
b.b_cfree9,						--存货自由项9
b.b_cfree10,					--存货自由项10
b.cInvDefine1
,b.cInvDefine2
,b.cInvDefine3
,b.cInvDefine4
,b.cInvDefine5
,b.cInvDefine6
,b.cInvDefine7
,b.cInvDefine8
,b.cInvDefine9
,b.cInvDefine10
,b.cInvDefine11
,b.cInvDefine12
,b.cInvDefine13
,b.cInvDefine14
,b.cInvDefine15
,b.cInvDefine16
,b.b_cinvname,--存货名称
b.b_cinvstd,          --规格
b.b_ccomunitname,
b.define22,--define22
b.define23,--define23
b.define24,--define24
b.define25,--define25
b.define26,--define26
b.define27,--define27
b.define28,--define28
b.define29,--define29
b.define30,--define30
b.define31,--define31
b.define32,--define32
b.define33,--define33
b.define34,--define34
b.define35,--define35
b.define36,--define36
b.define37--define37
,b.iinvexchrate
,b.auxunitcode
,b.auxunitname
,b.cpart
,b.iunitqty
,b.cperform
,b.iqty
,b.cbmemo
,b.coutsourced
,b.cbcloser
,b.cbclosedate
,b.illqty
,b.iwlyqty
,b.cinvname_cad
,b.irowno
from  V_EF_ProjectMRP   a
LEFT OUTER JOIN V_EF_ProjectMRPs b on  a.id=b.id					 

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

--
--drop view [dbo].[V_EF_ProjectMRPChanged]
--drop view [dbo].[V_EF_ProjectMRPChangeds]
--drop view [dbo].[V_List_EF_ProjectMRPChanged]
--
--select * from  [dbo].[V_EF_ProjectMRPChanged]
--select * from [dbo].[V_EF_ProjectMRPChangeds]
--select * from [dbo].[V_List_EF_ProjectMRPChanged]

 
print '1 项目需求计划变更表头视图 dbo.V_EF_ProjectMRPChanged... '
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_EF_ProjectMRPChanged]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_EF_ProjectMRPChanged]
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
CREATE VIEW [dbo].[V_EF_ProjectMRPChanged]
AS
SELECT
a.id,
a.ccode,
a.ddate,
a.cmaker,
a.cmakerddate,
a.cmodifer,
a.cmodiferDate,
a.cmodifier,
a.dmoddate,
convert(nvarchar,a.dmodifysystime,120) as dmodifysystime,
a.cverifier,
a.dverifydate,
a.ccloser,
a.dcloserdate,
a.vt_id,
CONVERT(char, CONVERT(money, a.ufts), 2) AS ufts ,
a.cvouchtype,
a.t_cdepcode,					--部门编码
e.cDepName as t_cdepname,		--部门名称 
a.t_cpersoncode,				--人员编码
i.cPersonName  as t_cpersonname,  --人员名称 
a.t_ccuscode,					--客户编码
f.cCusName as t_ccusname,		--客户名称
a.t_cvencode,					--供应商编码
g.cVenName  as t_cvenname,		--供应商名称
a.t_cwhcode,					--仓库编码
h.cWhName  as t_cwhname,		--仓库名称
a.t_cinvcode,					----物料号（存货编码）
a.t_cfree1,						--存货自由项1
a.t_cfree2,						--存货自由项2
a.t_cfree3,						--存货自由项3
a.t_cfree4,						--存货自由项4
a.t_cfree5,						--存货自由项5
a.t_cfree6,						--存货自由项6
a.t_cfree7,						--存货自由项7
a.t_cfree8,						--存货自由项8
a.t_cfree9,						--存货自由项9
a.t_cfree10,					--存货自由项10
b.cInvName as t_cinvname, 		--名称（存货名称）
a.ireturncount,
a.iswfcontrolled,
a.iverifystate,
a.VoucherId,
a.VoucherCode,
a.VoucherType,
a.define1,
a.define2,
a.define3,
a.define4,
a.define5,
a.define6,
a.define7,
a.define8,
a.define9,
a.define10,
a.define11,
a.define12,
a.define13,
a.define14,
a.define15,
a.define16
,a.citem_class
,a.citem_cname
,a.citemcode
,a.citemname
,im.ccusname
,im.citemth
,im.citemgg
,a.ipqty
,a.cmemo
,a.cpmrpcode
,case when ISNULL(a.dnverifytime,'')='' then a.dverifydate else convert(nvarchar,a.dnverifytime,120) end as dnverifytime
from  EF_ProjectMRPChanged  a
LEFT OUTER JOIN inventory b on  a.t_cinvcode=b.cinvcode					-- inventory存货档案表
LEFT OUTER JOIN Department e on a.t_cdepcode=e.cDepCode					--部门表关联
LEFT OUTER JOIN Person  i on a.t_cpersoncode=i.cPersonCode				--人员表关联
LEFT OUTER JOIN Customer f on a.t_ccuscode=f.cCusCode					--客户表关联
LEFT OUTER JOIN Vendor g on a.t_cvencode=g.cVenCode						--供应商关联
LEFT OUTER JOIN Warehouse  h on a.t_cwhcode=h.cWhCode					--仓库表关联
LEFT OUTER JOIN EF_V_fitemss97 im on a.citemcode=im.citemcode
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


print '2 项目需求计划变更表体视图 dbo.V_EF_ProjectMRPChangeds... '
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_EF_ProjectMRPChangeds]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_EF_ProjectMRPChangeds]
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
CREATE VIEW [dbo].[V_EF_ProjectMRPChangeds]
AS
SELECT
a.autoid,
a.id,
a.b_cdepcode,					--部门编码
e.cDepName as b_cdepname,		--部门名称 
a.b_cpersoncode,				--人员编码
i.cPersonName  as b_cpersonname,  --人员名称 
a.b_ccuscode,					--客户编码
f.cCusName as b_ccusname,		--客户名称
a.b_cvencode,					--供应商编码
g.cVenName  as b_cvenname,		--供应商名称
a.b_cwhcode,					--仓库编码
h.cWhName  as b_cwhname,		--仓库名称
a.b_cinvcode,					----物料号（存货编码）
b.cInvName as b_cinvname, 		--名称（存货名称）
b.cInvStd as b_cinvstd,          --规格
U.ccomunitname as b_ccomunitname,  --主计量单位
a.b_cfree1,						--存货自由项1
a.b_cfree2,						--存货自由项2
a.b_cfree3,						--存货自由项3
a.b_cfree4,						--存货自由项4
a.b_cfree5,						--存货自由项5
a.b_cfree6,						--存货自由项6
a.b_cfree7,						--存货自由项7
a.b_cfree8,						--存货自由项8
a.b_cfree9,						--存货自由项9
a.b_cfree10,					--存货自由项10
b.cInvDefine1
,b.cInvDefine2
,b.cInvDefine3
,b.cInvDefine4
,b.cInvDefine5
,b.cInvDefine6
,b.cInvDefine7
,b.cInvDefine8
,b.cInvDefine9
,b.cInvDefine10
,b.cInvDefine11
,b.cInvDefine12
,b.cInvDefine13
,b.cInvDefine14
,b.cInvDefine15
,b.cInvDefine16
,a.define22,
a.define23,
a.define24,
a.define25,
a.define26,
a.define27,
a.define28,
a.define29,
a.define30,
a.define31,
a.define32,
a.define33,
a.define34,
a.define35,
a.define36,
a.define37

,a.iinvexchrate
,a.auxunitcode
,uf.cComUnitName as auxunitname
,a.cpart
,a.iunitqty
,a.iunitqtyold
,a.cperform
,a.iqty
,a.iqtyold
,a.cbmemo
,a.coutsourced
,a.cbcloser
,a.cbclosedate
,a.cpmrpid
,a.cpmrpautoid
,a.irowno
from  EF_ProjectMRPChangeds   a
LEFT OUTER JOIN inventory b on  a.b_cinvcode=b.cinvcode					-- inventory存货档案表
LEFT OUTER JOIN ComputationUnit U ON U.cComunitCode=B.cComUnitCode
LEFT OUTER JOIN ComputationUnit UF on UF.cComunitCode=A.AuxUnitCode
LEFT OUTER JOIN Department e on a.b_cdepcode=e.cDepCode					--部门表关联
LEFT OUTER JOIN Person  i on a.b_cpersoncode=i.cPersonCode				--人员表关联
LEFT OUTER JOIN Customer f on a.b_ccuscode=f.cCusCode					--客户表关联
LEFT OUTER JOIN Vendor g on a.b_cvencode=g.cVenCode						--供应商关联
LEFT OUTER JOIN Warehouse  h on a.b_cwhcode=h.cWhCode					--仓库表关联
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

print '3 项目需求计划变更列表视图 dbo.V_List_EF_ProjectMRPChanged... '
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_List_EF_ProjectMRPChanged]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_List_EF_ProjectMRPChanged]
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
CREATE VIEW [dbo].[V_List_EF_ProjectMRPChanged]
AS
SELECT
a.id,--id
a.ccode,--单据编号
a.ddate,--单据日期
a.cmaker,--制单人
a.cmakerddate,--制单日期
a.cmodifer,--变更人
a.cmodiferDate,--变更日期
a.cmodifier,
a.dmoddate,
a.dmodifysystime,
a.cverifier,--审核人
a.dverifydate,--审核日期
a.ccloser,
a.dcloserdate,
a.vt_id,--vt_id
a.ufts,--ufts
a.cvouchtype,--cvouchtype
a.t_cdepcode,--部门编码
a.t_cdepname,--部门名称
a.t_cpersoncode,--人员编码
a.t_cpersonname,--人员名称
a.t_ccuscode,--客户编码
a.t_ccusname,--客户名称
a.t_cvencode,--供应商编码
a.t_cvenname,--供应商名称
a.t_cwhcode,--仓库编码
a.t_cwhname,--仓库名称
a.t_cinvcode,--存货编码
a.t_cfree1,						--存货自由项1
a.t_cfree2,						--存货自由项2
a.t_cfree3,						--存货自由项3
a.t_cfree4,						--存货自由项4
a.t_cfree5,						--存货自由项5
a.t_cfree6,						--存货自由项6
a.t_cfree7,						--存货自由项7
a.t_cfree8,						--存货自由项8
a.t_cfree9,						--存货自由项9
a.t_cfree10,					--存货自由项10
a.t_cinvname,--存货名称
a.ireturncount,--ireturncount
a.iswfcontrolled,--iswfcontrolled
a.iverifystate,--iverifystate
a.VoucherId,--VoucherId
a.VoucherCode,--VoucherCode
a.VoucherType,--VoucherType
a.define1,--define1
a.define2,--define2
a.define3,--define3
a.define4,--define4
a.define5,--define5
a.define6,--define6
a.define7,--define7
a.define8,--define8
a.define9,--define9
a.define10,--define10
a.define11,--define11
a.define12,--define12
a.define13,--define13
a.define14,--define14
a.define15,--define15
a.define16,--define16
a.citem_class
,a.citem_cname
,a.citemcode
,a.citemname
,a.ccusname
,a.citemth
,a.citemgg
,a.ipqty
,a.cmemo
,a.cpmrpcode
,a.dnverifytime
,
b.autoid,--autoid
--b.id,--id
b.b_cdepcode,--部门编码
b.b_cdepname,--部门名称
b.b_cpersoncode,--人员编码
b.b_cpersonname,--人员名称
b.b_ccuscode,--客户编码
b.b_ccusname,--客户名称
b.b_cvencode,--供应商编码
b.b_cvenname,--供应商名称
b.b_cwhcode,--仓库编码
b.b_cwhname,--仓库名称
b.b_cinvcode,--存货编码
b.b_cfree1,						--存货自由项1
b.b_cfree2,						--存货自由项2
b.b_cfree3,						--存货自由项3
b.b_cfree4,						--存货自由项4
b.b_cfree5,						--存货自由项5
b.b_cfree6,						--存货自由项6
b.b_cfree7,						--存货自由项7
b.b_cfree8,						--存货自由项8
b.b_cfree9,						--存货自由项9
b.b_cfree10,					--存货自由项10
b.cInvDefine1
,b.cInvDefine2
,b.cInvDefine3
,b.cInvDefine4
,b.cInvDefine5
,b.cInvDefine6
,b.cInvDefine7
,b.cInvDefine8
,b.cInvDefine9
,b.cInvDefine10
,b.cInvDefine11
,b.cInvDefine12
,b.cInvDefine13
,b.cInvDefine14
,b.cInvDefine15
,b.cInvDefine16
,b.b_cinvname,--存货名称
b.b_cinvstd,          --规格
b.b_ccomunitname,
b.define22,--define22
b.define23,--define23
b.define24,--define24
b.define25,--define25
b.define26,--define26
b.define27,--define27
b.define28,--define28
b.define29,--define29
b.define30,--define30
b.define31,--define31
b.define32,--define32
b.define33,--define33
b.define34,--define34
b.define35,--define35
b.define36,--define36
b.define37--define37
,b.iinvexchrate
,b.auxunitcode
,b.auxunitname
,b.cpart
,b.iunitqty
,b.iunitqtyold
,b.cperform
,b.iqty
,b.iqtyold
,b.cbmemo
,b.coutsourced
,b.cbcloser
,b.cbclosedate
,b.cpmrpid
,b.cpmrpautoid
,b.irowno
from  V_EF_ProjectMRPChanged   a
LEFT OUTER JOIN V_EF_ProjectMRPChangeds b on  a.id=b.id					 

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

--
--drop view [dbo].[V_EF_ProcurementPlan]
--drop view [dbo].[V_EF_ProcurementPlans]
--drop view [dbo].[V_List_EF_ProcurementPlan]
--
--select * from  [dbo].[V_EF_ProcurementPlan]
--select * from [dbo].[V_EF_ProcurementPlans]
--select * from [dbo].[V_List_EF_ProcurementPlan]

 
print '1 采购计划表头视图 dbo.V_EF_ProcurementPlan... '
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_EF_ProcurementPlan]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_EF_ProcurementPlan]
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
CREATE VIEW [dbo].[V_EF_ProcurementPlan]
AS
SELECT
a.id,
a.ccode,
a.ddate,
a.cmaker,
a.cmakerddate,
a.cmodifer,
a.cmodiferDate,
a.cmodifier,
a.dmoddate,
convert(nvarchar,a.dmodifysystime,120) as dmodifysystime,
a.cverifier,
a.dverifydate,
a.ccloser,
a.dcloserdate,
a.vt_id,
CONVERT(char, CONVERT(money, a.ufts), 2) AS ufts ,
a.cvouchtype,
a.t_cdepcode,					--部门编码
e.cDepName as t_cdepname,		--部门名称 
a.t_cpersoncode,				--人员编码
i.cPersonName  as t_cpersonname,  --人员名称 
a.t_ccuscode,					--客户编码
f.cCusName as t_ccusname,		--客户名称
a.t_cvencode,					--供应商编码
g.cVenName  as t_cvenname,		--供应商名称
a.t_cwhcode,					--仓库编码
h.cWhName  as t_cwhname,		--仓库名称
a.t_cinvcode,					----物料号（存货编码）
a.t_cfree1,						--存货自由项1
a.t_cfree2,						--存货自由项2
a.t_cfree3,						--存货自由项3
a.t_cfree4,						--存货自由项4
a.t_cfree5,						--存货自由项5
a.t_cfree6,						--存货自由项6
a.t_cfree7,						--存货自由项7
a.t_cfree8,						--存货自由项8
a.t_cfree9,						--存货自由项9
a.t_cfree10,					--存货自由项10
b.cInvName as t_cinvname, 		--名称（存货名称）
a.ireturncount,
a.iswfcontrolled,
a.iverifystate,
a.VoucherId,
a.VoucherCode,
a.VoucherType,
a.define1,
a.define2,
a.define3,
a.define4,
a.define5,
a.define6,
a.define7,
a.define8,
a.define9,
a.define10,
a.define11,
a.define12,
a.define13,
a.define14,
a.define15,
a.define16
,a.citem_class
,a.citem_cname
,a.citemcode
,a.citemname
,im.ccusname
,im.citemth
,im.citemgg
,a.ipqty
,a.cmemo
,a.cpmrpcode
,p.Total as iprintcount
,case when ISNULL(a.dnverifytime,'')='' then a.dverifydate else convert(nvarchar,a.dnverifytime,120) end as dnverifytime
from  EF_ProcurementPlan  a
LEFT OUTER JOIN inventory b on  a.t_cinvcode=b.cinvcode					-- inventory存货档案表
LEFT OUTER JOIN Department e on a.t_cdepcode=e.cDepCode					--部门表关联
LEFT OUTER JOIN Person  i on a.t_cpersoncode=i.cPersonCode				--人员表关联
LEFT OUTER JOIN Customer f on a.t_ccuscode=f.cCusCode					--客户表关联
LEFT OUTER JOIN Vendor g on a.t_cvencode=g.cVenCode						--供应商关联
LEFT OUTER JOIN Warehouse  h on a.t_cwhcode=h.cWhCode					--仓库表关联
LEFT OUTER JOIN EF_V_fitemss97 im on a.citemcode=im.citemcode
LEFT OUTER JOIN EF_PrintPolicy_VCH P ON A.ccode=p.vouchercode AND a.cVouchType=P.vouchertype
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


print '2 采购计划表体视图 dbo.V_EF_ProcurementPlans... '
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_EF_ProcurementPlans]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_EF_ProcurementPlans]
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
CREATE VIEW [dbo].[V_EF_ProcurementPlans]
AS
SELECT
a.autoid,
a.id,
a.b_cdepcode,					--部门编码
e.cDepName as b_cdepname,		--部门名称 
a.b_cpersoncode,				--人员编码
i.cPersonName  as b_cpersonname,  --人员名称 
a.b_ccuscode,					--客户编码
f.cCusName as b_ccusname,		--客户名称
a.b_cvencode,					--供应商编码
g.cVenName  as b_cvenname,		--供应商名称
a.b_cwhcode,					--仓库编码
h.cWhName  as b_cwhname,		--仓库名称
a.b_cinvcode,					----物料号（存货编码）
b.cInvName as b_cinvname, 		--名称（存货名称）
b.cInvStd as b_cinvstd,          --规格
U.ccomunitname as b_ccomunitname,  --主计量单位
a.b_cfree1,						--存货自由项1
a.b_cfree2,						--存货自由项2
a.b_cfree3,						--存货自由项3
a.b_cfree4,						--存货自由项4
a.b_cfree5,						--存货自由项5
a.b_cfree6,						--存货自由项6
a.b_cfree7,						--存货自由项7
a.b_cfree8,						--存货自由项8
a.b_cfree9,						--存货自由项9
a.b_cfree10,					--存货自由项10
b.cInvDefine1
,b.cInvDefine2
,b.cInvDefine3
,b.cInvDefine4
,b.cInvDefine5
,b.cInvDefine6
,b.cInvDefine7
,b.cInvDefine8
,b.cInvDefine9
,b.cInvDefine10
,b.cInvDefine11
,b.cInvDefine12
,b.cInvDefine13
,b.cInvDefine14
,b.cInvDefine15
,b.cInvDefine16
,a.define22,
a.define23,
a.define24,
a.define25,
a.define26,
a.define27,
a.define28,
a.define29,
a.define30,
a.define31,
a.define32,
a.define33,
a.define34,
a.define35,
a.define36,
a.define37

,a.iinvexchrate
,a.auxunitcode
,uf.cComUnitName as auxunitname
,a.cbmemo
,a.coutsourced
,a.imrpqty
,a.imrpqtyl
,a.isafenum
,a.istock
,a.iinqty
,a.ioutqty
,a.ikyl
,a.iminqty
,a.ijhqty
,a.ijyqty
,a.isjqty
,a.cbcloser
,a.cbclosedate
,a.iqglqty
,isnull(a.isjqty,0)-isnull(a.iqglqty,0) as iwqgqty--未请购量
,a.cpmrpautoid
,a.irowno
from  EF_ProcurementPlans   a
LEFT OUTER JOIN inventory b on  a.b_cinvcode=b.cinvcode					-- inventory存货档案表
LEFT OUTER JOIN ComputationUnit U ON U.cComunitCode=B.cComUnitCode
LEFT OUTER JOIN ComputationUnit UF on UF.cComunitCode=A.AuxUnitCode
LEFT OUTER JOIN Department e on a.b_cdepcode=e.cDepCode					--部门表关联
LEFT OUTER JOIN Person  i on a.b_cpersoncode=i.cPersonCode				--人员表关联
LEFT OUTER JOIN Customer f on a.b_ccuscode=f.cCusCode					--客户表关联
LEFT OUTER JOIN Vendor g on a.b_cvencode=g.cVenCode						--供应商关联
LEFT OUTER JOIN Warehouse  h on a.b_cwhcode=h.cWhCode					--仓库表关联
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

print '3 采购计划列表视图 dbo.V_List_EF_ProcurementPlan... '
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_List_EF_ProcurementPlan]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_List_EF_ProcurementPlan]
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
CREATE VIEW [dbo].[V_List_EF_ProcurementPlan]
AS
SELECT
a.id,--id
a.ccode,--单据编号
a.ddate,--单据日期
a.cmaker,--制单人
a.cmakerddate,--制单日期
a.cmodifer,--变更人
a.cmodiferDate,--变更日期
a.cmodifier,
a.dmoddate,
a.dmodifysystime,
a.cverifier,--审核人
a.dverifydate,--审核日期
a.ccloser,
a.dcloserdate,
a.vt_id,--vt_id
a.ufts,--ufts
a.cvouchtype,--cvouchtype
a.t_cdepcode,--部门编码
a.t_cdepname,--部门名称
a.t_cpersoncode,--人员编码
a.t_cpersonname,--人员名称
a.t_ccuscode,--客户编码
a.t_ccusname,--客户名称
a.t_cvencode,--供应商编码
a.t_cvenname,--供应商名称
a.t_cwhcode,--仓库编码
a.t_cwhname,--仓库名称
a.t_cinvcode,--存货编码
a.t_cfree1,						--存货自由项1
a.t_cfree2,						--存货自由项2
a.t_cfree3,						--存货自由项3
a.t_cfree4,						--存货自由项4
a.t_cfree5,						--存货自由项5
a.t_cfree6,						--存货自由项6
a.t_cfree7,						--存货自由项7
a.t_cfree8,						--存货自由项8
a.t_cfree9,						--存货自由项9
a.t_cfree10,					--存货自由项10
a.t_cinvname,--存货名称
a.ireturncount,--ireturncount
a.iswfcontrolled,--iswfcontrolled
a.iverifystate,--iverifystate
a.VoucherId,--VoucherId
a.VoucherCode,--VoucherCode
a.VoucherType,--VoucherType
a.define1,--define1
a.define2,--define2
a.define3,--define3
a.define4,--define4
a.define5,--define5
a.define6,--define6
a.define7,--define7
a.define8,--define8
a.define9,--define9
a.define10,--define10
a.define11,--define11
a.define12,--define12
a.define13,--define13
a.define14,--define14
a.define15,--define15
a.define16,--define16
a.citem_class
,a.citem_cname
,a.citemcode
,a.citemname
,a.ccusname
,a.citemth
,a.citemgg
,a.ipqty
,a.cmemo
,a.cpmrpcode
,a.iprintcount
,a.dnverifytime
,
b.autoid,--autoid
--b.id,--id
b.b_cdepcode,--部门编码
b.b_cdepname,--部门名称
b.b_cpersoncode,--人员编码
b.b_cpersonname,--人员名称
b.b_ccuscode,--客户编码
b.b_ccusname,--客户名称
b.b_cvencode,--供应商编码
b.b_cvenname,--供应商名称
b.b_cwhcode,--仓库编码
b.b_cwhname,--仓库名称
b.b_cinvcode,--存货编码
b.b_cfree1,						--存货自由项1
b.b_cfree2,						--存货自由项2
b.b_cfree3,						--存货自由项3
b.b_cfree4,						--存货自由项4
b.b_cfree5,						--存货自由项5
b.b_cfree6,						--存货自由项6
b.b_cfree7,						--存货自由项7
b.b_cfree8,						--存货自由项8
b.b_cfree9,						--存货自由项9
b.b_cfree10,					--存货自由项10
b.cInvDefine1
,b.cInvDefine2
,b.cInvDefine3
,b.cInvDefine4
,b.cInvDefine5
,b.cInvDefine6
,b.cInvDefine7
,b.cInvDefine8
,b.cInvDefine9
,b.cInvDefine10
,b.cInvDefine11
,b.cInvDefine12
,b.cInvDefine13
,b.cInvDefine14
,b.cInvDefine15
,b.cInvDefine16
,b.b_cinvname,--存货名称
b.b_cinvstd,          --规格
b.b_ccomunitname,
b.define22,--define22
b.define23,--define23
b.define24,--define24
b.define25,--define25
b.define26,--define26
b.define27,--define27
b.define28,--define28
b.define29,--define29
b.define30,--define30
b.define31,--define31
b.define32,--define32
b.define33,--define33
b.define34,--define34
b.define35,--define35
b.define36,--define36
b.define37--define37
,b.iinvexchrate
,b.auxunitcode
,b.auxunitname
,b.cbmemo
,b.coutsourced
,b.imrpqty
,b.imrpqtyl
,b.isafenum
,b.istock
,b.iinqty
,b.ioutqty
,b.ikyl
,b.iminqty
,b.ijhqty
,b.ijyqty
,b.isjqty
,b.cbcloser
,b.cbclosedate
,b.iqglqty
,b.iwqgqty--未请购量
,b.cpmrpautoid
,b.irowno
from  V_EF_ProcurementPlan   a
LEFT OUTER JOIN V_EF_ProcurementPlans b on  a.id=b.id					 

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

--
--drop view [dbo].[V_EF_Bid]
--drop view [dbo].[V_EF_Bids]
--drop view [dbo].[V_List_EF_Bid]
--
--select * from  [dbo].[V_EF_Bid]
--select * from [dbo].[V_EF_Bids]
--select * from [dbo].[V_List_EF_Bid]

 
print '1 项目投标报价表头视图 dbo.V_EF_Bid... '
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_EF_Bid]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_EF_Bid]
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
CREATE VIEW [dbo].[V_EF_Bid]
AS
SELECT
a.id,
a.ccode,
a.ddate,
a.cmaker,
a.cmakerddate,
a.cmodifer,
a.cmodiferDate,
a.cmodifier,
a.dmoddate,
convert(nvarchar,a.dmodifysystime,120) as dmodifysystime,
a.cverifier,
a.dverifydate,
a.ccloser,
a.dcloserdate,
a.vt_id,
CONVERT(char, CONVERT(money, a.ufts), 2) AS ufts ,
a.cvouchtype,
a.t_cdepcode,					--部门编码
e.cDepName as t_cdepname,		--部门名称 
a.t_cpersoncode,				--人员编码
i.cPersonName  as t_cpersonname,  --人员名称 
a.t_ccuscode,					--客户编码
f.cCusName as t_ccusname,		--客户名称
a.t_cvencode,					--供应商编码
g.cVenName  as t_cvenname,		--供应商名称
a.t_cwhcode,					--仓库编码
h.cWhName  as t_cwhname,		--仓库名称
a.t_cinvcode,					----物料号（存货编码）
a.t_cfree1,						--存货自由项1
a.t_cfree2,						--存货自由项2
a.t_cfree3,						--存货自由项3
a.t_cfree4,						--存货自由项4
a.t_cfree5,						--存货自由项5
a.t_cfree6,						--存货自由项6
a.t_cfree7,						--存货自由项7
a.t_cfree8,						--存货自由项8
a.t_cfree9,						--存货自由项9
a.t_cfree10,					--存货自由项10
b.cInvName as t_cinvname, 		--名称（存货名称）
a.ireturncount,
a.iswfcontrolled,
a.iverifystate,
a.VoucherId,
a.VoucherCode,
a.VoucherType,
a.define1,
a.define2,
a.define3,
a.define4,
a.define5,
a.define6,
a.define7,
a.define8,
a.define9,
a.define10,
a.define11,
a.define12,
a.define13,
a.define14,
a.define15,
a.define16
,a.citem_class
,a.citem_cname
,a.citemcode
,a.citemname
,im.ccusname
,im.citemth
,im.citemgg
,a.ipqty
,a.cmemo
,p.Total as iprintcount
,case when ISNULL(a.dnverifytime,'')='' then a.dverifydate else convert(nvarchar,a.dnverifytime,120) end as dnverifytime
from  EF_Bid  a
LEFT OUTER JOIN inventory b on  a.t_cinvcode=b.cinvcode					-- inventory存货档案表
LEFT OUTER JOIN Department e on a.t_cdepcode=e.cDepCode					--部门表关联
LEFT OUTER JOIN Person  i on a.t_cpersoncode=i.cPersonCode				--人员表关联
LEFT OUTER JOIN Customer f on a.t_ccuscode=f.cCusCode					--客户表关联
LEFT OUTER JOIN Vendor g on a.t_cvencode=g.cVenCode						--供应商关联
LEFT OUTER JOIN Warehouse  h on a.t_cwhcode=h.cWhCode					--仓库表关联
LEFT OUTER JOIN EF_V_fitemss97 im on a.citemcode=im.citemcode
LEFT OUTER JOIN EF_PrintPolicy_VCH P ON A.ccode=p.vouchercode AND a.cVouchType=P.vouchertype
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


print '2 项目投标报价表体视图 dbo.V_EF_Bids... '
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_EF_Bids]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_EF_Bids]
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
CREATE VIEW [dbo].[V_EF_Bids]
AS
SELECT
a.autoid,
a.id,
a.b_cdepcode,					--部门编码
e.cDepName as b_cdepname,		--部门名称 
a.b_cpersoncode,				--人员编码
i.cPersonName  as b_cpersonname,  --人员名称 
a.b_ccuscode,					--客户编码
f.cCusName as b_ccusname,		--客户名称
a.b_cvencode,					--供应商编码
g.cVenName  as b_cvenname,		--供应商名称
a.b_cwhcode,					--仓库编码
h.cWhName  as b_cwhname,		--仓库名称
a.b_cinvcode,					----物料号（存货编码）
b.cInvName as b_cinvname, 		--名称（存货名称）
b.cInvStd as b_cinvstd,          --规格
U.ccomunitname as b_ccomunitname,  --主计量单位
a.b_cfree1,						--存货自由项1
a.b_cfree2,						--存货自由项2
a.b_cfree3,						--存货自由项3
a.b_cfree4,						--存货自由项4
a.b_cfree5,						--存货自由项5
a.b_cfree6,						--存货自由项6
a.b_cfree7,						--存货自由项7
a.b_cfree8,						--存货自由项8
a.b_cfree9,						--存货自由项9
a.b_cfree10,					--存货自由项10
b.cInvDefine1
,b.cInvDefine2
,b.cInvDefine3
,b.cInvDefine4
,b.cInvDefine5
,b.cInvDefine6
,b.cInvDefine7
,b.cInvDefine8
,b.cInvDefine9
,b.cInvDefine10
,b.cInvDefine11
,b.cInvDefine12
,b.cInvDefine13
,b.cInvDefine14
,b.cInvDefine15
,b.cInvDefine16
,a.define22,
a.define23,
a.define24,
a.define25,
a.define26,
a.define27,
a.define28,
a.define29,
a.define30,
a.define31,
a.define32,
a.define33,
a.define34,
a.define35,
a.define36,
a.define37

,a.iinvexchrate
,a.auxunitcode
,uf.cComUnitName as auxunitname
,a.cmaterialclass
,a.icb
,a.ibj
,a.cbmemo
,a.cbcloser
,a.cbclosedate
,a.irowno
from  EF_Bids   a
LEFT OUTER JOIN inventory b on  a.b_cinvcode=b.cinvcode					-- inventory存货档案表
LEFT OUTER JOIN ComputationUnit U ON U.cComunitCode=B.cComUnitCode
LEFT OUTER JOIN ComputationUnit UF on UF.cComunitCode=A.AuxUnitCode
LEFT OUTER JOIN Department e on a.b_cdepcode=e.cDepCode					--部门表关联
LEFT OUTER JOIN Person  i on a.b_cpersoncode=i.cPersonCode				--人员表关联
LEFT OUTER JOIN Customer f on a.b_ccuscode=f.cCusCode					--客户表关联
LEFT OUTER JOIN Vendor g on a.b_cvencode=g.cVenCode						--供应商关联
LEFT OUTER JOIN Warehouse  h on a.b_cwhcode=h.cWhCode					--仓库表关联

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

print '3 项目投标报价列表视图 dbo.V_List_EF_Bid... '
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_List_EF_Bid]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_List_EF_Bid]
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
CREATE VIEW [dbo].[V_List_EF_Bid]
AS
SELECT
a.id,--id
a.ccode,--单据编号
a.ddate,--单据日期
a.cmaker,--制单人
a.cmakerddate,--制单日期
a.cmodifer,--变更人
a.cmodiferDate,--变更日期
a.cmodifier,
a.dmoddate,
a.dmodifysystime,
a.cverifier,--审核人
a.dverifydate,--审核日期
a.ccloser,
a.dcloserdate,
a.vt_id,--vt_id
a.ufts,--ufts
a.cvouchtype,--cvouchtype
a.t_cdepcode,--部门编码
a.t_cdepname,--部门名称
a.t_cpersoncode,--人员编码
a.t_cpersonname,--人员名称
a.t_ccuscode,--客户编码
a.t_ccusname,--客户名称
a.t_cvencode,--供应商编码
a.t_cvenname,--供应商名称
a.t_cwhcode,--仓库编码
a.t_cwhname,--仓库名称
a.t_cinvcode,--存货编码
a.t_cfree1,						--存货自由项1
a.t_cfree2,						--存货自由项2
a.t_cfree3,						--存货自由项3
a.t_cfree4,						--存货自由项4
a.t_cfree5,						--存货自由项5
a.t_cfree6,						--存货自由项6
a.t_cfree7,						--存货自由项7
a.t_cfree8,						--存货自由项8
a.t_cfree9,						--存货自由项9
a.t_cfree10,					--存货自由项10
a.t_cinvname,--存货名称
a.ireturncount,--ireturncount
a.iswfcontrolled,--iswfcontrolled
a.iverifystate,--iverifystate
a.VoucherId,--VoucherId
a.VoucherCode,--VoucherCode
a.VoucherType,--VoucherType
a.define1,--define1
a.define2,--define2
a.define3,--define3
a.define4,--define4
a.define5,--define5
a.define6,--define6
a.define7,--define7
a.define8,--define8
a.define9,--define9
a.define10,--define10
a.define11,--define11
a.define12,--define12
a.define13,--define13
a.define14,--define14
a.define15,--define15
a.define16,--define16
a.citem_class
,a.citem_cname
,a.citemcode
,a.citemname
,a.ccusname
,a.citemth
,a.citemgg
,a.ipqty
,a.cmemo
,a.iprintcount
,a.dnverifytime
,
b.autoid,--autoid
--b.id,--id
b.b_cdepcode,--部门编码
b.b_cdepname,--部门名称
b.b_cpersoncode,--人员编码
b.b_cpersonname,--人员名称
b.b_ccuscode,--客户编码
b.b_ccusname,--客户名称
b.b_cvencode,--供应商编码
b.b_cvenname,--供应商名称
b.b_cwhcode,--仓库编码
b.b_cwhname,--仓库名称
b.b_cinvcode,--存货编码
b.b_cfree1,						--存货自由项1
b.b_cfree2,						--存货自由项2
b.b_cfree3,						--存货自由项3
b.b_cfree4,						--存货自由项4
b.b_cfree5,						--存货自由项5
b.b_cfree6,						--存货自由项6
b.b_cfree7,						--存货自由项7
b.b_cfree8,						--存货自由项8
b.b_cfree9,						--存货自由项9
b.b_cfree10,					--存货自由项10
b.cInvDefine1
,b.cInvDefine2
,b.cInvDefine3
,b.cInvDefine4
,b.cInvDefine5
,b.cInvDefine6
,b.cInvDefine7
,b.cInvDefine8
,b.cInvDefine9
,b.cInvDefine10
,b.cInvDefine11
,b.cInvDefine12
,b.cInvDefine13
,b.cInvDefine14
,b.cInvDefine15
,b.cInvDefine16
,b.b_cinvname,--存货名称
b.b_cinvstd,          --规格
b.b_ccomunitname,
b.define22,--define22
b.define23,--define23
b.define24,--define24
b.define25,--define25
b.define26,--define26
b.define27,--define27
b.define28,--define28
b.define29,--define29
b.define30,--define30
b.define31,--define31
b.define32,--define32
b.define33,--define33
b.define34,--define34
b.define35,--define35
b.define36,--define36
b.define37--define37
,b.iinvexchrate
,b.auxunitcode
,b.auxunitname
,b.cmaterialclass
,b.icb
,b.ibj
,b.cbcloser
,b.cbclosedate
,B.cbmemo
,b.irowno
from  V_EF_Bid   a
LEFT OUTER JOIN V_EF_Bids b on  a.id=b.id					 

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

if (object_id('materialappvouchs_insert', 'tr') is not null)
drop trigger materialappvouchs_insert
go
--create trigger materialappvouchs_insert
--on materialappvouchs
--for insert 
--as
----定义变量
--declare @autoid nvarchar(50)='';
----在inserted表中查询已经插入记录信息
--select @autoid = cdefine23 from inserted;
--if @autoid<>''
--begin
--	update EF_ProjectMRPs set iLLQty=(select SUM(iquantity) as iqty from materialappvouchs where cdefine23=@autoid) where autoid=@autoid
--end
--go

if (object_id('materialappvouchs_update', 'tr') is not null)
drop trigger materialappvouchs_update
go
--create trigger materialappvouchs_update
--on materialappvouchs
--for update
--as
----定义变量
--declare @autoid nvarchar(50)='';
--if (update(iquantity))
--begin
--	select @autoid = cdefine23 from inserted;
--	if @autoid<>''
--	begin
--		update EF_ProjectMRPs set iLLQty=(select SUM(iquantity) as iqty from materialappvouchs where cdefine23=@autoid) where autoid=@autoid
--	end
--end

--go

if (object_id('materialappvouchs_delete', 'tr') is not null)
drop trigger materialappvouchs_delete
go
--create trigger materialappvouchs_delete
--on materialappvouchs
--for delete
--as
----定义变量
--declare @autoid nvarchar(50)='';
--select @autoid = cdefine23 from deleted;
--if @autoid<>''
--begin
--	update EF_ProjectMRPs set iLLQty=(select SUM(iquantity) as iqty from materialappvouchs where cdefine23=@autoid) where autoid=@autoid
--end
--go
/*=========================== View EF_V_PRT_XQCGZX =============================*/
print 'EF_V_PRT_XQCGZX' 
if exists (select * from sysobjects where id = object_id(N'[dbo].[EF_V_PRT_XQCGZX]') and sysstat & 0xf = 2)
     drop view [dbo].[EF_V_PRT_XQCGZX]
GO

create view EF_V_PRT_XQCGZX
AS
--需求采购执行统计表
select m.ddate,m.citemcode,m.citemname,m.ccusname,m.ccode,m.cpart,m.citemth,m.b_cinvcode,m.b_cinvname,m.b_cinvstd,m.b_ccomunitname,m.cperform
,m.cbmemo,m.iunitqty,m.ipqty,m.iqty,m.illqty,m.iwlyqty
,p.istock,p.isafenum,p.ikyl,p.ijyqty,p.iqglqty,p.b_cpersoncode as cgybm,p.b_cpersonname as cgy
,dd.iddqty,dd.fddmoney,dd.fddprice
,dh.idhqty,dh.fdhmoney,dh.fdhprice
,rk.irkqty,rk.frkmoney,rk.frkprice
 from V_List_EF_ProjectMRP m
 left outer join V_List_EF_ProcurementPlan p on m.autoid=p.cpmrpautoid and m.ccode=p.cpmrpcode and m.b_cinvcode=p.b_cinvcode
 left outer join (select SUM(iQuantity) as iddqty,SUM(iNatSum) as fddmoney,SUM(iNatSum)/SUM(iQuantity) as fddprice,cDefine23 
	from PO_Podetails where ISNULL(cDefine23,'')<>'' group by cDefine23) dd on p.autoid=dd.cDefine23
left outer join (select SUM(iQuantity) as idhqty,SUM(iTaxPrice) as fdhmoney,SUM(iTaxPrice )/SUM(iQuantity) as fdhprice,cDefine23 
	from PU_ArrivalVouchs where ISNULL(cDefine23,'')<>'' group by cDefine23) dh on p.autoid=dh.cDefine23
left outer join (select SUM(iQuantity) as irkqty,SUM(iTaxPrice) as frkmoney,SUM(iTaxPrice )/SUM(iQuantity) as frkprice,cDefine23 
	from rdrecords01 where ISNULL(cDefine23,'')<>'' group by cDefine23) rk on p.autoid=dh.cDefine23

GO
