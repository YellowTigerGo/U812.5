
--select * From ufsystem..ua_subsys_base where csub_id ='EP'
--delete  From ufsystem..ua_subsys_base where csub_id ='EP'
--预制产品模块信息表eee
print '1、预制产品模块信息表'
IF  NOT EXISTS (select * From ufsystem..ua_subsys_base where csub_id ='EP')
BEGIN
	insert into ufsystem..ua_subsys_base ([cSub_Id],[cSub_Name],[iTasks],[bInstalled],[iVersion],[cObjCreate],[dStart],[nType],[cEntType],[localeid],[iOrder])
	values('EP','项目需求计划',0,'0',8.9,'CreateCom',Null,64,Null,'en-US',30)
	insert into ufsystem..ua_subsys_base ([cSub_Id],[cSub_Name],[iTasks],[bInstalled],[iVersion],[cObjCreate],[dStart],[nType],[cEntType],[localeid],[iOrder])
	values('EP','项目需求计划',0,'0',8.9,'CreateCom',Null,64,Null,'zh-CN',30)
	insert into ufsystem..ua_subsys_base ([cSub_Id],[cSub_Name],[iTasks],[bInstalled],[iVersion],[cObjCreate],[dStart],[nType],[cEntType],[localeid],[iOrder])
	values('EP','项目需求计划',0,'0',8.9,'CreateCom',Null,64,Null,'zh-TW',30)
END
GO
------------------------------------------------------------------------------------------------------------------------

DELETE sa_menuconfig WHERE [menuid]='EP0101'
insert into sa_menuconfig ([menuid],[helpid],[functionid],[parameters],[toolbarname],[authid],[defaultstr],[condition])
values('EP0101',Null,'other','Interface_demo.cls_Show','',Null,Null,Null)
GO
DELETE UFSYSTEM..ua_idt WHERE [id]='EP0101'
insert into UFSYSTEM..ua_idt ([id],[assembly],[catalogtype],[type],[class],[entrypoint],[parameter],[reserved])
values('EP0101','EFMain.clsProductFacade',0,0,Null,Null,Null,Null)
GO
--------------------------------当前工具版本 V12.12.1   脚本创建时间:2018-03-31 9:54:47-----
-- 创建菜单脚本
GO

-- select *  from  UA_Menu where    cMenu_id='EP'
IF  NOT EXISTS (select *  from  UA_Menu where cMenu_id='EP') 
BEGIN 
delete from  UA_Menu where    cMenu_id='EP'
insert into UA_Menu ([cMenu_Id],[cMenu_Name],[cMenu_Eng],[cSub_Id],[IGrade],[cSupMenu_Id],[bEndGrade],[cAuth_Id],[iOrder],[iImgIndex],[Paramters],[Depends],[Flag],[IsWebFlag],[cImageName],[cMenuType])
values('EP','项目需求计划',Null,Null,1,'SCMG','0','','-8000',0,Null,Null,Null,Null,Null,Null)
END 
GO

-- select *  from  UA_Menu where    cMenu_id='EP01'
IF  NOT EXISTS (select *  from  UA_Menu where cMenu_id='EP01') 
BEGIN 
delete from  UA_Menu where    cMenu_id='EP01'
insert into UA_Menu ([cMenu_Id],[cMenu_Name],[cMenu_Eng],[cSub_Id],[IGrade],[cSupMenu_Id],[bEndGrade],[cAuth_Id],[iOrder],[iImgIndex],[Paramters],[Depends],[Flag],[IsWebFlag],[cImageName],[cMenuType])
values('EP01','选项',Null,Null,1,'EP','0','','100',2,Null,Null,Null,Null,Null,Null)
END 
GO

-- select *  from  UA_Menu where    cMenu_id='EP02'
IF  NOT EXISTS (select *  from  UA_Menu where cMenu_id='EP02') 
BEGIN 
delete from  UA_Menu where    cMenu_id='EP02'
insert into UA_Menu ([cMenu_Id],[cMenu_Name],[cMenu_Eng],[cSub_Id],[IGrade],[cSupMenu_Id],[bEndGrade],[cAuth_Id],[iOrder],[iImgIndex],[Paramters],[Depends],[Flag],[IsWebFlag],[cImageName],[cMenuType])
values('EP02','业务处理',Null,Null,1,'EP','0','','200',4,Null,Null,Null,Null,Null,Null)
END 
GO

-- select *  from  UA_Menu where    cMenu_id='EP04'
IF  NOT EXISTS (select *  from  UA_Menu where cMenu_id='EP04') 
BEGIN 
delete from  UA_Menu where    cMenu_id='EP04'
insert into UA_Menu ([cMenu_Id],[cMenu_Name],[cMenu_Eng],[cSub_Id],[IGrade],[cSupMenu_Id],[bEndGrade],[cAuth_Id],[iOrder],[iImgIndex],[Paramters],[Depends],[Flag],[IsWebFlag],[cImageName],[cMenuType])
values('EP04','分析报表',Null,Null,1,'EP','0','','300',4,Null,Null,Null,Null,Null,Null)
END 
GO

-- select *  from  UA_Menu where    cMenu_id='EP040110'
IF  NOT EXISTS (select *  from  UA_Menu where cMenu_id='EP040110') 
BEGIN 
delete from  UA_Menu where    cMenu_id='EP040110'
insert into UA_Menu ([cMenu_Id],[cMenu_Name],[cMenu_Eng],[cSub_Id],[IGrade],[cSupMenu_Id],[bEndGrade],[cAuth_Id],[iOrder],[iImgIndex],[Paramters],[Depends],[Flag],[IsWebFlag],[cImageName],[cMenuType])
values('EP040110','需求采购执行统计表','4','PU',2,'EP04','1','PU[__]3fcdfbcf-9190-410c-95b0-5be765b1b04a_01','77',0,Null,Null,Null,Null,Null,Null)
END 
GO

-- select *  from  UA_Menu where    cMenu_id='EP0101'
IF  NOT EXISTS (select *  from  UA_Menu where cMenu_id='EP0101') 
BEGIN 
delete from  UA_Menu where    cMenu_id='EP0101'
insert into UA_Menu ([cMenu_Id],[cMenu_Name],[cMenu_Eng],[cSub_Id],[IGrade],[cSupMenu_Id],[bEndGrade],[cAuth_Id],[iOrder],[iImgIndex],[Paramters],[Depends],[Flag],[IsWebFlag],[cImageName],[cMenuType])
values('EP0101','选项设置',Null,Null,2,'EP01','1','EP0101','100',4,Null,Null,Null,Null,Null,Null)
END 
GO

-- select *  from  UA_Menu where    cMenu_id='EP0401'
IF  NOT EXISTS (select *  from  UA_Menu where cMenu_id='EP0401') 
BEGIN 
delete from  UA_Menu where    cMenu_id='EP0401'
insert into UA_Menu ([cMenu_Id],[cMenu_Name],[cMenu_Eng],[cSub_Id],[IGrade],[cSupMenu_Id],[bEndGrade],[cAuth_Id],[iOrder],[iImgIndex],[Paramters],[Depends],[Flag],[IsWebFlag],[cImageName],[cMenuType])
values('EP0401','材料成本量差分析',Null,Null,2,'EP04','1','EP040101','100',4,Null,Null,Null,Null,Null,Null)
END 
GO

-- select *  from  UA_Menu where    cMenu_id='EP0201'
IF  NOT EXISTS (select *  from  UA_Menu where cMenu_id='EP0201') 
BEGIN 
delete from  UA_Menu where    cMenu_id='EP0201'
insert into UA_Menu ([cMenu_Id],[cMenu_Name],[cMenu_Eng],[cSub_Id],[IGrade],[cSupMenu_Id],[bEndGrade],[cAuth_Id],[iOrder],[iImgIndex],[Paramters],[Depends],[Flag],[IsWebFlag],[cImageName],[cMenuType])
values('EP0201','项目需求计划单',Null,Null,2,'EP02','1','EP020101','100',4,Null,Null,Null,Null,Null,Null)
END 
GO

-- select *  from  UA_Menu where    cMenu_id='EP0202'
IF  NOT EXISTS (select *  from  UA_Menu where cMenu_id='EP0202') 
BEGIN 
delete from  UA_Menu where    cMenu_id='EP0202'
insert into UA_Menu ([cMenu_Id],[cMenu_Name],[cMenu_Eng],[cSub_Id],[IGrade],[cSupMenu_Id],[bEndGrade],[cAuth_Id],[iOrder],[iImgIndex],[Paramters],[Depends],[Flag],[IsWebFlag],[cImageName],[cMenuType])
values('EP0202','项目需求计划单列表',Null,Null,2,'EP02','1','EP020201','101',4,Null,Null,Null,Null,Null,Null)
END 
GO

-- select *  from  UA_Menu where    cMenu_id='EP0203'
IF  NOT EXISTS (select *  from  UA_Menu where cMenu_id='EP0203') 
BEGIN 
delete from  UA_Menu where    cMenu_id='EP0203'
insert into UA_Menu ([cMenu_Id],[cMenu_Name],[cMenu_Eng],[cSub_Id],[IGrade],[cSupMenu_Id],[bEndGrade],[cAuth_Id],[iOrder],[iImgIndex],[Paramters],[Depends],[Flag],[IsWebFlag],[cImageName],[cMenuType])
values('EP0203','项目需求计划变更单',Null,Null,2,'EP02','1','EP020301','110',4,Null,Null,Null,Null,Null,Null)
END 
GO

-- select *  from  UA_Menu where    cMenu_id='EP0402'
IF  NOT EXISTS (select *  from  UA_Menu where cMenu_id='EP0402') 
BEGIN 
delete from  UA_Menu where    cMenu_id='EP0402'
insert into UA_Menu ([cMenu_Id],[cMenu_Name],[cMenu_Eng],[cSub_Id],[IGrade],[cSupMenu_Id],[bEndGrade],[cAuth_Id],[iOrder],[iImgIndex],[Paramters],[Depends],[Flag],[IsWebFlag],[cImageName],[cMenuType])
values('EP0402','材料成本价差分析',Null,Null,2,'EP04','1','EP040201','110',4,Null,Null,Null,Null,Null,Null)
END 
GO

-- select *  from  UA_Menu where    cMenu_id='EP0204'
IF  NOT EXISTS (select *  from  UA_Menu where cMenu_id='EP0204') 
BEGIN 
delete from  UA_Menu where    cMenu_id='EP0204'
insert into UA_Menu ([cMenu_Id],[cMenu_Name],[cMenu_Eng],[cSub_Id],[IGrade],[cSupMenu_Id],[bEndGrade],[cAuth_Id],[iOrder],[iImgIndex],[Paramters],[Depends],[Flag],[IsWebFlag],[cImageName],[cMenuType])
values('EP0204','项目需求计划变更单列表',Null,Null,2,'EP02','1','EP020401','111',4,Null,Null,Null,Null,Null,Null)
END 
GO

-- select *  from  UA_Menu where    cMenu_id='EP0205'
IF  NOT EXISTS (select *  from  UA_Menu where cMenu_id='EP0205') 
BEGIN 
delete from  UA_Menu where    cMenu_id='EP0205'
insert into UA_Menu ([cMenu_Id],[cMenu_Name],[cMenu_Eng],[cSub_Id],[IGrade],[cSupMenu_Id],[bEndGrade],[cAuth_Id],[iOrder],[iImgIndex],[Paramters],[Depends],[Flag],[IsWebFlag],[cImageName],[cMenuType])
values('EP0205','采购计划单',Null,Null,2,'EP02','1','EP020501','120',4,Null,Null,Null,Null,Null,Null)
END 
GO

-- select *  from  UA_Menu where    cMenu_id='EP0206'
IF  NOT EXISTS (select *  from  UA_Menu where cMenu_id='EP0206') 
BEGIN 
delete from  UA_Menu where    cMenu_id='EP0206'
insert into UA_Menu ([cMenu_Id],[cMenu_Name],[cMenu_Eng],[cSub_Id],[IGrade],[cSupMenu_Id],[bEndGrade],[cAuth_Id],[iOrder],[iImgIndex],[Paramters],[Depends],[Flag],[IsWebFlag],[cImageName],[cMenuType])
values('EP0206','采购计划单列表',Null,Null,2,'EP02','1','EP020601','121',4,Null,Null,Null,Null,Null,Null)
END 
GO

-- select *  from  UA_Menu where    cMenu_id='EP0207'
IF  NOT EXISTS (select *  from  UA_Menu where cMenu_id='EP0207') 
BEGIN 
delete from  UA_Menu where    cMenu_id='EP0207'
insert into UA_Menu ([cMenu_Id],[cMenu_Name],[cMenu_Eng],[cSub_Id],[IGrade],[cSupMenu_Id],[bEndGrade],[cAuth_Id],[iOrder],[iImgIndex],[Paramters],[Depends],[Flag],[IsWebFlag],[cImageName],[cMenuType])
values('EP0207','项目投标报价单',Null,Null,2,'EP02','1','EP020701','130',4,Null,Null,Null,Null,Null,Null)
END 
GO

-- select *  from  UA_Menu where    cMenu_id='EP0208'
IF  NOT EXISTS (select *  from  UA_Menu where cMenu_id='EP0208') 
BEGIN 
delete from  UA_Menu where    cMenu_id='EP0208'
insert into UA_Menu ([cMenu_Id],[cMenu_Name],[cMenu_Eng],[cSub_Id],[IGrade],[cSupMenu_Id],[bEndGrade],[cAuth_Id],[iOrder],[iImgIndex],[Paramters],[Depends],[Flag],[IsWebFlag],[cImageName],[cMenuType])
values('EP0208','项目投标报价单列表',Null,Null,2,'EP02','1','EP020801','131',4,Null,Null,Null,Null,Null,Null)
END 
GO
