--------------------------------当前工具版本 V12.12.1   脚本创建时间:2018-04-27 18:14:53-----
-- 以下脚本在不影响用户已分配权限的情况下将更新系统权限信息
-- select *  from  ua_auth_base where    cAuth_Id='EP01'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP01') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP01'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP01','EP',1,'EP','0','1000',Null,Null,Null,Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP01'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP01') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP01'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP01','选项')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP01'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP01') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP01'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('en-US','EP01','选项')
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP01'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP01') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP01'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-TW','EP01','选项')
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP02'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP02') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP02'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP02','EP',1,'EP','0','3040',Null,Null,Null,Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP02'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP02') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP02'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP02','业务处理')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP02'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP02') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP02'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP02'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP02') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP02'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP04'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP04') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP04'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP04','EP',1,'EP','0','3050',Null,Null,Null,Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP04'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP04') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP04'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP04','分析报表')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP04'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP04') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP04'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('en-US','EP04','分析报表')
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP04'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP04') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP04'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-TW','EP04','分析报表')
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP0101'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP0101') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP0101'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP0101','EP',2,'EP01','1','1000',Null,Null,'EP01',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP0101'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP0101') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP0101'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP0101','选项设置')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP0101'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP0101') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP0101'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('en-US','EP0101','选项设置')
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP0101'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP0101') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP0101'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-TW','EP0101','选项设置')
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP0201'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP0201') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP0201'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP0201','EP',2,'EP02','0','3050',Null,Null,'EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP0201'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP0201') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP0201'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP0201','项目需求计划单')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP0201'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP0201') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP0201'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP0201'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP0201') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP0201'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP0202'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP0202') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP0202'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP0202','EP',2,'EP02','0','3050',Null,Null,'EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP0202'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP0202') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP0202'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP0202','项目需求计划单列表')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP0202'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP0202') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP0202'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP0202'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP0202') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP0202'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP0203'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP0203') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP0203'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP0203','EP',2,'EP02','0','3050',Null,Null,'EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP0203'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP0203') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP0203'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP0203','项目需求计划变更单')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP0203'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP0203') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP0203'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP0203'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP0203') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP0203'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP0204'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP0204') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP0204'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP0204','EP',2,'EP02','0','3050',Null,Null,'EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP0204'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP0204') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP0204'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP0204','项目需求计划变更单列表')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP0204'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP0204') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP0204'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP0204'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP0204') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP0204'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP0205'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP0205') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP0205'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP0205','EP',2,'EP02','0','3050',Null,Null,'EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP0205'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP0205') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP0205'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP0205','采购计划单')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP0205'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP0205') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP0205'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP0205'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP0205') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP0205'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP0206'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP0206') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP0206'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP0206','EP',2,'EP02','0','3050',Null,Null,'EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP0206'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP0206') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP0206'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP0206','采购计划单列表')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP0206'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP0206') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP0206'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP0206'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP0206') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP0206'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP0207'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP0207') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP0207'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP0207','EP',2,'EP02','0','3050',Null,Null,'EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP0207'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP0207') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP0207'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP0207','项目投标报价单')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP0207'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP0207') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP0207'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP0207'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP0207') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP0207'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP0208'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP0208') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP0208'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP0208','EP',2,'EP02','0','3050',Null,Null,'EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP0208'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP0208') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP0208'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP0208','项目投标报价单列表')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP0208'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP0208') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP0208'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP0208'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP0208') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP0208'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020101'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020101') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020101'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020101','EP',3,'EP0201','1','3060',Null,Null,'EP0201,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020101'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020101') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020101'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020101','项目需求计划单_查询')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020101'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020101') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020101'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020101'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020101') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020101'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020102'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020102') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020102'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020102','EP',3,'EP0201','1','3150',Null,Null,'EP0201,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020102'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020102') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020102'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020102','项目需求计划单_增加')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020102'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020102') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020102'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020102'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020102') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020102'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020103'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020103') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020103'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020103','EP',3,'EP0201','1','3160',Null,Null,'EP0201,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020103'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020103') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020103'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020103','项目需求计划单_编辑')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020103'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020103') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020103'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020103'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020103') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020103'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020104'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020104') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020104'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020104','EP',3,'EP0201','1','3170',Null,Null,'EP0201,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020104'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020104') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020104'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020104','项目需求计划单_保存')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020104'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020104') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020104'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020104'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020104') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020104'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020105'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020105') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020105'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020105','EP',3,'EP0201','1','3180',Null,Null,'EP0201,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020105'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020105') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020105'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020105','项目需求计划单_删除')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020105'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020105') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020105'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020105'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020105') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020105'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020106'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020106') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020106'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020106','EP',3,'EP0201','1','3190',Null,Null,'EP0201,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020106'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020106') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020106'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020106','项目需求计划单_审核')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020106'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020106') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020106'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020106'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020106') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020106'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020107'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020107') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020107'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020107','EP',3,'EP0201','1','3190',Null,Null,'EP0201,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020107'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020107') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020107'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020107','项目需求计划单_弃审')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020107'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020107') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020107'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020107'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020107') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020107'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020108'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020108') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020108'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020108','EP',3,'EP0201','1','3190',Null,Null,'EP0201,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020108'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020108') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020108'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020108','项目需求计划单_打印')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020108'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020108') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020108'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020108'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020108') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020108'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020109'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020109') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020109'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020109','EP',3,'EP0201','1','3190',Null,Null,'EP0201,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020109'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020109') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020109'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020109','项目需求计划单_输出')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020109'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020109') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020109'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020109'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020109') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020109'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020110'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020110') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020110'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020110','EP',3,'EP0201','1','3190',Null,Null,'EP0201,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020110'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020110') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020110'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020110','项目需求计划单_提交')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020110'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020110') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020110'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020110'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020110') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020110'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020111'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020111') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020111'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020111','EP',3,'EP0201','1','3190',Null,Null,'EP0201,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020111'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020111') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020111'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020111','项目需求计划单_撤销')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020111'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020111') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020111'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020111'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020111') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020111'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020201'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020201') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020201'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020201','EP',3,'EP0202','1','3060',Null,Null,'EP0202,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020201'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020201') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020201'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020201','项目需求计划单列表_查询')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020201'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020201') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020201'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020201'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020201') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020201'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020301'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020301') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020301'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020301','EP',3,'EP0203','1','3060',Null,Null,'EP0203,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020301'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020301') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020301'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020301','项目需求计划变更单_查询')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020301'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020301') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020301'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020301'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020301') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020301'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020302'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020302') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020302'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020302','EP',3,'EP0203','1','3150',Null,Null,'EP0203,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020302'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020302') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020302'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020302','项目需求计划变更单_增加')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020302'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020302') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020302'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020302'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020302') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020302'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020303'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020303') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020303'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020303','EP',3,'EP0203','1','3160',Null,Null,'EP0203,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020303'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020303') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020303'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020303','项目需求计划变更单_编辑')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020303'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020303') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020303'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020303'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020303') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020303'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020304'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020304') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020304'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020304','EP',3,'EP0203','1','3170',Null,Null,'EP0203,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020304'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020304') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020304'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020304','项目需求计划变更单_保存')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020304'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020304') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020304'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020304'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020304') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020304'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020305'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020305') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020305'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020305','EP',3,'EP0203','1','3180',Null,Null,'EP0203,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020305'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020305') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020305'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020305','项目需求计划变更单_删除')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020305'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020305') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020305'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020305'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020305') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020305'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020306'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020306') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020306'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020306','EP',3,'EP0203','1','3190',Null,Null,'EP0203,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020306'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020306') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020306'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020306','项目需求计划变更单_审核')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020306'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020306') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020306'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020306'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020306') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020306'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020401'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020401') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020401'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020401','EP',3,'EP0204','1','3060',Null,Null,'EP0204,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020401'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020401') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020401'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020401','项目需求计划变更单列表_查询')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020401'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020401') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020401'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020401'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020401') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020401'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020501'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020501') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020501'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020501','EP',3,'EP0205','1','3060',Null,Null,'EP0205,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020501'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020501') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020501'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020501','采购计划单_查询')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020501'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020501') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020501'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020501'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020501') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020501'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020502'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020502') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020502'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020502','EP',3,'EP0205','1','3150',Null,Null,'EP0205,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020502'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020502') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020502'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020502','采购计划单_增加')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020502'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020502') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020502'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020502'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020502') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020502'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020503'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020503') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020503'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020503','EP',3,'EP0205','1','3160',Null,Null,'EP0205,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020503'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020503') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020503'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020503','采购计划单_编辑')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020503'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020503') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020503'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020503'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020503') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020503'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020504'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020504') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020504'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020504','EP',3,'EP0205','1','3170',Null,Null,'EP0205,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020504'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020504') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020504'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020504','采购计划单_保存')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020504'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020504') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020504'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020504'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020504') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020504'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020505'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020505') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020505'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020505','EP',3,'EP0205','1','3180',Null,Null,'EP0205,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020505'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020505') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020505'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020505','采购计划单_删除')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020505'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020505') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020505'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020505'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020505') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020505'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020506'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020506') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020506'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020506','EP',3,'EP0205','1','3190',Null,Null,'EP0205,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020506'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020506') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020506'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020506','采购计划单_审核')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020506'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020506') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020506'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020506'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020506') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020506'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020601'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020601') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020601'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020601','EP',3,'EP0206','1','3060',Null,Null,'EP0206,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020601'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020601') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020601'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020601','采购计划单列表_查询')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020601'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020601') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020601'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020601'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020601') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020601'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020701'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020701') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020701'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020701','EP',3,'EP0207','1','3060',Null,Null,'EP0207,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020701'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020701') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020701'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020701','项目投标报价单_查询')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020701'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020701') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020701'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020701'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020701') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020701'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020702'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020702') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020702'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020702','EP',3,'EP0207','1','3150',Null,Null,'EP0207,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020702'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020702') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020702'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020702','项目投标报价单_增加')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020702'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020702') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020702'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020702'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020702') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020702'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020703'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020703') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020703'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020703','EP',3,'EP0207','1','3160',Null,Null,'EP0207,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020703'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020703') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020703'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020703','项目投标报价单_编辑')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020703'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020703') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020703'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020703'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020703') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020703'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020704'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020704') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020704'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020704','EP',3,'EP0207','1','3170',Null,Null,'EP0207,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020704'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020704') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020704'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020704','项目投标报价单_保存')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020704'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020704') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020704'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020704'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020704') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020704'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020705'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020705') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020705'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020705','EP',3,'EP0207','1','3180',Null,Null,'EP0207,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020705'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020705') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020705'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020705','项目投标报价单_删除')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020705'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020705') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020705'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020705'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020705') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020705'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020706'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020706') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020706'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020706','EP',3,'EP0207','1','3190',Null,Null,'EP0207,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020706'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020706') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020706'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020706','项目投标报价单_审核')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020706'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020706') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020706'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020706'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020706') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020706'
END 
GO

-- select *  from  ua_auth_base where    cAuth_Id='EP020801'
IF NOT EXISTS (select *  from  UA_HoldAuth where cAuth_Id='EP020801') 
BEGIN 
delete from  ua_auth_base where    cAuth_Id='EP020801'
insert into ua_auth_base ([cAuth_Id],[cSub_Id],[iGrade],[cSupAuth_Id],[bEndGrade],[iOrder],[cAcc_Id],[cAuthType],[cAllSupAuths],[irepnum],[cRepellent],[cRepellentModule],[cNotRepellent],[cRepInDB],[cRepModInDB],[cNotRepInDB])
values('EP020801','EP',3,'EP0208','1','3060',Null,Null,'EP0208,EP02',Null,Null,Null,Null,Null,Null,Null)
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020801'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020801') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-CN' and  cAuth_Id='EP020801'
insert into ua_auth_lang ([localeid],[cAuth_ID],[cAuth_Name])
values('zh-CN','EP020801','项目投标报价单列表_查询')
END 
GO
-- select *  from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020801'
IF NOT EXISTS (select *  from  UA_HoldAuth where      cAuth_Id='EP020801') 
BEGIN 
delete from  ua_auth_lang where localeid='en-US' and  cAuth_Id='EP020801'
END 
GO
-- select *  from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020801'
IF NOT EXISTS (select *  from  UA_HoldAuth where   cAuth_Id='EP020801') 
BEGIN 
delete from  ua_auth_lang where localeid='zh-TW' and  cAuth_Id='EP020801'
END 
GO






--处理权限BUG 的修正处理
exec ufsystem..sp_CalculateSupAuth
GO
execute ufsystem..sp_CalculateiAuthtype
GO

