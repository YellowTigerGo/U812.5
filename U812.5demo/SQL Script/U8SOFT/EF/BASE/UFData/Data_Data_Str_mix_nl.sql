

--drop table [dbo].[EF_ProjectMRP]
--drop table [dbo].[EF_ProjectMRPs]


--select * from syscolumns where id=object_id('EF_ProjectMRP')

--1 项目需求计划主表
--select * from  [EF_ProjectMRP]
--2 项目需求计划子表
--select * from  [EF_ProjectMRPs]



/******************Contract structure *********************/
print '1 dbo.EF_ProjectMRP  项目需求计划主表 ...'
if not exists (select * from sysobjects where id = object_id('dbo.EF_ProjectMRP') and sysstat & 0xf = 3)
BEGIN
CREATE TABLE [EF_ProjectMRP] (
---------------------------------------------------------------------------------------------------------------------------
--单据主表标准字段
	[id] [bigint] NOT NULL ,							--主表ID
	[ccode] [nvarchar] (20)  NULL ,					--单据编码
	[ddate] [datetime] NULL ,  						--单据日期
	[cmaker] [nvarchar] (20)  NULL ,  				--制单人
	[cmakerddate] [datetime] NULL ,  				--制单日期
	[cmodifer] [nvarchar] (30)  NULL ,				--变更人
	[cmodiferDate] [datetime] NULL ,				--变更日期
	cmodifier [nvarchar] (30)  NULL ,				--修改人
	dmoddate [datetime] NULL ,						--修改日期
	dmodifysystime [datetime] NULL ,				--修改时间
	[cverifier] [nvarchar] (20)  NULL ,  			--审核人
	[dverifydate] [datetime] NULL  			,  		--审核日期
	[ccloser]  [nvarchar] (20)  NULL,				--关闭人
	[dcloserdate]  [datetime] NULL,                 --关闭日期
	[vt_id] [int]  NULL ,							--显示模板号
	[ufts] [timestamp] NULL ,						--时间戳
	[cvouchtype]	[nvarchar](50) NULL,			--单据类型
	[t_cdepcode]	[nvarchar](50) NULL,			--部门编码
	[t_cpersoncode] [nvarchar](50) NULL,			--人员编码
	[t_ccuscode]	[nvarchar](50) NULL,			--客户编码
	[t_cvencode]	[nvarchar](50) NULL,			--供应商编码
	[t_cwhcode]		[nvarchar](50) NULL,			--仓库编码
	[t_cinvcode]	[nvarchar](50) NULL,			--存货编码
	[t_cfree1] [nvarchar](20) NULL,					--存货自由项1
	[t_cfree2] [nvarchar](20) NULL,					--存货自由项2
	[t_cfree3] [nvarchar](20) NULL,					--存货自由项3
	[t_cfree4] [nvarchar](20) NULL,					--存货自由项4
	[t_cfree5] [nvarchar](20) NULL,					--存货自由项5
	[t_cfree6] [nvarchar](20) NULL,					--存货自由项6
	[t_cfree7] [nvarchar](20) NULL,					--存货自由项7
	[t_cfree8] [nvarchar](20) NULL,					--存货自由项8
	[t_cfree9] [nvarchar](20) NULL,					--存货自由项9
	[t_cfree10] [nvarchar](20) NULL,				--存货自由项10
--------------------------------------------------------------------------------------------------------------------------
--审批流专用
	[ireturncount] [int] NULL ,					--打回次数(审批流专用)
	[iswfcontrolled] [int] NULL ,				--审批流启用标志 0 未启用 1启用 2提交 
	[iverifystate] [int] NULL ,					--审批流状态
	[VoucherId] [int] NULL ,					--主表关键字=ID
	[VoucherCode] [nvarchar] (30)  NULL ,		--单据编号=ccode
	[VoucherType] [nvarchar] (30)  NULL ,		--单据类型号=CardNumber
---------------------------------------------------------------------------------------------------------------------------
	[define1] [nvarchar] (20)  NULL ,			--表头自定义项1
	[define2] [nvarchar] (20)  NULL ,			--表头自定义项2
	[define3] [nvarchar] (20)  NULL ,			--表头自定义项3
	[define4] [datetime] NULL ,					--表头自定义项4
	[define5] [int] NULL ,						--表头自定义项5
	[define6] [datetime] NULL ,					--表头自定义项6
	[define7] [float] NULL ,					--表头自定义项7
	[define8] [nvarchar] (20)  NULL ,			--表头自定义项8
	[define9] [nvarchar] (20)  NULL ,			--表头自定义项9
	[define10] [nvarchar] (60)  NULL ,			--表头自定义项10
	[define11] [nvarchar] (120)  NULL ,			--表头自定义项11
	[define12] [nvarchar] (120)  NULL ,			--表头自定义项12
	[define13] [nvarchar] (120)  NULL ,			--表头自定义项13
	[define14] [nvarchar] (120)  NULL ,			--表头自定义项14
	[define15] [int] NULL ,						--表头自定义项15
	[define16] [float] NULL ,					--表头自定义项16
-----------------------------------------------------------------------------------------------------------
--以上部分为必须字段,以下部分根据业务需要设置 --ahzzd	20100505
	cItem_class [nvarchar] (60) NULL,
	citem_cname [nvarchar] (120) NULL,
	cItemCode [nvarchar] (60) NULL,
	cItemName [nvarchar] (120) NULL,
	iPQty decimal(17,6) NULL,--产量
	cMemo [nvarchar] (200)  NULL ,
-----------------------------------------------------------------------------------------------------------
	CONSTRAINT [PK_EF_ProjectMRP] PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
) ON [PRIMARY]

END

GO

/*=======================EF_ProjectMRP add field cPPcode ============================*/
print 'dbo.EF_ProjectMRP add field cPPcode...'
if exists (select * from sysobjects where id = object_id('dbo.EF_ProjectMRP') and sysstat & 0xf = 3)
begin
  if not exists (select c.name,c.id from syscolumns c,sysobjects o 
                 where c.id=o.id and o.xtype='U' and o.name='EF_ProjectMRP' and c.name ='cPPcode') 
     alter table EF_ProjectMRP add cPPcode nvarchar(20) null 
end

GO

/*=======================EF_ProjectMRP add field cmodifier ============================*/
print 'dbo.EF_ProjectMRP add field cmodifier...'
if exists (select * from sysobjects where id = object_id('dbo.EF_ProjectMRP') and sysstat & 0xf = 3)
begin
  if not exists (select c.name,c.id from syscolumns c,sysobjects o 
                 where c.id=o.id and o.xtype='U' and o.name='EF_ProjectMRP' and c.name ='cmodifier') 
     alter table EF_ProjectMRP add cmodifier [nvarchar] (30) null 
end

GO
/*=======================EF_ProjectMRP add field dmoddate ============================*/
print 'dbo.EF_ProjectMRP add field dmoddate...'
if exists (select * from sysobjects where id = object_id('dbo.EF_ProjectMRP') and sysstat & 0xf = 3)
begin
  if not exists (select c.name,c.id from syscolumns c,sysobjects o 
                 where c.id=o.id and o.xtype='U' and o.name='EF_ProjectMRP' and c.name ='dmoddate') 
     alter table EF_ProjectMRP add dmoddate [datetime] null 
end

GO
/*=======================EF_ProjectMRP add field dmodifysystime ============================*/
print 'dbo.EF_ProjectMRP add field dmodifysystime...'
if exists (select * from sysobjects where id = object_id('dbo.EF_ProjectMRP') and sysstat & 0xf = 3)
begin
  if not exists (select c.name,c.id from syscolumns c,sysobjects o 
                 where c.id=o.id and o.xtype='U' and o.name='EF_ProjectMRP' and c.name ='dmodifysystime') 
     alter table EF_ProjectMRP add dmodifysystime [datetime] null 
end

GO
/*=======================EF_ProjectMRP add field dnverifytime ============================*/
print 'dbo.EF_ProjectMRP add field dnverifytime...'
if exists (select * from sysobjects where id = object_id('dbo.EF_ProjectMRP') and sysstat & 0xf = 3)
begin
  if not exists (select c.name,c.id from syscolumns c,sysobjects o 
                 where c.id=o.id and o.xtype='U' and o.name='EF_ProjectMRP' and c.name ='dnverifytime') 
     alter table EF_ProjectMRP add dnverifytime [datetime] null 
end

GO

/******************contracts structure *********************/
print '2 dbo.EF_ProjectMRPs   项目需求计划子表...'
if not exists (select * from sysobjects where id = object_id('dbo.EF_ProjectMRPs') and sysstat & 0xf = 3)
BEGIN
CREATE TABLE [EF_ProjectMRPs] (
---------------------------------------------------------------------------------------------------------------------------
--单据子表标准字段
	[autoid] [bigint] NOT NULL ,						--子表关键字
	[id] [bigint] NOT NULL ,							--主表关键字
	[b_cdepcode] [nvarchar](50) NULL,				--部门编码
	[b_cpersoncode] [nvarchar](50) NULL,			--人员编码
	[b_ccuscode] [nvarchar](50) NULL,				--客户编码
	[b_cvencode] [nvarchar](50) NULL,				--供应商编码
	[b_cwhcode] [nvarchar](50) NULL,				--仓库编码
	[b_cinvcode] [nvarchar](50) NULL,				--存货编码
	[b_cfree1] [nvarchar](20) NULL,					--存货自由项1
	[b_cfree2] [nvarchar](20) NULL,					--存货自由项2
	[b_cfree3] [nvarchar](20) NULL,					--存货自由项3
	[b_cfree4] [nvarchar](20) NULL,					--存货自由项4
	[b_cfree5] [nvarchar](20) NULL,					--存货自由项5
	[b_cfree6] [nvarchar](20) NULL,					--存货自由项6
	[b_cfree7] [nvarchar](20) NULL,					--存货自由项7
	[b_cfree8] [nvarchar](20) NULL,					--存货自由项8
	[b_cfree9] [nvarchar](20) NULL,					--存货自由项9
	[b_cfree10] [nvarchar](20) NULL,				--存货自由项10
	[define22] [nvarchar] (20)  NULL ,				--表体自定义项1
	[define23] [nvarchar] (20)  NULL ,				--表体自定义项2
	[define24] [nvarchar] (20)  NULL ,				--表体自定义项3
	[define25] [nvarchar] (20)  NULL ,				--表体自定义项4
	[define26] [float] NULL ,						--表体自定义项5
	[define27] [float] NULL ,						--表体自定义项6
	[define28] [nvarchar] (20)  NULL ,				--表体自定义项7
	[define29] [nvarchar] (20)  NULL ,				--表体自定义项8
	[define30] [nvarchar] (20)  NULL ,				--表体自定义项9
	[define31] [nvarchar] (20)  NULL ,				--表体自定义项10
	[define32] [nvarchar] (20)  NULL ,				--表体自定义项11
	[define33] [nvarchar] (20)  NULL ,				--表体自定义项12
	[define34] [int] NULL ,							--表体自定义项13
	[define35] [int] NULL ,							--表体自定义项14
	[define36] [datetime] NULL ,					--表体自定义项15
	[define37] [datetime] NULL ,					--表体自定义项16
-----------------------------------------------------------------------------------------------------------
--以上部分为必须字段,以下部分根据业务需要设置
	iinvexchrate decimal(17,6) NULL,--换算率
	AuxUnitCode [nvarchar] (20)  NULL ,--辅计量单位编码 
	cPart [nvarchar] (20)  NULL ,--产品部件
	iUnitQty decimal(17,6) NULL,--单台用量
	cPerform [nvarchar] (50)  NULL ,--材质/性能
	iQty decimal(17,6) NULL,--总数（重）量
	cbMemo [nvarchar] (200)  NULL ,--表体备注
	cOutsourced [nvarchar] (20)  NULL ,--外协/外购
	cbCloser [nvarchar] (20)  NULL ,--关闭人
	cbCloseDate [datetime] NULL ,--关闭日期
	iLLQty decimal(17,6) NULL,--累计领料量
	CONSTRAINT [PK_EF_ProjectMRPs] PRIMARY KEY  CLUSTERED 
	(
		[autoid]
	)  ON [PRIMARY] 
) ON [PRIMARY]
END

GO

/*=======================EF_ProjectMRPs add field cinvname_CAD ============================*/
print 'dbo.EF_ProjectMRPs add field cinvname_CAD...'
if exists (select * from sysobjects where id = object_id('dbo.EF_ProjectMRPs') and sysstat & 0xf = 3)
begin
  if not exists (select c.name,c.id from syscolumns c,sysobjects o 
                 where c.id=o.id and o.xtype='U' and o.name='EF_ProjectMRPs' and c.name ='cinvname_CAD') 
     alter table EF_ProjectMRPs add cinvname_CAD nvarchar(100) null 
end

GO

/*=======================EF_ProjectMRPs add field iRowNo ============================*/
print 'dbo.EF_ProjectMRPs add field iRowNo...'
if exists (select * from sysobjects where id = object_id('dbo.EF_ProjectMRPs') and sysstat & 0xf = 3)
begin
  if not exists (select c.name,c.id from syscolumns c,sysobjects o 
                 where c.id=o.id and o.xtype='U' and o.name='EF_ProjectMRPs' and c.name ='iRowNo') 
     alter table EF_ProjectMRPs add iRowNo int null 
end

GO

--drop table [dbo].[EF_ProjectMRPChanged]
--drop table [dbo].[EF_ProjectMRPChangeds]


--select * from syscolumns where id=object_id('EF_ProjectMRPChanged')

--1 项目需求计划变更主表
--select * from  [EF_ProjectMRPChanged]
--2 项目需求计划变更子表
--select * from  [EF_ProjectMRPChangeds]



/******************Contract structure *********************/
print '1 dbo.EF_ProjectMRPChanged  项目需求计划变更主表 ...'
if not exists (select * from sysobjects where id = object_id('dbo.EF_ProjectMRPChanged') and sysstat & 0xf = 3)
BEGIN
CREATE TABLE [EF_ProjectMRPChanged] (
---------------------------------------------------------------------------------------------------------------------------
--单据主表标准字段
	[id] bigint NOT NULL ,							--主表ID
	[ccode] [nvarchar] (20)  NULL ,					--单据编码
	[ddate] [datetime] NULL ,  						--单据日期
	[cmaker] [nvarchar] (20)  NULL ,  				--制单人
	[cmakerddate] [datetime] NULL ,  				--制单日期
	[cmodifer] [nvarchar] (30)  NULL ,				--变更人
	[cmodiferDate] [datetime] NULL ,				--变更日期
	cmodifier [nvarchar] (30)  NULL ,				--修改人
	dmoddate [datetime] NULL ,						--修改日期
	dmodifysystime [datetime] NULL ,				--修改时间
	[cverifier] [nvarchar] (20)  NULL ,  			--审核人
	[dverifydate] [datetime] NULL  			,  		--审核日期
	[ccloser]  [nvarchar] (20)  NULL,				--关闭人
	[dcloserdate]  [datetime] NULL,                 --关闭日期
	[vt_id] [int]  NULL ,							--显示模板号
	[ufts] [timestamp] NULL ,						--时间戳
	[cvouchtype]	[nvarchar](50) NULL,			--单据类型
	[t_cdepcode]	[nvarchar](50) NULL,			--部门编码
	[t_cpersoncode] [nvarchar](50) NULL,			--人员编码
	[t_ccuscode]	[nvarchar](50) NULL,			--客户编码
	[t_cvencode]	[nvarchar](50) NULL,			--供应商编码
	[t_cwhcode]		[nvarchar](50) NULL,			--仓库编码
	[t_cinvcode]	[nvarchar](50) NULL,			--存货编码
	[t_cfree1] [nvarchar](20) NULL,					--存货自由项1
	[t_cfree2] [nvarchar](20) NULL,					--存货自由项2
	[t_cfree3] [nvarchar](20) NULL,					--存货自由项3
	[t_cfree4] [nvarchar](20) NULL,					--存货自由项4
	[t_cfree5] [nvarchar](20) NULL,					--存货自由项5
	[t_cfree6] [nvarchar](20) NULL,					--存货自由项6
	[t_cfree7] [nvarchar](20) NULL,					--存货自由项7
	[t_cfree8] [nvarchar](20) NULL,					--存货自由项8
	[t_cfree9] [nvarchar](20) NULL,					--存货自由项9
	[t_cfree10] [nvarchar](20) NULL,				--存货自由项10
--------------------------------------------------------------------------------------------------------------------------
--审批流专用
	[ireturncount] [int] NULL ,					--打回次数(审批流专用)
	[iswfcontrolled] [int] NULL ,				--审批流启用标志 0 未启用 1启用 2提交 
	[iverifystate] [int] NULL ,					--审批流状态
	[VoucherId] [int] NULL ,					--主表关键字=ID
	[VoucherCode] [nvarchar] (30)  NULL ,		--单据编号=ccode
	[VoucherType] [nvarchar] (30)  NULL ,		--单据类型号=CardNumber
---------------------------------------------------------------------------------------------------------------------------
	[define1] [nvarchar] (20)  NULL ,			--表头自定义项1
	[define2] [nvarchar] (20)  NULL ,			--表头自定义项2
	[define3] [nvarchar] (20)  NULL ,			--表头自定义项3
	[define4] [datetime] NULL ,					--表头自定义项4
	[define5] [int] NULL ,						--表头自定义项5
	[define6] [datetime] NULL ,					--表头自定义项6
	[define7] [float] NULL ,					--表头自定义项7
	[define8] [nvarchar] (20)  NULL ,			--表头自定义项8
	[define9] [nvarchar] (20)  NULL ,			--表头自定义项9
	[define10] [nvarchar] (60)  NULL ,			--表头自定义项10
	[define11] [nvarchar] (120)  NULL ,			--表头自定义项11
	[define12] [nvarchar] (120)  NULL ,			--表头自定义项12
	[define13] [nvarchar] (120)  NULL ,			--表头自定义项13
	[define14] [nvarchar] (120)  NULL ,			--表头自定义项14
	[define15] [int] NULL ,						--表头自定义项15
	[define16] [float] NULL ,					--表头自定义项16
-----------------------------------------------------------------------------------------------------------
--以上部分为必须字段,以下部分根据业务需要设置 --ahzzd	20100505
	cItem_class [nvarchar] (60) NULL,
	citem_cname [nvarchar] (120) NULL,
	cItemCode [nvarchar] (60) NULL,
	cItemName [nvarchar] (120) NULL,
	iPQty decimal(17,6) NULL,--产量
	cMemo [nvarchar] (200)  NULL ,
	cPMRPCode [nvarchar] (60) NULL,--项目需求计划单号
-----------------------------------------------------------------------------------------------------------
	CONSTRAINT [PK_EF_ProjectMRPChanged] PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
) ON [PRIMARY]

END

GO



/*=======================EF_ProjectMRPChanged add field cmodifier ============================*/
print 'dbo.EF_ProjectMRPChanged add field cmodifier...'
if exists (select * from sysobjects where id = object_id('dbo.EF_ProjectMRPChanged') and sysstat & 0xf = 3)
begin
  if not exists (select c.name,c.id from syscolumns c,sysobjects o 
                 where c.id=o.id and o.xtype='U' and o.name='EF_ProjectMRPChanged' and c.name ='cmodifier') 
     alter table EF_ProjectMRPChanged add cmodifier [nvarchar] (30) null 
end

GO
/*=======================EF_ProjectMRPChanged add field dmoddate ============================*/
print 'dbo.EF_ProjectMRPChanged add field dmoddate...'
if exists (select * from sysobjects where id = object_id('dbo.EF_ProjectMRPChanged') and sysstat & 0xf = 3)
begin
  if not exists (select c.name,c.id from syscolumns c,sysobjects o 
                 where c.id=o.id and o.xtype='U' and o.name='EF_ProjectMRPChanged' and c.name ='dmoddate') 
     alter table EF_ProjectMRPChanged add dmoddate [datetime] null 
end

GO
/*=======================EF_ProjectMRPChanged add field dmodifysystime ============================*/
print 'dbo.EF_ProjectMRPChanged add field dmodifysystime...'
if exists (select * from sysobjects where id = object_id('dbo.EF_ProjectMRPChanged') and sysstat & 0xf = 3)
begin
  if not exists (select c.name,c.id from syscolumns c,sysobjects o 
                 where c.id=o.id and o.xtype='U' and o.name='EF_ProjectMRPChanged' and c.name ='dmodifysystime') 
     alter table EF_ProjectMRPChanged add dmodifysystime [datetime] null 
end

GO
/*=======================EF_ProjectMRPChanged add field dnverifytime ============================*/
print 'dbo.EF_ProjectMRPChanged add field dnverifytime...'
if exists (select * from sysobjects where id = object_id('dbo.EF_ProjectMRPChanged') and sysstat & 0xf = 3)
begin
  if not exists (select c.name,c.id from syscolumns c,sysobjects o 
                 where c.id=o.id and o.xtype='U' and o.name='EF_ProjectMRPChanged' and c.name ='dnverifytime') 
     alter table EF_ProjectMRPChanged add dnverifytime [datetime] null 
end

GO

/******************contracts structure *********************/
print '2 dbo.EF_ProjectMRPChangeds   项目需求计划变更子表...'
--drop table EF_ProjectMRPChangeds
if not exists (select * from sysobjects where id = object_id('dbo.EF_ProjectMRPChangeds') and sysstat & 0xf = 3)
BEGIN
CREATE TABLE [EF_ProjectMRPChangeds] (
---------------------------------------------------------------------------------------------------------------------------
--单据子表标准字段
	[autoid] bigint NOT NULL ,						--子表关键字
	[id] bigint NOT NULL ,							--主表关键字
	[b_cdepcode] [nvarchar](50) NULL,				--部门编码
	[b_cpersoncode] [nvarchar](50) NULL,			--人员编码
	[b_ccuscode] [nvarchar](50) NULL,				--客户编码
	[b_cvencode] [nvarchar](50) NULL,				--供应商编码
	[b_cwhcode] [nvarchar](50) NULL,				--仓库编码
	[b_cinvcode] [nvarchar](50) NULL,				--存货编码
	[b_cfree1] [nvarchar](20) NULL,					--存货自由项1
	[b_cfree2] [nvarchar](20) NULL,					--存货自由项2
	[b_cfree3] [nvarchar](20) NULL,					--存货自由项3
	[b_cfree4] [nvarchar](20) NULL,					--存货自由项4
	[b_cfree5] [nvarchar](20) NULL,					--存货自由项5
	[b_cfree6] [nvarchar](20) NULL,					--存货自由项6
	[b_cfree7] [nvarchar](20) NULL,					--存货自由项7
	[b_cfree8] [nvarchar](20) NULL,					--存货自由项8
	[b_cfree9] [nvarchar](20) NULL,					--存货自由项9
	[b_cfree10] [nvarchar](20) NULL,				--存货自由项10
	[define22] [nvarchar] (20)  NULL ,				--表体自定义项1
	[define23] [nvarchar] (20)  NULL ,				--表体自定义项2
	[define24] [nvarchar] (20)  NULL ,				--表体自定义项3
	[define25] [nvarchar] (20)  NULL ,				--表体自定义项4
	[define26] [float] NULL ,						--表体自定义项5
	[define27] [float] NULL ,						--表体自定义项6
	[define28] [nvarchar] (20)  NULL ,				--表体自定义项7
	[define29] [nvarchar] (20)  NULL ,				--表体自定义项8
	[define30] [nvarchar] (20)  NULL ,				--表体自定义项9
	[define31] [nvarchar] (20)  NULL ,				--表体自定义项10
	[define32] [nvarchar] (20)  NULL ,				--表体自定义项11
	[define33] [nvarchar] (20)  NULL ,				--表体自定义项12
	[define34] [int] NULL ,							--表体自定义项13
	[define35] [int] NULL ,							--表体自定义项14
	[define36] [datetime] NULL ,					--表体自定义项15
	[define37] [datetime] NULL ,					--表体自定义项16
-----------------------------------------------------------------------------------------------------------
--以上部分为必须字段,以下部分根据业务需要设置
	iinvexchrate decimal(17,6) NULL,--换算率
	AuxUnitCode [nvarchar] (20)  NULL ,--辅计量单位编码 
	cPart [nvarchar] (20)  NULL ,--产品部件
	iUnitQty decimal(17,6) NULL,--单台用量
	iUnitQtyOld  decimal(17,6) NULL,--原单台用量
	cPerform [nvarchar] (50)  NULL ,--材质/性能
	iQty decimal(17,6) NULL,--总数（重）量
	iQtyOld  decimal(17,6) NULL,--原总数（重）量
	cbMemo [nvarchar] (200)  NULL ,--表体备注
	cOutsourced [nvarchar] (20)  NULL ,--外协/外购
	cbCloser [nvarchar] (20)  NULL ,--关闭人
	cbCloseDate [datetime] NULL ,--关闭日期
	cPMRPid bigint NULL ,
	cPMRPautoid bigint NULL ,

	CONSTRAINT [PK_EF_ProjectMRPChangeds] PRIMARY KEY  CLUSTERED 
	(
		[autoid]
	)  ON [PRIMARY] 
) ON [PRIMARY]
END

GO

/*=======================EF_ProjectMRPChangeds add field iRowNo ============================*/
print 'dbo.EF_ProjectMRPChangeds add field iRowNo...'
if exists (select * from sysobjects where id = object_id('dbo.EF_ProjectMRPChangeds') and sysstat & 0xf = 3)
begin
  if not exists (select c.name,c.id from syscolumns c,sysobjects o 
                 where c.id=o.id and o.xtype='U' and o.name='EF_ProjectMRPChangeds' and c.name ='iRowNo') 
     alter table EF_ProjectMRPChangeds add iRowNo int null 
end

GO

--drop table [dbo].[EF_ProcurementPlan]
--drop table [dbo].[EF_ProcurementPlans]


--select * from syscolumns where id=object_id('EF_ProcurementPlan')

--1 采购计划主表
--select * from  [EF_ProcurementPlan]
--2 采购计划子表
--select * from  [EF_ProcurementPlans]



/******************Contract structure *********************/
print '1 dbo.EF_ProcurementPlan  采购计划主表 ...'
if not exists (select * from sysobjects where id = object_id('dbo.EF_ProcurementPlan') and sysstat & 0xf = 3)
BEGIN
CREATE TABLE [EF_ProcurementPlan] (
---------------------------------------------------------------------------------------------------------------------------
--单据主表标准字段
	[id] [bigint] NOT NULL ,							--主表ID
	[ccode] [nvarchar] (20)  NULL ,					--单据编码
	[ddate] [datetime] NULL ,  						--单据日期
	[cmaker] [nvarchar] (20)  NULL ,  				--制单人
	[cmakerddate] [datetime] NULL ,  				--制单日期
	[cmodifer] [nvarchar] (30)  NULL ,				--变更人
	[cmodiferDate] [datetime] NULL ,				--变更日期
	cmodifier [nvarchar] (30)  NULL ,				--修改人
	dmoddate [datetime] NULL ,						--修改日期
	dmodifysystime [datetime] NULL ,				--修改时间
	[cverifier] [nvarchar] (20)  NULL ,  			--审核人
	[dverifydate] [datetime] NULL  			,  		--审核日期
	[ccloser]  [nvarchar] (20)  NULL,				--关闭人
	[dcloserdate]  [datetime] NULL,                 --关闭日期
	[vt_id] [int]  NULL ,							--显示模板号
	[ufts] [timestamp] NULL ,						--时间戳
	[cvouchtype]	[nvarchar](50) NULL,			--单据类型
	[t_cdepcode]	[nvarchar](50) NULL,			--部门编码
	[t_cpersoncode] [nvarchar](50) NULL,			--人员编码
	[t_ccuscode]	[nvarchar](50) NULL,			--客户编码
	[t_cvencode]	[nvarchar](50) NULL,			--供应商编码
	[t_cwhcode]		[nvarchar](50) NULL,			--仓库编码
	[t_cinvcode]	[nvarchar](50) NULL,			--存货编码
	[t_cfree1] [nvarchar](20) NULL,					--存货自由项1
	[t_cfree2] [nvarchar](20) NULL,					--存货自由项2
	[t_cfree3] [nvarchar](20) NULL,					--存货自由项3
	[t_cfree4] [nvarchar](20) NULL,					--存货自由项4
	[t_cfree5] [nvarchar](20) NULL,					--存货自由项5
	[t_cfree6] [nvarchar](20) NULL,					--存货自由项6
	[t_cfree7] [nvarchar](20) NULL,					--存货自由项7
	[t_cfree8] [nvarchar](20) NULL,					--存货自由项8
	[t_cfree9] [nvarchar](20) NULL,					--存货自由项9
	[t_cfree10] [nvarchar](20) NULL,				--存货自由项10
--------------------------------------------------------------------------------------------------------------------------
--审批流专用
	[ireturncount] [int] NULL ,					--打回次数(审批流专用)
	[iswfcontrolled] [int] NULL ,				--审批流启用标志 0 未启用 1启用 2提交 
	[iverifystate] [int] NULL ,					--审批流状态
	[VoucherId] [int] NULL ,					--主表关键字=ID
	[VoucherCode] [nvarchar] (30)  NULL ,		--单据编号=ccode
	[VoucherType] [nvarchar] (30)  NULL ,		--单据类型号=CardNumber
---------------------------------------------------------------------------------------------------------------------------
	[define1] [nvarchar] (20)  NULL ,			--表头自定义项1
	[define2] [nvarchar] (20)  NULL ,			--表头自定义项2
	[define3] [nvarchar] (20)  NULL ,			--表头自定义项3
	[define4] [datetime] NULL ,					--表头自定义项4
	[define5] [int] NULL ,						--表头自定义项5
	[define6] [datetime] NULL ,					--表头自定义项6
	[define7] [float] NULL ,					--表头自定义项7
	[define8] [nvarchar] (20)  NULL ,			--表头自定义项8
	[define9] [nvarchar] (20)  NULL ,			--表头自定义项9
	[define10] [nvarchar] (60)  NULL ,			--表头自定义项10
	[define11] [nvarchar] (120)  NULL ,			--表头自定义项11
	[define12] [nvarchar] (120)  NULL ,			--表头自定义项12
	[define13] [nvarchar] (120)  NULL ,			--表头自定义项13
	[define14] [nvarchar] (120)  NULL ,			--表头自定义项14
	[define15] [int] NULL ,						--表头自定义项15
	[define16] [float] NULL ,					--表头自定义项16
-----------------------------------------------------------------------------------------------------------
--以上部分为必须字段,以下部分根据业务需要设置 --ahzzd	20100505
	cItem_class [nvarchar] (60) NULL,
	citem_cname [nvarchar] (120) NULL,
	cItemCode [nvarchar] (60) NULL,
	cItemName [nvarchar] (120) NULL,
	iPQty decimal(17,6) NULL,--产量
	cMemo [nvarchar] (200)  NULL ,
	cPMRPCode [nvarchar] (60) NULL,--清单编号
-----------------------------------------------------------------------------------------------------------
	CONSTRAINT [PK_EF_ProcurementPlan] PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
) ON [PRIMARY]

END

GO


/*=======================EF_ProcurementPlan add field cmodifier ============================*/
print 'dbo.EF_ProcurementPlan add field cmodifier...'
if exists (select * from sysobjects where id = object_id('dbo.EF_ProcurementPlan') and sysstat & 0xf = 3)
begin
  if not exists (select c.name,c.id from syscolumns c,sysobjects o 
                 where c.id=o.id and o.xtype='U' and o.name='EF_ProcurementPlan' and c.name ='cmodifier') 
     alter table EF_ProcurementPlan add cmodifier [nvarchar] (30) null 
end

GO
/*=======================EF_ProcurementPlan add field dmoddate ============================*/
print 'dbo.EF_ProcurementPlan add field dmoddate...'
if exists (select * from sysobjects where id = object_id('dbo.EF_ProcurementPlan') and sysstat & 0xf = 3)
begin
  if not exists (select c.name,c.id from syscolumns c,sysobjects o 
                 where c.id=o.id and o.xtype='U' and o.name='EF_ProcurementPlan' and c.name ='dmoddate') 
     alter table EF_ProcurementPlan add dmoddate [datetime] null 
end

GO
/*=======================EF_ProcurementPlan add field dmodifysystime ============================*/
print 'dbo.EF_ProcurementPlan add field dmodifysystime...'
if exists (select * from sysobjects where id = object_id('dbo.EF_ProcurementPlan') and sysstat & 0xf = 3)
begin
  if not exists (select c.name,c.id from syscolumns c,sysobjects o 
                 where c.id=o.id and o.xtype='U' and o.name='EF_ProcurementPlan' and c.name ='dmodifysystime') 
     alter table EF_ProcurementPlan add dmodifysystime [datetime] null 
end

GO
/*=======================EF_ProcurementPlan add field dnverifytime ============================*/
print 'dbo.EF_ProcurementPlan add field dnverifytime...'
if exists (select * from sysobjects where id = object_id('dbo.EF_ProcurementPlan') and sysstat & 0xf = 3)
begin
  if not exists (select c.name,c.id from syscolumns c,sysobjects o 
                 where c.id=o.id and o.xtype='U' and o.name='EF_ProcurementPlan' and c.name ='dnverifytime') 
     alter table EF_ProcurementPlan add dnverifytime [datetime] null 
end

GO

/******************contracts structure *********************/
print '2 dbo.EF_ProcurementPlans   采购计划子表...'
--drop table EF_ProcurementPlans
if not exists (select * from sysobjects where id = object_id('dbo.EF_ProcurementPlans') and sysstat & 0xf = 3)
BEGIN
CREATE TABLE [EF_ProcurementPlans] (
---------------------------------------------------------------------------------------------------------------------------
--单据子表标准字段
	[autoid] [bigint] NOT NULL ,						--子表关键字
	[id] [bigint] NOT NULL ,							--主表关键字
	[b_cdepcode] [nvarchar](50) NULL,				--部门编码
	[b_cpersoncode] [nvarchar](50) NULL,			--人员编码
	[b_ccuscode] [nvarchar](50) NULL,				--客户编码
	[b_cvencode] [nvarchar](50) NULL,				--供应商编码
	[b_cwhcode] [nvarchar](50) NULL,				--仓库编码
	[b_cinvcode] [nvarchar](50) NULL,				--存货编码
	[b_cfree1] [nvarchar](20) NULL,					--存货自由项1
	[b_cfree2] [nvarchar](20) NULL,					--存货自由项2
	[b_cfree3] [nvarchar](20) NULL,					--存货自由项3
	[b_cfree4] [nvarchar](20) NULL,					--存货自由项4
	[b_cfree5] [nvarchar](20) NULL,					--存货自由项5
	[b_cfree6] [nvarchar](20) NULL,					--存货自由项6
	[b_cfree7] [nvarchar](20) NULL,					--存货自由项7
	[b_cfree8] [nvarchar](20) NULL,					--存货自由项8
	[b_cfree9] [nvarchar](20) NULL,					--存货自由项9
	[b_cfree10] [nvarchar](20) NULL,				--存货自由项10
	[define22] [nvarchar] (20)  NULL ,				--表体自定义项1
	[define23] [nvarchar] (20)  NULL ,				--表体自定义项2
	[define24] [nvarchar] (20)  NULL ,				--表体自定义项3
	[define25] [nvarchar] (20)  NULL ,				--表体自定义项4
	[define26] [float] NULL ,						--表体自定义项5
	[define27] [float] NULL ,						--表体自定义项6
	[define28] [nvarchar] (20)  NULL ,				--表体自定义项7
	[define29] [nvarchar] (20)  NULL ,				--表体自定义项8
	[define30] [nvarchar] (20)  NULL ,				--表体自定义项9
	[define31] [nvarchar] (20)  NULL ,				--表体自定义项10
	[define32] [nvarchar] (20)  NULL ,				--表体自定义项11
	[define33] [nvarchar] (20)  NULL ,				--表体自定义项12
	[define34] [int] NULL ,							--表体自定义项13
	[define35] [int] NULL ,							--表体自定义项14
	[define36] [datetime] NULL ,					--表体自定义项15
	[define37] [datetime] NULL ,					--表体自定义项16
-----------------------------------------------------------------------------------------------------------
--以上部分为必须字段,以下部分根据业务需要设置
	iinvexchrate decimal(17,6) NULL,--换算率
	AuxUnitCode [nvarchar] (20)  NULL ,--辅计量单位编码 
	cPerform [nvarchar] (50)  NULL ,--材质/性能
	cbMemo [nvarchar] (200)  NULL ,--表体备注
	cOutsourced [nvarchar] (20)  NULL ,--外协/外购
	iMRPQty decimal(17,6) NULL,--需求计划量
	iMRPQtyL  decimal(17,6) NULL,--累计需求计划量
	iSafeNum  decimal(17,6) NULL,--安全库存
	iStock  decimal(17,6) NULL,--现存量
	iInQty  decimal(17,6) NULL,--预计入库量
	iOutQty  decimal(17,6) NULL,--预计出库量
	iKYL  decimal(17,6) NULL,--可用量
	iMinQty  decimal(17,6) NULL,--采购经济批量
	iJHQty   decimal(17,6) NULL,--采购计划量
	iJYQty   decimal(17,6) NULL,--建议采购量
	iSJQty   decimal(17,6) NULL,--实际申请量
	cbCloser [nvarchar] (20)  NULL ,--关闭人
	cbCloseDate [datetime] NULL ,--关闭日期
	iQgLQty decimal(17,6) NULL,--累计请购量
	cPMRPautoid bigint NULL ,
	CONSTRAINT [PK_EF_ProcurementPlans] PRIMARY KEY  CLUSTERED 
	(
		[autoid]
	)  ON [PRIMARY] 
) ON [PRIMARY]
END

GO


/*=======================EF_ProcurementPlans add field cPMRPautoid ============================*/
print 'dbo.EF_ProcurementPlans add field cPMRPautoid...'
if exists (select * from sysobjects where id = object_id('dbo.EF_ProcurementPlans') and sysstat & 0xf = 3)
begin
  if not exists (select c.name,c.id from syscolumns c,sysobjects o 
                 where c.id=o.id and o.xtype='U' and o.name='EF_ProcurementPlans' and c.name ='cPMRPautoid') 
     alter table EF_ProcurementPlans add cPMRPautoid bigint null 
end

GO
/*=======================EF_ProcurementPlans add field iRowNo ============================*/
print 'dbo.EF_ProcurementPlans add field iRowNo...'
if exists (select * from sysobjects where id = object_id('dbo.EF_ProcurementPlans') and sysstat & 0xf = 3)
begin
  if not exists (select c.name,c.id from syscolumns c,sysobjects o 
                 where c.id=o.id and o.xtype='U' and o.name='EF_ProcurementPlans' and c.name ='iRowNo') 
     alter table EF_ProcurementPlans add iRowNo int null 
end

GO

--drop table [dbo].[EF_Bid]
--drop table [dbo].[EF_Bids]


--select * from syscolumns where id=object_id('EF_Bid')

--1 项目投标报价主表
--select * from  [EF_Bid]
--2 项目投标报价子表
--select * from  [EF_Bids]



/******************Contract structure *********************/
print '1 dbo.EF_Bid  项目投标报价主表 ...'
if not exists (select * from sysobjects where id = object_id('dbo.EF_Bid') and sysstat & 0xf = 3)
BEGIN
CREATE TABLE [EF_Bid] (
---------------------------------------------------------------------------------------------------------------------------
--单据主表标准字段
	[id] [bigint] NOT NULL ,							--主表ID
	[ccode] [nvarchar] (20)  NULL ,					--单据编码
	[ddate] [datetime] NULL ,  						--单据日期
	[cmaker] [nvarchar] (20)  NULL ,  				--制单人
	[cmakerddate] [datetime] NULL ,  				--制单日期
	[cmodifer] [nvarchar] (30)  NULL ,				--变更人
	[cmodiferDate] [datetime] NULL ,				--变更日期
	cmodifier [nvarchar] (30)  NULL ,				--修改人
	dmoddate [datetime] NULL ,						--修改日期
	dmodifysystime [datetime] NULL ,				--修改时间
	[cverifier] [nvarchar] (20)  NULL ,  			--审核人
	[dverifydate] [datetime] NULL  			,  		--审核日期
	[ccloser]  [nvarchar] (20)  NULL,				--关闭人
	[dcloserdate]  [datetime] NULL,                 --关闭日期
	[vt_id] [int]  NULL ,							--显示模板号
	[ufts] [timestamp] NULL ,						--时间戳
	[cvouchtype]	[nvarchar](50) NULL,			--单据类型
	[t_cdepcode]	[nvarchar](50) NULL,			--部门编码
	[t_cpersoncode] [nvarchar](50) NULL,			--人员编码
	[t_ccuscode]	[nvarchar](50) NULL,			--客户编码
	[t_cvencode]	[nvarchar](50) NULL,			--供应商编码
	[t_cwhcode]		[nvarchar](50) NULL,			--仓库编码
	[t_cinvcode]	[nvarchar](50) NULL,			--存货编码
	[t_cfree1] [nvarchar](20) NULL,					--存货自由项1
	[t_cfree2] [nvarchar](20) NULL,					--存货自由项2
	[t_cfree3] [nvarchar](20) NULL,					--存货自由项3
	[t_cfree4] [nvarchar](20) NULL,					--存货自由项4
	[t_cfree5] [nvarchar](20) NULL,					--存货自由项5
	[t_cfree6] [nvarchar](20) NULL,					--存货自由项6
	[t_cfree7] [nvarchar](20) NULL,					--存货自由项7
	[t_cfree8] [nvarchar](20) NULL,					--存货自由项8
	[t_cfree9] [nvarchar](20) NULL,					--存货自由项9
	[t_cfree10] [nvarchar](20) NULL,				--存货自由项10
--------------------------------------------------------------------------------------------------------------------------
--审批流专用
	[ireturncount] [int] NULL ,					--打回次数(审批流专用)
	[iswfcontrolled] [int] NULL ,				--审批流启用标志 0 未启用 1启用 2提交 
	[iverifystate] [int] NULL ,					--审批流状态
	[VoucherId] [int] NULL ,					--主表关键字=ID
	[VoucherCode] [nvarchar] (30)  NULL ,		--单据编号=ccode
	[VoucherType] [nvarchar] (30)  NULL ,		--单据类型号=CardNumber
---------------------------------------------------------------------------------------------------------------------------
	[define1] [nvarchar] (20)  NULL ,			--表头自定义项1
	[define2] [nvarchar] (20)  NULL ,			--表头自定义项2
	[define3] [nvarchar] (20)  NULL ,			--表头自定义项3
	[define4] [datetime] NULL ,					--表头自定义项4
	[define5] [int] NULL ,						--表头自定义项5
	[define6] [datetime] NULL ,					--表头自定义项6
	[define7] [float] NULL ,					--表头自定义项7
	[define8] [nvarchar] (20)  NULL ,			--表头自定义项8
	[define9] [nvarchar] (20)  NULL ,			--表头自定义项9
	[define10] [nvarchar] (60)  NULL ,			--表头自定义项10
	[define11] [nvarchar] (120)  NULL ,			--表头自定义项11
	[define12] [nvarchar] (120)  NULL ,			--表头自定义项12
	[define13] [nvarchar] (120)  NULL ,			--表头自定义项13
	[define14] [nvarchar] (120)  NULL ,			--表头自定义项14
	[define15] [int] NULL ,						--表头自定义项15
	[define16] [float] NULL ,					--表头自定义项16
-----------------------------------------------------------------------------------------------------------
--以上部分为必须字段,以下部分根据业务需要设置 --ahzzd	20100505
	cItem_class [nvarchar] (60) NULL,
	citem_cname [nvarchar] (120) NULL,
	cItemCode [nvarchar] (60) NULL,
	cItemName [nvarchar] (120) NULL,
	iPQty decimal(17,6) NULL,--产量
	cMemo [nvarchar] (200)  NULL ,
-----------------------------------------------------------------------------------------------------------
	CONSTRAINT [PK_EF_Bid] PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
) ON [PRIMARY]

END

GO


/*=======================EF_Bid add field cmodifier ============================*/
print 'dbo.EF_Bid add field cmodifier...'
if exists (select * from sysobjects where id = object_id('dbo.EF_Bid') and sysstat & 0xf = 3)
begin
  if not exists (select c.name,c.id from syscolumns c,sysobjects o 
                 where c.id=o.id and o.xtype='U' and o.name='EF_Bid' and c.name ='cmodifier') 
     alter table EF_Bid add cmodifier [nvarchar] (30) null 
end

GO
/*=======================EF_Bid add field dmoddate ============================*/
print 'dbo.EF_Bid add field dmoddate...'
if exists (select * from sysobjects where id = object_id('dbo.EF_Bid') and sysstat & 0xf = 3)
begin
  if not exists (select c.name,c.id from syscolumns c,sysobjects o 
                 where c.id=o.id and o.xtype='U' and o.name='EF_Bid' and c.name ='dmoddate') 
     alter table EF_Bid add dmoddate [datetime] null 
end

GO
/*=======================EF_Bid add field dmodifysystime ============================*/
print 'dbo.EF_Bid add field dmodifysystime...'
if exists (select * from sysobjects where id = object_id('dbo.EF_Bid') and sysstat & 0xf = 3)
begin
  if not exists (select c.name,c.id from syscolumns c,sysobjects o 
                 where c.id=o.id and o.xtype='U' and o.name='EF_Bid' and c.name ='dmodifysystime') 
     alter table EF_Bid add dmodifysystime [datetime] null 
end

GO
/*=======================EF_Bid add field dnverifytime ============================*/
print 'dbo.EF_Bid add field dnverifytime...'
if exists (select * from sysobjects where id = object_id('dbo.EF_Bid') and sysstat & 0xf = 3)
begin
  if not exists (select c.name,c.id from syscolumns c,sysobjects o 
                 where c.id=o.id and o.xtype='U' and o.name='EF_Bid' and c.name ='dnverifytime') 
     alter table EF_Bid add dnverifytime [datetime] null 
end

GO



/******************contracts structure *********************/
print '2 dbo.EF_Bids   项目投标报价子表...'
if not exists (select * from sysobjects where id = object_id('dbo.EF_Bids') and sysstat & 0xf = 3)
BEGIN
CREATE TABLE [EF_Bids] (
---------------------------------------------------------------------------------------------------------------------------
--单据子表标准字段
	[autoid] [bigint] NOT NULL ,						--子表关键字
	[id] [bigint] NOT NULL ,							--主表关键字
	[b_cdepcode] [nvarchar](50) NULL,				--部门编码
	[b_cpersoncode] [nvarchar](50) NULL,			--人员编码
	[b_ccuscode] [nvarchar](50) NULL,				--客户编码
	[b_cvencode] [nvarchar](50) NULL,				--供应商编码
	[b_cwhcode] [nvarchar](50) NULL,				--仓库编码
	[b_cinvcode] [nvarchar](50) NULL,				--存货编码
	[b_cfree1] [nvarchar](20) NULL,					--存货自由项1
	[b_cfree2] [nvarchar](20) NULL,					--存货自由项2
	[b_cfree3] [nvarchar](20) NULL,					--存货自由项3
	[b_cfree4] [nvarchar](20) NULL,					--存货自由项4
	[b_cfree5] [nvarchar](20) NULL,					--存货自由项5
	[b_cfree6] [nvarchar](20) NULL,					--存货自由项6
	[b_cfree7] [nvarchar](20) NULL,					--存货自由项7
	[b_cfree8] [nvarchar](20) NULL,					--存货自由项8
	[b_cfree9] [nvarchar](20) NULL,					--存货自由项9
	[b_cfree10] [nvarchar](20) NULL,				--存货自由项10
	[define22] [nvarchar] (20)  NULL ,				--表体自定义项1
	[define23] [nvarchar] (20)  NULL ,				--表体自定义项2
	[define24] [nvarchar] (20)  NULL ,				--表体自定义项3
	[define25] [nvarchar] (20)  NULL ,				--表体自定义项4
	[define26] [float] NULL ,						--表体自定义项5
	[define27] [float] NULL ,						--表体自定义项6
	[define28] [nvarchar] (20)  NULL ,				--表体自定义项7
	[define29] [nvarchar] (20)  NULL ,				--表体自定义项8
	[define30] [nvarchar] (20)  NULL ,				--表体自定义项9
	[define31] [nvarchar] (20)  NULL ,				--表体自定义项10
	[define32] [nvarchar] (20)  NULL ,				--表体自定义项11
	[define33] [nvarchar] (20)  NULL ,				--表体自定义项12
	[define34] [int] NULL ,							--表体自定义项13
	[define35] [int] NULL ,							--表体自定义项14
	[define36] [datetime] NULL ,					--表体自定义项15
	[define37] [datetime] NULL ,					--表体自定义项16
-----------------------------------------------------------------------------------------------------------
--以上部分为必须字段,以下部分根据业务需要设置
	iinvexchrate decimal(17,6) NULL,--换算率
	AuxUnitCode [nvarchar] (20)  NULL ,--辅计量单位编码 
	cMaterialClass [nvarchar] (50)  NULL ,--材料类别
	iCB decimal(17,6) NULL,--采购成本
	iBJ decimal(17,6) NULL,--投标报价
	cbMemo [nvarchar] (200)  NULL ,--表体备注
	cbCloser [nvarchar] (20)  NULL ,--关闭人
	cbCloseDate [datetime] NULL ,--关闭日期
	CONSTRAINT [PK_EF_Bids] PRIMARY KEY  CLUSTERED 
	(
		[autoid]
	)  ON [PRIMARY] 
) ON [PRIMARY]
END

GO

/*=======================EF_Bids add field iRowNo ============================*/
print 'dbo.EF_Bids add field iRowNo...'
if exists (select * from sysobjects where id = object_id('dbo.EF_Bids') and sysstat & 0xf = 3)
begin
  if not exists (select c.name,c.id from syscolumns c,sysobjects o 
                 where c.id=o.id and o.xtype='U' and o.name='EF_Bids' and c.name ='iRowNo') 
     alter table EF_Bids add iRowNo int null 
end

GO
