

--drop table [dbo].[EF_ProjectMRP]
--drop table [dbo].[EF_ProjectMRPs]


--select * from syscolumns where id=object_id('EF_ProjectMRP')

--1 ��Ŀ����ƻ�����
--select * from  [EF_ProjectMRP]
--2 ��Ŀ����ƻ��ӱ�
--select * from  [EF_ProjectMRPs]



/******************Contract structure *********************/
print '1 dbo.EF_ProjectMRP  ��Ŀ����ƻ����� ...'
if not exists (select * from sysobjects where id = object_id('dbo.EF_ProjectMRP') and sysstat & 0xf = 3)
BEGIN
CREATE TABLE [EF_ProjectMRP] (
---------------------------------------------------------------------------------------------------------------------------
--���������׼�ֶ�
	[id] [bigint] NOT NULL ,							--����ID
	[ccode] [nvarchar] (20)  NULL ,					--���ݱ���
	[ddate] [datetime] NULL ,  						--��������
	[cmaker] [nvarchar] (20)  NULL ,  				--�Ƶ���
	[cmakerddate] [datetime] NULL ,  				--�Ƶ�����
	[cmodifer] [nvarchar] (30)  NULL ,				--�����
	[cmodiferDate] [datetime] NULL ,				--�������
	cmodifier [nvarchar] (30)  NULL ,				--�޸���
	dmoddate [datetime] NULL ,						--�޸�����
	dmodifysystime [datetime] NULL ,				--�޸�ʱ��
	[cverifier] [nvarchar] (20)  NULL ,  			--�����
	[dverifydate] [datetime] NULL  			,  		--�������
	[ccloser]  [nvarchar] (20)  NULL,				--�ر���
	[dcloserdate]  [datetime] NULL,                 --�ر�����
	[vt_id] [int]  NULL ,							--��ʾģ���
	[ufts] [timestamp] NULL ,						--ʱ���
	[cvouchtype]	[nvarchar](50) NULL,			--��������
	[t_cdepcode]	[nvarchar](50) NULL,			--���ű���
	[t_cpersoncode] [nvarchar](50) NULL,			--��Ա����
	[t_ccuscode]	[nvarchar](50) NULL,			--�ͻ�����
	[t_cvencode]	[nvarchar](50) NULL,			--��Ӧ�̱���
	[t_cwhcode]		[nvarchar](50) NULL,			--�ֿ����
	[t_cinvcode]	[nvarchar](50) NULL,			--�������
	[t_cfree1] [nvarchar](20) NULL,					--���������1
	[t_cfree2] [nvarchar](20) NULL,					--���������2
	[t_cfree3] [nvarchar](20) NULL,					--���������3
	[t_cfree4] [nvarchar](20) NULL,					--���������4
	[t_cfree5] [nvarchar](20) NULL,					--���������5
	[t_cfree6] [nvarchar](20) NULL,					--���������6
	[t_cfree7] [nvarchar](20) NULL,					--���������7
	[t_cfree8] [nvarchar](20) NULL,					--���������8
	[t_cfree9] [nvarchar](20) NULL,					--���������9
	[t_cfree10] [nvarchar](20) NULL,				--���������10
--------------------------------------------------------------------------------------------------------------------------
--������ר��
	[ireturncount] [int] NULL ,					--��ش���(������ר��)
	[iswfcontrolled] [int] NULL ,				--���������ñ�־ 0 δ���� 1���� 2�ύ 
	[iverifystate] [int] NULL ,					--������״̬
	[VoucherId] [int] NULL ,					--����ؼ���=ID
	[VoucherCode] [nvarchar] (30)  NULL ,		--���ݱ��=ccode
	[VoucherType] [nvarchar] (30)  NULL ,		--�������ͺ�=CardNumber
---------------------------------------------------------------------------------------------------------------------------
	[define1] [nvarchar] (20)  NULL ,			--��ͷ�Զ�����1
	[define2] [nvarchar] (20)  NULL ,			--��ͷ�Զ�����2
	[define3] [nvarchar] (20)  NULL ,			--��ͷ�Զ�����3
	[define4] [datetime] NULL ,					--��ͷ�Զ�����4
	[define5] [int] NULL ,						--��ͷ�Զ�����5
	[define6] [datetime] NULL ,					--��ͷ�Զ�����6
	[define7] [float] NULL ,					--��ͷ�Զ�����7
	[define8] [nvarchar] (20)  NULL ,			--��ͷ�Զ�����8
	[define9] [nvarchar] (20)  NULL ,			--��ͷ�Զ�����9
	[define10] [nvarchar] (60)  NULL ,			--��ͷ�Զ�����10
	[define11] [nvarchar] (120)  NULL ,			--��ͷ�Զ�����11
	[define12] [nvarchar] (120)  NULL ,			--��ͷ�Զ�����12
	[define13] [nvarchar] (120)  NULL ,			--��ͷ�Զ�����13
	[define14] [nvarchar] (120)  NULL ,			--��ͷ�Զ�����14
	[define15] [int] NULL ,						--��ͷ�Զ�����15
	[define16] [float] NULL ,					--��ͷ�Զ�����16
-----------------------------------------------------------------------------------------------------------
--���ϲ���Ϊ�����ֶ�,���²��ָ���ҵ����Ҫ���� --ahzzd	20100505
	cItem_class [nvarchar] (60) NULL,
	citem_cname [nvarchar] (120) NULL,
	cItemCode [nvarchar] (60) NULL,
	cItemName [nvarchar] (120) NULL,
	iPQty decimal(17,6) NULL,--����
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
print '2 dbo.EF_ProjectMRPs   ��Ŀ����ƻ��ӱ�...'
if not exists (select * from sysobjects where id = object_id('dbo.EF_ProjectMRPs') and sysstat & 0xf = 3)
BEGIN
CREATE TABLE [EF_ProjectMRPs] (
---------------------------------------------------------------------------------------------------------------------------
--�����ӱ��׼�ֶ�
	[autoid] [bigint] NOT NULL ,						--�ӱ�ؼ���
	[id] [bigint] NOT NULL ,							--����ؼ���
	[b_cdepcode] [nvarchar](50) NULL,				--���ű���
	[b_cpersoncode] [nvarchar](50) NULL,			--��Ա����
	[b_ccuscode] [nvarchar](50) NULL,				--�ͻ�����
	[b_cvencode] [nvarchar](50) NULL,				--��Ӧ�̱���
	[b_cwhcode] [nvarchar](50) NULL,				--�ֿ����
	[b_cinvcode] [nvarchar](50) NULL,				--�������
	[b_cfree1] [nvarchar](20) NULL,					--���������1
	[b_cfree2] [nvarchar](20) NULL,					--���������2
	[b_cfree3] [nvarchar](20) NULL,					--���������3
	[b_cfree4] [nvarchar](20) NULL,					--���������4
	[b_cfree5] [nvarchar](20) NULL,					--���������5
	[b_cfree6] [nvarchar](20) NULL,					--���������6
	[b_cfree7] [nvarchar](20) NULL,					--���������7
	[b_cfree8] [nvarchar](20) NULL,					--���������8
	[b_cfree9] [nvarchar](20) NULL,					--���������9
	[b_cfree10] [nvarchar](20) NULL,				--���������10
	[define22] [nvarchar] (20)  NULL ,				--�����Զ�����1
	[define23] [nvarchar] (20)  NULL ,				--�����Զ�����2
	[define24] [nvarchar] (20)  NULL ,				--�����Զ�����3
	[define25] [nvarchar] (20)  NULL ,				--�����Զ�����4
	[define26] [float] NULL ,						--�����Զ�����5
	[define27] [float] NULL ,						--�����Զ�����6
	[define28] [nvarchar] (20)  NULL ,				--�����Զ�����7
	[define29] [nvarchar] (20)  NULL ,				--�����Զ�����8
	[define30] [nvarchar] (20)  NULL ,				--�����Զ�����9
	[define31] [nvarchar] (20)  NULL ,				--�����Զ�����10
	[define32] [nvarchar] (20)  NULL ,				--�����Զ�����11
	[define33] [nvarchar] (20)  NULL ,				--�����Զ�����12
	[define34] [int] NULL ,							--�����Զ�����13
	[define35] [int] NULL ,							--�����Զ�����14
	[define36] [datetime] NULL ,					--�����Զ�����15
	[define37] [datetime] NULL ,					--�����Զ�����16
-----------------------------------------------------------------------------------------------------------
--���ϲ���Ϊ�����ֶ�,���²��ָ���ҵ����Ҫ����
	iinvexchrate decimal(17,6) NULL,--������
	AuxUnitCode [nvarchar] (20)  NULL ,--��������λ���� 
	cPart [nvarchar] (20)  NULL ,--��Ʒ����
	iUnitQty decimal(17,6) NULL,--��̨����
	cPerform [nvarchar] (50)  NULL ,--����/����
	iQty decimal(17,6) NULL,--�������أ���
	cbMemo [nvarchar] (200)  NULL ,--���屸ע
	cOutsourced [nvarchar] (20)  NULL ,--��Э/�⹺
	cbCloser [nvarchar] (20)  NULL ,--�ر���
	cbCloseDate [datetime] NULL ,--�ر�����
	iLLQty decimal(17,6) NULL,--�ۼ�������
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

--1 ��Ŀ����ƻ��������
--select * from  [EF_ProjectMRPChanged]
--2 ��Ŀ����ƻ�����ӱ�
--select * from  [EF_ProjectMRPChangeds]



/******************Contract structure *********************/
print '1 dbo.EF_ProjectMRPChanged  ��Ŀ����ƻ�������� ...'
if not exists (select * from sysobjects where id = object_id('dbo.EF_ProjectMRPChanged') and sysstat & 0xf = 3)
BEGIN
CREATE TABLE [EF_ProjectMRPChanged] (
---------------------------------------------------------------------------------------------------------------------------
--���������׼�ֶ�
	[id] bigint NOT NULL ,							--����ID
	[ccode] [nvarchar] (20)  NULL ,					--���ݱ���
	[ddate] [datetime] NULL ,  						--��������
	[cmaker] [nvarchar] (20)  NULL ,  				--�Ƶ���
	[cmakerddate] [datetime] NULL ,  				--�Ƶ�����
	[cmodifer] [nvarchar] (30)  NULL ,				--�����
	[cmodiferDate] [datetime] NULL ,				--�������
	cmodifier [nvarchar] (30)  NULL ,				--�޸���
	dmoddate [datetime] NULL ,						--�޸�����
	dmodifysystime [datetime] NULL ,				--�޸�ʱ��
	[cverifier] [nvarchar] (20)  NULL ,  			--�����
	[dverifydate] [datetime] NULL  			,  		--�������
	[ccloser]  [nvarchar] (20)  NULL,				--�ر���
	[dcloserdate]  [datetime] NULL,                 --�ر�����
	[vt_id] [int]  NULL ,							--��ʾģ���
	[ufts] [timestamp] NULL ,						--ʱ���
	[cvouchtype]	[nvarchar](50) NULL,			--��������
	[t_cdepcode]	[nvarchar](50) NULL,			--���ű���
	[t_cpersoncode] [nvarchar](50) NULL,			--��Ա����
	[t_ccuscode]	[nvarchar](50) NULL,			--�ͻ�����
	[t_cvencode]	[nvarchar](50) NULL,			--��Ӧ�̱���
	[t_cwhcode]		[nvarchar](50) NULL,			--�ֿ����
	[t_cinvcode]	[nvarchar](50) NULL,			--�������
	[t_cfree1] [nvarchar](20) NULL,					--���������1
	[t_cfree2] [nvarchar](20) NULL,					--���������2
	[t_cfree3] [nvarchar](20) NULL,					--���������3
	[t_cfree4] [nvarchar](20) NULL,					--���������4
	[t_cfree5] [nvarchar](20) NULL,					--���������5
	[t_cfree6] [nvarchar](20) NULL,					--���������6
	[t_cfree7] [nvarchar](20) NULL,					--���������7
	[t_cfree8] [nvarchar](20) NULL,					--���������8
	[t_cfree9] [nvarchar](20) NULL,					--���������9
	[t_cfree10] [nvarchar](20) NULL,				--���������10
--------------------------------------------------------------------------------------------------------------------------
--������ר��
	[ireturncount] [int] NULL ,					--��ش���(������ר��)
	[iswfcontrolled] [int] NULL ,				--���������ñ�־ 0 δ���� 1���� 2�ύ 
	[iverifystate] [int] NULL ,					--������״̬
	[VoucherId] [int] NULL ,					--����ؼ���=ID
	[VoucherCode] [nvarchar] (30)  NULL ,		--���ݱ��=ccode
	[VoucherType] [nvarchar] (30)  NULL ,		--�������ͺ�=CardNumber
---------------------------------------------------------------------------------------------------------------------------
	[define1] [nvarchar] (20)  NULL ,			--��ͷ�Զ�����1
	[define2] [nvarchar] (20)  NULL ,			--��ͷ�Զ�����2
	[define3] [nvarchar] (20)  NULL ,			--��ͷ�Զ�����3
	[define4] [datetime] NULL ,					--��ͷ�Զ�����4
	[define5] [int] NULL ,						--��ͷ�Զ�����5
	[define6] [datetime] NULL ,					--��ͷ�Զ�����6
	[define7] [float] NULL ,					--��ͷ�Զ�����7
	[define8] [nvarchar] (20)  NULL ,			--��ͷ�Զ�����8
	[define9] [nvarchar] (20)  NULL ,			--��ͷ�Զ�����9
	[define10] [nvarchar] (60)  NULL ,			--��ͷ�Զ�����10
	[define11] [nvarchar] (120)  NULL ,			--��ͷ�Զ�����11
	[define12] [nvarchar] (120)  NULL ,			--��ͷ�Զ�����12
	[define13] [nvarchar] (120)  NULL ,			--��ͷ�Զ�����13
	[define14] [nvarchar] (120)  NULL ,			--��ͷ�Զ�����14
	[define15] [int] NULL ,						--��ͷ�Զ�����15
	[define16] [float] NULL ,					--��ͷ�Զ�����16
-----------------------------------------------------------------------------------------------------------
--���ϲ���Ϊ�����ֶ�,���²��ָ���ҵ����Ҫ���� --ahzzd	20100505
	cItem_class [nvarchar] (60) NULL,
	citem_cname [nvarchar] (120) NULL,
	cItemCode [nvarchar] (60) NULL,
	cItemName [nvarchar] (120) NULL,
	iPQty decimal(17,6) NULL,--����
	cMemo [nvarchar] (200)  NULL ,
	cPMRPCode [nvarchar] (60) NULL,--��Ŀ����ƻ�����
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
print '2 dbo.EF_ProjectMRPChangeds   ��Ŀ����ƻ�����ӱ�...'
--drop table EF_ProjectMRPChangeds
if not exists (select * from sysobjects where id = object_id('dbo.EF_ProjectMRPChangeds') and sysstat & 0xf = 3)
BEGIN
CREATE TABLE [EF_ProjectMRPChangeds] (
---------------------------------------------------------------------------------------------------------------------------
--�����ӱ��׼�ֶ�
	[autoid] bigint NOT NULL ,						--�ӱ�ؼ���
	[id] bigint NOT NULL ,							--����ؼ���
	[b_cdepcode] [nvarchar](50) NULL,				--���ű���
	[b_cpersoncode] [nvarchar](50) NULL,			--��Ա����
	[b_ccuscode] [nvarchar](50) NULL,				--�ͻ�����
	[b_cvencode] [nvarchar](50) NULL,				--��Ӧ�̱���
	[b_cwhcode] [nvarchar](50) NULL,				--�ֿ����
	[b_cinvcode] [nvarchar](50) NULL,				--�������
	[b_cfree1] [nvarchar](20) NULL,					--���������1
	[b_cfree2] [nvarchar](20) NULL,					--���������2
	[b_cfree3] [nvarchar](20) NULL,					--���������3
	[b_cfree4] [nvarchar](20) NULL,					--���������4
	[b_cfree5] [nvarchar](20) NULL,					--���������5
	[b_cfree6] [nvarchar](20) NULL,					--���������6
	[b_cfree7] [nvarchar](20) NULL,					--���������7
	[b_cfree8] [nvarchar](20) NULL,					--���������8
	[b_cfree9] [nvarchar](20) NULL,					--���������9
	[b_cfree10] [nvarchar](20) NULL,				--���������10
	[define22] [nvarchar] (20)  NULL ,				--�����Զ�����1
	[define23] [nvarchar] (20)  NULL ,				--�����Զ�����2
	[define24] [nvarchar] (20)  NULL ,				--�����Զ�����3
	[define25] [nvarchar] (20)  NULL ,				--�����Զ�����4
	[define26] [float] NULL ,						--�����Զ�����5
	[define27] [float] NULL ,						--�����Զ�����6
	[define28] [nvarchar] (20)  NULL ,				--�����Զ�����7
	[define29] [nvarchar] (20)  NULL ,				--�����Զ�����8
	[define30] [nvarchar] (20)  NULL ,				--�����Զ�����9
	[define31] [nvarchar] (20)  NULL ,				--�����Զ�����10
	[define32] [nvarchar] (20)  NULL ,				--�����Զ�����11
	[define33] [nvarchar] (20)  NULL ,				--�����Զ�����12
	[define34] [int] NULL ,							--�����Զ�����13
	[define35] [int] NULL ,							--�����Զ�����14
	[define36] [datetime] NULL ,					--�����Զ�����15
	[define37] [datetime] NULL ,					--�����Զ�����16
-----------------------------------------------------------------------------------------------------------
--���ϲ���Ϊ�����ֶ�,���²��ָ���ҵ����Ҫ����
	iinvexchrate decimal(17,6) NULL,--������
	AuxUnitCode [nvarchar] (20)  NULL ,--��������λ���� 
	cPart [nvarchar] (20)  NULL ,--��Ʒ����
	iUnitQty decimal(17,6) NULL,--��̨����
	iUnitQtyOld  decimal(17,6) NULL,--ԭ��̨����
	cPerform [nvarchar] (50)  NULL ,--����/����
	iQty decimal(17,6) NULL,--�������أ���
	iQtyOld  decimal(17,6) NULL,--ԭ�������أ���
	cbMemo [nvarchar] (200)  NULL ,--���屸ע
	cOutsourced [nvarchar] (20)  NULL ,--��Э/�⹺
	cbCloser [nvarchar] (20)  NULL ,--�ر���
	cbCloseDate [datetime] NULL ,--�ر�����
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

--1 �ɹ��ƻ�����
--select * from  [EF_ProcurementPlan]
--2 �ɹ��ƻ��ӱ�
--select * from  [EF_ProcurementPlans]



/******************Contract structure *********************/
print '1 dbo.EF_ProcurementPlan  �ɹ��ƻ����� ...'
if not exists (select * from sysobjects where id = object_id('dbo.EF_ProcurementPlan') and sysstat & 0xf = 3)
BEGIN
CREATE TABLE [EF_ProcurementPlan] (
---------------------------------------------------------------------------------------------------------------------------
--���������׼�ֶ�
	[id] [bigint] NOT NULL ,							--����ID
	[ccode] [nvarchar] (20)  NULL ,					--���ݱ���
	[ddate] [datetime] NULL ,  						--��������
	[cmaker] [nvarchar] (20)  NULL ,  				--�Ƶ���
	[cmakerddate] [datetime] NULL ,  				--�Ƶ�����
	[cmodifer] [nvarchar] (30)  NULL ,				--�����
	[cmodiferDate] [datetime] NULL ,				--�������
	cmodifier [nvarchar] (30)  NULL ,				--�޸���
	dmoddate [datetime] NULL ,						--�޸�����
	dmodifysystime [datetime] NULL ,				--�޸�ʱ��
	[cverifier] [nvarchar] (20)  NULL ,  			--�����
	[dverifydate] [datetime] NULL  			,  		--�������
	[ccloser]  [nvarchar] (20)  NULL,				--�ر���
	[dcloserdate]  [datetime] NULL,                 --�ر�����
	[vt_id] [int]  NULL ,							--��ʾģ���
	[ufts] [timestamp] NULL ,						--ʱ���
	[cvouchtype]	[nvarchar](50) NULL,			--��������
	[t_cdepcode]	[nvarchar](50) NULL,			--���ű���
	[t_cpersoncode] [nvarchar](50) NULL,			--��Ա����
	[t_ccuscode]	[nvarchar](50) NULL,			--�ͻ�����
	[t_cvencode]	[nvarchar](50) NULL,			--��Ӧ�̱���
	[t_cwhcode]		[nvarchar](50) NULL,			--�ֿ����
	[t_cinvcode]	[nvarchar](50) NULL,			--�������
	[t_cfree1] [nvarchar](20) NULL,					--���������1
	[t_cfree2] [nvarchar](20) NULL,					--���������2
	[t_cfree3] [nvarchar](20) NULL,					--���������3
	[t_cfree4] [nvarchar](20) NULL,					--���������4
	[t_cfree5] [nvarchar](20) NULL,					--���������5
	[t_cfree6] [nvarchar](20) NULL,					--���������6
	[t_cfree7] [nvarchar](20) NULL,					--���������7
	[t_cfree8] [nvarchar](20) NULL,					--���������8
	[t_cfree9] [nvarchar](20) NULL,					--���������9
	[t_cfree10] [nvarchar](20) NULL,				--���������10
--------------------------------------------------------------------------------------------------------------------------
--������ר��
	[ireturncount] [int] NULL ,					--��ش���(������ר��)
	[iswfcontrolled] [int] NULL ,				--���������ñ�־ 0 δ���� 1���� 2�ύ 
	[iverifystate] [int] NULL ,					--������״̬
	[VoucherId] [int] NULL ,					--����ؼ���=ID
	[VoucherCode] [nvarchar] (30)  NULL ,		--���ݱ��=ccode
	[VoucherType] [nvarchar] (30)  NULL ,		--�������ͺ�=CardNumber
---------------------------------------------------------------------------------------------------------------------------
	[define1] [nvarchar] (20)  NULL ,			--��ͷ�Զ�����1
	[define2] [nvarchar] (20)  NULL ,			--��ͷ�Զ�����2
	[define3] [nvarchar] (20)  NULL ,			--��ͷ�Զ�����3
	[define4] [datetime] NULL ,					--��ͷ�Զ�����4
	[define5] [int] NULL ,						--��ͷ�Զ�����5
	[define6] [datetime] NULL ,					--��ͷ�Զ�����6
	[define7] [float] NULL ,					--��ͷ�Զ�����7
	[define8] [nvarchar] (20)  NULL ,			--��ͷ�Զ�����8
	[define9] [nvarchar] (20)  NULL ,			--��ͷ�Զ�����9
	[define10] [nvarchar] (60)  NULL ,			--��ͷ�Զ�����10
	[define11] [nvarchar] (120)  NULL ,			--��ͷ�Զ�����11
	[define12] [nvarchar] (120)  NULL ,			--��ͷ�Զ�����12
	[define13] [nvarchar] (120)  NULL ,			--��ͷ�Զ�����13
	[define14] [nvarchar] (120)  NULL ,			--��ͷ�Զ�����14
	[define15] [int] NULL ,						--��ͷ�Զ�����15
	[define16] [float] NULL ,					--��ͷ�Զ�����16
-----------------------------------------------------------------------------------------------------------
--���ϲ���Ϊ�����ֶ�,���²��ָ���ҵ����Ҫ���� --ahzzd	20100505
	cItem_class [nvarchar] (60) NULL,
	citem_cname [nvarchar] (120) NULL,
	cItemCode [nvarchar] (60) NULL,
	cItemName [nvarchar] (120) NULL,
	iPQty decimal(17,6) NULL,--����
	cMemo [nvarchar] (200)  NULL ,
	cPMRPCode [nvarchar] (60) NULL,--�嵥���
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
print '2 dbo.EF_ProcurementPlans   �ɹ��ƻ��ӱ�...'
--drop table EF_ProcurementPlans
if not exists (select * from sysobjects where id = object_id('dbo.EF_ProcurementPlans') and sysstat & 0xf = 3)
BEGIN
CREATE TABLE [EF_ProcurementPlans] (
---------------------------------------------------------------------------------------------------------------------------
--�����ӱ��׼�ֶ�
	[autoid] [bigint] NOT NULL ,						--�ӱ�ؼ���
	[id] [bigint] NOT NULL ,							--����ؼ���
	[b_cdepcode] [nvarchar](50) NULL,				--���ű���
	[b_cpersoncode] [nvarchar](50) NULL,			--��Ա����
	[b_ccuscode] [nvarchar](50) NULL,				--�ͻ�����
	[b_cvencode] [nvarchar](50) NULL,				--��Ӧ�̱���
	[b_cwhcode] [nvarchar](50) NULL,				--�ֿ����
	[b_cinvcode] [nvarchar](50) NULL,				--�������
	[b_cfree1] [nvarchar](20) NULL,					--���������1
	[b_cfree2] [nvarchar](20) NULL,					--���������2
	[b_cfree3] [nvarchar](20) NULL,					--���������3
	[b_cfree4] [nvarchar](20) NULL,					--���������4
	[b_cfree5] [nvarchar](20) NULL,					--���������5
	[b_cfree6] [nvarchar](20) NULL,					--���������6
	[b_cfree7] [nvarchar](20) NULL,					--���������7
	[b_cfree8] [nvarchar](20) NULL,					--���������8
	[b_cfree9] [nvarchar](20) NULL,					--���������9
	[b_cfree10] [nvarchar](20) NULL,				--���������10
	[define22] [nvarchar] (20)  NULL ,				--�����Զ�����1
	[define23] [nvarchar] (20)  NULL ,				--�����Զ�����2
	[define24] [nvarchar] (20)  NULL ,				--�����Զ�����3
	[define25] [nvarchar] (20)  NULL ,				--�����Զ�����4
	[define26] [float] NULL ,						--�����Զ�����5
	[define27] [float] NULL ,						--�����Զ�����6
	[define28] [nvarchar] (20)  NULL ,				--�����Զ�����7
	[define29] [nvarchar] (20)  NULL ,				--�����Զ�����8
	[define30] [nvarchar] (20)  NULL ,				--�����Զ�����9
	[define31] [nvarchar] (20)  NULL ,				--�����Զ�����10
	[define32] [nvarchar] (20)  NULL ,				--�����Զ�����11
	[define33] [nvarchar] (20)  NULL ,				--�����Զ�����12
	[define34] [int] NULL ,							--�����Զ�����13
	[define35] [int] NULL ,							--�����Զ�����14
	[define36] [datetime] NULL ,					--�����Զ�����15
	[define37] [datetime] NULL ,					--�����Զ�����16
-----------------------------------------------------------------------------------------------------------
--���ϲ���Ϊ�����ֶ�,���²��ָ���ҵ����Ҫ����
	iinvexchrate decimal(17,6) NULL,--������
	AuxUnitCode [nvarchar] (20)  NULL ,--��������λ���� 
	cPerform [nvarchar] (50)  NULL ,--����/����
	cbMemo [nvarchar] (200)  NULL ,--���屸ע
	cOutsourced [nvarchar] (20)  NULL ,--��Э/�⹺
	iMRPQty decimal(17,6) NULL,--����ƻ���
	iMRPQtyL  decimal(17,6) NULL,--�ۼ�����ƻ���
	iSafeNum  decimal(17,6) NULL,--��ȫ���
	iStock  decimal(17,6) NULL,--�ִ���
	iInQty  decimal(17,6) NULL,--Ԥ�������
	iOutQty  decimal(17,6) NULL,--Ԥ�Ƴ�����
	iKYL  decimal(17,6) NULL,--������
	iMinQty  decimal(17,6) NULL,--�ɹ���������
	iJHQty   decimal(17,6) NULL,--�ɹ��ƻ���
	iJYQty   decimal(17,6) NULL,--����ɹ���
	iSJQty   decimal(17,6) NULL,--ʵ��������
	cbCloser [nvarchar] (20)  NULL ,--�ر���
	cbCloseDate [datetime] NULL ,--�ر�����
	iQgLQty decimal(17,6) NULL,--�ۼ��빺��
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

--1 ��ĿͶ�걨������
--select * from  [EF_Bid]
--2 ��ĿͶ�걨���ӱ�
--select * from  [EF_Bids]



/******************Contract structure *********************/
print '1 dbo.EF_Bid  ��ĿͶ�걨������ ...'
if not exists (select * from sysobjects where id = object_id('dbo.EF_Bid') and sysstat & 0xf = 3)
BEGIN
CREATE TABLE [EF_Bid] (
---------------------------------------------------------------------------------------------------------------------------
--���������׼�ֶ�
	[id] [bigint] NOT NULL ,							--����ID
	[ccode] [nvarchar] (20)  NULL ,					--���ݱ���
	[ddate] [datetime] NULL ,  						--��������
	[cmaker] [nvarchar] (20)  NULL ,  				--�Ƶ���
	[cmakerddate] [datetime] NULL ,  				--�Ƶ�����
	[cmodifer] [nvarchar] (30)  NULL ,				--�����
	[cmodiferDate] [datetime] NULL ,				--�������
	cmodifier [nvarchar] (30)  NULL ,				--�޸���
	dmoddate [datetime] NULL ,						--�޸�����
	dmodifysystime [datetime] NULL ,				--�޸�ʱ��
	[cverifier] [nvarchar] (20)  NULL ,  			--�����
	[dverifydate] [datetime] NULL  			,  		--�������
	[ccloser]  [nvarchar] (20)  NULL,				--�ر���
	[dcloserdate]  [datetime] NULL,                 --�ر�����
	[vt_id] [int]  NULL ,							--��ʾģ���
	[ufts] [timestamp] NULL ,						--ʱ���
	[cvouchtype]	[nvarchar](50) NULL,			--��������
	[t_cdepcode]	[nvarchar](50) NULL,			--���ű���
	[t_cpersoncode] [nvarchar](50) NULL,			--��Ա����
	[t_ccuscode]	[nvarchar](50) NULL,			--�ͻ�����
	[t_cvencode]	[nvarchar](50) NULL,			--��Ӧ�̱���
	[t_cwhcode]		[nvarchar](50) NULL,			--�ֿ����
	[t_cinvcode]	[nvarchar](50) NULL,			--�������
	[t_cfree1] [nvarchar](20) NULL,					--���������1
	[t_cfree2] [nvarchar](20) NULL,					--���������2
	[t_cfree3] [nvarchar](20) NULL,					--���������3
	[t_cfree4] [nvarchar](20) NULL,					--���������4
	[t_cfree5] [nvarchar](20) NULL,					--���������5
	[t_cfree6] [nvarchar](20) NULL,					--���������6
	[t_cfree7] [nvarchar](20) NULL,					--���������7
	[t_cfree8] [nvarchar](20) NULL,					--���������8
	[t_cfree9] [nvarchar](20) NULL,					--���������9
	[t_cfree10] [nvarchar](20) NULL,				--���������10
--------------------------------------------------------------------------------------------------------------------------
--������ר��
	[ireturncount] [int] NULL ,					--��ش���(������ר��)
	[iswfcontrolled] [int] NULL ,				--���������ñ�־ 0 δ���� 1���� 2�ύ 
	[iverifystate] [int] NULL ,					--������״̬
	[VoucherId] [int] NULL ,					--����ؼ���=ID
	[VoucherCode] [nvarchar] (30)  NULL ,		--���ݱ��=ccode
	[VoucherType] [nvarchar] (30)  NULL ,		--�������ͺ�=CardNumber
---------------------------------------------------------------------------------------------------------------------------
	[define1] [nvarchar] (20)  NULL ,			--��ͷ�Զ�����1
	[define2] [nvarchar] (20)  NULL ,			--��ͷ�Զ�����2
	[define3] [nvarchar] (20)  NULL ,			--��ͷ�Զ�����3
	[define4] [datetime] NULL ,					--��ͷ�Զ�����4
	[define5] [int] NULL ,						--��ͷ�Զ�����5
	[define6] [datetime] NULL ,					--��ͷ�Զ�����6
	[define7] [float] NULL ,					--��ͷ�Զ�����7
	[define8] [nvarchar] (20)  NULL ,			--��ͷ�Զ�����8
	[define9] [nvarchar] (20)  NULL ,			--��ͷ�Զ�����9
	[define10] [nvarchar] (60)  NULL ,			--��ͷ�Զ�����10
	[define11] [nvarchar] (120)  NULL ,			--��ͷ�Զ�����11
	[define12] [nvarchar] (120)  NULL ,			--��ͷ�Զ�����12
	[define13] [nvarchar] (120)  NULL ,			--��ͷ�Զ�����13
	[define14] [nvarchar] (120)  NULL ,			--��ͷ�Զ�����14
	[define15] [int] NULL ,						--��ͷ�Զ�����15
	[define16] [float] NULL ,					--��ͷ�Զ�����16
-----------------------------------------------------------------------------------------------------------
--���ϲ���Ϊ�����ֶ�,���²��ָ���ҵ����Ҫ���� --ahzzd	20100505
	cItem_class [nvarchar] (60) NULL,
	citem_cname [nvarchar] (120) NULL,
	cItemCode [nvarchar] (60) NULL,
	cItemName [nvarchar] (120) NULL,
	iPQty decimal(17,6) NULL,--����
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
print '2 dbo.EF_Bids   ��ĿͶ�걨���ӱ�...'
if not exists (select * from sysobjects where id = object_id('dbo.EF_Bids') and sysstat & 0xf = 3)
BEGIN
CREATE TABLE [EF_Bids] (
---------------------------------------------------------------------------------------------------------------------------
--�����ӱ��׼�ֶ�
	[autoid] [bigint] NOT NULL ,						--�ӱ�ؼ���
	[id] [bigint] NOT NULL ,							--����ؼ���
	[b_cdepcode] [nvarchar](50) NULL,				--���ű���
	[b_cpersoncode] [nvarchar](50) NULL,			--��Ա����
	[b_ccuscode] [nvarchar](50) NULL,				--�ͻ�����
	[b_cvencode] [nvarchar](50) NULL,				--��Ӧ�̱���
	[b_cwhcode] [nvarchar](50) NULL,				--�ֿ����
	[b_cinvcode] [nvarchar](50) NULL,				--�������
	[b_cfree1] [nvarchar](20) NULL,					--���������1
	[b_cfree2] [nvarchar](20) NULL,					--���������2
	[b_cfree3] [nvarchar](20) NULL,					--���������3
	[b_cfree4] [nvarchar](20) NULL,					--���������4
	[b_cfree5] [nvarchar](20) NULL,					--���������5
	[b_cfree6] [nvarchar](20) NULL,					--���������6
	[b_cfree7] [nvarchar](20) NULL,					--���������7
	[b_cfree8] [nvarchar](20) NULL,					--���������8
	[b_cfree9] [nvarchar](20) NULL,					--���������9
	[b_cfree10] [nvarchar](20) NULL,				--���������10
	[define22] [nvarchar] (20)  NULL ,				--�����Զ�����1
	[define23] [nvarchar] (20)  NULL ,				--�����Զ�����2
	[define24] [nvarchar] (20)  NULL ,				--�����Զ�����3
	[define25] [nvarchar] (20)  NULL ,				--�����Զ�����4
	[define26] [float] NULL ,						--�����Զ�����5
	[define27] [float] NULL ,						--�����Զ�����6
	[define28] [nvarchar] (20)  NULL ,				--�����Զ�����7
	[define29] [nvarchar] (20)  NULL ,				--�����Զ�����8
	[define30] [nvarchar] (20)  NULL ,				--�����Զ�����9
	[define31] [nvarchar] (20)  NULL ,				--�����Զ�����10
	[define32] [nvarchar] (20)  NULL ,				--�����Զ�����11
	[define33] [nvarchar] (20)  NULL ,				--�����Զ�����12
	[define34] [int] NULL ,							--�����Զ�����13
	[define35] [int] NULL ,							--�����Զ�����14
	[define36] [datetime] NULL ,					--�����Զ�����15
	[define37] [datetime] NULL ,					--�����Զ�����16
-----------------------------------------------------------------------------------------------------------
--���ϲ���Ϊ�����ֶ�,���²��ָ���ҵ����Ҫ����
	iinvexchrate decimal(17,6) NULL,--������
	AuxUnitCode [nvarchar] (20)  NULL ,--��������λ���� 
	cMaterialClass [nvarchar] (50)  NULL ,--�������
	iCB decimal(17,6) NULL,--�ɹ��ɱ�
	iBJ decimal(17,6) NULL,--Ͷ�걨��
	cbMemo [nvarchar] (200)  NULL ,--���屸ע
	cbCloser [nvarchar] (20)  NULL ,--�ر���
	cbCloseDate [datetime] NULL ,--�ر�����
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
