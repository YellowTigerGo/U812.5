
delete accinformation where [cID]='EP01'
insert into accinformation ([cSysID],[cID],[cName],[cCaption],[cType],[cValue],[cDefault],[bVisible],[bEnable])
values('EP','EP01','bqgcjh','�빺�����ƻ�','Boolean','0','0','0','0')
delete accinformation where [cID]='EP02'
insert into accinformation ([cSysID],[cID],[cName],[cCaption],[cType],[cValue],[cDefault],[bVisible],[bEnable])
values('EP','EP02','bllcjh','���������ƻ�','Boolean','0','0','0','0')
go

--Insert into the Table aa_columndic_base

delete aa_columndic_base where ckey ='v_ef_projectmrp'
GO
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRP','ccode','ccode',2,'�嵥���',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRP','ccusname','ccusname',3,'�ͻ�',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRP','citem_class','citem_class',4,'��Ŀ�������',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRP','citem_cname','citem_cname',5,'��Ŀ����',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRP','citemcode','citemcode',6,'��Ŀ����',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRP','citemgg','citemgg',7,'��Ʒ���',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRP','citemname','citemname',8,'��Ŀ',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRP','citemth','citemth',9,'��Ʒͼ��',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRP','cmaker','cmaker',10,'�Ƶ���',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRP','cmakerddate','cmakerddate',11,'�Ƶ�����',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRP','cmemo','cmemo',12,'��ע',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRP','cverifier','cverifier',13,'�����',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRP','ddate','ddate',14,'��������',Null,'1','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRP','dverifydate','dverifydate',15,'�������',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRP','id','id',14,'id',Null,'0','0',1500,1,Null,'1',0,'1','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRP','ipqty','ipqty',16,'����',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRP','selcol','selcol',1,'ѡ��',Null,'1','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
--Insert into the Table aa_columndic_base

--Insert into the Table aa_columndic_base

delete aa_columndic_base where ckey='v_ef_projectmrps'

GO
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRPs','autoid','autoid',15,'autoid',Null,'0','0',1500,1,Null,'1',0,'1','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRPs','b_ccomunitname','b_ccomunitname',2,'��������λ',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRPs','b_cinvcode','b_cinvcode',3,'���ϱ���',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRPs','b_cinvname','b_cinvname',4,'�������',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRPs','b_cinvstd','b_cinvstd',5,'���',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRPs','cbclosedate','cbclosedate',6,'�ر�����',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRPs','cbcloser','cbcloser',7,'�ر���',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRPs','cbmemo','cbmemo',8,'���屸ע',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRPs','coutsourced','coutsourced',9,'��Э�⹺',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRPs','cpart','cpart',10,'��Ʒ����',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRPs','cperform','cperform',11,'��������',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRPs','id','id',14,'id',Null,'0','0',1500,1,Null,'1',0,'1','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRPs','iqty','iqty',12,'�������أ���',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRPs','iunitqty','iunitqty',13,'��̨����',Null,'0','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')
insert into aa_columndic_base ([cKey],[cFld],[cQryField],[iColPos],[cCaption],[cCaptionPar],[bFixed],[bDisp],[iColWidth],[iAlign],[cOrder],[bLock],[iMergeCols],[bMustSel],[bNeedSum],[LocaleID],[IsEnum],[EnumType],[EnumTypeString],[bFilter],[bMerge],[CanModify],[ReferType],[bHideInColSet],[cSumType],[iFieldType],[bExtended],[EnumShowType],[iLinkType],[bMustInput],[bCanModifyMustInput],[bCanModifyOrder])
values('V_EF_ProjectMRPs','selcol','selcol',1,'ѡ��',Null,'1','1',1500,1,Null,'0',0,'0','0','zh-cn','0',Null,Null,'0','0','0',Null,'0',Null,Null,Null,Null,Null,'0','0','0')

--Insert into the Table sa_refervoucherconfig

delete sa_refervoucherconfig where cardnum='ep0203'

GO
insert into sa_refervoucherconfig ([cardnum],[referkey],[maincolumnkey],[detailcolumnkey],[toolbarkey],[subsysid],[filtername],[maindatasource],[detaildatasource],[selecttype],[fillkey],[mainuniquekey],[pagesize],[defaultfilter],[filltype],[captionresid],[detailuniquekey],[buttonspara],[uniqueflds],[condition],[setdetailvalues],[alertflds],[uniquefldsB],[alertfldsB],[itype],[defaultsubfilter],[iGetdataOrder],[iFillTimes])
values('EP0203','EP0201','V_EF_ProjectMRP','V_EF_ProjectMRPs',Null,'EP','EP[__]EPSARefVouch0203','V_EF_ProjectMRP','V_EF_ProjectMRPs',Null,Null,'id','20',' isnull(cverifier,'''')<>'''' and isnull(ccloser,'''')='''' ',Null,Null,'autoid',Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null)

--Insert into the Table aa_columndic_base
go

declare @AppSysID int,@AppTypeID int,@AppTagID int,@EndpointID nvarchar(50),@MsgTypeID int,@MsgTypeCategoryID int,@MsgFilterID int,@ParentAppTypeID int 
select @AppSysID=IB_AppSys.ID from IB_AppSys,IB_Entities where IB_Entities.ID=IB_AppSys.EntityID and IB_Entities.EntityTag='U8API' 
if @AppSysID is null 
begin 
insert into IB_Entities(EntityTag,FriendName)values('U8API','U8API') 
insert into IB_AppSys(AppSystem,FriendName,EntityID)values('U8API','U8API',@@identity) 
set @AppSysID=@@identity 
end 
set @ParentAppTypeID=0 
set @AppTypeID=0 
select @AppTypeID=ID from IB_AppType where AppType='PU' and AppSysID=@AppSysID 
if @AppTypeID is null or @AppTypeID=0 
begin 
insert into IB_AppType(AppType,FriendName,AppSysID)values('PU','�ɹ�����',@AppSysID) 
select @AppTypeID=@@identity 
if @ParentAppTypeID>0 
begin 
update IB_AppType set ParentID=@ParentAppTypeID where ID=@AppTypeID 
end 
end 
set @ParentAppTypeID=@AppTypeID 
set @AppTypeID=0 
select @AppTypeID=ID from IB_AppType where AppType='PurchaseRequisition' and AppSysID=@AppSysID 
if @AppTypeID is null or @AppTypeID=0 
begin 
insert into IB_AppType(AppType,FriendName,AppSysID)values('PurchaseRequisition','�빺��',@AppSysID) 
select @AppTypeID=@@identity 
if @ParentAppTypeID>0 
begin 
update IB_AppType set ParentID=@ParentAppTypeID where ID=@AppTypeID 
end 
end 
set @ParentAppTypeID=@AppTypeID 
select @AppTagID=ID from IB_AppTag where AppTag='PurchaseRequisition_Delete_After' and AppTypeID=@AppTypeID 
if @AppTagID is null 
begin 
insert into IB_AppTag(AppTag,FriendName,AppTypeID,ExtendProperties,Description,Customize,IsPlugin) 
 values('PurchaseRequisition_Delete_After','PurchaseRequisition_Delete_After',@AppTypeID,'','',0,1) 
set @AppTagID=@@identity 
end 
else 
begin 
update IB_AppTag set ExtendProperties='',Description='',Customize=0 where ID=@AppTagID 
end 
select top 1 @EndpointID=ID from IB_EndPoint where AppTagID=@AppTagID 
if @EndpointID is null 
begin 
insert into IB_EndPoint(Address,ProtocolID,ProtocolParams,AppTagID) 
 values('','MSDCOM_RPC','<?xml version="1.0" encoding="utf-8"?><momEndPointProtocol name="MSDCOM_RPC"><runtimeParameters><param name="DLLFilePath" value="%U8SOFT%\EF\EFWXZG\EF_Refer.dll" description="" display="true" /><param name="ProgID" value="EF_Refer.clsVoucherPlugin" description="" display="true" /><param name="Server" value="." description="" display="false" /><param name="ClassName" value="clsVoucherPlugin" description="" display="true" /><param name="MethodName" value="PurchaseRequisition_Delete_After" description="" display="true" /><param name="ComPlusTransaction" value="False" description="" display="true" /><param name="IsPlugin" value="true" description="" display="true" /></runtimeParameters></momEndPointProtocol>',@AppTagID) 
end 
else 
begin 
update IB_EndPoint set Address='',ProtocolID='MSDCOM_RPC',ProtocolParams='<?xml version="1.0" encoding="utf-8"?><momEndPointProtocol name="MSDCOM_RPC"><runtimeParameters><param name="DLLFilePath" value="%U8SOFT%\EF\EFWXZG\EF_Refer.dll" description="" display="true" /><param name="ProgID" value="EF_Refer.clsVoucherPlugin" description="" display="true" /><param name="Server" value="." description="" display="false" /><param name="ClassName" value="clsVoucherPlugin" description="" display="true" /><param name="MethodName" value="PurchaseRequisition_Delete_After" description="" display="true" /><param name="ComPlusTransaction" value="False" description="" display="true" /><param name="IsPlugin" value="true" description="" display="true" /></runtimeParameters></momEndPointProtocol>' where ID=@EndpointID 
end 
select @MsgTypeCategoryID=ID from MOM_MsgTypeCategories where MsgTypeCategory='����¼�' 
select @MsgFilterID=ID from IB_MsgFilter where FilterName='XPathFilter' 
select @MsgTypeID=ID from IB_MsgType where MsgType='Delete_After' and AppTypeID=@AppTypeID 
if @MsgTypeID is null 
begin 
insert into IB_MsgType(MsgType,FriendName,MsgSchema,AppTypeID,MsgTypeCategoryID,MsgFilterID,ExtendProperties) values('Delete_After','ɾ�����¼�','<?xml version="1.0" encoding="utf-8" ?><serviceInterface transactionType=""><operation name="" ><parameters><parameter index="0" name="domhead" type="MSXML2.IXMLDOMDocument2" direction="inout" desc="��ͷ" optional="false" byRef="true" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /><parameter index="1" name="dombody" type="MSXML2.IXMLDOMDocument2" direction="inout" desc="����" optional="false" byRef="true" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /><parameter index="2" name="errmsg" type="string" direction="inout" desc="���ش�����Ϣ" optional="false" byRef="false" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /><parameter type="bool" direction="retval" desc="����ֵ: true:�ɹ�, false: ʧ��" byRef="false" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /></parameters></operation></serviceInterface>',@AppTypeID,@MsgTypeCategoryID,@MsgFilterID,'') 
 set @MsgTypeID=@@identity 
end 
else 
begin 
update IB_MsgType set MsgSchema='<?xml version="1.0" encoding="utf-8" ?><serviceInterface transactionType=""><operation name="" ><parameters><parameter index="0" name="domhead" type="MSXML2.IXMLDOMDocument2" direction="inout" desc="��ͷ" optional="false" byRef="true" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /><parameter index="1" name="dombody" type="MSXML2.IXMLDOMDocument2" direction="inout" desc="����" optional="false" byRef="true" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /><parameter index="2" name="errmsg" type="string" direction="inout" desc="���ش�����Ϣ" optional="false" byRef="false" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /><parameter type="bool" direction="retval" desc="����ֵ: true:�ɹ�, false: ʧ��" byRef="false" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /></parameters></operation></serviceInterface>',ExtendProperties='' where ID=@MsgTypeID 
end 
if (select count(*) from IB_Event_Plugin where AppTagID=@AppTagID and MsgTypeID=@MsgTypeID) = 0 
begin 
insert into IB_Event_Plugin(AppTagID,MsgTypeID,AccID,OrderNO,IsSyncOrAsync,Description,Disabled,UnVisible,UnDeleted) values(@AppTagID,@MsgTypeID,'',1,0,'',0,0,0)  
end 
go
declare @AppSysID int,@AppTypeID int,@AppTagID int,@EndpointID nvarchar(50),@MsgTypeID int,@MsgTypeCategoryID int,@MsgFilterID int,@ParentAppTypeID int 
select @AppSysID=IB_AppSys.ID from IB_AppSys,IB_Entities where IB_Entities.ID=IB_AppSys.EntityID and IB_Entities.EntityTag='U8API' 
if @AppSysID is null 
begin 
insert into IB_Entities(EntityTag,FriendName)values('U8API','U8API') 
insert into IB_AppSys(AppSystem,FriendName,EntityID)values('U8API','U8API',@@identity) 
set @AppSysID=@@identity 
end 
set @ParentAppTypeID=0 
set @AppTypeID=0 
select @AppTypeID=ID from IB_AppType where AppType='PU' and AppSysID=@AppSysID 
if @AppTypeID is null or @AppTypeID=0 
begin 
insert into IB_AppType(AppType,FriendName,AppSysID)values('PU','�ɹ�����',@AppSysID) 
select @AppTypeID=@@identity 
if @ParentAppTypeID>0 
begin 
update IB_AppType set ParentID=@ParentAppTypeID where ID=@AppTypeID 
end 
end 
set @ParentAppTypeID=@AppTypeID 
set @AppTypeID=0 
select @AppTypeID=ID from IB_AppType where AppType='PurchaseRequisition' and AppSysID=@AppSysID 
if @AppTypeID is null or @AppTypeID=0 
begin 
insert into IB_AppType(AppType,FriendName,AppSysID)values('PurchaseRequisition','�빺��',@AppSysID) 
select @AppTypeID=@@identity 
if @ParentAppTypeID>0 
begin 
update IB_AppType set ParentID=@ParentAppTypeID where ID=@AppTypeID 
end 
end 
set @ParentAppTypeID=@AppTypeID 
select @AppTagID=ID from IB_AppTag where AppTag='PurchaseRequisition_Save_After' and AppTypeID=@AppTypeID 
if @AppTagID is null 
begin 
insert into IB_AppTag(AppTag,FriendName,AppTypeID,ExtendProperties,Description,Customize,IsPlugin) 
 values('PurchaseRequisition_Save_After','PurchaseRequisition_Save_After',@AppTypeID,'','',0,1) 
set @AppTagID=@@identity 
end 
else 
begin 
update IB_AppTag set ExtendProperties='',Description='',Customize=0 where ID=@AppTagID 
end 
select top 1 @EndpointID=ID from IB_EndPoint where AppTagID=@AppTagID 
if @EndpointID is null 
begin 
insert into IB_EndPoint(Address,ProtocolID,ProtocolParams,AppTagID) 
 values('','MSDCOM_RPC','<?xml version="1.0" encoding="utf-8"?><momEndPointProtocol name="MSDCOM_RPC"><runtimeParameters><param name="DLLFilePath" value="%U8SOFT%\EF\EFWXZG\EF_Refer.dll" description="" display="true" /><param name="ProgID" value="EF_Refer.clsVoucherPlugin" description="" display="true" /><param name="Server" value="." description="" display="false" /><param name="ClassName" value="clsVoucherPlugin" description="" display="true" /><param name="MethodName" value="PurchaseRequisition_Save_After" description="" display="true" /><param name="ComPlusTransaction" value="False" description="" display="true" /><param name="IsPlugin" value="true" description="" display="true" /></runtimeParameters></momEndPointProtocol>',@AppTagID) 
end 
else 
begin 
update IB_EndPoint set Address='',ProtocolID='MSDCOM_RPC',ProtocolParams='<?xml version="1.0" encoding="utf-8"?><momEndPointProtocol name="MSDCOM_RPC"><runtimeParameters><param name="DLLFilePath" value="%U8SOFT%\EF\EFWXZG\EF_Refer.dll" description="" display="true" /><param name="ProgID" value="EF_Refer.clsVoucherPlugin" description="" display="true" /><param name="Server" value="." description="" display="false" /><param name="ClassName" value="clsVoucherPlugin" description="" display="true" /><param name="MethodName" value="PurchaseRequisition_Save_After" description="" display="true" /><param name="ComPlusTransaction" value="False" description="" display="true" /><param name="IsPlugin" value="true" description="" display="true" /></runtimeParameters></momEndPointProtocol>' where ID=@EndpointID 
end 
select @MsgTypeCategoryID=ID from MOM_MsgTypeCategories where MsgTypeCategory='����¼�' 
select @MsgFilterID=ID from IB_MsgFilter where FilterName='XPathFilter' 
select @MsgTypeID=ID from IB_MsgType where MsgType='Save_After' and AppTypeID=@AppTypeID 
if @MsgTypeID is null 
begin 
insert into IB_MsgType(MsgType,FriendName,MsgSchema,AppTypeID,MsgTypeCategoryID,MsgFilterID,ExtendProperties) values('Save_After','������¼�','<?xml version="1.0" encoding="utf-8" ?><serviceInterface transactionType=""><operation name="" ><parameters><parameter index="0" name="domhead" type="MSXML2.IXMLDOMDocument2" direction="inout" desc="��ͷ" optional="false" byRef="true" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /><parameter index="1" name="dombody" type="MSXML2.IXMLDOMDocument2" direction="inout" desc="����" optional="false" byRef="true" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /><parameter index="2" name="errmsg" type="string" direction="inout" desc="���ش�����Ϣ" optional="false" byRef="false" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /><parameter type="bool" direction="retval" desc="����ֵ: true:�ɹ�, false: ʧ��" byRef="false" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /></parameters></operation></serviceInterface>',@AppTypeID,@MsgTypeCategoryID,@MsgFilterID,'') 
 set @MsgTypeID=@@identity 
end 
else 
begin 
update IB_MsgType set MsgSchema='<?xml version="1.0" encoding="utf-8" ?><serviceInterface transactionType=""><operation name="" ><parameters><parameter index="0" name="domhead" type="MSXML2.IXMLDOMDocument2" direction="inout" desc="��ͷ" optional="false" byRef="true" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /><parameter index="1" name="dombody" type="MSXML2.IXMLDOMDocument2" direction="inout" desc="����" optional="false" byRef="true" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /><parameter index="2" name="errmsg" type="string" direction="inout" desc="���ش�����Ϣ" optional="false" byRef="false" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /><parameter type="bool" direction="retval" desc="����ֵ: true:�ɹ�, false: ʧ��" byRef="false" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /></parameters></operation></serviceInterface>',ExtendProperties='' where ID=@MsgTypeID 
end 
if (select count(*) from IB_Event_Plugin where AppTagID=@AppTagID and MsgTypeID=@MsgTypeID) = 0 
begin 
insert into IB_Event_Plugin(AppTagID,MsgTypeID,AccID,OrderNO,IsSyncOrAsync,Description,Disabled,UnVisible,UnDeleted) values(@AppTagID,@MsgTypeID,'',1,0,'',0,0,0)  
end 
go

declare @AppSysID int,@AppTypeID int,@AppTagID int,@EndpointID nvarchar(50),@MsgTypeID int,@MsgTypeCategoryID int,@MsgFilterID int,@ParentAppTypeID int 
select @AppSysID=IB_AppSys.ID from IB_AppSys,IB_Entities where IB_Entities.ID=IB_AppSys.EntityID and IB_Entities.EntityTag='U8API' 
if @AppSysID is null 
begin 
insert into IB_Entities(EntityTag,FriendName)values('U8API','U8API') 
insert into IB_AppSys(AppSystem,FriendName,EntityID)values('U8API','U8API',@@identity) 
set @AppSysID=@@identity 
end 
set @ParentAppTypeID=0 
set @AppTypeID=0 
select @AppTypeID=ID from IB_AppType where AppType='ST' and AppSysID=@AppSysID 
if @AppTypeID is null or @AppTypeID=0 
begin 
insert into IB_AppType(AppType,FriendName,AppSysID)values('ST','������',@AppSysID) 
select @AppTypeID=@@identity 
if @ParentAppTypeID>0 
begin 
update IB_AppType set ParentID=@ParentAppTypeID where ID=@AppTypeID 
end 
end 
set @ParentAppTypeID=@AppTypeID 
set @AppTypeID=0 
select @AppTypeID=ID from IB_AppType where AppType='MaterialOut' and AppSysID=@AppSysID 
if @AppTypeID is null or @AppTypeID=0 
begin 
insert into IB_AppType(AppType,FriendName,AppSysID)values('MaterialOut','���ϳ��ⵥ',@AppSysID) 
select @AppTypeID=@@identity 
if @ParentAppTypeID>0 
begin 
update IB_AppType set ParentID=@ParentAppTypeID where ID=@AppTypeID 
end 
end 
set @ParentAppTypeID=@AppTypeID 
select @AppTagID=ID from IB_AppTag where AppTag='MaterialOutSave_After' and AppTypeID=@AppTypeID 
if @AppTagID is null 
begin 
insert into IB_AppTag(AppTag,FriendName,AppTypeID,ExtendProperties,Description,Customize,IsPlugin) 
 values('MaterialOutSave_After','MaterialOutSave_After',@AppTypeID,'','',0,1) 
set @AppTagID=@@identity 
end 
else 
begin 
update IB_AppTag set ExtendProperties='',Description='',Customize=0 where ID=@AppTagID 
end 
select top 1 @EndpointID=ID from IB_EndPoint where AppTagID=@AppTagID 
if @EndpointID is null 
begin 
insert into IB_EndPoint(Address,ProtocolID,ProtocolParams,AppTagID) 
 values('','MSDCOM_RPC','<?xml version="1.0" encoding="utf-8"?><momEndPointProtocol name="MSDCOM_RPC"><runtimeParameters><param name="DLLFilePath" value="%U8SOFT%\EF\EFWXZG\EF_Refer.dll" description="" display="true" /><param name="ProgID" value="EF_Refer.clsVoucherPlugin" description="" display="true" /><param name="Server" value="." description="" display="false" /><param name="ClassName" value="clsVoucherPlugin" description="" display="true" /><param name="MethodName" value="MaterialOut_Save_After" description="" display="true" /><param name="ComPlusTransaction" value="False" description="" display="true" /><param name="IsPlugin" value="true" description="" display="true" /></runtimeParameters></momEndPointProtocol>',@AppTagID) 
end 
else 
begin 
update IB_EndPoint set Address='',ProtocolID='MSDCOM_RPC',ProtocolParams='<?xml version="1.0" encoding="utf-8"?><momEndPointProtocol name="MSDCOM_RPC"><runtimeParameters><param name="DLLFilePath" value="%U8SOFT%\EF\EFWXZG\EF_Refer.dll" description="" display="true" /><param name="ProgID" value="EF_Refer.clsVoucherPlugin" description="" display="true" /><param name="Server" value="." description="" display="false" /><param name="ClassName" value="clsVoucherPlugin" description="" display="true" /><param name="MethodName" value="MaterialOut_Save_After" description="" display="true" /><param name="ComPlusTransaction" value="False" description="" display="true" /><param name="IsPlugin" value="true" description="" display="true" /></runtimeParameters></momEndPointProtocol>' where ID=@EndpointID 
end 
select @MsgTypeCategoryID=ID from MOM_MsgTypeCategories where MsgTypeCategory='����¼�' 
select @MsgFilterID=ID from IB_MsgFilter where FilterName='XPathFilter' 
select @MsgTypeID=ID from IB_MsgType where MsgType='Save_After' and AppTypeID=@AppTypeID 
if @MsgTypeID is null 
begin 
insert into IB_MsgType(MsgType,FriendName,MsgSchema,AppTypeID,MsgTypeCategoryID,MsgFilterID,ExtendProperties) values('Save_After','������¼�','<?xml version="1.0" encoding="utf-8" ?><serviceInterface transactionType=""><operation name="" ><parameters><parameter index="0" name="domhead" type="MSXML2.IXMLDOMDocument2" direction="inout" desc="��ͷ" optional="false" byRef="true" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /><parameter index="1" name="dombody" type="MSXML2.IXMLDOMDocument2" direction="inout" desc="����" optional="false" byRef="true" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /><parameter index="2" name="errmsg" type="string" direction="inout" desc="���ش�����Ϣ" optional="false" byRef="false" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /><parameter type="bool" direction="retval" desc="����ֵ: true:�ɹ�, false: ʧ��" byRef="false" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /></parameters></operation></serviceInterface>',@AppTypeID,@MsgTypeCategoryID,@MsgFilterID,'') 
 set @MsgTypeID=@@identity 
end 
else 
begin 
update IB_MsgType set MsgSchema='<?xml version="1.0" encoding="utf-8" ?><serviceInterface transactionType=""><operation name="" ><parameters><parameter index="0" name="domhead" type="MSXML2.IXMLDOMDocument2" direction="inout" desc="��ͷ" optional="false" byRef="true" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /><parameter index="1" name="dombody" type="MSXML2.IXMLDOMDocument2" direction="inout" desc="����" optional="false" byRef="true" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /><parameter index="2" name="errmsg" type="string" direction="inout" desc="���ش�����Ϣ" optional="false" byRef="false" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /><parameter type="bool" direction="retval" desc="����ֵ: true:�ɹ�, false: ʧ��" byRef="false" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /></parameters></operation></serviceInterface>',ExtendProperties='' where ID=@MsgTypeID 
end 
if (select count(*) from IB_Event_Plugin where AppTagID=@AppTagID and MsgTypeID=@MsgTypeID) = 0 
begin 
insert into IB_Event_Plugin(AppTagID,MsgTypeID,AccID,OrderNO,IsSyncOrAsync,Description,Disabled,UnVisible,UnDeleted) values(@AppTagID,@MsgTypeID,'',1,0,'',0,0,0)  
end 
go
declare @AppSysID int,@AppTypeID int,@AppTagID int,@EndpointID nvarchar(50),@MsgTypeID int,@MsgTypeCategoryID int,@MsgFilterID int,@ParentAppTypeID int 
select @AppSysID=IB_AppSys.ID from IB_AppSys,IB_Entities where IB_Entities.ID=IB_AppSys.EntityID and IB_Entities.EntityTag='U8API' 
if @AppSysID is null 
begin 
insert into IB_Entities(EntityTag,FriendName)values('U8API','U8API') 
insert into IB_AppSys(AppSystem,FriendName,EntityID)values('U8API','U8API',@@identity) 
set @AppSysID=@@identity 
end 
set @ParentAppTypeID=0 
set @AppTypeID=0 
select @AppTypeID=ID from IB_AppType where AppType='ST' and AppSysID=@AppSysID 
if @AppTypeID is null or @AppTypeID=0 
begin 
insert into IB_AppType(AppType,FriendName,AppSysID)values('ST','������',@AppSysID) 
select @AppTypeID=@@identity 
if @ParentAppTypeID>0 
begin 
update IB_AppType set ParentID=@ParentAppTypeID where ID=@AppTypeID 
end 
end 
set @ParentAppTypeID=@AppTypeID 
set @AppTypeID=0 
select @AppTypeID=ID from IB_AppType where AppType='MaterialOut' and AppSysID=@AppSysID 
if @AppTypeID is null or @AppTypeID=0 
begin 
insert into IB_AppType(AppType,FriendName,AppSysID)values('MaterialOut','���ϳ��ⵥ',@AppSysID) 
select @AppTypeID=@@identity 
if @ParentAppTypeID>0 
begin 
update IB_AppType set ParentID=@ParentAppTypeID where ID=@AppTypeID 
end 
end 
set @ParentAppTypeID=@AppTypeID 
select @AppTagID=ID from IB_AppTag where AppTag='MaterialOutDelete_After' and AppTypeID=@AppTypeID 
if @AppTagID is null 
begin 
insert into IB_AppTag(AppTag,FriendName,AppTypeID,ExtendProperties,Description,Customize,IsPlugin) 
 values('MaterialOutDelete_After','MaterialOutDelete_After',@AppTypeID,'','',0,1) 
set @AppTagID=@@identity 
end 
else 
begin 
update IB_AppTag set ExtendProperties='',Description='',Customize=0 where ID=@AppTagID 
end 
select top 1 @EndpointID=ID from IB_EndPoint where AppTagID=@AppTagID 
if @EndpointID is null 
begin 
insert into IB_EndPoint(Address,ProtocolID,ProtocolParams,AppTagID) 
 values('','MSDCOM_RPC','<?xml version="1.0" encoding="utf-8"?><momEndPointProtocol name="MSDCOM_RPC"><runtimeParameters><param name="DLLFilePath" value="%U8SOFT%\EF\EFWXZG\EF_Refer.dll" description="" display="true" /><param name="ProgID" value="EF_Refer.clsVoucherPlugin" description="" display="true" /><param name="Server" value="." description="" display="false" /><param name="ClassName" value="clsVoucherPlugin" description="" display="true" /><param name="MethodName" value="MaterialOut_Delete_After" description="" display="true" /><param name="ComPlusTransaction" value="False" description="" display="true" /><param name="IsPlugin" value="true" description="" display="true" /></runtimeParameters></momEndPointProtocol>',@AppTagID) 
end 
else 
begin 
update IB_EndPoint set Address='',ProtocolID='MSDCOM_RPC',ProtocolParams='<?xml version="1.0" encoding="utf-8"?><momEndPointProtocol name="MSDCOM_RPC"><runtimeParameters><param name="DLLFilePath" value="%U8SOFT%\EF\EFWXZG\EF_Refer.dll" description="" display="true" /><param name="ProgID" value="EF_Refer.clsVoucherPlugin" description="" display="true" /><param name="Server" value="." description="" display="false" /><param name="ClassName" value="clsVoucherPlugin" description="" display="true" /><param name="MethodName" value="MaterialOut_Delete_After" description="" display="true" /><param name="ComPlusTransaction" value="False" description="" display="true" /><param name="IsPlugin" value="true" description="" display="true" /></runtimeParameters></momEndPointProtocol>' where ID=@EndpointID 
end 
select @MsgTypeCategoryID=ID from MOM_MsgTypeCategories where MsgTypeCategory='����¼�' 
select @MsgFilterID=ID from IB_MsgFilter where FilterName='XPathFilter' 
select @MsgTypeID=ID from IB_MsgType where MsgType='Delete_After' and AppTypeID=@AppTypeID 
if @MsgTypeID is null 
begin 
insert into IB_MsgType(MsgType,FriendName,MsgSchema,AppTypeID,MsgTypeCategoryID,MsgFilterID,ExtendProperties) values('Delete_After','ɾ�����¼�','<?xml version="1.0" encoding="utf-8" ?><serviceInterface transactionType=""><operation name="" ><parameters><parameter index="0" name="domhead" type="MSXML2.IXMLDOMDocument2" direction="inout" desc="��ͷ" optional="false" byRef="true" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /><parameter index="1" name="dombody" type="MSXML2.IXMLDOMDocument2" direction="inout" desc="����" optional="false" byRef="true" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /><parameter index="2" name="errmsg" type="string" direction="inout" desc="���ش�����Ϣ" optional="false" byRef="false" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /><parameter type="bool" direction="retval" desc="����ֵ: true:�ɹ�, false: ʧ��" byRef="false" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /></parameters></operation></serviceInterface>',@AppTypeID,@MsgTypeCategoryID,@MsgFilterID,'') 
 set @MsgTypeID=@@identity 
end 
else 
begin 
update IB_MsgType set MsgSchema='<?xml version="1.0" encoding="utf-8" ?><serviceInterface transactionType=""><operation name="" ><parameters><parameter index="0" name="domhead" type="MSXML2.IXMLDOMDocument2" direction="inout" desc="��ͷ" optional="false" byRef="true" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /><parameter index="1" name="dombody" type="MSXML2.IXMLDOMDocument2" direction="inout" desc="����" optional="false" byRef="true" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /><parameter index="2" name="errmsg" type="string" direction="inout" desc="���ش�����Ϣ" optional="false" byRef="false" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /><parameter type="bool" direction="retval" desc="����ֵ: true:�ɹ�, false: ʧ��" byRef="false" uapMetaType="" uapMetaID="" uapMetaName="" isBoHead="false" isBoBody="false" /></parameters></operation></serviceInterface>',ExtendProperties='' where ID=@MsgTypeID 
end 
if (select count(*) from IB_Event_Plugin where AppTagID=@AppTagID and MsgTypeID=@MsgTypeID) = 0 
begin 
insert into IB_Event_Plugin(AppTagID,MsgTypeID,AccID,OrderNO,IsSyncOrAsync,Description,Disabled,UnVisible,UnDeleted) values(@AppTagID,@MsgTypeID,'',1,0,'',0,0,0)  
end 
go
