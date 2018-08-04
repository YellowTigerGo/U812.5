IF not EXISTS ( Select cBizObjectId From AuditBizObjects Where  cBizObjectId=N'EP0201'  ) Insert Into AuditBizObjects (cCategoryId,cBizObjectId,cBizObjectName,cBizObjectDesc,cClassName ) Values (N'PU',N'EP0201',N'项目需求计划',N'项目需求计划',N'采购管理')
IF not EXISTS ( Select cBizObjectId From AuditBizObjects_Lang Where  cBizObjectId=N'EP0201'  And Lang_ID=N'zh-CN'  ) Insert Into AuditBizObjects_Lang (cBizObjectId,cBizObjectName,cBizObjectDesc,Lang_ID ) Values (N'EP0201',N'项目需求计划',N'项目需求计划',N'zh-CN')
IF not EXISTS ( Select cBizEventId From AuditBizEvents Where  cBizEventId=N'EP0201.Audit'  ) Insert Into AuditBizEvents (cBizObjectId,cBizEventId,cBizEventName,cBizEventDesc,cBizEventURL,bPluginEnabled,iTaskType ) Values (N'EP0201',N'EP0201.Audit',N'审批',N'审批',N'<?xml version="1.0" encoding="utf-8"?>
<Command>
  <param Name="id" Value="EP0201"/>
  <param Name="name" Value="项目需求计划单"/>
  <param Name="authID" Value="EP020106"/> 
  <param Name="cmdLine" Value=""/>
  <param Name="docType" Value=""/> 
  <param Name="docID" Value=""/> 
  <param Name="subFunction" Value=" "/> 
  <param Name="SubSysID" Value="EP"/> 
</Command>',1,1)
IF not EXISTS ( Select cBizEventId From AuditBizEvents Where  cBizEventId=N'EP0201.Return'  ) Insert Into AuditBizEvents (cBizObjectId,cBizEventId,cBizEventName,cBizEventDesc,cBizEventURL,bPluginEnabled,iTaskType ) Values (N'EP0201',N'EP0201.Return',N'打回',N'打回',N'<?xml version="1.0" encoding="utf-8"?>
<Command>
  <param Name="id" Value="EP0201"/>
  <param Name="name" Value="项目需求计划单"/>
  <param Name="authID" Value="EP020107"/> 
  <param Name="cmdLine" Value=""/>
  <param Name="docType" Value=""/> 
  <param Name="docID" Value=""/> 
  <param Name="subFunction" Value=" "/> 
  <param Name="SubSysID" Value="EP"/> 
</Command>
',0,2)
IF not EXISTS ( Select cBizEventId From AuditBizEvents Where  cBizEventId=N'EP0201.Submit'  ) Insert Into AuditBizEvents (cBizObjectId,cBizEventId,cBizEventName,cBizEventDesc,cBizEventURL,bPluginEnabled,iTaskType ) Values (N'EP0201',N'EP0201.Submit',N'提交',N'提交',N'<?xml version="1.0" encoding="utf-8"?>
<Command>
  <param Name="id" Value="EP0201"/>
  <param Name="name" Value="项目需求计划单"/>
  <param Name="authID" Value="EP020106"/> 
  <param Name="cmdLine" Value=""/>
  <param Name="docType" Value=""/> 
  <param Name="docID" Value=""/> 
  <param Name="subFunction" Value=" "/> 
  <param Name="SubSysID" Value="EP"/> 
</Command>',0,0)
IF not EXISTS ( Select cMsgTypeId From AuditMsgTypeConfig Where  cMsgTypeId=N'EP0201|EP0201.Audit'  ) Insert Into AuditMsgTypeConfig (cMsgTypeId,cMsgTypeName,cMsgTypeDesc,cBizObjectId,cCompBizEntityId,MOMRegPath,MsgSchema,cKeyFidldPath,cBizEventId ) Values (N'EP0201|EP0201.Audit',N'项目需求计划的审批消息',N'',N'EP0201',N'EP0201.Audit',N'',N'',N'Audit/BusinessObjects/V_List_EF_ProjectMRP/VoucherType;Audit/BusinessObjects/V_List_EF_ProjectMRP/VoucherId;Audit/BusinessObjects/V_List_EF_ProjectMRP/VoucherCode',N'EP0201.Audit')
IF not EXISTS ( Select cBizEventId From AuditBizEvents_Lang Where  cBizEventId=N'EP0201.Audit'  And Lang_ID=N'zh-CN'  ) Insert Into AuditBizEvents_Lang ( cBizEventId ,cBizEventName,cBizEventDesc,Lang_ID   ) Values (N'EP0201.Audit',N'审批',N'审批',N'zh-CN')
IF not EXISTS ( Select cBizEntityId From AuditBizEntities Where  cBizEntityId=N'EP0201.Audit'  ) Insert Into AuditBizEntities (cBizEntityId,cBizEventId,cBizDataViewId,cBizDataViewName,cBizDataViewDesc,cBizFunctionName,cBizFunctionDesc,cBizQueryName,cBizQueryDesc,cResultTableName ) Values (N'EP0201.Audit',N'EP0201.Audit',N'95370b66-2e41-4780-8c93-cba7ffa770da',N'V_List_EF_ProjectMRPBO',N'项目需求计划单',N'V_List_EF_ProjectMRP',N'项目需求计划单查询',N'V_List_EF_ProjectMRP',N'项目需求计划单',N'V_List_EF_ProjectMRP')
IF not EXISTS ( Select cBizEventId From AuditEventPlugins Where  cBizEventId=N'EP0201.Audit'  And cPluginId=N'EFWorkFlowSrv.clsSAWorkFlowSrv'  ) Insert Into AuditEventPlugins (  cBizEventId,cPluginId,iExecuteOrder,bFinalAuditPlugin ) Values (N'EP0201.Audit',N'EFWorkFlowSrv.clsSAWorkFlowSrv',0,1)
IF not EXISTS ( Select cPluginId From AuditPlugins Where  cPluginId=N'EFWorkFlowSrv.clsSAWorkFlowSrv'  ) Insert Into AuditPlugins ( cPluginId,cPluginName,cPluginDesc,cAssembly,cTypeName,iPluginType ) Values (N'EFWorkFlowSrv.clsSAWorkFlowSrv',N'自定义插件',N'自定义插件',N'efworkflowsrv.dll',N'efworkflowsrv.clsSAWorkFlowSrv',0)
IF not EXISTS ( Select cBizEventId,cParamName From AuditBizPriParams Where   cBizEventId=N'EP0201.Audit'  And cParamName=N'VoucherCode'  ) Insert Into AuditBizPriParams ( cBizEventId,cParamName, cParamDesc ) Values (N'EP0201.Audit',N'VoucherCode',N'单据编号')
IF not EXISTS ( Select cBizEventId,cParamName From AuditBizPriParams Where   cBizEventId=N'EP0201.Audit'  And cParamName=N'VoucherId'  ) Insert Into AuditBizPriParams ( cBizEventId,cParamName, cParamDesc ) Values (N'EP0201.Audit',N'VoucherId',N'单据号')
IF not EXISTS ( Select cBizEventId,cParamName From AuditBizPriParams Where   cBizEventId=N'EP0201.Audit'  And cParamName=N'VoucherType'  ) Insert Into AuditBizPriParams ( cBizEventId,cParamName, cParamDesc ) Values (N'EP0201.Audit',N'VoucherType',N'单据类型')
IF not EXISTS ( Select cBizEntityId From AuditBizEntities Where  cBizEntityId=N'EP0201.Return'  ) Insert Into AuditBizEntities (cBizEntityId,cBizEventId,cBizDataViewId,cBizDataViewName,cBizDataViewDesc,cBizFunctionName,cBizFunctionDesc,cBizQueryName,cBizQueryDesc,cResultTableName ) Values (N'EP0201.Return',N'EP0201.Return',N'95370b66-2e41-4780-8c93-cba7ffa770da',N'V_List_EF_ProjectMRPBO',N'项目需求计划单',N'V_List_EF_ProjectMRP',N'项目需求计划单查询',N'V_List_EF_ProjectMRP',N'项目需求计划单',N'V_List_EF_ProjectMRP')
IF not EXISTS ( Select cBizEventId From AuditEventPlugins Where  cBizEventId=N'EP0201.Return'  And cPluginId=N'EFWorkFlowSrv.clsSAWorkFlowSrv'  ) Insert Into AuditEventPlugins (  cBizEventId,cPluginId,iExecuteOrder,bFinalAuditPlugin ) Values (N'EP0201.Return',N'EFWorkFlowSrv.clsSAWorkFlowSrv',0,0)
IF not EXISTS ( Select cPluginId From AuditPlugins Where  cPluginId=N'EFWorkFlowSrv.clsSAWorkFlowSrv'  ) Insert Into AuditPlugins ( cPluginId,cPluginName,cPluginDesc,cAssembly,cTypeName,iPluginType ) Values (N'EFWorkFlowSrv.clsSAWorkFlowSrv',N'自定义插件',N'自定义插件',N'efworkflowsrv.dll',N'efworkflowsrv.clsSAWorkFlowSrv',0)
IF not EXISTS ( Select cBizEventId,cParamName From AuditBizPriParams Where   cBizEventId=N'EP0201.Return'  And cParamName=N'VoucherCode'  ) Insert Into AuditBizPriParams ( cBizEventId,cParamName, cParamDesc ) Values (N'EP0201.Return',N'VoucherCode',N'单据编号')
IF not EXISTS ( Select cBizEventId,cParamName From AuditBizPriParams Where   cBizEventId=N'EP0201.Return'  And cParamName=N'VoucherId'  ) Insert Into AuditBizPriParams ( cBizEventId,cParamName, cParamDesc ) Values (N'EP0201.Return',N'VoucherId',N'单据号')
IF not EXISTS ( Select cBizEventId,cParamName From AuditBizPriParams Where   cBizEventId=N'EP0201.Return'  And cParamName=N'VoucherType'  ) Insert Into AuditBizPriParams ( cBizEventId,cParamName, cParamDesc ) Values (N'EP0201.Return',N'VoucherType',N'单据类型')
IF not EXISTS ( Select cBizEventId,cParamName From AuditBizPriParams_Lang Where   cBizEventId=N'EP0201.Return'  And cParamName=N'VoucherId'  And Lang_ID=N'zh-CN'  ) Insert Into AuditBizPriParams_Lang ( cBizEventId,cParamName, cParamDesc,Lang_ID ) Values (N'EP0201.Return',N'VoucherId',N'单据号',N'zh-CN')
IF not EXISTS ( Select cBizEventId,cParamName From AuditBizPriParams_Lang Where   cBizEventId=N'EP0201.Return'  And cParamName=N'VoucherType'  And Lang_ID=N'zh-CN'  ) Insert Into AuditBizPriParams_Lang ( cBizEventId,cParamName, cParamDesc,Lang_ID ) Values (N'EP0201.Return',N'VoucherType',N'单据类型',N'zh-CN')
IF not EXISTS ( Select cBizEventId,cParamName From AuditBizPriParams_Lang Where   cBizEventId=N'EP0201.Return'  And cParamName=N'VoucherCode'  And Lang_ID=N'zh-CN'  ) Insert Into AuditBizPriParams_Lang ( cBizEventId,cParamName, cParamDesc,Lang_ID ) Values (N'EP0201.Return',N'VoucherCode',N'单据编号',N'zh-CN')
IF not EXISTS ( Select cMsgTypeId From AuditMsgTypeConfig Where  cMsgTypeId=N'EP0201|EP0201.Submit'  ) Insert Into AuditMsgTypeConfig (cMsgTypeId,cMsgTypeName,cMsgTypeDesc,cBizObjectId,cCompBizEntityId,MOMRegPath,MsgSchema,cKeyFidldPath,cBizEventId ) Values (N'EP0201|EP0201.Submit',N'项目需求计划单的提交消息',N'',N'EP0201',N'EP0201.Submit',N'',N'',N'Audit/BusinessObjects/V_List_EF_ProjectMRP/VoucherType;Audit/BusinessObjects/V_List_EF_ProjectMRP/VoucherId;Audit/BusinessObjects/V_List_EF_ProjectMRP/VoucherCode',N'EP0201.Submit')
IF not EXISTS ( Select cBizEventId From AuditBizEvents_Lang Where  cBizEventId=N'EP0201.Submit'  And Lang_ID=N'zh-CN'  ) Insert Into AuditBizEvents_Lang ( cBizEventId ,cBizEventName,cBizEventDesc,Lang_ID   ) Values (N'EP0201.Submit',N'提交',N'提交',N'zh-CN')
IF not EXISTS ( Select cBizEntityId From AuditBizEntities Where  cBizEntityId=N'EP0201.Submit'  ) Insert Into AuditBizEntities (cBizEntityId,cBizEventId,cBizDataViewId,cBizDataViewName,cBizDataViewDesc,cBizFunctionName,cBizFunctionDesc,cBizQueryName,cBizQueryDesc,cResultTableName ) Values (N'EP0201.Submit',N'EP0201.Submit',N'95370b66-2e41-4780-8c93-cba7ffa770da',N'V_List_EF_ProjectMRPBO',N'项目需求计划单',N'V_List_EF_ProjectMRP',N'项目需求计划单查询',N'V_List_EF_ProjectMRP',N'项目需求计划单',N'V_List_EF_ProjectMRP')
IF not EXISTS ( Select cCategoryId From cClassName_Lang Where  cCategoryId=N'PU'  ) Insert Into cClassName_Lang (cClassName,Lang_ID,cClassNameDesc,cCategoryId ) Values (N'采购管理',N'zh-CN',N'采购管理',N'PU')
