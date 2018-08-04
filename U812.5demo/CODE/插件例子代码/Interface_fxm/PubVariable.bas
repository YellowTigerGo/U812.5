Attribute VB_Name = "PubVariable"
'生产订单主表字段
'Public moid As String                       'MoId 生产令单表头ID
'Public MoCode As String
'Public CreateDate As String
'Public CreateTime As String
'Public CreateUser As String
'Public ModifyDate As String
'Public ModifyTime As String
'Public ModifyUser As String
'Public UpdCount As String
'Public Define1 As String
'Public Define2 As String
'Public Define3 As String
'Public Define4 As String
'Public Define5 As String
'Public Define6 As String
'Public Define7 As String

Public isTrans As Boolean

'生产订单子件资料子表插入数据字段
Public moa_AllocateId As Integer                    'AllocateId 生产令单的所有子件用料资料ID
Public moa_MoDId As Integer                         'MoDId 生产订单明细ID
Public moa_SortSeq As Integer                       'SortSeq序号
Public moa_OpSeq As String                          '工序行号
Public moa_ComponentId As Integer                   'ComponentId 子件物料ID
Public moa_FVFlag As Integer                        'FVFlag 固定/变动批量(0/1)
Public moa_BaseQtyN As Double                       'BaseQtyN 基本用量－分子
Public moa_BaseQtyD As Double                       'BaseQtyD 基本用量－分母
Public moa_ParentScrap As Double                    'ParentScrap 母件损耗率
Public moa_CompScrap As Double                      'CompScrap 子件损耗率
Public moa_Qty As Double                            '数量
Public moa_IssQty As Double                         'IssQty已领量
Public moa_DeclaredQty As Double                    '报检数量
Public moa_StartDemDate As String                   'StartDemDate 开始需求日期
Public moa_EndDemDate As String                     'EndDemDate 结束始需求日期
Public moa_WhCode As String                         '仓库代码
Public moa_LotNo As String                          '批号
Public moa_WIPType As Integer                       'WIPType WIP属性(1入库/2工序/3领料/5基于BOM)
Public moa_ByproductFlag As Integer                  'ByproductFlag 是否联副产品 boolean类型数据
Public moa_ProductType As Integer                   '产出类型(1:空/2:联产品/3:副产品)
Public moa_QcFlag As Integer                        'QcFlag 检验否 boolean类型数据
Public moa_Offset As String                         'Offset 偏置期
Public moa_InvCode As String                        '存货编码
Public moa_Free1 As String                          '自由项1
Public moa_Free2 As String                          '自由项2
Public moa_Free3 As String                          '自由项3
Public moa_Free4 As String                          '自由项4
Public moa_Free5 As String                          '自由项5
Public moa_Free6 As String                          '自由项6
Public moa_Free7 As String                          '自由项7
Public moa_Free8 As String                          '自由项8
Public moa_Free9 As String                          '自由项9
Public moa_Free10 As String                         '自由项10
Public moa_Define22 As String                       '自定义项1
Public moa_Define23 As String                       '自定义项2
Public moa_Define24 As String                       '自定义项3
Public moa_Define25 As String                       '自定义项4
Public moa_Define26 As Double                       '自定义项5
Public moa_Define27 As Double                       '自定义项6
Public moa_Define28 As String                       '自定义项7
Public moa_Define29 As String                       '自定义项8
Public moa_Define30 As String                       '自定义项9
Public moa_Define31 As String                       '自定义项10
Public moa_Define32 As String                       '自定义项11
Public moa_Define33 As String                       '自定义项12
Public moa_Define34 As Integer                      '自定义项13
Public moa_Define35 As Integer                      '自定义项14
Public moa_Define36 As String                       '自定义项15 日期型数据
Public moa_Define37 As String                       '自定义项16 日期型数据
Public moa_OpComponentId As Integer                 'OpComponentId BOM子件资料ID
Public moa_AuxUnitCode  As String                   'AuxUnitCode 辅助计量单位
Public moa_ChangeRate As Double                     'ChangeRate换算率
Public moa_AuxBaseQtyN As Double                    'AuxBaseQtyN 辅助基本用量
Public moa_AuxQty As Double                         'AuxQty 应领辅助量
Public moa_ReplenishQty As Double                   'ReplenishQty 补料量
Public moa_Remark As String                         '备注
Public moa_TransQty As Double                       '已调拨量
Public moa_SoType As Integer                        'SoType需求跟踪方式 0无来源 1销售订单行 3出口订单行 4需求分类 5 销售订单 6出口订单
Public moa_SoCode As String                         '需求跟踪号
Public moa_SoSeq As Integer                         '需求跟踪行号
Public moa_SoDId As String                          '需求跟踪DId
Public moa_DemandCode As String                     '需求分类单号

''BOM展开
'Public bom_OpComponentId    As String
'Public bom_OpSeq            As String
'Public bom_CompId           As String
'Public bom_UnitId           As String
'Public bom_BaseQtyN         As String
'Public bom_BaseQtyD         As String
'Public bom_ParentScrap      As String
'Public bom_CompScrap        As String
'Public bom_FVQty            As Integer
'Public bom_Cqty             As Double
'Public bom_Cqty1            As Double
'Public bom_UseQty           As Double
'Public bom_Offset           As Double
'Public bom_WIPtype          As Integer
'Public bom_WhCode           As String
'Public bom_InvCode          As String
'Public bom_Free1            As String
'Public bom_Free2            As String
'Public bom_Free3            As String
'Public bom_Free4            As String
'Public bom_Free5            As String
'Public bom_Free6            As String
'Public bom_Free7            As String
'Public bom_Free8            As String
'Public bom_Free9            As String
'Public bom_Free10           As String
'Public bom_Dept             As String
'Public bom_DepName          As String
'Public bom_ByproductFlag    As Integer
'Public bom_AccuCostFlag     As Integer
'Public bom_SubFlag          As Integer
'Public bom_BomType          As Integer
'Public bom_iGrade           As Integer
'Public bom_DemDate          As String
'Public bom_AuxUnitCode      As String
'Public bom_ChangeRate       As Double
'Public bom_AuxBaseQtyN      As Double
'Public bom_AuxCqty          As String
'Public bom_AuxUseQty        As Double
'Public bom_AuxUnitName      As String
'Public bom_Define1      As String
'Public bom_Define2      As String
'Public bom_Define3      As String
'Public bom_Define4      As String
'Public bom_Define5      As String
'Public bom_Define6      As String
'Public bom_Define7      As String
'Public bom_Define8      As String
'Public bom_Define9      As String
'Public bom_Define10     As String
'Public bom_Define11     As String
'Public bom_Define12     As String
'Public bom_Define13     As String
'Public bom_Define14     As String
'Public bom_Define15     As String
'Public bom_Define16     As String
'Public bom_Define22     As String
'Public bom_Define23     As String
'Public bom_Define24     As String
'Public bom_Define25     As String
'Public bom_Define26     As String
'Public bom_Define27     As String
'Public bom_Define28     As String
'Public bom_Define29     As String
'Public bom_Define30     As String
'Public bom_Define31     As String
'Public bom_Define32     As String
'Public bom_Define33     As String
'Public bom_Define34     As String
'Public bom_Define35     As String
'Public bom_Define36     As String
'Public bom_Define37     As String
'Public bom_CCode        As String
'Public bom_cInvCode       As String
'Public bom_mInvCode       As String
'Public bom_mQty       As Double


