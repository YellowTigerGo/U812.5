Attribute VB_Name = "PubVariable"
'�������������ֶ�
'Public moid As String                       'MoId �������ͷID
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

'���������Ӽ������ӱ���������ֶ�
Public moa_AllocateId As Integer                    'AllocateId ������������Ӽ���������ID
Public moa_MoDId As Integer                         'MoDId ����������ϸID
Public moa_SortSeq As Integer                       'SortSeq���
Public moa_OpSeq As String                          '�����к�
Public moa_ComponentId As Integer                   'ComponentId �Ӽ�����ID
Public moa_FVFlag As Integer                        'FVFlag �̶�/�䶯����(0/1)
Public moa_BaseQtyN As Double                       'BaseQtyN ��������������
Public moa_BaseQtyD As Double                       'BaseQtyD ������������ĸ
Public moa_ParentScrap As Double                    'ParentScrap ĸ�������
Public moa_CompScrap As Double                      'CompScrap �Ӽ������
Public moa_Qty As Double                            '����
Public moa_IssQty As Double                         'IssQty������
Public moa_DeclaredQty As Double                    '��������
Public moa_StartDemDate As String                   'StartDemDate ��ʼ��������
Public moa_EndDemDate As String                     'EndDemDate ����ʼ��������
Public moa_WhCode As String                         '�ֿ����
Public moa_LotNo As String                          '����
Public moa_WIPType As Integer                       'WIPType WIP����(1���/2����/3����/5����BOM)
Public moa_ByproductFlag As Integer                  'ByproductFlag �Ƿ�������Ʒ boolean��������
Public moa_ProductType As Integer                   '��������(1:��/2:����Ʒ/3:����Ʒ)
Public moa_QcFlag As Integer                        'QcFlag ����� boolean��������
Public moa_Offset As String                         'Offset ƫ����
Public moa_InvCode As String                        '�������
Public moa_Free1 As String                          '������1
Public moa_Free2 As String                          '������2
Public moa_Free3 As String                          '������3
Public moa_Free4 As String                          '������4
Public moa_Free5 As String                          '������5
Public moa_Free6 As String                          '������6
Public moa_Free7 As String                          '������7
Public moa_Free8 As String                          '������8
Public moa_Free9 As String                          '������9
Public moa_Free10 As String                         '������10
Public moa_Define22 As String                       '�Զ�����1
Public moa_Define23 As String                       '�Զ�����2
Public moa_Define24 As String                       '�Զ�����3
Public moa_Define25 As String                       '�Զ�����4
Public moa_Define26 As Double                       '�Զ�����5
Public moa_Define27 As Double                       '�Զ�����6
Public moa_Define28 As String                       '�Զ�����7
Public moa_Define29 As String                       '�Զ�����8
Public moa_Define30 As String                       '�Զ�����9
Public moa_Define31 As String                       '�Զ�����10
Public moa_Define32 As String                       '�Զ�����11
Public moa_Define33 As String                       '�Զ�����12
Public moa_Define34 As Integer                      '�Զ�����13
Public moa_Define35 As Integer                      '�Զ�����14
Public moa_Define36 As String                       '�Զ�����15 ����������
Public moa_Define37 As String                       '�Զ�����16 ����������
Public moa_OpComponentId As Integer                 'OpComponentId BOM�Ӽ�����ID
Public moa_AuxUnitCode  As String                   'AuxUnitCode ����������λ
Public moa_ChangeRate As Double                     'ChangeRate������
Public moa_AuxBaseQtyN As Double                    'AuxBaseQtyN ������������
Public moa_AuxQty As Double                         'AuxQty Ӧ�츨����
Public moa_ReplenishQty As Double                   'ReplenishQty ������
Public moa_Remark As String                         '��ע
Public moa_TransQty As Double                       '�ѵ�����
Public moa_SoType As Integer                        'SoType������ٷ�ʽ 0����Դ 1���۶����� 3���ڶ����� 4������� 5 ���۶��� 6���ڶ���
Public moa_SoCode As String                         '������ٺ�
Public moa_SoSeq As Integer                         '��������к�
Public moa_SoDId As String                          '�������DId
Public moa_DemandCode As String                     '������൥��

''BOMչ��
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


