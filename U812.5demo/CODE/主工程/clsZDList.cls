VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsZDList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' 软件著作权: 用友软件集团
' 系统名称：应收、应付系统
' 功能说明: 单据集合
' 作者：翟万达
'------------------------------------------------------------------------------------
Option Explicit

'局部变量，保存集合
Private mCol As Dictionary

Public Function Add(objNewMember As clsZD) As clsZD
    mCol.Add mCol.Count + 1, objNewMember
    Set Add = objNewMember
End Function

Public Property Get Item(Index As Variant) As clsZD
    '引用集合中的一个元素时使用。
    'Index 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
    Set Item = mCol(Index)
End Property

Public Property Get Count() As Long
    '检索集合中的元素数时使用。语法：Debug.Print x.Count
    Count = mCol.Count
End Property

Public Sub Remove(Index As Variant)
    '删除集合中的元素时使用。
    'Index 包含索引或关键字，这是为什么要声明为 Variant 的原因
    '语法：x.Remove(xyz)
    mCol.Remove Index
End Sub

'Public Property Get NewEnum() As IUnknown
'    '本属性允许用 For...Each 语法枚举该集合。
'    Set NewEnum = mCol.[_NewEnum]
'End Property

Private Sub Class_Initialize()
    '创建类后创建集合
    Set mCol = New Dictionary
End Sub

Private Sub Class_Terminate()
    '类终止后破坏集合
    Set mCol = Nothing
End Sub







