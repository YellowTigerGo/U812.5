VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFLList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mCPZ As New Dictionary
'局部变量保存属性值
Private mvarCFL As clsFL '本地复本
' 增加一条分录
Public Sub AddFL(CtmpFL As clsFL)
    mCPZ.Add mCPZ.Count + 1, CtmpFL
End Sub
Public Property Set CFL(ByVal vData As Object)
'当把对象赋值给属性时在 Set 语句左边使用。
'Syntax: Set x.CFL = Form1
    Set mvarCFL = vData
End Property
Public Property Get Item(Index As Variant) As clsFL
Attribute Item.VB_UserMemId = 0
    '引用集合中的一个元素时使用。
    'Index 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCPZ(Index)
End Property
Public Property Get Count() As Long
    '检索集合中的元素数时使用。语法：Debug.Print x.Count
    Count = mCPZ.Count
End Property
Public Sub Remove(Index As Variant)
    mCPZ.Remove Index
End Sub
Public Property Get CFL() As clsFL
    Set CFL = mvarCFL
End Property
Private Sub Class_Initialize()
    '创建类后创建集合
    Set mCPZ = New Dictionary
End Sub
Private Sub Class_Terminate()
    '类终止后破坏集合
    Set mCPZ = Nothing
End Sub

