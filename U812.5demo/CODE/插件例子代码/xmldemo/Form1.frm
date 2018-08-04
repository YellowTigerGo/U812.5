VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   12870
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   2160
      TabIndex        =   0
      Top             =   3360
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim MOrder As IXMLDOMNode
Dim Order As IXMLDOMNode
Dim OrderDetail As IXMLDOMNode
Dim MOrderDetail As IXMLDOMNode
Dim Allocate As IXMLDOMNode
Dim dom As New DOMDocument
Dim domxml As New DOMDocument
 
Set MOrder = domxml.createElement("MOrder").cloneNode(True)
dom.appendChild MOrder  '增加根节点 MOrder

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'创建节点
Set Order = GET_space_Orderxml
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'节点赋值
SET_IXMLDOMNode_text Order, "ID", 1234
'节点追加
dom.selectSingleNode("//MOrder").appendChild Order '增加MOrder节点的下已节点 MOrder

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'创建节点
Set OrderDetail = GET_space_OrderDetailxml
'节点赋值

'节点追加
dom.selectSingleNode("//MOrder").appendChild OrderDetail '增加MOrder节点的下已节点 OrderDetail


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
  
  '创建节点
Set MOrderDetail = GET_space_MOrderDetailxml
'节点赋值

'节点追加
dom.selectSingleNode("//MOrder").appendChild MOrderDetail '增加MOrder节点的下已节点 MOrderDetail


'创建节点
Set Allocate = GET_space_Allocatexml
'节点赋值

'节点追加
dom.selectSingleNode("//MOrder").appendChild Allocate '增加MOrder节点的下已节点 Allocate


  
  
 
dom.save "c:\dom.xml"


End Sub

Public Function GET_space_Orderxml() As IXMLDOMNode
Dim Order As IXMLDOMNode
Dim domxml As New DOMDocument
Set Order = domxml.createElement("Order").cloneNode(True)
Order.appendChild domxml.createElement("ID").cloneNode(True)
Order.appendChild domxml.createElement("MoCode").cloneNode(True)
Order.appendChild domxml.createElement("CreateDate").cloneNode(True)
Order.appendChild domxml.createElement("CreateUser").cloneNode(True)
Order.appendChild domxml.createElement("Define_1").cloneNode(True)
Order.appendChild domxml.createElement("Define_2").cloneNode(True)
Order.appendChild domxml.createElement("Define_3").cloneNode(True)
Order.appendChild domxml.createElement("Define_4").cloneNode(True)
Order.appendChild domxml.createElement("Define_5").cloneNode(True)
Order.appendChild domxml.createElement("Define_6").cloneNode(True)
Order.appendChild domxml.createElement("Define_7").cloneNode(True)
Order.appendChild domxml.createElement("Define_8").cloneNode(True)
Order.appendChild domxml.createElement("Define_9").cloneNode(True)
Order.appendChild domxml.createElement("Define_10").cloneNode(True)
Order.appendChild domxml.createElement("Define_11").cloneNode(True)
Order.appendChild domxml.createElement("Define_12").cloneNode(True)
Order.appendChild domxml.createElement("Define_13").cloneNode(True)
Order.appendChild domxml.createElement("Define_14").cloneNode(True)
Order.appendChild domxml.createElement("Define_15").cloneNode(True)
Order.appendChild domxml.createElement("Define_16").cloneNode(True)
Set GET_space_Orderxml = Order
End Function



Public Function GET_space_OrderDetailxml() As IXMLDOMNode
Dim Order As IXMLDOMNode
Dim domxml As New DOMDocument
Set Order = domxml.createElement("OrderDetail").cloneNode(True)
Order.appendChild domxml.createElement("ID").cloneNode(True)
Order.appendChild domxml.createElement("MoCode").cloneNode(True)
Order.appendChild domxml.createElement("CreateDate").cloneNode(True)
Order.appendChild domxml.createElement("CreateUser").cloneNode(True)
Order.appendChild domxml.createElement("Define_1").cloneNode(True)
Order.appendChild domxml.createElement("Define_2").cloneNode(True)
Order.appendChild domxml.createElement("Define_3").cloneNode(True)
Order.appendChild domxml.createElement("Define_4").cloneNode(True)
Order.appendChild domxml.createElement("Define_5").cloneNode(True)
Order.appendChild domxml.createElement("Define_6").cloneNode(True)
Order.appendChild domxml.createElement("Define_7").cloneNode(True)
Order.appendChild domxml.createElement("Define_8").cloneNode(True)
Order.appendChild domxml.createElement("Define_9").cloneNode(True)
Order.appendChild domxml.createElement("Define_10").cloneNode(True)
Order.appendChild domxml.createElement("Define_11").cloneNode(True)
Order.appendChild domxml.createElement("Define_12").cloneNode(True)
Order.appendChild domxml.createElement("Define_13").cloneNode(True)
Order.appendChild domxml.createElement("Define_14").cloneNode(True)
Order.appendChild domxml.createElement("Define_15").cloneNode(True)
Order.appendChild domxml.createElement("Define_16").cloneNode(True)
Set GET_space_OrderDetailxml = Order
End Function


Public Function GET_space_MOrderDetailxml() As IXMLDOMNode
Dim Order As IXMLDOMNode
Dim domxml As New DOMDocument
Set Order = domxml.createElement("MOOrderDetail").cloneNode(True)
Order.appendChild domxml.createElement("ID").cloneNode(True)
Order.appendChild domxml.createElement("MoCode").cloneNode(True)
Order.appendChild domxml.createElement("CreateDate").cloneNode(True)
Order.appendChild domxml.createElement("CreateUser").cloneNode(True)
Order.appendChild domxml.createElement("Define_1").cloneNode(True)
Order.appendChild domxml.createElement("Define_2").cloneNode(True)
Order.appendChild domxml.createElement("Define_3").cloneNode(True)
Order.appendChild domxml.createElement("Define_4").cloneNode(True)
Order.appendChild domxml.createElement("Define_5").cloneNode(True)
Order.appendChild domxml.createElement("Define_6").cloneNode(True)
Order.appendChild domxml.createElement("Define_7").cloneNode(True)
Order.appendChild domxml.createElement("Define_8").cloneNode(True)
Order.appendChild domxml.createElement("Define_9").cloneNode(True)
Order.appendChild domxml.createElement("Define_10").cloneNode(True)
Order.appendChild domxml.createElement("Define_11").cloneNode(True)
Order.appendChild domxml.createElement("Define_12").cloneNode(True)
Order.appendChild domxml.createElement("Define_13").cloneNode(True)
Order.appendChild domxml.createElement("Define_14").cloneNode(True)
Order.appendChild domxml.createElement("Define_15").cloneNode(True)
Order.appendChild domxml.createElement("Define_16").cloneNode(True)
Set GET_space_MOrderDetailxml = Order
End Function


Public Function GET_space_Allocatexml() As IXMLDOMNode
Dim Order As IXMLDOMNode
Dim domxml As New DOMDocument
Set Order = domxml.createElement("Allocate").cloneNode(True)
Order.appendChild domxml.createElement("ID").cloneNode(True)
Order.appendChild domxml.createElement("MoCode").cloneNode(True)
Order.appendChild domxml.createElement("CreateDate").cloneNode(True)
Order.appendChild domxml.createElement("CreateUser").cloneNode(True)
Order.appendChild domxml.createElement("Define_1").cloneNode(True)
Order.appendChild domxml.createElement("Define_2").cloneNode(True)
Order.appendChild domxml.createElement("Define_3").cloneNode(True)
Order.appendChild domxml.createElement("Define_4").cloneNode(True)
Order.appendChild domxml.createElement("Define_5").cloneNode(True)
Order.appendChild domxml.createElement("Define_6").cloneNode(True)
Order.appendChild domxml.createElement("Define_7").cloneNode(True)
Order.appendChild domxml.createElement("Define_8").cloneNode(True)
Order.appendChild domxml.createElement("Define_9").cloneNode(True)
Order.appendChild domxml.createElement("Define_10").cloneNode(True)
Order.appendChild domxml.createElement("Define_11").cloneNode(True)
Order.appendChild domxml.createElement("Define_12").cloneNode(True)
Order.appendChild domxml.createElement("Define_13").cloneNode(True)
Order.appendChild domxml.createElement("Define_14").cloneNode(True)
Order.appendChild domxml.createElement("Define_15").cloneNode(True)
Order.appendChild domxml.createElement("Define_16").cloneNode(True)
Set GET_space_Allocatexml = Order
End Function


Public Sub SET_IXMLDOMNode_text(Nodexml As IXMLDOMNode, NodeName As String, NodeText As String)
    If Not Nodexml.selectSingleNode("//" & NodeName) Is Nothing Then
        Nodexml.selectSingleNode("//" & NodeName).Text = NodeText
    End If
End Sub
 
