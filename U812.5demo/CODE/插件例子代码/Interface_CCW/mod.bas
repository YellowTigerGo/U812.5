Attribute VB_Name = "mod"
'函数功能 ：得到DOM对象中指定元素的值
'domBody  dom 对象
'sKey   关键字名称
'R      行号

Public gstrVoucherType '全局变量，记录当前操作的单据cardnum 号
Public gstrKeyName '全局变量，记录当前操作的按钮名称
Public Function GetBodyItemValue(ByVal domBody As DOMDocument, ByVal sKey As String, ByVal R As Long) As String
    sKey = LCase(sKey)
    If Not domBody.selectNodes("//z:row").Item(R).Attributes.getNamedItem(sKey) Is Nothing Then
        GetBodyItemValue = domBody.selectNodes("//z:row").Item(R).Attributes.getNamedItem(sKey).nodeValue
    Else
        GetBodyItemValue = ""
    End If
End Function


Public Function GetHeadItemValue(ByVal domHead As DOMDocument, ByVal sKey As String) As String
    sKey = LCase(sKey)
    If Not domHead.selectSingleNode("//z:row").Attributes.getNamedItem(sKey) Is Nothing Then
        GetHeadItemValue = domHead.selectSingleNode("//z:row").Attributes.getNamedItem(sKey).nodeValue
    Else
        GetHeadItemValue = ""
    End If
End Function
