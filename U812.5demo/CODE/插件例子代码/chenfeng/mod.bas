Attribute VB_Name = "mod"
'函数功能 ：得到DOM对象中指定元素的值
'domBody  dom 对象
'sKey   关键字名称
'R      行号
Public Function GetBodyItemValue(ByVal domBody As DOMDocument, ByVal sKey As String, ByVal R As Long) As String
    sKey = LCase(sKey)
    If Not domBody.selectNodes("//z:row").Item(R).Attributes.getNamedItem(sKey) Is Nothing Then
        GetBodyItemValue = domBody.selectNodes("//z:row").Item(R).Attributes.getNamedItem(sKey).nodeValue
    Else
        GetBodyItemValue = ""
    End If
End Function
