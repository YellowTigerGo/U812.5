Attribute VB_Name = "mod"
'�������� ���õ�DOM������ָ��Ԫ�ص�ֵ
'domBody  dom ����
'sKey   �ؼ�������
'R      �к�

Public gstrVoucherType 'ȫ�ֱ�������¼��ǰ�����ĵ���cardnum ��
Public gstrKeyName 'ȫ�ֱ�������¼��ǰ�����İ�ť����
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
