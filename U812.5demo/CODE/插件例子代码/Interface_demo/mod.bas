Attribute VB_Name = "mod"
'函数功能 ：得到DOM对象中指定元素的值
'domBody  dom 对象
'sKey   关键字名称
'R      行号

Public DBconn As ADODB.Connection
Public g_oLogin As Object

Public Function GetElementValue(ByVal ele As IXMLDOMElement, ByVal Skey As String) As String
    If Not ele.Attributes.getNamedItem(Skey) Is Nothing Then
        GetElementValue = ele.Attributes.getNamedItem(Skey).nodeValue
    Else
        GetElementValue = ""
    End If
End Function

Public Function GetHeadItemValue(ByVal domHead As DOMDocument, ByVal Skey As String) As String
    Skey = LCase(Skey)
    If Not domHead.selectSingleNode("//z:row").Attributes.getNamedItem(Skey) Is Nothing Then
        GetHeadItemValue = domHead.selectSingleNode("//z:row").Attributes.getNamedItem(Skey).nodeValue
    Else
        GetHeadItemValue = ""
    End If
End Function

Public Function SetHeadItemValue(ByVal domHead As DOMDocument, ByVal Skey As String, ByVal value As Variant) As Boolean
    Skey = LCase(Skey)
    If Not domHead.selectSingleNode("//z:row").Attributes.getNamedItem(Skey) Is Nothing Then
        domHead.selectSingleNode("//z:row").Attributes.getNamedItem(Skey).nodeValue = value
        SetHeadItemValue = True
    Else
        SetHeadItemValue = False
    End If
End Function

Public Function GetBodyItemValue(ByVal domBody As DOMDocument, ByVal Skey As String, ByVal R As Long) As String
    Skey = LCase(Skey)
    If Not domBody.selectNodes("//z:row").Item(R).Attributes.getNamedItem(Skey) Is Nothing Then
        GetBodyItemValue = domBody.selectNodes("//z:row").Item(R).Attributes.getNamedItem(Skey).nodeValue
    Else
        GetBodyItemValue = ""
    End If
End Function

Public Function SetBodyItemValue(ByVal domBody As DOMDocument, ByVal Skey As String, ByVal R As Long, ByVal value As Variant) As Boolean
    Skey = LCase(Skey)
    If Not domBody.selectNodes("//z:row").Item(R).Attributes.getNamedItem(Skey) Is Nothing Then
        domBody.selectNodes("//z:row").Item(R).Attributes.getNamedItem(Skey).nodeValue = value
        SetBodyItemValue = True
    Else
        SetBodyItemValue = False
    End If
End Function

Public Function GetNodeValue(ByVal node As IXMLDOMNode, ByVal Skey As String) As String
'    Skey = LCase(Skey)
    If Not node.Attributes.getNamedItem(Skey) Is Nothing Then
        GetNodeValue = node.Attributes.getNamedItem(Skey).nodeValue
    Else
        GetNodeValue = ""
    End If
End Function

Public Sub SetNodeValue(ByVal node As IXMLDOMNode, ByVal Skey As String, value As String)
'    Skey = LCase(Skey)
    If Not node.Attributes.getNamedItem(Skey) Is Nothing Then
        node.Attributes.getNamedItem(Skey).nodeValue = value
    End If
End Sub

Public Sub FormatDom(SourceDom As DOMDocument, DistDom As DOMDocument, Optional editprop As String = "")
Dim element As IXMLDOMElement
Dim ele_head As IXMLDOMElement
Dim ele_body As IXMLDOMElement
Dim nd  As IXMLDOMNode
Dim tempnd As IXMLDOMNode
Dim ndheadlist As IXMLDOMNodeList
Dim ndbodylist As IXMLDOMNodeList
 
DistDom.loadXML SourceDom.xml
Dim Filedname As String
'格式部分
 Set ndheadlist = SourceDom.selectNodes("//s:Schema/s:ElementType/s:AttributeType")
 
 '数据部分
 
 
 Set ndbodylist = DistDom.selectNodes("//rs:data/z:row")
 
 For Each ele_body In ndbodylist
    For Each ele_head In ndheadlist
        Filedname = ele_head.Attributes.getNamedItem("name").nodeValue
        If ele_body.Attributes.getNamedItem(Filedname) Is Nothing Then
            '若没有当前元素，就增加当前元素
                ele_body.setAttribute Filedname, ""
 
        End If
            
            
            Select Case ele_head.lastChild.Attributes.getNamedItem("dt:type").nodeValue
            
            Case "number", "float", "boolean"
                If UCase(ele_body.Attributes.getNamedItem(Filedname).nodeValue) = UCase("false") Then
                    ele_body.setAttribute Filedname, 0
                End If
            Case Else
            
                If UCase(ele_body.Attributes.getNamedItem(Filedname).nodeValue) = UCase("否") Then
                    ele_body.setAttribute Filedname, 0
                End If
 
            End Select
       
        
        
'         Debug.Print Filedname & "=" & ele_head.selectSingleNode("//s:datatype").Attributes.getNamedItem("dt:type").nodeValue
        
        
        
        
        
    Next
    If editprop <> "" Then
        ele_body.setAttribute "editprop", editprop
    End If
Next
End Sub


Public Function str2Dbl(ByVal val As String) As Double
    On Error GoTo hErr
    If Len(val) > 0 Then
        str2Dbl = CDbl(val)
    End If
    Exit Function
hErr:
    str2Dbl = 0
End Function

'向上取整
Public Function UpLng(a As Double) As Long
    If a = 0 Then
        UpLng = 0
    Else
        If a <= CLng(a) Then
            UpLng = a
        Else
            UpLng = a + 1
        End If
    End If
    
End Function

'向下取整
Public Function DownLng(a As Double) As Long
    If a = 0 Then
        DownLng = 0
    Else
        If a >= CLng(a) Then
            DownLng = CLng(a)
        Else
            DownLng = CLng(a) - 1
        End If
    End If
    
End Function

'取系统配置信息 chenliangc
Public Function getAccinformation(strSysID As String, strName As String, conn As Object) As String
    Dim rst As New ADODB.Recordset

    rst.CursorLocation = adUseClient
    rst.Open "Select cValue from accinformation where cSysID=N'" & strSysID & "' and cName=N'" & strName & "'", conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rst.EOF Then
        getAccinformation = ""
    Else
        If IsNull(rst(0)) Then
            getAccinformation = ""
        Else
            getAccinformation = rst(0)
        End If
    End If
    rst.Close
    Set rst = Nothing
End Function

'更新插入系统信息
Public Sub UpdateAccinfo(strSysID As String, strName As String, strValue As String, conn)
    Dim affeceted As Long
    conn.Execute "Update accinformation set cValue=N'" & strValue & "' where cSysId=N'" & strSysID & "' and cname=N'" & strName & "'", affeceted
    If affeceted = 0 Then
        conn.Execute "insert into accinformation(cValue,cSysId,cname) values(N'" & strValue & "' ,N'" & strSysID & "' ,N'" & strName & "')"
    End If
End Sub

