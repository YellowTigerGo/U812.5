Attribute VB_Name = "mdlRegeidt"


Option Explicit


Sub Main()
    On Error Resume Next '容错句
    Dim wsh As Object
    Dim paths As String
    Dim dom As New DOMDocument
    Dim nd As IXMLDOMElement
    Dim nodeValue As String
    Dim fso As New FileSystemObject
    Dim MyFolder As Folder
    Dim Item As Variant
    Dim Sub_id As String
    Dim Sub_name As String
    Dim FilesName As String
    Dim regeditstr As String
    Dim u8path As String
    Dim netpath As String
    Dim NetFile_dll As String
    Dim NetFile_tlb As String
    Dim Dom_TransferProducts As New DOMDocument
    Dim ID As String
    Dim have As Boolean
    Dim clsXML As Object
    Dim configname As String
    Dim functionname As String
    Dim dll_clsname As String
    Dim blnClient As Boolean    '是否客户端，默认为不是
    Dim blnServer As Boolean    '是否服务器端，默认为不是
    Dim strtmp As String
    
    '注册.net文件
    regread "HKEY_LOCAL_MACHINE\SOFTWARE\Ufsoft\WF\V8.700\Install\CurrentInstPath\", u8path
    regread "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\.NETFramework\InstallRoot", netpath
    
    regread "HKEY_LOCAL_MACHINE\SOFTWARE\Ufsoft\WF\V8.700\Install\Installed\Client\", strtmp
    If strtmp <> "" Then
        blnClient = True
    Else
        blnClient = False
    End If
    
    regread "HKEY_LOCAL_MACHINE\SOFTWARE\Ufsoft\WF\V8.700\Install\Installed\DBServer\", strtmp
    If strtmp <> "" Then
        blnServer = True
    Else
        blnServer = False
    End If
    
    '去掉文件的只读属性
'    fso.GetFile(u8path & "\Admin\TransferProducts.xml").Attributes = 0
'
'    Call Dom_TransferProducts.Load(u8path & "\Admin\TransferProducts.xml")
'    Set clsXML = CreateObject("EF_Base_Information.clsUserConfig")  '创建标准产品的接口文件生成类
    
    '清除系统缓存
    fso.DeleteFolder u8path & "\Cache", True
'    Set wsh = CreateObject("Wscript.Shell") '利用wshshell写注册表
    Set MyFolder = fso.GetFolder(App.Path)   'u8path & "\EF")
    '清除U8缓存文件
'    fso.DeleteFolder u8path & "\cache", True
    
    paths = App.Path
    If MyFolder.Files.Count > 0 Then
        For Each Item In MyFolder.Files
            FilesName = Item.Name
            If LCase(Right(FilesName, 11)) = LCase("_Config.Xml") Then
                 
                '去掉文件的只读属性
'                fso.GetFile(u8path & "\EF\" & FilesName).Attributes = 0
                
'                Sub_id = UCase(Left(FilesName, 2))
                If dom.Load(App.Path & "\" & FilesName) Then
'                    '1、写注册表
'                    Sub_name = dom.selectSingleNode("//subplugin").Attributes.getNamedItem("csubname").nodeValue
'                    regwrite Sub_id, Sub_name
'                    '2、写升级配置文件，得到当前配置文件中年度结转数据
'                    If Not dom.selectSingleNode("//pluginmanager/other/Product") Is Nothing And blnServer Then
'                        ID = dom.selectSingleNode("//pluginmanager/other/Product").Attributes.getNamedItem("ID").nodeValue
'                        For Each nd In Dom_TransferProducts.selectNodes("//Products/Product")
'                            If ID = nd.Attributes.getNamedItem("ID").nodeValue Then
'                                have = True
'                                Exit For
'                            Else
'                                have = False
'                            End If
'                        Next
'                        If have = False Then
'                        '根据将当前需要的年度结转的信息追加到产品的年度结转配置文件中
'                            Dom_TransferProducts.selectSingleNode("//Products").appendChild dom.selectSingleNode("//pluginmanager/other/Product").cloneNode(True)
'                            Call Dom_TransferProducts.save(u8path & "\Admin\TransferProducts.xml")
'                        End If
'                    End If
'                    '3、更新数据库升级的版本信息
'                    If Not dom.selectSingleNode("//pluginmanager/other/version") Is Nothing Then
'                        If Not dom.selectSingleNode("//pluginmanager/other/version").Attributes.getNamedItem("updateversion") Is Nothing Then
'                            If UCase(dom.selectSingleNode("//pluginmanager/other/version").Attributes.getNamedItem("updateversion").nodeValue) = "Y" Then
'                                dom.selectSingleNode("//pluginmanager/other/version").Attributes.getNamedItem("updateversion").nodeValue = "N" '防止影响其他模块
'                                dom.selectSingleNode("//pluginmanager/subplugin").Attributes.getNamedItem("ccurdbversion").nodeValue = Format(Now, "yyyymmddhh")
'                                Call dom.save(u8path & "\EF\" & FilesName)
'                            End If
'                        End If
'                    End If
                    '4、注册netdll转化成tlb
                    If Not dom.selectSingleNode("//pluginmanager/other/netdll/dll") Is Nothing Then
                        For Each nd In dom.selectNodes("//pluginmanager/other/netdll/dll")
                            NetFile_dll = u8path & nd.Attributes.getNamedItem("dllfilename").nodeValue
                            NetFile_tlb = u8path & nd.Attributes.getNamedItem("tlbfilename").nodeValue
                            regeditstr = netpath & "v2.0.50727\\regasm.exe /codebase /silent " & NetFile_dll & " /tlb:" & NetFile_tlb
                            Shell regeditstr
                        Next
                    End If

'                    '5、配置userconfig 文件
'                    'clsXML.uninstall objLogin, "userconfig", "Save", "pl_gmp_rdInterface.clsInterface"
'                    If Not dom.selectSingleNode("//pluginmanager/other/userconfig") Is Nothing And blnClient Then
'                        For Each nd In dom.selectNodes("//pluginmanager/other/userconfig/config")
'                            configname = ""
'                            functionname = ""
'                            dll_clsname = ""
'                            configname = nd.Attributes.getNamedItem("configname").nodeValue
'                            functionname = nd.Attributes.getNamedItem("functionname").nodeValue
'                            dll_clsname = nd.Attributes.getNamedItem("dll_clsname").nodeValue
'                            If Trim(configname) <> "" And Trim(functionname) <> "" And Trim(dll_clsname) <> "" Then
'                                installconfig u8path & configname, dll_clsname, functionname
'                            End If
'                        Next
'                    End If
                    
'                    'modify by suyong 20091109
'                    '6、配置结转文件，并复制脚本文件
'                    If Not dom.selectSingleNode("//pluginmanager/other/updatefiles") Is Nothing And blnServer Then
'                        For Each nd In dom.selectNodes("//pluginmanager/other/updatefiles")
'                            If Not nd.Attributes.getNamedItem("path") Is Nothing And Not nd.Attributes.getNamedItem("filename") Is Nothing Then
'                                Call ConvertFiles(dom, u8path, nd.Attributes.getNamedItem("path").Text, nd.Attributes.getNamedItem("filename").Text, u8path & "\EF\" & dom.selectSingleNode("//pluginmanager/subplugin").Attributes.getNamedItem("cid").Text)
'                            End If
'                        Next
'                    End If
'                    'modify end
                    '7 复制文件
                    If Not dom.selectSingleNode("//pluginmanager/other/copyfile") Is Nothing Then
                        For Each nd In dom.selectNodes("//pluginmanager/other/copyfile/file")
                            If Not nd.Attributes.getNamedItem("sourcefile") Is Nothing And Not nd.Attributes.getNamedItem("desfile") Is Nothing Then
                                 
                                If Left(nd.Attributes.getNamedItem("desfile").Text, 1) = "%" Then
                                     fso.CopyFile App.Path & "\" & nd.Attributes.getNamedItem("sourcefile").Text, fso.GetSpecialFolder(1) & Right(nd.Attributes.getNamedItem("desfile").Text, Len(nd.Attributes.getNamedItem("desfile").Text) - InStr(3, nd.Attributes.getNamedItem("desfile").Text, "%")), True
                                Else
                                    fso.CopyFile App.Path & "\" & nd.Attributes.getNamedItem("sourcefile").Text, u8path & nd.Attributes.getNamedItem("desfile").Text, True
                                End If
                            End If
                        Next
                    End If
                    
                End If
            End If
        Next
    End If
    
'
'
'netpath = netpath & "v2.0.50727\"
''文件管理.net文件注册
'
'NetFile_dll = u8path & "\UAP\Runtime\NewUAPList.dll"
'NetFile_tlb = u8path & "\UAP\Runtime\NewUAPList.tlb"
'regeditstr = netpath & "\regasm.exe /codebase /silent " & NetFile_dll & " /tlb:" & NetFile_tlb
'Shell regeditstr
''E:\WINDOWS\Microsoft.NET\Framework\v2.0.50727
'
'NetFile_dll = u8path & "\UAP\Runtime\NewFileBoxList.dll"
'NetFile_tlb = u8path & "\UAP\Runtime\NewFileBoxList.tlb"
'regeditstr = netpath & "regasm.exe /codebase /silent " & NetFile_dll & " /tlb:" & NetFile_tlb
'Shell regeditstr





    
'序号产品版本    模块号  插件启用标识    模块名称    版本
'1   U872       HE      GMP             GMP管理     V5.0
'2   U872       HH      ZZGL            资质审批    V5.0
'3   U872       LH      WLCJ            物料重计    V5.0
'4   U872       J9      QDGL            渠道管理    V5.0
'5   U872       J8      WJGL            文件管理    V5.0
'6   U872       J0      JSGL            结算管理    V5.0
'7   U872       GS      GSP             GSP质量管理 V5.0
'
'    wsh.regwrite "HKEY_LOCAL_MACHINE\SOFTWARE\Ufsoft\WF\V8.700\Install\Installed\HE\", "GMP管理", "REG_SZ"
'    wsh.regwrite "HKEY_LOCAL_MACHINE\SOFTWARE\Ufsoft\WF\V8.700\Install\Installed\HH\", "资质审批", "REG_SZ"
'    wsh.regwrite "HKEY_LOCAL_MACHINE\SOFTWARE\Ufsoft\WF\V8.700\Install\Installed\LH\", "物料重计", "REG_SZ"
'    wsh.regwrite "HKEY_LOCAL_MACHINE\SOFTWARE\Ufsoft\WF\V8.700\Install\Installed\J9\", "渠道管理", "REG_SZ"
'    wsh.regwrite "HKEY_LOCAL_MACHINE\SOFTWARE\Ufsoft\WF\V8.700\Install\Installed\J8\", "文件管理", "REG_SZ"
'    wsh.regwrite "HKEY_LOCAL_MACHINE\SOFTWARE\Ufsoft\WF\V8.700\Install\Installed\J0\", "结算管理", "REG_SZ"
'    wsh.regwrite "HKEY_LOCAL_MACHINE\SOFTWARE\Ufsoft\WF\V8.700\Install\Installed\GS\", "GSP质量管理", "REG_SZ"
End Sub

'modify by suyong 20091109
'将pluginmanager/subplugin/cupgradesql/install下的文件复制到sPath下
Private Sub ConvertFiles(dom As DOMDocument, sU8Path As String, sDesPath As String, sFileName As String, sSourcPath As String)
    Dim strXml As String
    Dim oXml As New DOMDocument
'    Dim nd As IXMLDOMElement
    Dim nd As IXMLDOMNode
    Dim fso As Object
    
    Dim Copy As Boolean
    
    strXml = ""
    
    On Error Resume Next
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not dom.selectSingleNode("//pluginmanager/subplugin/cupgradesql/install") Is Nothing Then
        strXml = dom.selectSingleNode("//pluginmanager/subplugin/cupgradesql/install").xml
        For Each nd In dom.selectNodes("//pluginmanager/subplugin/cupgradesql/install/ufsystem/sqlfile")
            If nd.Text <> "" Then
                '去掉文件的只读属性
                If Not fso.GetFile(sU8Path & sDesPath & nd.Text) Is Nothing Then fso.GetFile(sU8Path & sDesPath & nd.Text).Attributes = 0
                If nd.Attributes.getNamedItem("updatefiles") Is Nothing Then
                    Copy = True
                Else
                    If nd.Attributes.getNamedItem("updatefiles").nodeValue = "N" Then
                        Copy = False
                    Else
                        Copy = True
                    End If
                End If
                If Copy Then
                    fso.CopyFile sSourcPath & "\ufsystem\" & nd.Text, sU8Path & sDesPath & nd.Text, True
                Else
                    dom.selectSingleNode("//pluginmanager/subplugin/cupgradesql/install/ufsystem").removeChild nd
                End If
            End If
        Next
        
        For Each nd In dom.selectNodes("//pluginmanager/subplugin/cupgradesql/install/ufdata/sqlfile")
            If nd.Text <> "" Then
                '去掉文件的只读属性
                If Not fso.GetFile(sU8Path & sDesPath & nd.Text) Is Nothing Then fso.GetFile(sU8Path & sDesPath & nd.Text).Attributes = 0
                If nd.Attributes.getNamedItem("updatefiles") Is Nothing Then
                    Copy = True
                Else
                    If nd.Attributes.getNamedItem("updatefiles").nodeValue = "N" Then
                        Copy = False
                    Else
                        Copy = True
                    End If
                End If
                If Copy Then
                    fso.CopyFile sSourcPath & "\ufdata\" & nd.Text, sU8Path & sDesPath & nd.Text, True
                Else
                    dom.selectSingleNode("//pluginmanager/subplugin/cupgradesql/install/ufdata").removeChild nd
                End If
            End If
        Next
        For Each nd In dom.selectNodes("//pluginmanager/subplugin/cupgradesql/install/ufmeta/sqlfile")
            If nd.Text <> "" Then
                '去掉文件的只读属性
                If Not fso.GetFile(sU8Path & sDesPath & nd.Text) Is Nothing Then fso.GetFile(sU8Path & sDesPath & nd.Text).Attributes = 0
                If nd.Attributes.getNamedItem("updatefiles") Is Nothing Then
                    Copy = True
                Else
                    If nd.Attributes.getNamedItem("updatefiles").nodeValue = "N" Then
                        Copy = False
                    Else
                        Copy = True
                    End If
                End If
                If Copy Then
                    fso.CopyFile sSourcPath & "\ufmeta\" & nd.Text, sU8Path & sDesPath & nd.Text, True
                Else
                    dom.selectSingleNode("//pluginmanager/subplugin/cupgradesql/install/ufmeta").removeChild nd
                End If
            End If
        Next
        For Each nd In dom.selectNodes("//pluginmanager/subplugin/cupgradesql/install/u8workflow/sqlfile")
            If nd.Text <> "" Then
                '去掉文件的只读属性
                If Not fso.GetFile(sU8Path & sDesPath & nd.Text) Is Nothing Then fso.GetFile(sU8Path & sDesPath & nd.Text).Attributes = 0
                If nd.Attributes.getNamedItem("updatefiles") Is Nothing Then
                    Copy = True
                Else
                    If nd.Attributes.getNamedItem("updatefiles").nodeValue = "N" Then
                        Copy = False
                    Else
                        Copy = True
                    End If
                End If
                If Copy Then
                    fso.CopyFile sSourcPath & "\ufmeta\" & nd.Text, sU8Path & sDesPath & nd.Text, True
                Else
                    dom.selectSingleNode("//pluginmanager/subplugin/cupgradesql/install/u8workflow").removeChild nd
                End If
            End If
        Next
        For Each nd In dom.selectNodes("//pluginmanager/subplugin/cupgradesql/install/ufmom/sqlfile")
            If nd.Text <> "" Then
                '去掉文件的只读属性
                If Not fso.GetFile(sU8Path & sDesPath & nd.Text) Is Nothing Then fso.GetFile(sU8Path & sDesPath & nd.Text).Attributes = 0
                If nd.Attributes.getNamedItem("updatefiles") Is Nothing Then
                    Copy = True
                Else
                    If nd.Attributes.getNamedItem("updatefiles").nodeValue = "N" Then
                        Copy = False
                    Else
                        Copy = True
                    End If
                End If
                If Copy Then
                    fso.CopyFile sSourcPath & "\ufmom\" & nd.Text, sU8Path & sDesPath & nd.Text, True
                Else
                    dom.selectSingleNode("//pluginmanager/subplugin/cupgradesql/install/ufmom").removeChild nd
                End If
            End If
        Next
    End If
    strXml = dom.selectSingleNode("//pluginmanager/subplugin/cupgradesql/install").xml
    If strXml <> "" Then
        strXml = "<?xml version='1.0' encoding='gb2312'?>" & vbCrLf & Replace(Replace(strXml, "<install>", "<body>"), "</install>", "</body>")
        oXml.loadXML strXml
        oXml.save sU8Path & sDesPath & sFileName
    End If
End Sub
'modify end

Private Sub regwrite(ByVal Sub_id As String, ByVal Name As String)
Dim wsh As Object
    Set wsh = CreateObject("Wscript.Shell") '利用wshshell写注册表
    wsh.regwrite "HKEY_LOCAL_MACHINE\SOFTWARE\Ufsoft\WF\V8.700\Install\Installed\" & Sub_id & "\", Name, "REG_SZ"
End Sub

Private Sub regread(ByVal HKEYstr As String, ByRef HKEYVale As String)
Dim wsh As Object
    Set wsh = CreateObject("Wscript.Shell") '利用wshshell写注册表
    
        HKEYVale = wsh.regread(HKEYstr)
'    End If
End Sub

Public Function installconfig(ByVal configname As String, ByVal dll_clsname As String, ByVal functionname As String) As Boolean
    Dim domxml As New DOMDocument
    Dim domxmlef As New DOMDocument
    Dim tmpxml As New DOMDocument
    Dim strSql As String
    Dim ndRoot  As IXMLDOMNode
    Dim Node As IXMLDOMNode
    Dim Node2 As IXMLDOMNode
    Dim Nodefunction As IXMLDOMNode
    Dim NdList As IXMLDOMNodeList
    Dim NdListfunction As IXMLDOMNodeList
    Dim iHave As Boolean
    Dim ele As IXMLDOMElement
    Dim att As IXMLDOMAttribute
    Dim i As Integer
    Dim iNum As Integer
 
    iHave = False
    On Error Resume Next
 
'    GetIstallPath = "c:\u8soft"
    If domxml.Load(configname) = False Then
        If domxml Is Nothing Or domxml.xml = "" Then
            '1、创建cfig文件
            strSql = "<config> </config>"
            domxml.loadXML strSql
            domxml.save configname
        End If
    End If
    
    Set NdList = domxml.selectNodes("//userdll")
    If NdList.length = 0 Then
        '2、cfig文件没有节点，创建当前DLL节点
        Set ndRoot = domxml.selectSingleNode("//config")
        strSql = "<dll>  <userdll>" & dll_clsname & "</userdll>  <function><userfunction>" & functionname & "</userfunction></function> </dll>  "
        domxmlef.loadXML strSql
        Set Node = domxmlef.selectSingleNode("//")
        ndRoot.appendChild Node
        domxml.save configname
        installconfig = True
        Exit Function
    Else    '3、cfig文件有节点，并判断是否有当前DLL节点
        For i = 0 To NdList.length - 1
            If NdList.Item(i).nodeTypedValue = dll_clsname Then
                iNum = 1
                Exit For
            End If
        Next i
        If iNum = 0 Then    '4、cfig文件有节点，但没有有当前DLL节点
            Set ndRoot = domxml.selectSingleNode("//config")
            strSql = "<dll>  <userdll>" & dll_clsname & "</userdll>  <function><userfunction>" & functionname & "</userfunction></function> </dll>  "
            domxmlef.loadXML strSql
            Set Node = domxmlef.selectSingleNode("//")
            ndRoot.appendChild Node
            domxml.save configname
            installconfig = True
            Exit Function
        End If
    End If
    

    
    If iNum = 1 Then '5、cfig文件有当前DLL节点
        Set NdList = domxml.selectNodes("//config/dll")
        For Each Node In NdList
            If Node.selectSingleNode("userdll").Text = dll_clsname Then
                Set NdListfunction = Node.selectNodes("function/userfunction")
                For Each Nodefunction In NdListfunction
                    If Nodefunction.Text = functionname Then
                        iNum = 2
                        installconfig = True
                        Exit Function
                    End If
                Next Nodefunction
                If iNum = 1 Then
                    Set ndRoot = Node.selectSingleNode("//function")
                    strSql = " <function> <userfunction>" & functionname & "</userfunction> </function>  "
                    domxmlef.loadXML strSql
                    Set Node2 = domxmlef.selectSingleNode("//function/userfunction")
                    Node.selectSingleNode("function").appendChild Node2
                    domxml.save configname
                    installconfig = True
                    Exit Function
                End If
                
            End If
        Next
    End If
       
   Set domxml = Nothing
End Function


Public Function Uninstallconfig(ByVal configname As String, ByVal dll_clsname As String, ByVal functionname As String)

    Dim domxml As New DOMDocument
    Dim domxmlef As New DOMDocument
    Dim tmpxml As New DOMDocument
    Dim strSql As String
    Dim ndRoot  As IXMLDOMNode
    Dim Node As IXMLDOMNode
    Dim Node2 As IXMLDOMNode
    Dim Nodefunction As IXMLDOMNode
    Dim NdList As IXMLDOMNodeList
    Dim NdListfunction As IXMLDOMNodeList
    Dim iHave As Boolean
    Dim ele As IXMLDOMElement
    Dim att As IXMLDOMAttribute
    Dim i As Integer
    Dim iNum As Integer
    Dim GetIstallPath As String
    iHave = False
    On Error Resume Next
'    GetIstallPath = mlogin.GetIstallPath
'    GetIstallPath = "c:\u8soft"
    If domxml.Load(configname) = False Then
        Uninstallconfig = True
        Exit Function
'    Else
'        domxml.loadXML LCase(domxml.xml)
'        domxml.Save GetIstallPath + "\ufcomsql\" & configname & ".xml"
    End If
    
    Set NdList = domxml.selectNodes("//config/dll")
    For Each Node In NdList
        If Node.selectSingleNode("userdll").Text = dll_clsname Then
            Set NdListfunction = Node.selectNodes("function/userfunction")
            For Each Nodefunction In NdListfunction
                If Nodefunction.Text = functionname Then
                    If NdListfunction.length > 1 Then
                        Node.selectSingleNode("function").removeChild Nodefunction
                    Else
                        domxml.selectSingleNode("config").removeChild Node
                    End If
                    domxml.save configname
                    Uninstallconfig = True
                    Exit Function
                End If
            Next Nodefunction
        End If
    Next
    
   Set domxml = Nothing
End Function

