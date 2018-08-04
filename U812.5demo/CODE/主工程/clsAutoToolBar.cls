VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAutoToolBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Dim domTblConfig As New DOMDocument
Dim domButtonState As New DOMDocument

Public Function initToolBar(tbl As Toolbar, KeyValue As String, strErrorResId As String) As Boolean
    Dim rst As New ADODB.Recordset
    
    rst.CursorLocation = adUseClient
    rst.Open "select tblkey, imageid, keydownapplytask, keyupfreetask, buttonkey, " _
      & "buttoncaptionresid, tooltiptextresid,isnull(buttonstyle,'0') as buttonstyle, inivisible, inienable, buttonindex," _
      & "subindex, hotshift, hotkeycode, hotkeystate, editenable, buttonparas, downstatus," _
      & "upstatus, errstatus, buttontag, grouptype from sa_toolbarconfig with (nolock) where tblkey=N'" + KeyValue + "' order by buttonindex,subindex", DBconn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rst.Save domTblConfig, adPersistXML
    rst.Close
    rst.Open "select * from sa_buttonstate with (nolock) where tblkey=N'" + KeyValue + "'", DBconn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rst.Save domButtonState, adPersistXML
    rst.Close
    Set rst = Nothing
    With tbl
        .buttons.Clear
        Dim nod As IXMLDOMNode
        Dim subNod As IXMLDOMNode
        Dim lst As IXMLDOMNodeList
        Dim sublst As IXMLDOMNodeList
        Dim strButtonIndex As String
        
        Dim btn As Button
        Dim strValue As String
        Set lst = domTblConfig.selectNodes("//z:row[@buttonstyle='0' or @buttonstyle='5']")
        
        For Each nod In lst
            strButtonIndex = nod.Attributes.getNamedItem("buttonindex").Text
            Set btn = .buttons.Add(, nod.Attributes.getNamedItem("buttonkey").Text, GetResStringByResID(nod.Attributes.getNamedItem("buttoncaptionresid").Text, m_login.LanguageRegion), nod.Attributes.getNamedItem("buttonstyle").Text)
            If nod.Attributes.getNamedItem("inivisible") Is Nothing Then
                btn.Visible = False
            Else
                btn.Visible = CBool(nod.Attributes.getNamedItem("inivisible").Text)
            End If
            If nod.Attributes.getNamedItem("inienable") Is Nothing Then
                btn.Enabled = False
            Else
                btn.Enabled = CBool(nod.Attributes.getNamedItem("inienable").Text)
            End If
            If Not nod.Attributes.getNamedItem("tooltiptextresid") Is Nothing Then
                strValue = ""
                If Not nod.Attributes.getNamedItem("hotkeystate") Is Nothing Then
                    strValue = nod.Attributes.getNamedItem("hotkeystate").nodeValue
                End If
                btn.ToolTipText = GetResStringByResID(nod.Attributes.getNamedItem("tooltiptextresid").Text, m_login.LanguageRegion) & IIf(strValue = "", "", "(" + strValue + ")")
            End If
            If nod.Attributes.getNamedItem("buttonstyle").Text = "5" Then
                Set sublst = domTblConfig.selectNodes("//z:row[@buttonstyle='" + nod.Attributes.getNamedItem("buttonkey").Text + "']")
                For Each subNod In sublst
                    strValue = ""
                    If Not subNod.Attributes.getNamedItem("hotkeystate") Is Nothing Then
                        strValue = subNod.Attributes.getNamedItem("hotkeystate").nodeValue
                    End If
                    btn.ButtonMenus.Add , subNod.Attributes.getNamedItem("buttonkey").Text, GetResStringByResID(subNod.Attributes.getNamedItem("buttoncaptionresid").Text, m_login.LanguageRegion) & IIf(strValue = "", "", "(" + strValue + ")")
                    btn.ButtonMenus.Item(btn.ButtonMenus.Count).Visible = CBool(subNod.Attributes.getNamedItem("inivisible").Text)
                    btn.ButtonMenus.Item(btn.ButtonMenus.Count).Enabled = CBool(subNod.Attributes.getNamedItem("inienable").Text)
                    btn.ButtonMenus.Item(btn.ButtonMenus.Count).Tag = GetResStringByResID(subNod.Attributes.getNamedItem("tooltiptextresid").Text, m_login.LanguageRegion)
                Next
            End If
            If Not nod.Attributes.getNamedItem("buttontag") Is Nothing Then
                Dim strGroup As String
                strGroup = "ICOMMON"
                If Not nod.Attributes.getNamedItem("grouptype") Is Nothing Then
                    
                    strGroup = nod.Attributes.getNamedItem("grouptype").Text
                End If
                btn.Tag = CreatePortalToolbarTag(nod.Attributes.getNamedItem("buttontag").Text, strGroup, "PortalToolbar")
            End If
        Next
    End With
End Function

Private Sub Class_Terminate()
    Set domTblConfig = Nothing
    Set domButtonState = Nothing
End Sub

Public Function ButtonKeyDown(m_login As U8Login.clsLogin, strKeyValue As String, Optional blnMsg As Boolean = True) As Boolean
    Dim nod As IXMLDOMNode
    Set nod = domTblConfig.selectSingleNode("//z:row[@buttonkey='" + strKeyValue + "']")
    ButtonKeyDown = True
    If nod Is Nothing Then Exit Function
    If Not nod.Attributes.getNamedItem("keydownapplytask") Is Nothing Then
        ButtonKeyDown = LockItem(nod.Attributes.getNamedItem("keydownapplytask").Text, True, blnMsg)
    End If
End Function

Public Function ReleaseKeyDown(m_login As U8Login.clsLogin, strKeyValue As String) As Boolean
    Dim nod As IXMLDOMNode
    Set nod = domTblConfig.selectSingleNode("//z:row[@buttonkey='" + strKeyValue + "']")
    If nod Is Nothing Then Exit Function
    ReleaseKeyDown = True
    If Not nod.Attributes.getNamedItem("keydownapplytask") Is Nothing Then
        ReleaseKeyDown = LockItem(nod.Attributes.getNamedItem("keydownapplytask").Text, False)
    End If
End Function

Public Sub ButtonKeyUp(m_login As U8Login.clsLogin, strKeyValue As String)
    Dim nod As IXMLDOMNode
    Set nod = domTblConfig.selectSingleNode("//z:row[@buttonkey='" + strKeyValue + "']")
    If nod Is Nothing Then Exit Sub
    If Not nod.Attributes.getNamedItem("keyupfreetask") Is Nothing Then
        LockItem nod.Attributes.getNamedItem("keyupfreetask").Text, False
    End If
End Sub

Public Function GetButtonStatus(strKeyValue As String, strName As String) As String
    GetButtonStatus = ""
    Dim nod As IXMLDOMNode
    Set nod = domTblConfig.selectSingleNode("//z:row[@buttonkey='" + strKeyValue + "']")
    If nod Is Nothing Then Exit Function
    If Not nod.Attributes.getNamedItem(strName) Is Nothing Then
        GetButtonStatus = GetString(nod.Attributes.getNamedItem(strName).Text)
    End If
End Function

Public Sub ChangeButtonState(Voucher As ctlVoucher, tbl As Toolbar, uftbl As UFToolbar, intVoucherState As VoucherStatusConstants)
    Dim lst As IXMLDOMNodeList
    Dim nod As IXMLDOMNode
    Dim strField As String
    Dim subNod As IXMLDOMNode
    On Error Resume Next
    For Each nod In domTblConfig.selectNodes("//z:row[@buttonstyle='0' or @buttonstyle='5']")
        If intVoucherState = VSNormalMode Then
            strField = "inienable"
        Else
            strField = "editenable"
        End If
        'Debug.Print nod.Attributes.getNamedItem("buttonkey").Text
        If nod.Attributes.getNamedItem(strField) Is Nothing Then
            tbl.buttons(nod.Attributes.getNamedItem("buttonkey").Text).Enabled = False
        Else
            If nod.Attributes.getNamedItem(strField).Text <> "2" Then
                tbl.buttons(nod.Attributes.getNamedItem("buttonkey").Text).Enabled = CBool(nod.Attributes.getNamedItem(strField).Text)
            End If
            If tbl.buttons(nod.Attributes.getNamedItem("buttonkey").Text).Style = tbrDropdown Then
                Set lst = domTblConfig.selectNodes("//z:row[@buttonstyle='" + nod.Attributes.getNamedItem("buttonkey").Text + "']")
                For Each subNod In lst
                    tbl.buttons(nod.Attributes.getNamedItem("buttonkey").Text).ButtonMenus.Item(subNod.Attributes.getNamedItem("buttonkey").Text).Enabled = CBool(subNod.Attributes.getNamedItem(strField).Text)
                Next
            End If
        End If
    Next
    SetButtonStateBySet Voucher, tbl
    uftbl.RefreshVisible
End Sub

Private Sub SetButtonStateBySet(Voucher As ctlVoucher, tbl As Toolbar)
    Dim lst As IXMLDOMNodeList
    Dim nod As IXMLDOMNode
    Dim varCheck As Variant
    Dim strCheck As String
    Dim varFieldName As Variant
    Dim strValue As String
    Dim blnEnable As Boolean
    Dim clsComp As New UsSaCompStr.clsCompStr
    On Error Resume Next
    For Each nod In domButtonState.selectNodes("//z:row")
        blnEnable = True
        'Debug.Print nod.Attributes.getNamedItem("checktype").Text
        'If nod.Attributes.getNamedItem("buttonkey").Text = "openorder" Then Debug.Assert 0
        strCheck = ReplaceVoucherItems(nod.Attributes.getNamedItem("checktype").Text, Voucher)
        strValue = clsComp.GetExp(strCheck)
        If CBool(strValue) Then
            blnEnable = False
        End If
        If nod.Attributes.getNamedItem("parentkey") Is Nothing Then
            'Debug.Print nod.Attributes.getNamedItem("buttonkey").Text
            If Not blnEnable Then tbl.buttons(nod.Attributes.getNamedItem("buttonkey").Text).Enabled = False
        Else
            If Not blnEnable Then tbl.buttons(nod.Attributes.getNamedItem("parentkey").Text).ButtonMenus.Item(nod.Attributes.getNamedItem("buttonkey").Text).Enabled = False
        End If
    Next
'    uftbl.RefreshVisible
End Sub

Private Function ReplaceVar(strItem As String, strKey As String) As String
    ReplaceVar = Replace(strItem, "@ckey", strKey)
End Function

Public Sub SetListButtonState(strKey As String, tbl As Toolbar, uftbl As UFToolbar)
    Dim lst As IXMLDOMNodeList
    Dim nod As IXMLDOMNode
    Dim varCheck As Variant
    Dim strCheck As String
    Dim varFieldName As Variant
    Dim strValue As String
    Dim blnEnable As Boolean
    Dim clsComp As New UsSaCompStr.clsCompStr
    On Error Resume Next
    For Each nod In domButtonState.selectNodes("//z:row")
        blnEnable = True
        strCheck = ReplaceVar(nod.Attributes.getNamedItem("checktype").Text, strKey)
        strValue = clsComp.GetExp(strCheck)
        If CBool(strValue) Then
            blnEnable = False
        End If
        Debug.Print nod.Attributes.getNamedItem("buttonkey").Text
        If nod.Attributes.getNamedItem("parentkey") Is Nothing Then
            If Not blnEnable Then tbl.buttons(nod.Attributes.getNamedItem("buttonkey").Text).Visible = False
        Else
            If Not blnEnable Then tbl.buttons(nod.Attributes.getNamedItem("parentkey").Text).ButtonMenus.Item(nod.Attributes.getNamedItem("buttonkey").Text).Visible = False
        End If
    Next
    uftbl.RefreshVisible
End Sub

Public Function GetKeyCodeByHotKey(KeyCode As Integer, Shift As Integer) As IXMLDOMElement
    Dim nod As IXMLDOMNode
    Set nod = domTblConfig.selectSingleNode("//z:row[@hotshift='" + CStr(Shift) + "' and @hotkeycode='" + CStr(KeyCode) + "']")
    If Not nod Is Nothing Then
        Set GetKeyCodeByHotKey = nod ' nod.Attributes.getNamedItem("buttonkey").Text
    End If
End Function

Public Function GetButtonPara(strKeyValue As String) As String
    Dim nod As IXMLDOMNode
    Set nod = domTblConfig.selectSingleNode("//z:row[@buttonkey='" + strKeyValue + "']")
    If nod Is Nothing Then Exit Function
    If Not nod.Attributes.getNamedItem("buttonparas") Is Nothing Then
        GetButtonPara = nod.Attributes.getNamedItem("buttonparas").Text
    End If
End Function
