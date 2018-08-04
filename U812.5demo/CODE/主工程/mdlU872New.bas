Attribute VB_Name = "mdlU872New"

'此模块主要存放与菜单、工具栏、单据、参照有关的变量、函数及方法
Option Explicit
Public domMenu As New DOMDocument    '//保存与菜单有关属性的对象
Public domLook As New DOMDocument

Public Function GetMenuConfig()
    Dim rst As New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.Open "select * from sa_MenuConfig", DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rst.Save domMenu, adPersistXML
    rst.Close
    Set rst = Nothing
End Function

Public Function GetResStringByResID(strResID As String, strLanguageRegion As String) As String
    If strResID <> "" Then
        GetResStringByResID = GetStringByLocalID(strResID, strLanguageRegion)
    End If
End Function

Public Function ReplaceVoucherItems(strsql As String, Voucher As ctlVoucher, Optional lngRow As Long) As String
    Dim lngPos1 As Integer
    Dim lngPos2 As Integer
    Dim strFieldName As String
    Dim varField As Variant
    Dim strValue As String
    
    strsql = Replace(strsql, "@VoucherStatus", CStr(Voucher.VoucherStatus))
    lngPos1 = InStr(1, strsql, "[")
    Do While lngPos1 > 0
        lngPos2 = InStr(lngPos1, strsql, "]")
        If lngPos2 <= 0 Then Exit Do
        strFieldName = Mid(strsql, lngPos1 + 1, lngPos2 - lngPos1 - 1)
        varField = Split(strFieldName, ",")
        If UBound(varField) = 1 Then
            strValue = GetVoucherItemValue(Voucher, CStr(varField(0)), CStr(varField(1)), lngRow)
            strsql = Replace(strsql, "[" + varField(0) + "," + varField(1) + "]", strValue)
        Else
            strsql = Replace(strsql, "[" + varField(0) & "]", varField(0) + "")
        End If
        lngPos1 = InStr(lngPos1 + Len(strValue), strsql, "[")
    Loop
    ReplaceVoucherItems = strsql
End Function


Public Function GetVoucherItemValue(Voucher As ctlVoucher, strSection As String, strFieldName As String, Optional lngRow As Long) As String
    If strSection = "B" Then
        GetVoucherItemValue = Voucher.bodyText(lngRow, strFieldName)
    End If
    If strSection = "T" Then
        GetVoucherItemValue = Voucher.headerText(strFieldName)
    End If
End Function

Public Sub SetVoucherItemValue(Voucher As ctlVoucher, strSection As String, strFieldName As String, strValue As String, Optional lngRow As Long)
    If strSection = "B" Then
        Voucher.bodyText(lngRow, strFieldName) = strValue
    End If
    If strSection = "T" Then
        Voucher.headerText(strFieldName) = strValue
    End If
End Sub

Public Function ReplaceSysPara(strSource As String) As String
    Dim lngPos1 As Integer
    Dim lngPos2 As Integer
    Dim strFieldName As String
    Dim varField As Variant
    
    lngPos1 = InStr(1, strSource, "@")
    Do While lngPos1 > 0
        lngPos2 = InStr(lngPos1, strSource, "=")
        If lngPos2 = 0 Then
            strFieldName = Mid(strSource, lngPos1)
            If Right(strFieldName, 1) = ")" Then
                strFieldName = Left(strFieldName, Len(strFieldName) - 1)
            End If
            If Right(strFieldName, 1) = """" Then
                strFieldName = Left(strFieldName, Len(strFieldName) - 1)
            End If
        Else
            strFieldName = Mid(strSource, lngPos1, lngPos2 - lngPos1)
        End If
        If Right(strFieldName, 1) = """" Then strFieldName = Left(strFieldName, Len(strFieldName) - 1)
        strSource = Replace(strSource, strFieldName, GetGlobalVariant(CStr(strFieldName)))
        lngPos1 = InStr(1, strSource, "@")
    Loop
    ReplaceSysPara = strSource
End Function

Public Function GetGlobalVariant(strName As String) As String
    Select Case LCase(strName)
        Case "@username"
            GetGlobalVariant = m_Login.cUserName
        Case "@curdate"
            GetGlobalVariant = m_Login.CurDate
        Case Else
            GetGlobalVariant = strName
    End Select
End Function

Public Function ShowPortalForm(Frm As Form, blnModel As Boolean) As String
    Dim sGuid As String
    Dim vfd As Object
    
    If blnModel Then
        Frm.Show 1
    Else
        sGuid = CreateGUID()
        Set vfd = g_business.CreateFormEnv(sGuid, Frm)
        Call g_business.ShowForm(Frm, "J6", sGuid, blnModel, True, vfd)
    End If
    ShowPortalForm = sGuid
End Function

'87X added 判断是否启用工作流
Public Function getIsWfControl(myConn As ADODB.Connection, bizObjectID As String, ByRef errMsg As String) As Boolean
    Dim clsisWfCtl As Object
    Dim strkey As String
    Select Case bizObjectID
        Case "01", "05"
            strkey = "01"
        Case "07", "14", "15"
            strkey = "07"
        Case "02", "04"
            strkey = "02"
        Case Else
            strkey = bizObjectID
    End Select
    Set clsisWfCtl = CreateObject("SCMWorkFlowCommon.clsWFController")
    Dim isWfCtl As Boolean
    Call clsisWfCtl.GetIsWFControlled(myConn, strkey, strkey & ".Submit", m_Login.cIYear, m_Login.cAcc_Id, isWfCtl, errMsg)
    getIsWfControl = isWfCtl
End Function

Public Sub ShowWorkFlowView(m_strFormGuid As String, strCardNumber As String, sViewID As String, Optional ByVal strMessageType As String = "SHOWVIEW")
'sViewID:="UFIDA.U8.Audit.AuditViews.TreatTaskViewPart",审批视图,
'sViewID:="UFIDA.U8.Audit.AuditHistoryView",审批进程表查,审时调用
'SHOWVIEW显示视图，HIDEVIEW隐藏视图
    Dim tsb As Object
    Dim strXml As String
    If Not (g_business Is Nothing) Then
        Set tsb = g_business.GetToolbarSubjectEx(m_strFormGuid)
    End If
    strXml = ""
    strXml = strXml & "<Message type='" & strMessageType & "'>"
    strXml = strXml & "   <Selection context='SA:" + strCardNumber + "'>"
    strXml = strXml & "      <Element typeName = 'ViewPart' viewID = '" & sViewID & "'  isFirstElement = 'true'/> "
    strXml = strXml & "   </Selection>"
    strXml = strXml & "</Message>"
    If Not (tsb Is Nothing) Then
           Call tsb.TransMessage(m_strFormGuid, strXml)
    End If
 
    If Not tsb Is Nothing Then Set tsb = Nothing
   
End Sub


'操作系统帮助调不出来，使用变量ContextID指定要显示的主题
Public Function ShowContextHelp(hwnd As Long, sHelpFile As String, lContextID As Long) As Long
'    MsgBox sHelpFile & "-----" & lContextID
    ShowContextHelp = htmlHelp(hwnd, sHelpFile, &HF, lContextID)
End Function
