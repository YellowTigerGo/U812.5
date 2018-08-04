Attribute VB_Name = "basPortalPublic"
Option Explicit

'产生GUID的API相关函数
Private Type guid
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4(8) As Byte
End Type

Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As guid) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long

' Show window
Public Const SW_SHOW = 5
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1

'门户传递进来的命令行参数，包含Login信息，860不用
Public g_cCommand As String
'门户MDI主窗体
Public g_oMainFrmProxy As Object
'判断产品是否可以注销
Public g_bCanExit As Boolean

'Frank Hawker(黄剑锋)
Public g_business As Object
Public g_bLogined As Boolean


Public Function InitDockinPortalEnv(sKey As String, CurForm As Form) As Object
    'Frank Hawker(黄剑锋)
    Dim vfd As Object
    If Not (g_business Is Nothing) Then
        Set vfd = g_business.CreateFormEnv(sKey, CurForm)
        'ShowWindow CurForm.hWnd, SW_HIDE
    End If
    Set InitDockinPortalEnv = vfd
End Function

Public Sub DockinPortal(sKey As String, vfd As Object, CurForm As Form)
    'Frank Hawker(黄剑锋)
    If Not (g_business Is Nothing) Then
        Call g_business.RegisterForm(sKey, CurForm, CreateVBFormDescriptor(CurForm, vfd))
        Call g_business.ShowInEditor(sKey, "UFIDA.U8.Portal.Proxy.editors.VoucherEditor")
        'ShowWindow CurForm.hWnd, SW_SHOW
    End If
End Sub

Public Sub DockinPortalModal(sKey As String, vfd As Object, CurForm As Form)
    'Frank Hawker(黄剑锋)
    If Not (g_business Is Nothing) Then
        Call g_business.RegisterForm(sKey, CurForm, CreateVBFormDescriptor(CurForm, vfd))
        Call g_business.ShowInEditor(sKey, "UFIDA.U8.Portal.Proxy.editors.DialogEditor")
        'ShowWindow CurForm.hWnd, SW_SHOW
    End If
End Sub

Private Function CreateVBFormDescriptor(fm As Form, vfd As Object) As Object
    If (vfd Is Nothing) Then
        Set vfd = CreateObject("UFPortalProxy.VBFormDescriptor")
    End If
    vfd.Name = fm.Name
    vfd.Title = fm.Caption
    vfd.handle = fm.hwnd
    Set CreateVBFormDescriptor = vfd
End Function

'示例，可以用如下方法来调用微软Toolbar的Tag属性，各产品可以Toolbar上Button的具体功能来设置图示
'分组，还有Toolbar的类型
Public Sub InitToolbarTag(TB As MSComctlLib.Toolbar)
    Dim i As Integer
    Dim tblMenu As ButtonMenu
    Dim strGroup As String
    
    For i = 1 To TB.buttons.Count
        If TB.buttons(i).Style <> tbrSeparator Then
            If TB.buttons(i).Tag = "" Then
                MsgBox TB.buttons(i).Key
            End If
            strGroup = GetButtonGroup(TB.buttons(i).Tag)
            TB.buttons(i).Tag = CreatePortalToolbarTag(TB.buttons(i).Tag, strGroup, "PortalToolbar")
        End If
    Next
End Sub
Private Function GetButtonGroup(strButtonKey As String) As String
'常用 (ICOMMON): 如打开?保存等
'编辑 (IEDIT): 如增行?拷贝单元格等
'处理 (IDEAL): 如审核?分单等
'查询 (ISEARCH): 如过滤?联查等
    GetButtonGroup = "ICOMMON"
    Select Case LCase(strButtonKey)
        Case "save", "add", "batchadd", "modify", "chenged", "erase", "copy", "addrow", "delrow", "cancel", "Erase"
            GetButtonGroup = "IEDIT"
            
        Case "approval query", "sure", "unsure", "seek"
            GetButtonGroup = "IDEAL"
            
        Case "ToFirst", "toprevious", "tonext", "tolast", "paint", "filter"
            GetButtonGroup = "ISEARCH"
    
    End Select
End Function

''image 按钮的图片，actionSet 分组, toolbarType 分组Toolbar的类型
Public Function CreatePortalToolbarTag(image As String, actionSet As String, toolbarType As String, Optional ByVal OldTag As String) As String
    Dim OldID As String
    Dim oldidend, oldidstart As Integer
    If OldTag <> "" Then
        oldidend = InStr(1, OldTag, "&&&IMAGE:")
        oldidstart = InStr(1, OldTag, "ID:") + 3
        OldID = Mid(OldTag, oldidstart, oldidend - oldidstart)
    End If
    If OldID <> "" Then
        CreatePortalToolbarTag = "ID:" & OldID & "&&&IMAGE:" & image & "&&&ACTIONSET:" & actionSet & "&&&TOOLBARTYPE:" & toolbarType
    Else
        CreatePortalToolbarTag = "ID:" & CreateGUID() & "&&&IMAGE:" & image & "&&&ACTIONSET:" & actionSet & "&&&TOOLBARTYPE:" & toolbarType
    End If
End Function



Public Function CreateGUID(Optional strRemoveChars As String = "") As String
    Dim udtGUID As guid
    Dim strGUID As String
    Dim bytGUID() As Byte
    Dim lngLen As Long
    Dim lngRetVal As Long
    Dim lngPos As Long

    'Initialize
    lngLen = 40
    bytGUID = String(lngLen, 0)

    'Create the GUID
    CoCreateGuid udtGUID

    'Convert the structure into a displayable string
    lngRetVal = StringFromGUID2(udtGUID, VarPtr(bytGUID(0)), lngLen)
    strGUID = bytGUID
    If (Asc(Mid$(strGUID, lngRetVal, 1)) = 0) Then
        lngRetVal = lngRetVal - 1
    End If

    'Trim the trailing characters
    strGUID = Left$(strGUID, lngRetVal)

    'Remove the unwanted characters
    For lngPos = 1 To Len(strRemoveChars)
        strGUID = Replace(strGUID, Mid(strRemoveChars, lngPos, 1), "")
    Next
    strGUID = Replace(strGUID, "-", "")
    strGUID = Replace(strGUID, "{", "")
    strGUID = Replace(strGUID, "}", "")
    CreateGUID = strGUID
End Function



Public Function GetActiveForm() As Object
    If Not (g_business Is Nothing) Then
        Set GetActiveForm = g_business.GetActiveForm
    Else
        Set GetActiveForm = g_oMainFrmProxy.GetActiveForm
    End If
End Function
 






