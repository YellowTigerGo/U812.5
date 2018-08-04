Attribute VB_Name = "basPortalPublic"
Option Explicit

'����GUID��API��غ���
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

'�Ż����ݽ����������в���������Login��Ϣ��860����
Public g_cCommand As String
'�Ż�MDI������
Public g_oMainFrmProxy As Object
'�жϲ�Ʒ�Ƿ����ע��
Public g_bCanExit As Boolean

'Frank Hawker(�ƽ���)
Public g_business As Object
Public g_bLogined As Boolean


Public Function InitDockinPortalEnv(sKey As String, CurForm As Form) As Object
    'Frank Hawker(�ƽ���)
    Dim vfd As Object
    If Not (g_business Is Nothing) Then
        Set vfd = g_business.CreateFormEnv(sKey, CurForm)
        'ShowWindow CurForm.hWnd, SW_HIDE
    End If
    Set InitDockinPortalEnv = vfd
End Function

Public Sub DockinPortal(sKey As String, vfd As Object, CurForm As Form)
    'Frank Hawker(�ƽ���)
    If Not (g_business Is Nothing) Then
        Call g_business.RegisterForm(sKey, CurForm, CreateVBFormDescriptor(CurForm, vfd))
        Call g_business.ShowInEditor(sKey, "UFIDA.U8.Portal.Proxy.editors.VoucherEditor")
        'ShowWindow CurForm.hWnd, SW_SHOW
    End If
End Sub

Public Sub DockinPortalModal(sKey As String, vfd As Object, CurForm As Form)
    'Frank Hawker(�ƽ���)
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

'ʾ�������������·���������΢��Toolbar��Tag���ԣ�����Ʒ����Toolbar��Button�ľ��幦��������ͼʾ
'���飬����Toolbar������
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
'���� (ICOMMON): ���?�����
'�༭ (IEDIT): ������?������Ԫ���
'���� (IDEAL): �����?�ֵ���
'��ѯ (ISEARCH): �����?�����
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

''image ��ť��ͼƬ��actionSet ����, toolbarType ����Toolbar������
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
 






