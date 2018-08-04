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


' Show window
'Public Const SW_SHOW = 5
Public Const SW_HIDE = 0
'Public Const SW_SHOWNORMAL = 1

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
    vfd.Handle = fm.hwnd
    Set CreateVBFormDescriptor = vfd
End Function

'示例，可以用如下方法来调用微软Toolbar的Tag属性，各产品可以Toolbar上Button的具体功能来设置图示
'分组，还有Toolbar的类型
Public Sub InitToolbarTag(tb As MSComctlLib.Toolbar)
    Dim i As Integer
    For i = 1 To tb.Buttons.Count
        If Not (tb.Buttons(i).Style = tbrSeparator) Then
            tb.Buttons(i).Tag = CreatePortalToolbarTag("ICON_NEW", "IPRINTABLE", "PortalToolbar")
        End If
    Next
End Sub

'image 按钮的图片，actionSet 分组, toolbarType 分组Toolbar的类型
Public Function CreatePortalToolbarTag(image As String, actionSet As String, toolbarType As String) As String
    CreatePortalToolbarTag = "ID:" & CreateGUID() & "&&&IMAGE:" & image & "&&&ACTIONSET:" & actionSet & "&&&TOOLBARTYPE:" & toolbarType
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
    If (asc(Mid$(strGUID, lngRetVal, 1)) = 0) Then
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

' **------------------------------------------------------------**
' 函数名    :gError_Proc
' 功能      :错误处理
' 返回值    :无
' 参数      :strFunctionName        As String       出错函数名
'            blnFlag                As Boolean      是否记录
' 功能说明  :错误处理，并记录错误到日志文件
' 备注      :默认参数为记录日志文件
' **------------------------------------------------------------**
Public Sub gError_Proc(ByVal strFunctionName As String, _
            Optional ByVal blnLog As Boolean = True)

    Dim strMsg                  As String
    Dim intFileNum              As Integer
    Dim strFileName             As String

    strMsg = "Position: " & strFunctionName & vbCrLf
    strMsg = strMsg & "Error Number: " & CStr(Err.Number) & vbCrLf
    strMsg = strMsg & "Error Description: " & Err.Description

    #If DEBUGVER >= 2 Then
        MsgBox strMsg, , "Unhandling error occor"
    #End If

    If blnLog Then
        strFileName = App.path
        If Left$(strFileName, 1) <> "\" Then strFileName = strFileName & "\"
        strFileName = strFileName & "error.log"

        '记录时间信息
        strMsg = String(70, "-") & vbCrLf
        strMsg = strMsg & "Time             : " & Format(Now(), "yyyy-MM-dd HH:mm:ss") & vbCrLf

        '记录程序信息
        strMsg = strMsg & "[Program Information]" & vbCrLf
        strMsg = strMsg & "  Name           : " & App.EXEName & vbCrLf
        strMsg = strMsg & "  Function       : " & strFunctionName & vbCrLf

        '记录错误信息
        strMsg = strMsg & "[VB Error]" & vbCrLf
        strMsg = strMsg & "  Source         : " & Err.Source & vbCrLf
        strMsg = strMsg & "  Number         : " & CStr(Err.Number) & vbCrLf
        strMsg = strMsg & "  Description    : " & Err.Description

        '假如日志文件超过2MB，则删除
        If Dir$(strFileName) <> vbNullString Then
            If FileLen(strFileName) > 2 * 1024# * 1024# Then
                Kill strFileName
            End If
        End If

        intFileNum = FreeFile

        Open strFileName For Append As #intFileNum
        Print #intFileNum, strMsg
        Close #intFileNum

    End If
End Sub


' **------------------------------------------------------------**
' 函数名    :RecordLogFile
' 功能      :日期记录
' 返回值    :无
' 参数      :strFunctionName        As String       出错函数名
'            blnFlag                As Boolean      是否记录
' 功能说明  :错误处理，并记录错误到日志文件
' 备注      :默认参数为记录日志文件
' **------------------------------------------------------------**
Public Sub RecordLogFile(ByVal strLogMsg As String, Optional strFncName As String = "", _
            Optional ByVal blnLog As Boolean = True, Optional sTimeFlag As String = "")

    Dim strMsg                  As String
    Dim intFileNum              As Integer
    Dim strFileName             As String
    Dim retnew As Long
    Dim retold As Long
    Dim PersistTime As Long
    Dim bFlagExist As Boolean
    
    Static timecol As Collection
    
    retnew = GetTickCount
    
    If (timecol Is Nothing) Then
        Set timecol = New Collection
    End If
    
    bFlagExist = IsExistsItem(timecol, strFncName & "." & sTimeFlag)
    
    If (sTimeFlag <> "") Then
        If (bFlagExist) Then
            retold = timecol(strFncName & "." & sTimeFlag)
        Else
            timecol.Add retnew, strFncName & "." & sTimeFlag
        End If
    End If

    

    #If DEBUGVER >= 1 Then
        Debug.Print "Time:" & Now & "    Message:" & strLogMsg & "  Function:" & strFncName
        If blnLog Then
            strFileName = App.path
            If Left$(strFileName, 1) <> "\" Then strFileName = strFileName & "\"
            strFileName = strFileName & "recordinfo(UFPortalProxy).log"

            '记录程序信息
            strMsg = strMsg & "[Log Information]" & vbCrLf
            strMsg = strMsg & "  Name           : " & App.EXEName & vbCrLf
            If (strFncName <> "") Then
                strMsg = strMsg & "  Function          : " & strFncName & vbCrLf
            Else
                strMsg = strMsg & "  Function          : Unspecified" & vbCrLf
            End If
            strMsg = strMsg & "  Message           : " & strLogMsg & vbCrLf
            strMsg = strMsg & "  Time             : " & Format(Now(), "yyyy-MM-dd HH:mm:ss") & vbCrLf
            If (bFlagExist) Then
                strMsg = strMsg & "  RunTime(" & sTimeFlag & ") :" & CStr(retnew - retold) & " Begin:" & CStr(retold) & " End:" & CStr(retnew)
                'timecol(strFncName & "." & sTimeFlag) = retnew
                timecol.Remove (strFncName & "." & sTimeFlag)
                timecol.Add retnew, strFncName & "." & sTimeFlag
            End If
            
            Debug.Print strMsg

            '假如日志文件超过2MB，则删除
            If Dir$(strFileName) <> vbNullString Then
                If FileLen(strFileName) > 2 * 1024# * 1024# Then
                    Kill strFileName
                End If
            End If

            intFileNum = FreeFile

            Open strFileName For Append As #intFileNum
            Print #intFileNum, strMsg
            Close #intFileNum
        End If
    #End If
End Sub


Public Function IsExistsItem(list As Collection, ByVal fldname As String) As Boolean
    If (ExistsItem(list, fldname)) Then
        IsExistsItem = True
        GoTo ExitFnc
    End If
    If (ExistsItemObj(list, fldname)) Then
        IsExistsItem = True
        GoTo ExitFnc
    End If
ExitFnc:
    Exit Function
ErrHandler:
    Call gError_Proc("mFunctions.IsExistsItem")
    GoTo ExitFnc
End Function

'是否集合存在此项目
Private Function ExistsItem(list As Collection, ByVal fldname As String) As Boolean
    Dim Key As String
    On Error GoTo err0
    Key = list(fldname)
    ExistsItem = True
    Exit Function
err0:
    ExistsItem = False
End Function

Private Function ExistsItemObj(list As Collection, ByVal fldname As String) As Boolean
    Dim Key As Object
    On Error GoTo err0
    Set Key = list(fldname)
    ExistsItemObj = True
    Exit Function
err0:
    ExistsItemObj = False
End Function



