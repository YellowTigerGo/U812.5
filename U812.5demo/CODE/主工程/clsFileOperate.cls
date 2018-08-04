VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileOperate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mobjFileClient As Object
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private m_objFile As U8FileManageClient
Private m_objU8Login As U8Login.clsLogin

Public Property Get U8Login()
    U8Login = m_objU8Login
End Property
'Public Property Get U8Login(ByVal vNewValue As Variant)
'    Set m_objU8Login = vNewValue
'End Property


Private Sub Class_Terminate()
    If Not m_objFile Is Nothing Then Set m_objFile = Nothing
    If Not m_objU8Login Is Nothing Then Set m_objU8Login = Nothing
End Sub

Public Function SetParam(objU8Login As Object)
    Set m_objU8Login = objU8Login
End Function


'得到文件名，不包含路径
Public Function GetFileName(ByVal strFileName As String) As String
    Dim strTemp() As String
    strTemp = Split(strFileName, "\")
    If UBound(strTemp) > 0 Then
        GetFileName = strTemp(UBound(strTemp))
    Else
        GetFileName = strFileName
    End If
End Function

'得到文件扩展名
Public Function GetFileExtension(ByVal strFileName As String) As String
    Dim pos As Long
    pos = InStr(1, strFileName, ".")
    If pos > 0 Then
        GetFileExtension = Right(strFileName, Len(strFileName) - pos)
    Else
        GetFileExtension = ""
    End If
End Function

'浏览文件
Public Function OpenFile(ByVal strLocalFileName As String, Optional ByVal strServerFileName As String = "") As Boolean
    OpenFile = False
    
    On Error GoTo lnErr
    '存在的本地的文件，直接显示；否则，下载到当前目录下显示
    Dim strDescFile As String
    Dim fs As New FileSystemObject
    
'    strDescFile = strLocalFileName
    strDescFile = IIf(Right(App.path, 1) = "\", App.path, App.path & "\") & strLocalFileName
    
    If fs.FileExists(strDescFile) Then
        ShellExecute GetDesktopWindow, "open", strDescFile, vbNullString, vbNullString, vbNormalFocus
    Else
        If Not SetFileServer Then GoTo lnErr
        m_objFile.ReadFile strServerFileName, strDescFile
        ShellExecute GetDesktopWindow, "open", strDescFile, vbNullString, vbNullString, vbNormalFocus
    End If
    OpenFile = True
    
lnExit:
    Exit Function
lnErr:
    MsgBox Err.Description
    GoTo lnExit
End Function


'保存到服务器文件，返回保存后的文件名
Public Function SaveFileToServer(ByVal strPathFileName As String) As String
    SaveFileToServer = ""
    
    On Error GoTo lnErr
    If strPathFileName = "" Then Exit Function
    If Not SetFileServer Then Exit Function
    SaveFileToServer = m_objFile.AddFile(strPathFileName, False)

lnExit:
    Exit Function
lnErr:
    MsgBox Err.Description
    GoTo lnExit
End Function

'从服务器删除文件
Public Function DeleteFileFromServer(ByVal strFileName As String) As Boolean
    DeleteFileFromServer = False
    On Error GoTo lnErr
    If strFileName = "" Then Exit Function
    If Not SetFileServer Then Exit Function
    DeleteFileFromServer = m_objFile.DeleteFile(strFileName)
    
lnExit:
    Exit Function
lnErr:
'    MsgBox Err.Description
    GoTo lnExit
End Function

Private Function SetFileServer() As Boolean
    On Error GoTo lnErr
    If m_objFile Is Nothing Then
        Set m_objFile = New U8FileManagerClient.U8FileManageClient
        Set m_objFile.LoginObject = m_objU8Login
    End If
    SetFileServer = True
    Exit Function
    
lnErr:
    '文件服务器错误,请设置文件服务器
    MsgBox "文件服务器错误,请设置文件服务器"
    'MsgBox GetResString("U8.CW.GL.CashFlow.00194") & vbCrLf & Err.Number & ":" & Err.Description, vbInformation, GetResString("U8.GL.ZZPZ.00003")
End Function
