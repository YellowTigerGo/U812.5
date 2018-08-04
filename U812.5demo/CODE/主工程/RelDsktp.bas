Attribute VB_Name = "RelativeDesktop"
Option Explicit
Public Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long           '1 = Windows 95. 2 = Windows NT
   szCSDVersion As String * 128
End Type
Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Function RegRead(ByVal cSubKey As String, ByVal cItem As String) As String
    RegRead = ""
    Dim hKey As Long
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, cSubKey, 0, KEY_QUERY_VALUE, hKey) = ERROR_SUCCESS Then ' 打开注册表键
        Dim cTemp As String * 128
        Dim nTemp As Long
        Dim nType As Long
        nType = REG_SZ
        nTemp = 128
        If RegQueryValueEx(hKey, cItem, 0, nType, ByVal cTemp, nTemp) = ERROR_SUCCESS Then       ' 获得/创建键值
            RegRead = Left(cTemp, InStr(1, cTemp, Chr(0)) - 1)
        End If
        RegCloseKey (hKey)                                 ' 关闭注册表键
    End If
End Function

Public Function IsWindow9X() As Boolean
   Dim osi As OSVERSIONINFO
   osi.dwOSVersionInfoSize = Len(osi)
   GetVersionExA osi
   IsWindow9X = osi.dwPlatformId = 1
End Function

'判断如果在9X上运行是否有足够的资源
Public Function HaveSufficeResources() As Boolean
    If IsWindow9X() Then
        Dim oSR As Object
        Set oSR = CreateObject("prjSR.clsSR")
        HaveSufficeResources = oSR.SystemResources > 15
        Set oSR = Nothing
        If Not HaveSufficeResources Then MsgBox GetString("U8.SA.xsglsql.reldsktp.00134"), vbExclamation 'zh-CN：系统资源不足，请退出不用的程序或功能，再试！
    Else
        HaveSufficeResources = True
    End If
End Function
