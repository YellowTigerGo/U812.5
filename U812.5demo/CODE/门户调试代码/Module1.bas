Attribute VB_Name = "Module1"
Option Explicit
'×¢²á±íÏà¹Ø
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As Long, ByRef lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValueExW Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As Long, ByVal cbData As Long) As Long                                 ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Const CINSTLANGID = "SOFTWARE\UfSoft\WF\v8.700\"
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const KEY_QUERY_VALUE = &H1
Public Const REG_SZ = 1
Public Const REG_BINARY = 3                     ' Free form binary
Public Const ERROR_SUCCESS = 0&
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const KEY_SET_VALUE = &H2
Public Function RegRead(ByVal cSubKey As String, ByVal cItem As String, Optional ByVal nMainKey As Long = HKEY_LOCAL_MACHINE) As String
    RegRead = ""
    Dim hKey As Long
    If RegOpenKeyEx(nMainKey, cSubKey, 0, KEY_QUERY_VALUE, hKey) = ERROR_SUCCESS Then ' ¡ä¨°?a¡Á¡é2¨¢¡À¨ª?¨¹
        Dim cTemp As String
        Dim nTemp As Long
        Dim nType As Long
        nTemp = 256
        cTemp = Space(nTemp)
        nType = REG_SZ
        
        If RegQueryValueExW(hKey, StrPtr(cItem), 0, nType, StrPtr(cTemp), nTemp) = ERROR_SUCCESS Then       ' ??¦Ì?/¡ä¡ä?¡§?¨¹?¦Ì
            If nTemp = 0 Then
                RegRead = ""
            Else
                RegRead = Left(cTemp, InStr(1, cTemp, Chr(0)) - 1)
            End If
        End If
        RegCloseKey (hKey)
    End If
End Function
Public Sub RegWrite(ByVal cSubKey As String, ByVal cItem As String, ByVal cValue As String)
    Dim hKey As Long
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, cSubKey, 0, KEY_SET_VALUE, hKey) <> ERROR_SUCCESS Then ' ¡ä¨°?a¡Á¡é2¨¢¡À¨ª?¨¹
        If RegCreateKey(HKEY_LOCAL_MACHINE, cSubKey, hKey) Then
            Exit Sub
        End If
    End If
    Dim nType As Long
    nType = REG_SZ
    'RegSetValueEx hKey, cItem, 0, nType, ByVal cValue, LenB(cValue)           ' ??¦Ì?/¡ä¡ä?¡§?¨¹?¦Ì
    RegSetValueExW hKey, StrPtr(cItem), 0, nType, StrPtr(cValue), LenB(cValue)           ' ??¦Ì?/¡ä¡ä?¡§?¨¹?¦Ì
    RegCloseKey (hKey)                                 ' 1?¡À?¡Á¡é2¨¢¡À¨ª?¨¹
End Sub
 'RegWrite CINSTLANGID & "Portal", "Language", g_LangID
