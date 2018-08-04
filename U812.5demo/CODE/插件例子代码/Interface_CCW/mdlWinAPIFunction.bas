Attribute VB_Name = "mdlWinAPIFunction"
Declare Function GetPrivateProfileString Lib "kernel32" Alias _
"GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, _
ByVal lpDefault As String, _
ByVal lpReturnedString As String, _
ByVal nSize As Long, _
ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib "kernel32" Alias _
"WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'Declare Function ShellExecute Lib "shell32.dll" Alias _
'"ShellExecuteA" ( byval hwnd as Long ,_
'ref lpOperation as String,ref string lpFile,ref string lpParameters,ref string lpDirectory,ulong nShowCmd))

Public Function GetIniParam(NomFichier As String, NomSection As String, NomVariable As String) As String
    Dim ReadString As String * 255
    Dim returnv As String
    Dim mResultLen As Integer
    
    mResultLen = GetPrivateProfileString(NomSection, NomVariable, "(Unassigned)", ReadString, Len(ReadString) - 1, NomFichier)
    
    If IsNull(ReadString) Or Left$(ReadString, 12) = "(Unassigned)" Then
        Dim Tempvalue As Variant
        Dim Message As String
        Message = "配置文件 " & NomFichier & " 不存在."
        returnv = ""
    Else
        returnv = Left$(ReadString, InStr(ReadString, Chr$(0)) - 1)
    End If
    
    GetIniParam = returnv
End Function

Public Function WriteWinIniParam(NomDuIni As String, sLaSection As String, sNouvelleCle As String, sNouvelleValeur As String)
    Dim iSucccess As Integer
    
    iSucccess = WritePrivateProfileString(sLaSection, sNouvelleCle, sNouvelleValeur, NomDuIni)
    
    If iSucccess = 0 Then
        WriteWinIniParam = False
    Else
        WriteWinIniParam = True
    End If
End Function


