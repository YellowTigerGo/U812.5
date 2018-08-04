Attribute VB_Name = "ModEncrypt"

Option Explicit

Private Const BASE64CHR As String = "AbclmndefghiBCD23+EFGH/45IJPQRSvwxTUV6789WXYZajkoK01LMNOpqrstuyz="
Private Const mCODE As String = "C79D82IJKNQRLABMSFGH1TU3Z04VW6XYO5EP"
Private mvarCodeString(0 To 34)
Private psBase64Chr(0 To 63) As String
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal numBytes As Long)

Private Type SYSTEM_INFO
    wProcessorArchitecture As Integer
    wReserved As Integer
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type

Private Declare Function apiGetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (lpRootPathName As String, lpVolumeNameBuffer As Any, ByVal nVolumeNameSize As Long, ByRef lpVolumeSerialNumber As Long, lpMaximumComponentLength As Any, lpFileSystemFlags As Any, lpFileSystemNameBuffer As Any, ByVal nFileSystemNameSize As Long) As Boolean
Private Declare Sub apiGetSystemInfo Lib "kernel32" Alias "GetSystemInfo" (lpSystemInfo As SYSTEM_INFO)


'转换字符串为BASE64字符串，转换3个8比特(0 to 255)的字符成为4个6比特(0 to 63),不足以=填充
Public Function ZipBase64String(str2Encode As String) As String
    Dim lCtr As Long
    Dim lPtr As Long
    Dim lLen As Long
    Dim sValue() As Byte
    Dim sEncoded As String
    Dim Bits8(1 To 3) As Byte
    Dim Bits6(1 To 4) As Byte
    Dim i As Integer
    
    Call PasswordInitialize
    
    '获得3个字符
    sValue = StrConv(str2Encode, vbFromUnicode)
    For lCtr = 1 To UBound(sValue) + 1 Step 3
    
    For i = 1 To 3
    If lCtr + i - 2 <= UBound(sValue) Then
    Bits8(i) = sValue(lCtr + i - 2)
    lLen = 3
    Else
    Bits8(i) = 0
    lLen = lLen - 1
    End If
    Next
    
    '转换字符串为数组，然后转换为4个6位(0-63)
    Bits6(1) = (Bits8(1) And &HFC) \ 4
    Bits6(2) = (Bits8(1) And &H3) * &H10 + (Bits8(2) And &HF0) \ &H10
    Bits6(3) = (Bits8(2) And &HF) * 4 + (Bits8(3) And &HC0) \ &H40
    Bits6(4) = Bits8(3) And &H3F
    
    '添加4个新字符
    For lPtr = 1 To lLen + 1
    sEncoded = sEncoded & psBase64Chr(Bits6(lPtr))
    Next
    Next
    
    '不足4位，以=填充
    Select Case lLen + 1
    Case 2: sEncoded = sEncoded & "=="
    Case 3: sEncoded = sEncoded & "="
    Case 4:
    End Select
    ZipBase64String = sEncoded
End Function


'解压BASE64字符串
Public Function UnZipBase64String(str2Decode As String) As String
    Dim lPtr As Long
    Dim iValue As Integer
    Dim iLen As Integer
    Dim iCtr As Integer
    Dim Bits(1 To 4) As Byte
    Dim strDecode As String
    Dim str As String
    Dim Output() As Byte
    Dim iIndex As Long
    Dim lFrom As Long
    Dim lTo As Long
    
    Call PasswordInitialize
    
    '除去回车
    str = Replace(str2Decode, vbCrLf, "")
    
    '每4个字符一组（4个字符表示3个字）
    For lPtr = 1 To Len(str) Step 4
    iLen = 4
    For iCtr = 0 To 3
    
    '找字符在BASE64字符串中的位置
    iValue = InStr(1, BASE64CHR, Mid$(str, lPtr + iCtr, 1), vbBinaryCompare)
    Select Case iValue 'A~Za~z0~9+/
    Case 1 To 64:
    Bits(iCtr + 1) = iValue - 1
    Case 65 '=
    iLen = iCtr
    Exit For
    
    '没有发现
    Case 0: Exit Function
    End Select
    Next
    
    '转换4个6比特数成为3个8比特数
    Bits(1) = Bits(1) * &H4 + (Bits(2) And &H30) \ &H10
    Bits(2) = (Bits(2) And &HF) * &H10 + (Bits(3) And &H3C) \ &H4
    Bits(3) = (Bits(3) And &H3) * &H40 + Bits(4)
    
    '计算数组的起始位置
    lFrom = lTo
    lTo = lTo + (iLen - 1) - 1
    
    '重新定义输出数组
    ReDim Preserve Output(0 To lTo)
    
    For iIndex = lFrom To lTo
    Output(iIndex) = Bits(iIndex - lFrom + 1)
    Next
    lTo = lTo + 1
    
    Next
    
    UnZipBase64String = StrConv(Output, vbUnicode)
End Function
'初始化
Public Sub PasswordInitialize()
    Dim iPtr As Integer
    '初始化 BASE64数组
    For iPtr = 0 To 63
    psBase64Chr(iPtr) = Mid$(BASE64CHR, iPtr + 1, 1)
    Next
End Sub

'初始化
Public Sub CodePasswordInitialize()
    Dim intI As Integer
    For intI = 0 To 34
        mvarCodeString(intI) = Mid(mCODE, intI + 1, 1)
    Next
End Sub

