VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------
' 文 件 名: clsWords.cls
'
' 功 能: Word文档操作
'
' 作 者: 王良彬    时  间:     2002
' 北就用友软件集团 版权所有 Copyright(c) 2002
'--------------------------------------------


Public cTableName               As String
Public FullName                 As String               '当前编辑文档路径
Public imode                    As Integer
Public iOpenMode                As OpenState            '打开方式
Private bReadOnly               As Boolean              '是否只读
Private cVouchType              As String               '单据编号
Private cComputerName           As String               '计算机名
Private strFileTmp              As String               '临时文件名
Public WithEvents clsWd         As Word.Application     '
Attribute clsWd.VB_VarHelpID = -1

Enum OpenState
    NewOpen = 0         '新建方式
    EditOpen = 1        '修改方式
    QueryOpen = 2       '查询方式
End Enum


'退出事件
Public Event Quit()
'保存前触发事件
Public Event SaveBefore()

'<api声明
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'/api声明>



'-----------------------------------------------------------
'功能：把用二进制打开的文件strfilename 存放到数据库中的一个字段
'
'参数：strFileName要保存到数据库的文档路径， rsWrite 已打开的要保存文档的recordset， strField 要保存到的字段
'
'返回：保存是否成功
'
'-----------------------------------------------------------
Public Function funFileToDB(ByVal strFileName As String, rsWrite As Object, ByVal strField As String) As Boolean
  Dim lFilelen      As Long         '文件长度
  Dim Offset        As Long         '位移
  Dim ChunkSize     As Long         '一次保存的块大小
  Dim TheBytes()    As Byte         '保存到数据库的临时动态数组
  
  '判断文件、字段名、记录集是否合法
  If strFileName = "" Or strField = "" Or rsWrite.State = 0 Then
    funFileToDB = False
    Exit Function
  End If
  
  On Error GoTo m_eRr
  lFilelen = FileLen(strFileName)                   '取文件长度
  ChunkSize = 1000                                  '设置固定块大小
  If ChunkSize > lFilelen Then ChunkSize = lFilelen
'  rsWrite.AddNew
  Offset = ChunkSize
  ReDim TheBytes(0 To ChunkSize - 1)
  Open strFileName For Binary Access Read As #2       'get source file len
  Do While Offset < lFilelen
    Get #2, , TheBytes                              '读文件内容
    rsWrite(strField).AppendChunk TheBytes          '写到数据库里
    If Offset + ChunkSize > lFilelen Then ChunkSize = lFilelen - Offset
    Offset = Offset + ChunkSize                     '重新计算位移
  Loop
  If ChunkSize <> 0 Then                            '如果 ChunkSize <> 0 则，做最后一次保存
    ReDim TheBytes(0 To ChunkSize - 1)
    Get #2, , TheBytes
    rsWrite(strField).AppendChunk TheBytes
  End If
  Close #2
'  rsWrite.UpdateBatch
  funFileToDB = True
  Exit Function
m_eRr:
  On Error Resume Next
  Close #2
  funFileToDB = False
  Err.Clear
End Function


'-----------------------------------------------------------
'功能：把数据库中的数据存放到一个文件里
'
'参数：strFileName存放数据库数据的文档路径， rsWrite 已打开的recordset， strField 数据库字段
'
'返回：保存是否成功
'
'-----------------------------------------------------------
Public Function funDBToFile(ByVal strFileName As String, rsRead As Object, ByVal strField As String) As Boolean
  
  Dim lTotalSize    As Long         '字段里存放数据大小
  Dim Offset        As Long         '位移
  Dim ChunkSize     As Long         '一次读取的块大小
  Dim TheBytes()    As Byte         '保存到临时数据的动态数组
  
  '判断文件、字段名、记录集是否合法
  If strFileName = "" Or strField = "" Or rsRead.State = 0 Then
    funDBToFile = False
    Exit Function
  End If
  On Error GoTo m_eRr
  'If Dir(strFileName) <> "" Then Kill strFileName
  ChunkSize = 1000                              '
  lTotalSize = rsRead(strField).ActualSize      '读取strField字段存放数据的大小
  If ChunkSize > lTotalSize Then ChunkSize = lTotalSize
  If ChunkSize - 1 <= 0 Then
    funDBToFile = False
    Exit Function
  End If
  ReDim thebyte(0 To ChunkSize - 1)                         '冲定义动态数组
  Offset = 0
  Open strFileName For Binary Access Write As #2            'get source file len
  Do While Offset < lTotalSize
    If Offset + ChunkSize > lTotalSize Then                 '是否已经读取完毕
      ChunkSize = lTotalSize - Offset                       '计算最后一次读取块大小
      ReDim TheBytes(0 To ChunkSize - 1)
      TheBytes = rsRead(strField).GetChunk(ChunkSize)       '读取数据到临时数组
      Put #2, , TheBytes                                    '写文件
    Else
      TheBytes = rsRead(strField).GetChunk(ChunkSize)       '读取数据到临时数组
      Put #2, , TheBytes                                    '写文件
    End If
    Offset = Offset + ChunkSize                             '重新计算位移
  Loop
  Close #2
  funDBToFile = True
  Exit Function
m_eRr:
  On Error Resume Next
  Close #2
  funDBToFile = False
  Err.Clear
End Function


'-----------------------------------------------------------
'功能：打开文件cFieldName，如果cFieldName=''则新建文件
'
'参数：sVouchType单据编号，cFielName要打开的文件路径
'
'返回：打开文件路径
'
'-----------------------------------------------------------
Public Function NewDocument(ByVal sVouchType As String, ByVal cFileName As String) As String
    Dim strMsg  As String               '提示信息
    Dim cFile   As String               '存放文件路径的临时变量
    Dim lngDlg  As Long
    Dim lErr    As Long                 '错误类型，1=无此文件 2=文件创建失败
    On Error GoTo errHandle
    
    strFileTmp = cFileName
    
rOpen:
    strMsg = "请检测是否安装了Word软件！"
    cVouchType = sVouchType
    cFile = "无此文件！"
    lErr = 1
    If Not clsWd Is Nothing Then Set clsWd = Nothing
    Set clsWd = CreateObject("Word.Application")
    lErr = 2
    strMsg = ""
    
    If cFileName = "" Then                  '新建文档打开
        cFile = DotFile(sVouchType)         '去模版文件路径
        If Dir(cFile) = "" Then
            clsWd.Documents.Add             '增加临时文件
        Else
            clsWd.Documents.Open cFile, , True      '打开模版文件
        End If
        clsWd.Visible = True
        FullName = ""
    Else                                    '打开编辑文件
        cFile = ""
        FullName = ""
        
        clsWd.Documents.Open cFileName, , bReadOnly, , , , , "gsp"
        
        clsWd.Visible = True
        FullName = cFileName
    End If
    
    NewDocument = FullName
    imode = 1
            
    Exit Function
errHandle:
    NewDocument = CStr(lErr)
    If Not clsWd Is Nothing Then clsWd.Quit
    If strMsg = "" Then
        MsgBox Err.Description, vbCritical, Msg_Title
    Else
        MsgBox strMsg, vbCritical, Msg_Title
        Exit Function
    End If
    FullName = ""
    cFileName = ""
    If cFile = "" Then
        If bReadOnly = False Then GoTo rOpen
    End If
End Function

'-----------------------------------------------------------
'功能：获得模版文件路径
'
'参数：单据模版号
'
'返回：模版文件路径
'
'-----------------------------------------------------------
Private Function DotFile(sKey As String) As String
    Select Case sKey
            Case "079"
                DotFile = App.path & "\Doc\" & "管理制度档案.doc"
            Case "080"
                DotFile = App.path & "\Doc\" & "标准分类档案.doc"
    End Select
End Function


'-----------------------------------------------------------
'功能：关闭Word文档触发事件
'
'参数：Doc 当前编辑文档，Cancel是否取消退出
'
'返回：
'
'-----------------------------------------------------------
Private Sub clsWd_DocumentBeforeClose(ByVal Doc As Word.Document, Cancel As Boolean)
    If bReadOnly = True Then clsWd.ActiveDocument.Saved = True
End Sub


'-----------------------------------------------------------
'功能：保存Word文档触发事件
'
'参数：Doc 当前编辑文档，Cancel是否取消保存，SaveAsUI是否显示另存为界面
'
'返回：
'
'-----------------------------------------------------------
Private Sub clsWd_DocumentBeforeSave(ByVal Doc As Word.Document, SaveAsUI As Boolean, Cancel As Boolean)
   
    Dim fs
    
    Cancel = True                                   '取消系统保存，因为不好判断当前临时文件路径
    If bReadOnly Then Exit Sub
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    strFileTmp = App.path + "\档案" & cVouchType & ".doc"
    On Error Resume Next
    Kill strFileTmp
    fs.CreateTextFile strFileTmp                    '创建一个文件，保存当前Word文档
    
    Doc.SaveAs strFileTmp
    FullName = strFileTmp
    
    Set fs = Nothing
    
'    If Trim(Doc.Path) = "" Or Doc.FullName = DotFile(cVouchType) Then
'        Cancel = True
'        If clsWd.Dialogs(wdDialogFileSaveAs).Show() <> 0 Then
'            FullName = FullPath(clsWd.ActiveDocument.FullName)
'        End If
'    Else
'        FullName = FullPath(Doc.FullName)
'    End If
    RaiseEvent SaveBefore
End Sub

'-----------------------------------------------------------
'功能：退出Word文档触发事件
'
'参数：
'
'返回：
'
'-----------------------------------------------------------
Private Sub clsWd_Quit()
    If imode <> 2 Then imode = 3
    RaiseEvent Quit
End Sub

'-----------------------------------------------------------
'功能：设置是否只读属性
'
'参数：V 是否只读
'
'返回：
'
'-----------------------------------------------------------
Public Property Let ReadOnly(V As Boolean)
    bReadOnly = V
End Property


'-----------------------------------------------------------
'功能：读取数据，保存在默认文件里
'
'参数：strVchType 单据编号， strID 档案编号， Opt
'
'返回：
'
'-----------------------------------------------------------
Public Function GetFile(strVchType As String, strID As String, ByVal Opt As Long) As String
    '************************************************************************
    Dim rs      As ADODB.Recordset
    Dim sSql    As String
    Dim strTmp  As String
    Dim fs, f
    
    GetFile = ""
    '默认临时文件路径
    strTmp = App.path + "\档案" & strVchType & ".DOC"
    If Dir(strTmp) = "" Then
        '不存在，则创建临时文件，并从数据库里读取数据
        Set fs = CreateObject("Scripting.FileSystemObject")
        On Error Resume Next
        Kill strTmp
        fs.CreateTextFile strTmp           '创建一个文件
        
        Set fs = Nothing
        
        Set rs = New ADODB.Recordset
        
        sSql = "select  * from gsp_archive where cvouchtype='" & strVchType & "' and ccode='" & strID & "'"
        rs.Open sSql, AdoCnn, adOpenStatic, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            If Not funDBToFile(strTmp, rs, "CFILE") Then
'                If Opt <> 0 Then Kill strTmp
                Kill strTmp
            Else
                GetFile = strTmp
            End If
        End If
    Else
        GetFile = strTmp
    End If
    Set rs = Nothing
End Function

'-----------------------------------------------------------
'功能：关闭Word文档
'
'参数：
'
'返回：
'
'-----------------------------------------------------------
Public Sub WordClose()
    On Error Resume Next
    clsWd.Documents.Close wdDoNotSaveChanges
End Sub
 


