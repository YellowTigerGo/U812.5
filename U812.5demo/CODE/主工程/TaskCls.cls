VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function TaskBegin(TaskNum As String, Optional cMsg, Optional ByVal bMsg As Boolean = True) As Boolean
'功能：功能，任务处理程序
'参数：功能，任务号
'返回：功能申请是否成功
On Error GoTo Err0
    Login1.ClearError
    If Login1.TaskExec(TaskNum, 1) Then
    Else
        TaskBegin = False
        'If bMsg Then Msg "您没有执行[" & AuthIDToName(TaskNum) & "]的权利！", vbInformation
        If bMsg Then Msg Login1.ShareString, vbInformation
        Login1.ClearError
        Exit Function
    End If
    If Login1.TaskExec(TaskNum, -1) Then
        TaskBegin = True
    Else
        TaskBegin = False
        If bMsg Then
            If IsMissing(cMsg) Then
                Msg Login1.ShareString & "", vbInformation
            Else
                cMsg = Login1.ShareString
            End If
        End If
    End If
    Exit Function
Err0:
    TaskBegin = False
End Function

Public Function TaskEnd(TaskNum As String) As Boolean
'功能：功能，任务处理程序
'参数：功能，任务号
'返回：
On Error GoTo Err0
    If Login1.TaskExec(TaskNum, 0) Then
        TaskEnd = True
    Else
        TaskEnd = False
    End If
    Login1.ClearError
    Exit Function
Err0:
    TaskEnd = False
End Function

