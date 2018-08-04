VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


''日志文件
Public Sub WrtDBlog(adoconn As Object, Optional ByVal loguser As String = "", Optional ByVal logType As String = "", Optional ByVal logtext As String = "")
    Dim strsql As String
    On Error Resume Next
    loguser = Left(loguser, 50)
    logType = Left(logType, 50)
    logtext = Replace(logtext, "'", "''")
    
    logtext = Left(logtext, 4000)
    strsql = "insert into ef_log ([loguser],[logType],[logtext],[logdatetime])" & vbCrLf & _
                "values('" & loguser & "','" & logType & "','" & logtext & "',GETDATE())"
    adoconn.Execute strsql
    
End Sub
