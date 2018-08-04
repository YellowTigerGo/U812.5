VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUserInterfaceInit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim domAuthConfig As New DOMDocument
Public clsAuth As Object     'New U8RowAuthsvr.clsRowAuth
Dim m_DBConn As Object
Public Sub Init(DBConn As ADODB.Connection, strUserID As String, strFormType As String, strKey As String)
    If clsAuth Is Nothing Then
        Set clsAuth = CreateObject("U8RowAuthsvr.clsRowAuth")
        clsAuth.Init DBConn.ConnectionString, strUserID
    End If
    Dim rst As New ADODB.Recordset
    
    Set m_DBConn = DBConn
    rst.Open "select * from sa_authconfig where formtype=N'" + strFormType + "' and ckey=N'" + strKey + "' and bcontrol=1", DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rst.Save domAuthConfig, adPersistXML
    rst.Close
    Set rst = Nothing
End Sub

Private Sub Class_Terminate()
    Set domAuthConfig = Nothing
    Set clsAuth = Nothing
End Sub
Public Function GetAuthStringByOperation(strOperationType As String, Optional cardsection As String = "") As String
    Dim lst As IXMLDOMNodeList
    Dim nod As IXMLDOMNode
    Dim strAuth As String
    If cardsection = "" Then
        Set lst = domAuthConfig.selectNodes("//z:row[@operationtype='" + LCase(strOperationType) + "']")
    Else
        Set lst = domAuthConfig.selectNodes("//z:row[@operationtype='" + LCase(strOperationType) + "'and ( @cardsection='" + LCase(cardsection) + "' or @cardsection='" + UCase(cardsection) + "'  )]")
    End If
'    ElseIf LCase(cardsection) = LCase("B") Then
'        Set lst = domAuthConfig.selectNodes("//z:row[@operationtype='" + LCase(strOperationType) + "' @cardsection='" + LCase(cardsection) + "' ]")
'    End If
    For Each nod In lst
        If getOptionConfig(nod.Attributes.getNamedItem("authname").Text) = False Then GetAuthStringByOperation = "": Exit Function
        strAuth = clsAuth.GetAuthStringCode(nod.Attributes.getNamedItem("authenumname").Text, , nod.Attributes.getNamedItem("authtype").Text)
        If strAuth = "" Then
        ElseIf strAuth = "1=2" Then
            If Not nod.Attributes.getNamedItem("hasnotauthstring") Is Nothing Then
                GetAuthStringByOperation = IIf(GetAuthStringByOperation = "", "", GetAuthStringByOperation & " and ") & "(" & nod.Attributes.getNamedItem("hasnotauthstring").Text & ")"
            End If
        Else
            GetAuthStringByOperation = IIf(GetAuthStringByOperation = "", "", GetAuthStringByOperation & " and ") & "(" & ReplaceAuthMark(nod.Attributes.getNamedItem("hasauthstring").Text, strAuth) & ")"
        End If
    Next
    GetAuthStringByOperation = ConvertSQLString(GetAuthStringByOperation)
End Function

'获取选项配置是否启用
Private Function getOptionConfig(ByVal sConfigName As String) As Boolean
    Dim strSQL As String
    Dim Rs As Object
    strSQL = "SELECT cValue FROM AccInformation WHERE cSysID='EF' and cName='" & VBA.Replace(sConfigName, "'", "''") & "'"
    Set Rs = m_DBConn.Execute(strSQL)
    If Rs.EOF Then
        getOptionConfig = False
    Else
        If Val(Rs(0) & "") = 1 Then
            getOptionConfig = True
        Else
            getOptionConfig = False
        End If
    End If
    
    'by ahzzd 20100926 暂时取消参数选项
    getOptionConfig = True
    
    Rs.Close: Set Rs = Nothing: Exit Function
End Function

Private Function ReplaceAuthMark(strSAuth As String, strDAuth As String) As String
    ReplaceAuthMark = Replace(strSAuth, "[authstr]", strDAuth)
End Function
Public Sub GetEnumReferAuth(strReferName As String, strCardSection As String, strAuthName As String, strAuthType As String)
    Dim nod As IXMLDOMNode
    Set nod = domAuthConfig.selectSingleNode("//z:row[@operationtype='" + strReferName + "' and @cardsection='" + strCardSection + "']")
    If nod Is Nothing Then
        strAuthName = ""
        strAuthType = ""
    Else
        strAuthName = nod.Attributes.getNamedItem("hasauthstring").Text
        strAuthType = nod.Attributes.getNamedItem("authtype").Text
    End If
End Sub
Public Function getAuthString(strAuthName As String) As String
    getAuthString = clsAuth.GetAuthStringCode(strAuthName)
    getAuthString = ConvertSQLString(getAuthString)
End Function
Public Function IsHoldAuth(cBusObjId As String, cValue As String, Optional cClassCode As String, Optional cFuncId As String) As Boolean
    IsHoldAuth = clsAuth.IsHoldAuth(cBusObjId, cValue, cClassCode, cFuncId)
End Function



