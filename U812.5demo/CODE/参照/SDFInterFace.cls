VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SDFInterFace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function RunCommand(ByVal objLogin As Object, ByVal objForm As Object, ByVal objVoucher As Object, _
                            ByVal sKey As String, ByVal VarentValue As Variant, ByVal other As String)
    
    Call ButtonClick(sKey, "", objForm, objVoucher)
  
End Function

'初始化
Public Function Init(ByVal objLogin As Object, ByVal objForm As Object, ByVal objVoucher As Object, msbar As Object)
'MsgBox "Init"
End Function

'在执行系统按钮只前调用,Cancel后会不再执行系统方法。
Public Function BeforeRunSysCommand(ByVal objLogin As Object, ByVal objForm As Object, ByVal objVoucher As Object, _
                            ByVal sKey As String, ByVal VarentValue As Variant, ByRef Cancel As Boolean, ByVal other As String)
    MsgBox "Save is Canceled"
    Cancel = True
End Function

Public Function ButtonClick(sButtonKey As String, sMenuKey As String, oForm As Object, octlVoucher As Object)

    Dim bSuccess As Boolean
    Dim sWhere As String
    Dim oDomHead As New DOMDocument
    Dim oDomBody As New DOMDocument
     
     
     
    If octlVoucher.headerText("bd_c") = True Then
        MsgBox GetString("U8.DZ.JA.Res2160"), vbInformation, GetString("U8.DZ.JA.Res030")
        Exit Function
    End If
    
    Set g_oLogin = octlVoucher.LoginObj
    
    
    octlVoucher.getVoucherDataXML oDomHead, oDomBody
     
    
'    Dim xoctlVoucher As Object
'
'    Set xoctlVoucher = CreateObject("UAPVoucherControl85.ctlVoucher")
'
'    xoctlVoucher.setVoucherDataXML oDomHead, oDomBody
   
   
    
    Set g_Conn = Nothing
    g_Conn.ConnectionString = g_oLogin.UfDbName
    g_Conn.CursorLocation = adUseClient
    g_Conn.Open

   ' GetMjInvDefine

    Select Case sButtonKey
       Case "btnpro", "_CUSTDEFbtnpro"
      Call getAuthString(g_Conn)
    iSinvCZ = True
'        If ReferVouchq Then
'            '调用添加按钮
'
'            Dim voucherfrom As Object
'
'            Call ProcessDatac(octlVoucher, oForm)
''            Call oForm.setVoucherStatesForHY
''            octlVoucher.row = octlVoucher.BodyRows
'    End If
    End Select
End Function

''参照单据
Public Function ReferVouchRD09f() As Boolean
    
    Dim Frm As New frmVouchRefers
    Dim clsReferVoucher As New clsReferVoucher
    
    ' 设置参照生单控件的属性
     clsReferVoucher.HelpID = "0"                                             '帮助 10151180
     clsReferVoucher.pageSize = 20                                                   '默认分页大小
     clsReferVoucher.strMainKey = "ID"                                               '主表唯一主键，作为和子表关联的依据
     clsReferVoucher.strDetailKey = "AutoID"                                         '子表唯一主键
     clsReferVoucher.FrmCaption = GetString("U8.SO.VOUCH.copytabst.00716")
     clsReferVoucher.FilterKey = "借出借用单运费参照" '"借出借用单参照"                                    '过滤器名称 SA26
     clsReferVoucher.FilterSubID = "ST"
     
     
     clsReferVoucher.HeadKey = "HY_JCGH00101"       '主表的列信息 AA_ColumnDic
     clsReferVoucher.BodyKey = "HY_JCGH00102"       '子表的列信息 ,若设置只有表头时该属性置空
     '添加自定义按钮
     'clsReferVoucher.Buttons = "<root><row type='toolbar' name='tlbTest' caption='自动匹配' index='26' /></root>"
     '多语时
     'strButtons ="<root><row type='toolbar' name='tlbTest' caption='{res:U8.SA.xsglsql_2.frmvouchrefers.automatching}' index='26' /></root>"
     'strButtons =ReplaceResId("<root><row type='toolbar' name='tlbTest' caption='{res:U8.SA.xsglsql_2.frmvouchrefers.automatching}' index='26' /></root>")
      clsReferVoucher.MainDataSource = "V_HY_DZ_BorrowOutSDRD09F"                           '主表数据源视图
      clsReferVoucher.DetailDataSource = "V_HY_DZ_BorrowOutsSDRD09F"                         '子表数据源试图
      clsReferVoucher.DefaultFilter = "" '"IsNULL(cVerifier,'')<>'' and IsNULL(cCloser,'')=''And IsNull(cSCloser,'')=''"           '默认过滤条件
      
    clsReferVoucher.strAuth = Replace(sMakeAuth_ALL, "HY_DZ_BorrowOut.id", "V_HY_DZ_BorrowOutSDRD09F.id") '数据权限SQL串, 这里不支持仓库权限控制
      
      
      clsReferVoucher.OtherFilter = ""                                               '其他过滤条件
      
      clsReferVoucher.HeadEnabled = False                                            '主表是否可编辑
      clsReferVoucher.BodyEnabled = False                                            '子表是否可编辑
      
      'clsReferVoucher.bSelectSingle = True                                           '表头是否只能取唯一记录
      
      clsReferVoucher.bSelectSingle = False                                           '表头是否只能取唯一记录
      clsReferVoucher.strCheckFlds = "cCODE" '"cType,bObjectCode" '
      clsReferVoucher.strCheckMsg = GetString("U8.DZ.JA.Res790")


    
    Set Frm.clsReferVoucher = clsReferVoucher
    If Frm.OpenFilter Then
        Frm.Show vbModal
    Else
        Frm.bcancel = True
    End If
    Set clsReferVoucher = Nothing
    

    If Not Frm.bcancel Then
        ReferVouchRD09f = True
        'Set Domhead = frmVouchRef.Domhead
        'Set Dombody = frmVouchRef.Dombody
        Set gDomReferHead = Frm.domHead
        Set gDomReferBody = Frm.domBody
    End If
    
    Unload Frm
    Set Frm = Nothing

End Function








