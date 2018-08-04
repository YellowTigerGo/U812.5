VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReportCallBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Public StrReportName As String

Public Sub CallBack_ValueChange(sKey As String, flt As UFGeneralFilter.FilterSrv)
    Dim sValue As String
    Dim strsql As String
    Dim strID As String
'    MsgBox StrReportName
    Select Case StrReportName
        Case "销售现存量查询"
            If LCase(sKey) = "isotype" Then
                strID = flt.FilterList.Item("isotype").varValue
                If strID = "" Then
                    strsql = ""
                Else
                    If strID = "2" Or strID = "3" Then
                        strID = "3"
                    End If
                    If strID = "1" Or strID = "5" Then
                        strID = "1"
                    End If
                    strsql = "idemandtype=" & strID
                End If
                flt.SetEditFilter "csocode", "", strsql
            End If
        Case "01", "07", "05", "02"
            If LCase(sKey) = "idemandtype" Then
                strID = flt.FilterList.Item("idemandtype").varValue
                If strID = "" Then
                    strsql = ""
                Else
                    If strID = "2" Or strID = "3" Then
                        strID = "3"
                    End If
                    If strID = "1" Or strID = "5" Then
                        strID = "1"
                    End If
                    strsql = "idemandtype=" & strID
                End If
                flt.SetEditFilter "cdemandcode", "", strsql
            End If
    End Select
End Sub
'当过滤初始化是回调此函数
Public Function CallBack_FilterInit(objFilter As UFGeneralFilter.FilterSrv)
If m_login Is Nothing Then Exit Function
    If UCase(m_login.cSub_Id) <> "SA" Then Exit Function
    'Dim objflt As UfHeronReport.clsFilter
    Dim objflt As UFGeneralFilter.FilterItem
    Dim tmp As String
    Dim i As Integer
    Dim ele As IXMLDOMElement
    Dim lst As IXMLDOMNodeList
    StrReportName = Replace(objFilter.FilterID, "SA[__]", "")
    
    ''''''''''''''''''''''''''''''''''''''''''''''''清空报表上次的过滤
'    Select Case StrReportName
'    Case "进销存统计表", "经营状况分析", "商品周转率分析", "市场分析", "委托代销明细账", "委托代销统计表", "销售成本明细账", "销售订单预警和报警", _
'         "销售计划执行报告", "销售结构分析", "销售毛利分析", "销售明细账", "销售收入明细账", "销售统计表", "销售现存量查询", "销售增长分析", _
'         "销售综合统计表", "信用余额表", "订单追溯", "畅适滞分析", "订单执行统计表", "动销分析", "发货单开票收款勾对表", "发货结算勾对表", _
'         "发货统计表", "发票使用明细表", "货龄分析（按发票）", "货龄分析（按收款）", "货物流向分析"
'         For I = 1 To objFilter.FilterList.count
'            If objFilter.FilterList(I).strSqlString <> "" Then objFilter.FilterList(I).strSqlString = ""
'         Next
'    End Select
    
    ''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    
    
    If Not domFilter Is Nothing Then
        Set lst = domFilter.documentElement.selectNodes("row[@type='refer']")
        For Each ele In lst
            objFilter.SetEditFilter ele.Attributes.getNamedItem("itemname").nodeValue, "", ele.Attributes.getNamedItem("expression").nodeValue
        Next
'        Set lst = domFilter.documentElement.selectNodes("row[@type='sql']")
'        If lst.length > 0 Then
'            For I = 1 To objFilter.FilterList.count
'                Set objFlt = objFilter.FilterList.item(I)
'                Set ele = domFilter.documentElement.selectSingleNode("row[@type='sql' and @itemname='" + LCase(objFlt.Name) + "']")
'                If Not ele Is Nothing Then
'                    objFlt.Expression = ele.Attributes.getNamedItem("expression").nodeValue
'                End If
'            Next
'        End If
        Exit Function
    End If
    Select Case StrReportName
        Case "SA26"
            If VoucherRefAgain = True And Not VoucherRefDomH Is Nothing Then '客户
                tmp = GetHeadItemValue(VoucherRefDomH, "ccuscode")
                objFilter.FilterList.Item("cCusCode").varValue = tmp
                If Not CBool(objFilter.FilterList.Item("cCusCode").IsSingle) Then
                    objFilter.FilterList.Item("cCusCode").varValue2 = tmp
'                    Call objFilter.DisableCondition("cCusCode__2", True)
                End If
                Call objFilter.DisableCondition("cCusCode", True)
                
                tmp = GetHeadItemValue(VoucherRefDomH, "ccusabbname")
                objFilter.FilterList.Item("ccusabbname").varValue = tmp
                If Not CBool(objFilter.FilterList.Item("ccusabbname").IsSingle) Then
                    objFilter.FilterList.Item("ccusabbname").varValue2 = tmp
'                    Call objFilter.DisableCondition("cCusCode__2", True)
                End If
                Call objFilter.DisableCondition("ccusabbname", True)
'                Set VoucherRefDomH = Nothing
'                If objflt.Name = "cCusCode" Then
'                    objflt.varValue = tmp
'                    If LCase(objflt.IsSingle) = "false" Or LCase(objflt.IsSingle) = "0" Then
'                        objflt.varValue2 = tmp
'                    End If
'                    Call objFilter.DisableCondition(objflt.Name, True)
'                End If
            End If
    Case "SARefVouchT5", "SARefVouchT8", "SARefVouchT2", "SARefVouchT1", "SARefVouchT4", "SARefVouchT7"
        
        '参照生单 过滤条件预制
        For i = 1 To objFilter.FilterList.Count
            Set objflt = objFilter.FilterList.Item(i)  '编辑状态 带入表头的 信息
            
            If InStr(1, objflt.Name, "cBustype") > 0 Then
                objflt.Expression = refCbustype
                
'                If objflt.varValue = "" Then
                    If InStr(1, refCbustype, ",") > 0 Then
                        objflt.varValue = Left(refCbustype, InStr(1, refCbustype, ",") - 1)
                    Else
                        objflt.varValue = refCbustype
                    End If
'                Else
'                    If InStr(1, refCbustype, objflt.varValue) = 0 Then
'                        If InStr(1, refCbustype, ",") > 0 Then
'                            objflt.varValue = Left(refCbustype, InStr(1, refCbustype, ",") - 1)
'                        Else
'                            objflt.varValue = refCbustype
'                        End If
'                    End If
'                End If
            End If
            If objflt.Name = "DispatchListType" Then
                If objflt.varValue = "" Then 'SA.BlueRed
                    objflt.varValue = "蓝字记录"
                End If
            End If
            If objflt.Name = "BillType" Then
                If objflt.varValue = "" Then
                    objflt.varValue = "否"
                End If
            End If
            If objflt.Name = "isShowOverOrder" Then
                If objflt.varValue = "" Then
                    objflt.varValue = "否"
                End If
            End If
            If VoucherRefAgain = True Then '客户
                If objflt.Name = "cCusCode" And StrReportName <> "SARefVouchT1" And StrReportName <> "SARefVouchT7" And StrReportName <> "SARefVouchT5" Then
                    tmp = GetHeadItemValue(VoucherRefDomH, "ccuscode")
                    objflt.varValue = tmp
                    If LCase(objflt.IsSingle) = "false" Or LCase(objflt.IsSingle) = "0" Then
                        objflt.varValue2 = tmp
                    End If
                    Call objFilter.DisableCondition(objflt.Name, True)
                    
                '部门
                ElseIf objflt.Name = "cDepCode" Then
                    tmp = GetHeadItemValue(VoucherRefDomH, "cDepCode")
                    objflt.varValue = tmp
                    If LCase(objflt.IsSingle) = "false" Or LCase(objflt.IsSingle) = "0" Then
                        objflt.varValue2 = tmp
                    End If
                    

                 '业务员
                ElseIf objflt.Name = "cPersonCode" Then
                    tmp = GetHeadItemValue(VoucherRefDomH, "cPersonCode")
                    objflt.varValue = tmp
                    If LCase(objflt.IsSingle) = "false" Or LCase(objflt.IsSingle) = "0" Then
                        objflt.varValue2 = tmp
                    End If
                   

                 '税率
                ElseIf InStr(1, objflt.Name, "iTaxRate") > 0 Then
                    tmp = GetHeadItemValue(VoucherRefDomH, "iTaxRate")
                    objflt.varValue = tmp
                    If LCase(objflt.IsSingle) = "false" Or LCase(objflt.IsSingle) = "0" Then
                        objflt.varValue2 = tmp
                    End If
                   Call objFilter.DisableCondition(objflt.Name, True)
                
                '业务类型
                ElseIf InStr(1, objflt.Name, "cBustype") > 0 Then
                    tmp = GetHeadItemValue(VoucherRefDomH, "cBustype") '委托代销结算单 没有业务类型
                    If tmp = "" And StrReportName = "SARefVouchT7" Then tmp = "委托代销"
                    objflt.varValue = tmp
                    If LCase(objflt.IsSingle) = "false" Or LCase(objflt.IsSingle) = "0" Then
                        objflt.varValue2 = tmp
                    End If
                    Call objFilter.DisableCondition(objflt.Name, True)
                ElseIf LCase(objflt.Name) = "cinvoicecompany" And (StrReportName = "SARefVouchT1" Or StrReportName = "SARefVouchT7" Or StrReportName = "SARefVouchT5") Then
                    tmp = GetHeadItemValue(VoucherRefDomH, "ccuscode")
                    objflt.varValue = tmp
                    If LCase(objflt.IsSingle) = "false" Or LCase(objflt.IsSingle) = "0" Then
                        objflt.varValue2 = tmp
                    End If
                    Call objFilter.DisableCondition(objflt.Name, True)
                End If
            End If

        Next
    Case "订单追溯"
         For i = 1 To objFilter.FilterList.Count
            Set objflt = objFilter.FilterList.Item(i)
            If objflt.Name = "HaveMout" Then
                If objflt.varValue = "" Then
                    objflt.varValue = "否"
                End If
            End If
        Next
    Case "货龄分析（按发票）", "货龄分析（按收款）"
        For i = 1 To objFilter.FilterList.Count
            Set objflt = objFilter.FilterList.Item(i)
            If objflt.Name = "cVouchDateType" Then
                If objflt.varValue = "" Then
                    objflt.varValue = "单据日期"
                End If
            End If
'            If objflt.Name = "设置账龄区间" And GetAp_billAge("SA") Then
'                objflt.varValue = GetString("U8.SA.xsglsql.frmage.01094")
'            End If
        Next
    Case "07fwsk"
         For i = 1 To objFilter.FilterList.Count
            Set objflt = objFilter.FilterList.Item(i)
            If objflt.Name = "cvouchtype" Then
                If objflt.varValue = "" Then
                    objflt.varValue = "26"
                End If
            End If
        Next
     Case "销售订单预警和报警"
          Dim strPreDay As String
          Dim strOverDay As String
          strPreDay = getAccinformation("SA", "iPreDay")
          strOverDay = getAccinformation("SA", "iOverDay")
          For i = 1 To objFilter.FilterList.Count
              
            If objFilter.FilterList(i).Name = "iPreDay" Then
                  '预期天数
                  objFilter.FilterList(i).varValue = strPreDay
            End If
            '逾期天数
            If objFilter.FilterList(i).Name = "iOverDay" Then
                objFilter.FilterList(i).varValue = strOverDay
            End If
            If objFilter.FilterList(i).Name = "dCurDate" Then
                '登陆日期
                objFilter.FilterList(i).varValue = m_login.CurDate
            End If
         Next
    Case "signinreferdispatch"
         For i = 1 To objFilter.FilterList.Count
            Set objflt = objFilter.FilterList.Item(i)
            If objflt.Name = "IsShowExcuted" Then
                If objflt.varValue = "" Then
                    objflt.varValue = "否"
                End If
            End If
        Next
    End Select
    

End Function

'当点过滤时回调此函数

Public Function CallBack_FilterClick(objFilter As UFGeneralFilter.FilterSrv, Cancel As Boolean)
      '当过滤条件界面过滤按纽按下时激活该事件
      'StrReportName  报表名称
      'ObjFilter  过滤条件名称
      'cancel  true  取消操作 false 继续'
      '
      Dim i As Long
      If m_login Is Nothing Then Exit Function

      If UCase(m_login.cSub_Id) <> "SA" Then Exit Function
      StrReportName = Replace(objFilter.FilterID, "SA[__]", "")
      Select Case StrReportName
        
      
      Case "货龄分析（按发票）"
      
            For i = 1 To objFilter.FilterList.Count
            
                  If objFilter.FilterList(i).Name = "ddate" Then
                  
                        If Trim(objFilter.FilterList(i).varValue) = vbNullString Then
                              Beep
                              MsgBox GetString("U8.SA.xsglsql.clsreportcallback.00095"), vbOKOnly + vbExclamation, StrReportName 'zh-CN：日期必须输入
                              Cancel = True
                        End If
                  End If
                  
            Next i
'            If GetAp_billAge("SA") And Cancel = False Then
'                Call objFilter.FilterArgs.Add("70d9ddab-130a-4e5f-a3c9-0c21544908d2", "ViewID") '3  交叉表
'            Else
'                Call objFilter.FilterArgs.Add("02651bd1-673a-42fb-8d83-cc9a769daff7", "ViewID")
'            End If
        Case "货龄分析（按收款）"
            For i = 1 To objFilter.FilterList.Count
            
                  If objFilter.FilterList(i).Name = "ddate" Then
                  
                        If Trim(objFilter.FilterList(i).varValue) = vbNullString Then
                              Beep
                              MsgBox GetString("U8.SA.xsglsql.clsreportcallback.00095"), vbOKOnly + vbExclamation, StrReportName 'zh-CN：日期必须输入
                              Cancel = True
                        End If
                  End If
                  
            Next i
'            If GetAp_billAge("AR") And Cancel = False Then
'                Call objFilter.FilterArgs.Add("904cdbfe-d279-405e-8639-32d509cc9d5d", "ViewID") '3 交叉表
''                MsgBox 3
'            Else
'                Call objFilter.FilterArgs.Add("33680aab-4543-4674-bfed-cf1a6f76c3c8", "ViewID")
''                MsgBox 1
'            End If
      Case "销售明细表"
      
            For i = 1 To objFilter.FilterList.Count
            
            
                  If objFilter.FilterList(i).Name = "cChecker" Then
                  
                        If objFilter.FilterList(i).varValue = "复核" Then
                        
                              objFilter.FilterList(i).strSqlString = " isnull (Salebillvouch.cChecker,'') <> '' "
                        Else
                              If objFilter.FilterList(i).varValue = "未复核" Then
                                    
                                    objFilter.FilterList(i).strSqlString = " isnull (Salebillvouch.cChecker,'') = '' "
                                    
                              End If
                              
                        End If
                        
                  End If
                  
            Next i
            
      Case "发货明细表"
      
            For i = 1 To objFilter.FilterList.Count
            
            
                    If objFilter.FilterList(i).Name = "Dispatchlist.cVerifier" Then
                  
                        If objFilter.FilterList(i).varValue = "审核" Then
                        
                              objFilter.FilterList(i).strSqlString = " isnull (Dispatchlist.cVerifier,'') <> '' "
                        Else
                              If objFilter.FilterList(i).varValue = "未审核" Then
                                    
                                    objFilter.FilterList(i).strSqlString = " isnull (Dispatchlist.cVerifier,'') = '' "
                                    
                              End If
                              
                        End If
                        
                    End If
                    
                    '退货标识 2004.08.16 by jzq
                    If objFilter.FilterList(i).Name = "Dispatchlist.bReturnFlag" Or objFilter.FilterList(i).Name = "bReturnFlag" Then
                        If objFilter.FilterList(i).varValue = "是" Or objFilter.FilterList(i).varValue = "1" Then
                            objFilter.FilterList(i).strSqlString = " isnull (Dispatchlist.bReturnFlag,0)=1 "
                        ElseIf objFilter.FilterList(i).varValue = "否" Or objFilter.FilterList(i).varValue = "0" Then
                            objFilter.FilterList(i).strSqlString = " isnull (Dispatchlist.bReturnFlag,0)=0 "
                        ElseIf Trim(objFilter.FilterList(i).varValue) <> "" Then   '输入其它值则忽略
                            objFilter.FilterList(i).varValue = vbNullString
                            objFilter.FilterList(i).strSqlString = ""
                        End If
                    End If
            Next i
        Case "01"  '发货单列表
            For i = 1 To objFilter.FilterList.Count
            
            
'                    If LCase(objFilter.FilterList(I).Name) = "cverifier" Then
'
'                        If objFilter.FilterList(I).varValue = "审核" Then
'
'                              objFilter.FilterList(I).strSqlString = " isnull (cVerifier,'') <> '' "
'                        Else
'                              If objFilter.FilterList(I).varValue = "未审核" Then
'
'                                    objFilter.FilterList(I).strSqlString = " isnull (cVerifier,'') = '' "
'
'                              End If
'
'                        End If
'
'                    End If
                    
                    '退货标识 2004.08.16 by jzq
                    If LCase(objFilter.FilterList(i).Name) = "breturnflag" Then
                        If objFilter.FilterList(i).varValue = "是" Or objFilter.FilterList(i).varValue = "1" Then
                            objFilter.FilterList(i).strSqlString = " isnull (bReturnFlag,0)=1 "
                        ElseIf objFilter.FilterList(i).varValue = "否" Or objFilter.FilterList(i).varValue = "0" Then
                            objFilter.FilterList(i).strSqlString = " isnull (bReturnFlag,0)=0 "
                        ElseIf Trim(objFilter.FilterList(i).varValue) <> "" Then   '输入其它值则忽略
                            objFilter.FilterList(i).varValue = vbNullString
                            objFilter.FilterList(i).strSqlString = ""
                        End If
                    End If
                    '是否审核 2008.09.19 by wangxiny
                    If LCase(objFilter.FilterList(i).Name) = "isverified" Then
                        If objFilter.FilterList(i).varValue = "是" Or objFilter.FilterList(i).varValue = "1" Then
                            objFilter.FilterList(i).strSqlString = " isnull (cVerifier,'')<>'' "
                        ElseIf objFilter.FilterList(i).varValue = "否" Or objFilter.FilterList(i).varValue = "0" Then
                            objFilter.FilterList(i).strSqlString = " isnull (cVerifier,'')='' "
                        ElseIf Trim(objFilter.FilterList(i).varValue) <> "" Then   '输入其它值则忽略
                            objFilter.FilterList(i).varValue = vbNullString
                            objFilter.FilterList(i).strSqlString = ""
                        End If
                    End If
            Next i
      Case "销售订货统计表"
      
            For i = 1 To objFilter.FilterList.Count
            
            
                  If objFilter.FilterList(i).Name = "是否审核" Then
                  
                        If objFilter.FilterList(i).varValue = "审核" Then
                        
                              objFilter.FilterList(i).strSqlString = " isnull (SO_SOMain.cVerifier,'') <> '' "
'                        Else
'                              If objFilter.FilterList(I).varValue = "未审核" Then
'
'                                    objFilter.FilterList(I).strSqlString = " isnull (SO_SOMain.cVerifier,'') = '' "
'
'                              End If
                        '-----------------------------------------------------
                        '屏蔽用户输入错误的是否审核状态 by jzq
                        ElseIf objFilter.FilterList(i).varValue = "未审核" Then
                            objFilter.FilterList(i).strSqlString = " isnull (SO_SOMain.cVerifier,'') = '' "
                        ElseIf Trim(objFilter.FilterList(i).varValue) <> "" Then   '输入其它值则提示错误，退出
                            MsgBox GetString("U8.SA.xsglsql.clsreportcallback.00108"), vbOKOnly + vbExclamation, StrReportName 'zh-CN：请在[是否审核]栏输入正确的值。
                            Cancel = True
                        '-----------------------------------------------------
                        End If
                        
                  End If
                  
                  If objFilter.FilterList(i).Name = "定单状态" Then
                  
                        If objFilter.FilterList(i).varValue = "关闭" Or objFilter.FilterList(i).varValue = "1" Then
                        
                              objFilter.FilterList(i).strSqlString = " isnull (SO_SOMain.cCloser,'') <> '' "
'                        Else
'                              If objFilter.FilterList(I).varValue = "打开" Then
'
'                                    objFilter.FilterList(I).strSqlString = " isnull (SO_SOMain.cCloser,'') = '' "
'
'                              End If
                        '----------------------------------------------------
                        '屏蔽用户输入错误的是否审核状态 by jzq
                        ElseIf objFilter.FilterList(i).varValue = "打开" Or objFilter.FilterList(i).varValue = "0" Then
                            objFilter.FilterList(i).strSqlString = " isnull (SO_SOMain.cCloser,'') = '' "
                        ElseIf Trim(objFilter.FilterList(i).varValue) <> "" Then   '输入其它值则提示错误，退出
                            MsgBox GetString("U8.SA.xsglsql.clsreportcallback.00112"), vbOKOnly + vbExclamation, StrReportName 'zh-CN：请在[订单状态]栏输入正确的值。
                            Cancel = True
                        '-----------------------------------------------------
                        End If
                        
                  End If
                  
            Next i
            
      Case "畅适滞分析"
            
            For i = 1 To objFilter.FilterList.Count
            
            
                  Select Case objFilter.FilterList(i).Name
                  
                  Case "ddate"
                  
                        If Trim(objFilter.FilterList(i).varValue) = vbNullString _
                              Or Trim(objFilter.FilterList(i).varValue2) = vbNullString Then
                              Beep
                              MsgBox GetString("U8.SA.xsglsql.clsreportcallback.00114"), vbOKOnly + vbExclamation, StrReportName 'zh-CN：必须输入日期区间
                              Cancel = True
                              Exit Function
                              
                        End If
                        
                  Case "平均周转次数"
                  
                        If Trim(objFilter.FilterList(i).varValue) = vbNullString Then
                              
                              Beep
                              MsgBox GetString("U8.SA.xsglsql.clsreportcallback.00116"), vbOKOnly + vbExclamation, StrReportName 'zh-CN：必须输入平均周转次数
                              Cancel = True
                              Exit Function
                              
                        End If
                        
                  Case "畅销"
                  
                        If Trim(objFilter.FilterList(i).varValue) = vbNullString Then
                              
                              Beep
                              MsgBox GetString("U8.SA.xsglsql.clsreportcallback.00118"), vbOKOnly + vbExclamation, StrReportName 'zh-CN：必须输入畅销
                              Cancel = True
                              Exit Function
                              
                        End If
                        
                  Case "滞销"
                  
                        If Trim(objFilter.FilterList(i).varValue) = vbNullString Then
                              
                              Beep
                              MsgBox GetString("U8.SA.xsglsql.clsreportcallback.00120"), vbOKOnly + vbExclamation, StrReportName 'zh-CN：必须输入滞销
                              Cancel = True
                              Exit Function
                              
                        End If
                        
                  End Select
                  
            Next i
            
      Case "商品周转率分析"
      
            For i = 1 To objFilter.FilterList.Count
            
                  If objFilter.FilterList(i).Name = "ddate" Then
                        '必须输入周转日期
                        If Trim(objFilter.FilterList(i).varValue) = vbNullString _
                              Or Trim(objFilter.FilterList(i).varValue2) = vbNullString Then
                              '分析的日期区间必须输入
                              Beep
                              MsgBox GetString("U8.SA.xsglsql.clsreportcallback.00122"), vbOKOnly + vbExclamation, StrReportName 'zh-CN：必须输入分析日期区间...
                              Cancel = True
                              Exit Function
                        Else
                              If CDate(objFilter.FilterList(i).varValue) > CDate(objFilter.FilterList(i).varValue2) Then
                              
                                    Beep
                                    MsgBox GetString("U8.SA.xsglsql.clsreportcallback.00123"), vbOKOnly + vbExclamation 'zh-CN：开始日期不能大于结束日期...
                                    Cancel = True
                                    Exit Function
                                    
                              End If
                              
                                    
                        End If
                  
                  End If
                  
            Next i
      Case "销售订单预警和报警"
          Dim strPreDay As String
          Dim strOverDay As String
        
          For i = 1 To objFilter.FilterList.Count
              
                    If objFilter.FilterList(i).Name = "iPreDay" Then
                          '预期天数
                          strPreDay = objFilter.FilterList(i).varValue
                          If Trim(objFilter.FilterList(i).varValue) <> vbNullString Then
                                If CDbl(strPreDay) < 0 Then
                                  MsgBox GetString("U8.SA.xsglsql.clsreportcallback.00125"), vbOKOnly + vbExclamation, StrReportName 'zh-CN：提前预警天数不能小于零!
                                  Cancel = True
                                End If
                          End If
                    End If
                    If objFilter.FilterList(i).Name = "iOverDay" Then
                          '逾期天数
                          strOverDay = objFilter.FilterList(i).varValue
                          If Trim(objFilter.FilterList(i).varValue) <> vbNullString Then
                                If CDbl(strOverDay) < 0 Then
                                  MsgBox GetString("U8.SA.xsglsql.clsreportcallback.00126"), vbOKOnly + vbExclamation, StrReportName 'zh-CN：逾期报警天数不能小于零!
                                  Cancel = True
                                End If
                          End If
                    End If
            Next i
            If strPreDay = "" And strOverDay = "" Then
                MsgBox GetString("U8.SA.xsglsql.clsreportcallback.00127"), vbOKOnly + vbExclamation, StrReportName 'zh-CN：提前预警天数和逾期报警天数不能同时为空!
                Cancel = True
            End If
            If val(strPreDay) > 1095 Or val(strOverDay) > 1095 Then    '天数不能超过三年 2008-10-22
                MsgBox GetString("U8.SA.xsglsql_2.orderalert.invaliddate"), vbOKOnly + vbExclamation, StrReportName 'zh-CN：提前预警天数和逾期报警天数不能同时为空!
                Cancel = True
            End If
    Case "SARefVouchT5", "SARefVouchT8", "SARefVouchT2", "SARefVouchT1", "SARefVouchT4", "SARefVouchT7"
        For i = 1 To objFilter.FilterList.Count
            If InStr(1, objFilter.FilterList(i).Name, "cBustype") > 0 Then
                If objFilter.FilterList(i).varValue = "" Then
                    MsgBox GetString("U8.SA.xsglsql_2.msgbox.0001")
                    'MsgBox "请选择正确(参照列表中存在)的业务类型!"
                    Cancel = True
                End If
            End If
            If objFilter.FilterList(i).Name = "DispatchListType" Then
                If objFilter.FilterList(i).varValue = "" Then
                    MsgBox GetString("U8.SA.xsglsql_2.msgbox.0002")
                    'MsgBox "请选择正确(参照列表中存在)的发货单类型"
                    Cancel = True
                End If
            End If
            If objFilter.FilterList(i).Name = "BillType" Then
                If objFilter.FilterList(i).varValue = "" Then
                    MsgBox GetString("U8.SA.xsglsql_2.msgbox.0003")
                    'MsgBox "蓝字记录是否按出库数开票选择有误,请重新选择"
                    Cancel = True
                End If
            End If
            
            
            
            
        Next
    Case "07fwsk"
         For i = 1 To objFilter.FilterList.Count
            If objFilter.FilterList(i).Name = "cvouchtype" Then
                If objFilter.FilterList(i).varValue = "" Then
                    MsgBox GetString("U8.SA.xsglsql_2.msgbox.0004")
                    'MsgBox "请选择正确(参照列表中存在)的发票类型"
                    Cancel = True
                End If
            End If
        Next
    Case "signinreferdispatch"
         For i = 1 To objFilter.FilterList.Count
            If objFilter.FilterList(i).Name = "IsShowExcuted" Then
                If objFilter.FilterList(i).varValue = "是" Then
                    objFilter.FilterList(i).strSqlString = " 1=1 "
                ElseIf objFilter.FilterList(i).varValue = "否" Then
                    objFilter.FilterList(i).strSqlString = " abs(isnull(fsumsignquantity,0))<abs(isnull(foutquantity,0)) "
                End If
            End If
        Next
'    Case "02", "05", "08", "09", "14", "15", "16", "17"
'        For i = 1 To objFilter.FilterList.count
'            If LCase(objFilter.FilterList(i).Name) = "isverified" Then
'                If objFilter.FilterList(i).varValue = "是" Or objFilter.FilterList(i).varValue = "1" Then
'                    objFilter.FilterList(i).strSqlString = " isnull (cVerifier,'')<>'' "
'                ElseIf objFilter.FilterList(i).varValue = "否" Or objFilter.FilterList(i).varValue = "0" Then
'                    objFilter.FilterList(i).strSqlString = " isnull (cVerifier,'')='' "
'                End If
'            End If
'        Next
'    Case "07"
'        For i = 1 To objFilter.FilterList.count
'            If LCase(objFilter.FilterList(i).Name) = "isverified" Then
'                If objFilter.FilterList(i).varValue = "是" Or objFilter.FilterList(i).varValue = "1" Then
'                    objFilter.FilterList(i).strSqlString = " isnull (ccheker,'')<>'' "
'                ElseIf objFilter.FilterList(i).varValue = "否" Or objFilter.FilterList(i).varValue = "0" Then
'                    objFilter.FilterList(i).strSqlString = " isnull (ccheker,'')='' "
'                End If
'            End If
'        Next
    End Select
    

End Function

Public Function CallBack_UserBrowse(ByVal sKey As String, objFilter As UFGeneralFilter.FilterSrv)

        If sKey = "设置账龄区间" Then
'            Set frmAge.mconn = New ADODB.Connection
'            frmAge.mconn.Open m_login.UfDbName
'            frmAge.Show vbModal
'            objFilter.setValue sKey, frmAge.StrResult
        End If
End Function
'获得枚举得值 从数据库中读取
Private Function GetStringEnumForMe(sEnumType As String, sEnumCode As String) As String
    Dim sql As String
    Dim m_rst As New ADODB.Recordset
    sql = "select enumname from v_aa_enum where enumtype=N'" + sEnumType + "' and enumcode=N'" + sEnumCode + "'"
    m_rst.Open sql, DBConn
    If m_rst.BOF And m_rst.EOF Then
        GetStringEnumForMe = sEnumCode
    Else
        If IsNull(m_rst.Fields("enumname").value) Then
            GetStringEnumForMe = ""
        Else
            GetStringEnumForMe = m_rst.Fields("enumname").value
        End If
    End If
    m_rst.Close
    Set m_rst = Nothing
End Function
'判断货龄区间是否有值
Private Function GetAp_billAge(subid As String) As Boolean
    Dim sql As String
    Dim m_rst As New ADODB.Recordset
    sql = "select cnum from Ap_billAge where cFlag =N'" + subid + "'"
    m_rst.Open sql, DBConn
    If m_rst.BOF And m_rst.EOF Then
        GetAp_billAge = False
    Else
        GetAp_billAge = True
    End If
    m_rst.Close
    Set m_rst = Nothing
End Function

Public Sub On_FilterOk(ByVal StrReportName As String, objFilter As clsReportFilter, Cancel As Boolean)
      '当过滤条件界面过滤按纽按下时激活该事件
      'StrReportName  报表名称
      'ObjFilter  过滤条件名称
      'cancel  true  取消操作 false 继续'
      '
      Dim i As Long
      If m_login Is Nothing Then Exit Sub
      
      If LCase(m_login.cSub_Id) <> "SA" Then Exit Sub
      Select Case StrReportName
      Case "销售明细表"
      
            For i = 1 To objFilter.FilterList.Count
            
            
                  If objFilter.FilterList(i).Name = "cChecker" Then
                  
                        If objFilter.FilterList(i).varValue = "复核" Then
                        
                              objFilter.FilterList(i).strSqlString = " isnull (Salebillvouch.cChecker,'') <> '' "
                        Else
                              If objFilter.FilterList(i).varValue = "未复核" Then
                                    
                                    objFilter.FilterList(i).strSqlString = " isnull (Salebillvouch.cChecker,'') = '' "
                                    
                              End If
                              
                        End If
                        
                  End If
                  
            Next i
            
      Case "发货明细表"
      
            For i = 1 To objFilter.FilterList.Count
            
            
                    If objFilter.FilterList(i).Name = "Dispatchlist.cVerifier" Then
                  
                        If objFilter.FilterList(i).varValue = "审核" Then
                        
                              objFilter.FilterList(i).strSqlString = " isnull (Dispatchlist.cVerifier,'') <> '' "
                        Else
                              If objFilter.FilterList(i).varValue = "未审核" Then
                                    
                                    objFilter.FilterList(i).strSqlString = " isnull (Dispatchlist.cVerifier,'') = '' "
                                    
                              End If
                              
                        End If
                        
                    End If
                    
                    '退货标识 2004.08.16 by jzq
                    If objFilter.FilterList(i).Name = "Dispatchlist.bReturnFlag" Or objFilter.FilterList(i).Name = "bReturnFlag" Then
                        If objFilter.FilterList(i).varValue = "是" Or objFilter.FilterList(i).varValue = "1" Then
                            objFilter.FilterList(i).strSqlString = " isnull (Dispatchlist.bReturnFlag,0)=1 "
                        ElseIf objFilter.FilterList(i).varValue = "否" Or objFilter.FilterList(i).varValue = "0" Then
                            objFilter.FilterList(i).strSqlString = " isnull (Dispatchlist.bReturnFlag,0)=0 "
                        ElseIf Trim(objFilter.FilterList(i).varValue) <> "" Then   '输入其它值则忽略
                            objFilter.FilterList(i).varValue = vbNullString
                            objFilter.FilterList(i).strSqlString = ""
                        End If
                    End If
            Next i
            
      Case "销售订货统计表"
      
            For i = 1 To objFilter.FilterList.Count
            
            
                  If objFilter.FilterList(i).Name = "是否审核" Then
                  
                        If objFilter.FilterList(i).varValue = "审核" Then
                        
                              objFilter.FilterList(i).strSqlString = " isnull (SO_SOMain.cVerifier,'') <> '' "
'                        Else
'                              If objFilter.FilterList(I).varValue = "未审核" Then
'
'                                    objFilter.FilterList(I).strSqlString = " isnull (SO_SOMain.cVerifier,'') = '' "
'
'                              End If
                        '-----------------------------------------------------
                        '屏蔽用户输入错误的是否审核状态 by jzq
                        ElseIf objFilter.FilterList(i).varValue = "未审核" Then
                            objFilter.FilterList(i).strSqlString = " isnull (SO_SOMain.cVerifier,'') = '' "
                        ElseIf Trim(objFilter.FilterList(i).varValue) <> "" Then   '输入其它值则提示错误，退出
                            MsgBox GetString("U8.SA.xsglsql.clsreportcallback.00108"), vbOKOnly + vbExclamation, StrReportName 'zh-CN：请在[是否审核]栏输入正确的值。
                            Cancel = True
                        '-----------------------------------------------------
                        End If
                        
                  End If
                  
                  If objFilter.FilterList(i).Name = "定单状态" Then
                  
                        If objFilter.FilterList(i).varValue = "关闭" Or objFilter.FilterList(i).varValue = "1" Then
                        
                              objFilter.FilterList(i).strSqlString = " isnull (SO_SOMain.cCloser,'') <> '' "
'                        Else
'                              If objFilter.FilterList(I).varValue = "打开" Then
'
'                                    objFilter.FilterList(I).strSqlString = " isnull (SO_SOMain.cCloser,'') = '' "
'
'                              End If
                        '----------------------------------------------------
                        '屏蔽用户输入错误的是否审核状态 by jzq
                        ElseIf objFilter.FilterList(i).varValue = "打开" Or objFilter.FilterList(i).varValue = "0" Then
                            objFilter.FilterList(i).strSqlString = " isnull (SO_SOMain.cCloser,'') = '' "
                        ElseIf Trim(objFilter.FilterList(i).varValue) <> "" Then   '输入其它值则提示错误，退出
                            MsgBox GetString("U8.SA.xsglsql.clsreportcallback.00112"), vbOKOnly + vbExclamation, StrReportName 'zh-CN：请在[订单状态]栏输入正确的值。
                            Cancel = True
                        '-----------------------------------------------------
                        End If
                        
                  End If
                  
            Next i
            
      Case "畅适滞分析"
            
            For i = 1 To objFilter.FilterList.Count
            
            
                  Select Case objFilter.FilterList(i).Name
                  
                  Case "ddate"
                  
                        If Trim(objFilter.FilterList(i).varValue) = vbNullString _
                              Or Trim(objFilter.FilterList(i).varValue2) = vbNullString Then
                              Beep
                              MsgBox GetString("U8.SA.xsglsql.clsreportcallback.00114"), vbOKOnly + vbExclamation, StrReportName 'zh-CN：必须输入日期区间
                              Cancel = True
                              Exit Sub
                              
                        End If
                        
                  Case "平均周转次数"
                  
                        If Trim(objFilter.FilterList(i).varValue) = vbNullString Then
                              
                              Beep
                              MsgBox GetString("U8.SA.xsglsql.clsreportcallback.00116"), vbOKOnly + vbExclamation, StrReportName 'zh-CN：必须输入平均周转次数
                              Cancel = True
                              Exit Sub
                              
                        End If
                        
                  Case "畅销"
                  
                        If Trim(objFilter.FilterList(i).varValue) = vbNullString Then
                              
                              Beep
                              MsgBox GetString("U8.SA.xsglsql.clsreportcallback.00118"), vbOKOnly + vbExclamation, StrReportName 'zh-CN：必须输入畅销
                              Cancel = True
                              Exit Sub
                              
                        End If
                        
                  Case "滞销"
                  
                        If Trim(objFilter.FilterList(i).varValue) = vbNullString Then
                              
                              Beep
                              MsgBox GetString("U8.SA.xsglsql.clsreportcallback.00120"), vbOKOnly + vbExclamation, StrReportName 'zh-CN：必须输入滞销
                              Cancel = True
                              Exit Sub
                              
                        End If
                        
                  End Select
                  
            Next i
            
      Case "商品周转率分析"
            For i = 1 To objFilter.FilterList.Count
            
                  If objFilter.FilterList(i).Name = "ddate" Then
                        '必须输入周转日期
                        If Trim(objFilter.FilterList(i).varValue) = vbNullString _
                              Or Trim(objFilter.FilterList(i).varValue2) = vbNullString Then
                              '分析的日期区间必须输入
                              Beep
                              MsgBox GetString("U8.SA.xsglsql.clsreportcallback.00122"), vbOKOnly + vbExclamation, StrReportName 'zh-CN：必须输入分析日期区间...
                              Cancel = True
                              Exit Sub
                        Else
                              If CDate(objFilter.FilterList(i).varValue) > CDate(objFilter.FilterList(i).varValue2) Then
                              
                                    Beep
                                    MsgBox GetString("U8.SA.xsglsql.clsreportcallback.00123"), vbOKOnly + vbExclamation 'zh-CN：开始日期不能大于结束日期...
                                    Cancel = True
                                    Exit Sub
                                    
                              End If
                              
                                    
                        End If
                  
                  End If
                  
            Next i
      End Select
      
End Sub




