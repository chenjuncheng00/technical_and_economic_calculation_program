Imports System.IO
Module 辅助程序
    Sub 打印数据用于调试(data_list As Array, n_data As Integer)
        'data_list：数据，列表
        'n_data：数据数量
        Dim str_txt As String = ""
        For i = 1 To n_data
            str_txt = str_txt & Str(data_list(i)) & "， "
        Next
        MsgBox(str_txt)
    End Sub
    Sub 解锁表格(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '解锁表格
        ExcelApp.ThisWorkbook.Worksheets("估算表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("指标数据").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("财务计划现金流量表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("资产负债表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("指标数据").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("总成本表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("成本税收表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入税收表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("经济增加值计算表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").unProtect(Password:="wscjc")
    End Sub
    Sub 锁定表格(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '锁定表格
        ExcelApp.ThisWorkbook.Worksheets("估算表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("指标数据").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("财务计划现金流量表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("资产负债表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("指标数据").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("总成本表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("成本税收表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("经济增加值计算表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Protect(Password:="wscjc")
    End Sub
    Sub 自保护程序(ExcelApp As Object)
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        '程序开始
        Try '异常处理，防止无文件
            '验证本地C盘是否存在白名单文件，用来区分是不是自己的电脑，是否需要进行自保护程序验证
            Dim fs As New FileStream("C:\Windows\WhiteList_CJC.txt", FileMode.Open)
            Dim sr As New StreamReader(fs)
            Dim strTemp As String
            strTemp = sr.ReadLine
            Dim WhiteList_PC As String = strTemp '获取TXT文本内容
            sr.Close()
            fs.Close()
            If WhiteList_PC <> "WhiteList_PC" Then '如果不是合格的白名单文件，验证失败，则进行自保护验证
                'Call 获取本地服务器版本信息并验证(ExcelApp)
                Call 获取本机MAC地址并验证(ExcelApp)
                Call 网络时间和本地时间交替验证(ExcelApp)
            End If
        Catch ex As Exception
            '发生任何异常，则进行自保护验证
            'Call 获取本地服务器版本信息并验证(ExcelApp)
            Call 获取本机MAC地址并验证(ExcelApp)
            Call 网络时间和本地时间交替验证(ExcelApp)
            Exit Sub
        End Try
    End Sub
    Sub 获取本机MAC地址并验证(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        '程序开始
        Dim wmiObjSet As WbemScripting.SWbemObjectSet
        Dim obj As WbemScripting.SWbemObject
        Dim MAC As String = Nothing
        Dim MACTEST As Integer = 0 'MAC地址检测，1代表通过，0代表不通过
        wmiObjSet = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_NetworkAdapterConfiguration")
        For Each obj In wmiObjSet
            MAC = obj.MACAddress
            'MAC地址白名单(macbook虚拟机)
            If MAC = "00:1C:42:23:91:BA" Then
                MACTEST = 1
                Exit For
            Else
                MACTEST = 0
            End If
        Next
        '—————————————————————————————————————————————————————————————————————————————————————————
        '针对MAC地址检测结果，进行不同操作
        If MACTEST = 0 Then '如果MAC地址检测不通过
            '保存表格的改动
            ExcelApp.Application.DisplayAlerts = False
            ExcelApp.ThisWorkbook.Save()
            ExcelApp.Application.DisplayAlerts = True
            '错误提示
            MsgBox("Local MAC Error！")
            '直接退出表格
            ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
            ExcelApp.Application.DisplayAlerts = False
            ExcelApp.Application.Quit()
            ExcelApp.Application.DisplayAlerts = True
        End If
    End Sub
    Sub 网络时间和本地时间交替验证(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        '————————————————————————————————————————————————————————————————————————————————————————
        '程序开始
        '先验证网络时间
        Dim PingStatus As Boolean = False 'ping的状态
        Dim A As String = "TimeOut" '网络状态默认值
        For i = 1 To 10 '最多Ping网络10次
            A = CStr(Pings()) '反回函数状态
            If A = "Success" Then
                A = "Success"
                Exit For '只要检测出ping通，就跳出循环
            End If
        Next
        If A = "Success" Then
            PingStatus = True
        Else
            PingStatus = False
        End If
        Dim strText As String
        If PingStatus = True Then '如果网络是通的
            With CreateObject("MSXML2.ServerXMLHTTP") '获取网络时间
                .Open("GET", "https://www.baidu.com/index.php", False)
                .send
                strText = .getResponseHeader("Date")
                Dim GetDate = DateAdd("h", 8, Split(Replace(strText, " GMT", ""), ",")(1)) '将获取的字符串格式GMT网络时间加8小时转成北京时间，并改成日期格式
                If GetDate >= #12/01/2022# Then '月/日/年，验证网络时间
                    '保存表格的改动
                    ExcelApp.Application.DisplayAlerts = False
                    ExcelApp.ThisWorkbook.Save()
                    ExcelApp.Application.DisplayAlerts = True
                    '错误提示
                    MsgBox("Web Date Error！")
                    '直接退出表格
                    ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
                    ExcelApp.Application.DisplayAlerts = False
                    ExcelApp.Application.Quit()
                    ExcelApp.Application.DisplayAlerts = True
                End If
            End With
        Else '如果网络不通，验证系统本地时间
            '进程暂停一段时间，Sleep的参数是毫秒(千分之一秒)
            Threading.Thread.Sleep(6000000)
            Dim Local_Time = Date.Now '获取系统本地时间
            Dim Dead_Time = Convert.ToDateTime("2021/12/01 01:00:00") '设定程序有效期截止时间
            If Date.Compare(Local_Time, Dead_Time) > 0 Then '大于0，说明系统本地时间大于程序有效期，程序不可以继续使用
                '保存表格的改动
                ExcelApp.Application.DisplayAlerts = False
                ExcelApp.ThisWorkbook.Save()
                ExcelApp.Application.DisplayAlerts = True
                '错误提示
                MsgBox("Internet Ping Error and Local Date Error！")
                '直接退出表格
                ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
                ExcelApp.Application.DisplayAlerts = False
                ExcelApp.Application.Quit()
                ExcelApp.Application.DisplayAlerts = True
            End If
        End If
    End Sub
    Sub 获取系统时间并验证(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        Dim Local_Time = Date.Now '获取系统本地时间
        Dim Dead_Time = Convert.ToDateTime("2021/12/01 01:00:00") '设定程序有效期截止时间
        If Date.Compare(Local_Time, Dead_Time) > 0 Then '大于0，说明系统本地时间大于程序有效期，程序不可以继续使用
            '保存表格的改动
            ExcelApp.Application.DisplayAlerts = False
            ExcelApp.ThisWorkbook.Save()
            ExcelApp.Application.DisplayAlerts = True
            '错误提示
            MsgBox("Local Date Error！")
            '直接退出表格
            ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
            ExcelApp.Application.DisplayAlerts = False
            ExcelApp.Application.Quit()
            ExcelApp.Application.DisplayAlerts = True
        End If
    End Sub
    Sub 程序联网验证(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        '程序开始
        Dim PingStatus As Boolean = False 'ping的状态
        Dim A As String = CStr(Pings()) '反回函数状态
        If A = "Success" Then
            PingStatus = True
        Else
            PingStatus = False
        End If
        Dim strText As String
        If PingStatus = True Then
            With CreateObject("MSXML2.ServerXMLHTTP")
                .Open("GET", "https://www.baidu.com/index.php", False)
                .send
                strText = .getResponseHeader("Date")
                Dim GetDate = DateAdd("h", 8, Split(Replace(strText, " GMT", ""), ",")(1)) '将获取的字符串格式GMT网络时间加8小时转成北京时间，并改成日期格式
                If GetDate >= #12/01/2021# Then '月/日/年
                    '保存表格的改动
                    ExcelApp.Application.DisplayAlerts = False
                    ExcelApp.ThisWorkbook.Save()
                    ExcelApp.Application.DisplayAlerts = True
                    '错误提示
                    MsgBox("Web Date Error！")
                    '直接退出表格
                    ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
                    ExcelApp.Application.DisplayAlerts = False
                    ExcelApp.Application.Quit()
                    ExcelApp.Application.DisplayAlerts = True
                End If
            End With
        Else
            MsgBox("Internet Ping Error")
            ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
            ExcelApp.Application.DisplayAlerts = False
            ExcelApp.Application.Quit()
            ExcelApp.Application.DisplayAlerts = True
        End If
    End Sub
    Function Pings() As String
        Dim _ping As New Net.NetworkInformation.Ping
        Dim _pingreply As Net.NetworkInformation.PingReply = _ping.Send("www.baidu.com") 'ping 百度
        Return _pingreply.Status.ToString()
        '返回以下信息
        'TimedOut 失败
        'Success  成功
    End Function
    Sub 获取本地服务器版本信息并验证(ExcelApp As Object)
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        '程序开始
        Try '异常处理，防止无文件
            Dim fs As New FileStream("\\192.168.9.201\user1\CJC-陈俊丞\综合能源多联供项目计算工具\技术经济分析计算程序\Version.txt", FileMode.Open)
            Dim sr As New StreamReader(fs)
            Dim strTemp As String
            strTemp = sr.ReadLine
            Dim FWQBBH_String As String = strTemp '服务器版本号，字符串格式
            sr.Close()
            fs.Close()
            Dim FWQBBH = CInt(FWQBBH_String) '将服务器版本号从字符串格式转为整数型格式
            If FWQBBH > 20190401 Then '如果服务器版本号大于设定的内置版本号信息，则退出程序
                '保存表格的改动
                ExcelApp.Application.DisplayAlerts = False
                ExcelApp.ThisWorkbook.Save()
                ExcelApp.Application.DisplayAlerts = True
                '错误提示
                MsgBox("Version Error！")
                '直接退出表格
                ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
                ExcelApp.Application.DisplayAlerts = False
                ExcelApp.Application.Quit()
                ExcelApp.Application.DisplayAlerts = True
            End If
        Catch ex As Exception
            '发生任何异常，直接退出
            '保存表格的改动
            ExcelApp.Application.DisplayAlerts = False
            ExcelApp.ThisWorkbook.Save()
            ExcelApp.Application.DisplayAlerts = True
            '错误提示
            MsgBox("Version Error！")
            '直接退出表格
            ExcelApp.ActiveWorkbook.Close(SaveChanges:=True)
            ExcelApp.Application.DisplayAlerts = False
            ExcelApp.Application.Quit()
            ExcelApp.Application.DisplayAlerts = True
            Exit Sub
        End Try
    End Sub
End Module
