Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.IO
<ComClass(Com技术经济分析计算程序.ClassId, Com技术经济分析计算程序.InterfaceId, Com技术经济分析计算程序.EventsId)>
Public Class Com技术经济分析计算程序
#Region "COM GUID"
    ' 这些 GUID 提供此类的 COM 标识 
    ' 及其 COM 接口。若更改它们，则现有的
    ' 客户端将不再能访问此类。
    Public Const ClassId As String = "de8978b4-8bff-40cb-8dd3-20d43f6a6fea"
    Public Const InterfaceId As String = "f9f114d5-c79c-462b-a470-00ede680b72f"
    Public Const EventsId As String = "a84fa039-2e81-4ccd-a62a-cc28e28fc11a"
#End Region
    ' 可创建的 COM 类必须具有一个不带参数的 Sub New() 
    ' 否则， 将不会在 
    ' COM 注册表中注册此类，且无法通过
    ' CreateObject 创建此类。
    Sub New()
        MyBase.New()
    End Sub
    Sub 一键出表()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————  
        Dim Form11 As New 经济评价表格导出
        Form11.ShowDialog() '窗口显示
        Form11.TopMost = True
        System.Windows.Forms.Application.DoEvents()
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        Call 锁定表格()
    End Sub
    Sub 估算表年限修改后系数计算()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '项目计算年数改变后改变的计算用系数
        Call 项目计算年限变化后改变相关系数()
        '折旧年数改变后改变的计算用系数
        Call 折旧计算年限变化后改变相关系数()
        '无形资产摊销年限改变后改变的计算用系数
        Call 无形资产摊销年限变化后改变相关系数()
        '年数总和法逐年折旧摊销系数
        Call 年数总和法逐年折旧摊销系数()
        '长期贷款还款年数改变后改变的计算用系数
        Call 长期贷款计算年限变化后改变相关系数()
        '判断是否处于长期贷款宽限期
        Call 是否处于长期贷款宽限期()
        '建设期可抵扣增值税计算系数
        Call 建设期增值税抵扣系数()
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '将部分设置重置回默认状态
        Call 将部分设置重置回默认状态()
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '计算流动资金
        Call 流动资金相关计算()
        '重新计算投资回收期
        Call 投资回收期计算()
        '锁定表格
        ExcelApp.ThisWorkbook.Worksheets("估算表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Protect(Password:="wscjc")
        '重新打开excel自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
    End Sub

    Sub 直接输入综合负荷率()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        '屏蔽屏幕更新，防止屏闪；手动计算，关闭excel的自动计算；解锁表格
        Call 计算前基本处理()
        '重新打开excel自动计算，仅部分区域
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("C3:R9").Calculate
        For i = 3 To 17
            '综合负荷率，前15年
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(102, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(5, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value = 0
            End If
        Next
        For i = 2 To 17
            '综合负荷率，后16年        
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(105, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(8, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i).Value = 0
            End If
        Next
        '检测建设期，将没有投产月份数量的年份负荷率设置为0
        For i = 3 To 17 '投资计划与资金筹措表列号
            '前15年
            If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value = 0 Then '如果当年投产的月份数=0                
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value = 0 '负荷率设置为0
            End If
        Next
        For i = 2 To 17 '投资计划与资金筹措表列号           
            '后16年
            If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 16).Value = 0 Then '如果当年投产的月份数=0                
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i).Value = 0 '负荷率设置为0
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '如果此时接入费是按照逐年达产率增加值计算的，则重新计算此时的接入费逐年系数
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(41, 18).Value = "逐年达产率增加值" Then
            '接入费计算模式（第2年到计算期最后一年，根据逐年达产率增加值计算）
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").unProtect(Password:="wscjc")
            '计算一次工作簿
            ExcelApp.Calculate()
            Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
            '第1年
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, 3).Value = 0
            '第1-15年
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= 2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i - 1).Value)
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
                End If
            Next
            '第16年
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, 18).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, 2).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, 17).Value)
            '第17-31年
            For i = 19 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= 2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15 - 1).Value)
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
                End If
            Next
            '将小于0的结果设置为0
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value < 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
                End If
            Next
        End If
        '————————————————————————————————————————————————————————————————————————————————————————        
        '计算流动资金
        Call 流动资金相关计算()
        '重新计算投资回收期
        Call 投资回收期计算()
        '锁定表格
        ExcelApp.ThisWorkbook.Worksheets("估算表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Protect(Password:="wscjc")
        '重新打开excel自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
    End Sub
    Sub 分投资逐次输入负荷率()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        '屏蔽屏幕更新，防止屏闪；手动计算，关闭excel的自动计算；解锁表格
        Call 计算前基本处理()
        '计算年限改变后，改变相应的逐年达产系数（前15年）
        For i = 3 To 17
            '第1次投资负荷率
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(102, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(507, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(447, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(507, i).Value = 0
            End If
            '第2次投资负荷率
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(102, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(512, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(452, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(512, i).Value = 0
            End If
            '第3次投资负荷率
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(102, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(517, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(457, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(517, i).Value = 0
            End If
            '第4次投资负荷率
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(102, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(522, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(462, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(522, i).Value = 0
            End If
            '第5次投资负荷率
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(102, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(527, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(467, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(527, i).Value = 0
            End If
            '第6次投资负荷率
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(102, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(532, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(472, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(532, i).Value = 0
            End If
            '第7次投资负荷率
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(102, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(537, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(477, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(537, i).Value = 0
            End If
            '第8次投资负荷率
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(102, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(542, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(482, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(542, i).Value = 0
            End If
            '第9次投资负荷率
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(102, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(547, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(487, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(547, i).Value = 0
            End If
            '第10次投资负荷率
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(102, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(552, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(492, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(552, i).Value = 0
            End If
        Next
        '计算年限改变后，改变相应的逐年达产系数（后16年）
        For i = 2 To 17
            '第1次投资负荷率
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(105, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(509, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(449, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(509, i).Value = 0
            End If
            '第2次投资负荷率
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(105, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(514, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(454, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(514, i).Value = 0
            End If
            '第3次投资负荷率
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(105, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(519, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(459, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(519, i).Value = 0
            End If
            '第4次投资负荷率
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(105, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(524, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(464, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(524, i).Value = 0
            End If
            '第5次投资负荷率
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(105, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(529, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(469, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(529, i).Value = 0
            End If
            '第6次投资负荷率
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(105, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(534, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(474, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(534, i).Value = 0
            End If
            '第7次投资负荷率
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(105, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(539, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(479, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(539, i).Value = 0
            End If
            '第8次投资负荷率
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(105, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(544, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(484, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(544, i).Value = 0
            End If
            '第9次投资负荷率
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(105, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(549, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(489, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(549, i).Value = 0
            End If
            '第10次投资负荷率          
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(105, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(554, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(494, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(554, i).Value = 0
            End If
        Next
        '重新打开excel自动计算，仅部分区域
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("A90:S650").Calculate
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("C3:R9").Calculate
        '————————————————————————————————————————————————————————————————————————————————————————        
        '前15年
        For i = 3 To 17
            '综合负荷率
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(102, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(5, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value = 0
            End If
        Next
        '后16年
        For i = 2 To 17
            '综合负荷率          
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(105, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(8, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i).Value = 0
            End If
        Next
        '检测建设期，将没有投产月份数量的年份负荷率设置为0
        For i = 3 To 17 '投资计划与资金筹措表列号
            '前15年
            If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value = 0 Then '如果当年投产的月份数=0                
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value = 0 '负荷率设置为0
            End If
        Next
        For i = 2 To 17 '投资计划与资金筹措表列号          
            '后16年
            If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 16).Value = 0 Then '如果当年投产的月份数=0                
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i).Value = 0 '负荷率设置为0
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '如果此时接入费是按照逐年达产率增加值计算的，则重新计算此时的接入费逐年系数
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(41, 18).Value = "逐年达产率增加值" Then
            '接入费计算模式（第2年到计算期最后一年，根据逐年达产率增加值计算）
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").unProtect(Password:="wscjc")
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").unProtect(Password:="wscjc")
            '计算一次工作簿
            ExcelApp.Calculate()
            Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
            '第1年
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, 3).Value = 0
            '第1-15年
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= 2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i - 1).Value)
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
                End If
            Next
            '第16年
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, 18).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, 2).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, 17).Value)
            '第17-31年
            For i = 19 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= 2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15 - 1).Value)
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
                End If
            Next
            '将小于0的结果设置为0
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value < 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
                End If
            Next
            '锁定表格
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Protect(Password:="wscjc")
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Protect(Password:="wscjc")
        End If
        '————————————————————————————————————————————————————————————————————————————————————————        
        '计算流动资金
        Call 流动资金相关计算()
        '重新计算投资回收期
        Call 投资回收期计算()
        '锁定表格
        ExcelApp.ThisWorkbook.Worksheets("估算表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Protect(Password:="wscjc")
        '重新打开excel自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
    End Sub

    Sub 清空输入的收入和成本数据()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '屏蔽屏幕更新
        ExcelApp.Application.ScreenUpdating = False
        '屏蔽事件
        ExcelApp.Application.EnableEvents = False
        '屏蔽自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationManual
        '————————————————————————————————————————————————————————————————————————————————————————
        '收入
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("E13:F25").Value = 0
        '成本
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("L14:M19").Value = 0
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("L21:M25").Value = 0
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("L27:M27").Value = 0
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(20, 19).Value = "直接输入" Then
            '直接输入的逐年负荷系数
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(95, 3).Value = 0
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(95, 4).Value = 0.35
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(95, 5).Value = 0.7
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("F95:Q95").Value = 1
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("B97:Q97").Value = 1
        Else
            '逐次输入的负荷达产率
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("F443:O443").Value = Nothing
            '第1次投资
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("C447:Q447").Value = Nothing
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("B449:Q449").Value = Nothing
            '第2次投资
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("C452:Q452").Value = Nothing
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("B454:Q454").Value = Nothing
            '第3次投资
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("C457:Q457").Value = Nothing
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("B459:Q459").Value = Nothing
            '第4次投资
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("C462:Q462").Value = Nothing
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("B464:Q464").Value = Nothing
            '第5次投资
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("C467:Q467").Value = Nothing
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("B469:Q469").Value = Nothing
            '第6次投资
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("C472:Q472").Value = Nothing
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("B474:Q474").Value = Nothing
            '第7次投资
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("C477:Q477").Value = Nothing
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("B479:Q479").Value = Nothing
            '第8次投资
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("C482:Q482").Value = Nothing
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("B484:Q484").Value = Nothing
            '第9次投资
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("C487:Q487").Value = Nothing
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("B489:Q489").Value = Nothing
            '第10次投资
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("C492:Q492").Value = Nothing
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("B494:Q494").Value = Nothing
        End If
        '计算一次
        ExcelApp.Calculate()
        '负荷率
        Call 直接输入综合负荷率()
        Call 分投资逐次输入负荷率()
        '隐藏收入和成本表格
        Call 隐藏总成本表中成本为0的行()
        Call 隐藏收入税收表中收入为0的行()
        '计算一次流动资金
        Call 流动资金相关计算()
        '计算一次回收期
        Call 投资回收期计算()
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '打开事件
        ExcelApp.Application.EnableEvents = True
        '打开自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
    End Sub

    Sub 反算临界点()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '屏蔽事件
        ExcelApp.Application.EnableEvents = False
        '屏蔽屏幕更新
        ExcelApp.Application.ScreenUpdating = False
        Dim FSLJDJSCSMax As Integer = InputBox("请输入反算临界点功能的最大计算次数，输入的数字越大，计算次数越多，计算速度越慢，计算精度越高。", "输入反算临界点最大计算次数", 200)
        '解锁表格
        ExcelApp.ThisWorkbook.Worksheets("估算表").Unprotect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Unprotect(Password:="wscjc")
        '实例化计算进度显示窗体
        Dim Form1 As New 计算进度显示
        '确定选择的需要反算的内部收益率类型
        '如果是资本金所得税后内部收益率，进行下列计算
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(41, 5).Value = "资本金所得税后内部收益率" Then
            Call 资本金税后收益率反算静态投资(FSLJDJSCSMax)
            Call 资本金税后收益率反算收入单价(FSLJDJSCSMax)
            Call 资本金税后收益率反算成本单价(FSLJDJSCSMax)
        End If
        '如果是全投资所得税后内部收益率，进行下列计算
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(41, 5).Value = "全投资所得税后内部收益率" Then
            Call 全投资税后收益率反算静态投资(FSLJDJSCSMax)
            Call 全投资税后收益率反算收入单价(FSLJDJSCSMax)
            Call 全投资税后收益率反算成本单价(FSLJDJSCSMax)
        End If
        '锁定表格
        ExcelApp.ThisWorkbook.Worksheets("估算表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Protect(Password:="wscjc")
        '打开事件
        ExcelApp.Application.EnableEvents = True
        '计算流动资金
        Call 流动资金相关计算()
        '计算回收期
        Call 投资回收期计算()
        '提醒计算完成
        Form1.Show()
        Form1.Label1.Text = "反算临界点已计算完成！"
        Form1.TopMost = True
        System.Windows.Forms.Application.DoEvents()
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
    End Sub
    Sub 清空反算临界点数据()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '解锁表格
        ExcelApp.ThisWorkbook.Worksheets("估算表").Unprotect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Unprotect(Password:="wscjc")
        '清空数据
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("D43:D67").ClearContents
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Range("G42:G67").ClearContents
        '锁定表格
        ExcelApp.ThisWorkbook.Worksheets("估算表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Protect(Password:="wscjc")
        MsgBox("反算临界点数据已清空！")
    End Sub
    Sub 敏感性分析计算()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————       
        Dim Form10 As New 选择单因素敏感性分析内容
        '显示敏感性分析选择窗口
        Form10.ShowDialog() '窗口显示
        Form10.TopMost = True
        System.Windows.Forms.Application.DoEvents()
    End Sub
    Sub 盈亏平衡点计算()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '本方法为，按照总成本最大年份进行计算
        '定义局部变量
        Dim nfxh As Integer = 0
        '解锁表格
        ExcelApp.ThisWorkbook.Worksheets("指标数据").Unprotect(Password:="wscjc")
        '寻找总成本费用最大值的年份序号，并记录下年固定成本和可变成本
        For i = 5 To 19 '前15年
            If ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(24, i).Value = ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(1, 4).Value Then
                nfxh = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(4, i).Value '记录下年份序号
                '记录下年固定成本、可变成本
                '年固定成本
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(56, 5).Value = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(26, i).Value
                '年可变成本
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(58, 5).Value = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(25, i).Value
                Exit For
            End If
        Next
        For i = 4 To 19 '后16年
            If ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(56, i).Value = ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(1, 4).Value Then
                nfxh = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(36, i).Value '记录下年份序号
                '记录下年固定成本、可变成本
                '年固定成本
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(56, 5).Value = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(58, i).Value
                '年可变成本
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(58, 5).Value = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(57, i).Value
                Exit For
            End If
        Next
        '记录选择的计算年份
        ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(55, 5).Value = nfxh
        '记录年销售收入、税金及附加
        For i = 5 To 19 '前15年
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(4, i).Value = nfxh Then
                '年销售收入
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(57, 5).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(5, i).Value
                '年税金及附加
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(59, 5).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(17, i).Value
            End If
        Next
        '记录年销售收入、税金及附加
        For i = 4 To 19 '后16年
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(27, i).Value = nfxh Then
                '年销售收入
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(57, 5).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(28, i).Value
                '年税金及附加
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(59, 5).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(40, i).Value
            End If
        Next
        '锁定表格
        ExcelApp.ThisWorkbook.Worksheets("指标数据").Protect(Password:="wscjc")
        MsgBox("盈亏平衡点（BEP）已计算完成！")
    End Sub
    Sub 复制敏感性分析图()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '解锁表格
        ExcelApp.ThisWorkbook.Worksheets("指标数据").Unprotect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("指标数据").ChartObjects("单因素敏感性分析图").Activate
        ExcelApp.ActiveChart.ChartArea.Copy()
        MsgBox("复制完成！")
    End Sub
    Sub 复制盈亏平衡分析图()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '解锁表格
        ExcelApp.ThisWorkbook.Worksheets("指标数据").Unprotect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("指标数据").ChartObjects("盈亏平衡分析图").Activate
        ExcelApp.ActiveChart.ChartArea.Copy()
        MsgBox("复制完成！")
    End Sub
    Sub 清空敏感性分析和盈亏平衡分析计算数据()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '解锁表格
        ExcelApp.ThisWorkbook.Worksheets("指标数据").Unprotect(Password:="wscjc")
        '敏感性分析计算结果
        ExcelApp.ThisWorkbook.Worksheets("指标数据").Range("K7:L141").ClearContents
        ExcelApp.ThisWorkbook.Worksheets("指标数据").Range("H7:H141").ClearContents
        '删除敏感性分析图表
        ExcelApp.ThisWorkbook.Worksheets("指标数据").ChartObjects("单因素敏感性分析图").Activate
        ExcelApp.ActiveChart.Parent.Delete
        '手动计算，关闭excel的自动计算,提高计算速度
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationManual
        '将敏感性分析变化率重置回5%，格式重置回0%
        For i = 7 To 137 Step 5
            For j = 1 To 5
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i + j - 1, 9).Value = 0.05 * j - 3 * 0.05
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i + j - 1, 9).NumberFormatLocal = "0%"
            Next
        Next
        '清空Excel内已有的输入
        For i = 7 To 137 Step 5
            ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i, 29).Value = 0
        Next
        '清空盈亏平衡分析计算数据
        ExcelApp.ThisWorkbook.Worksheets("指标数据").Range("E55:E59").ClearContents
        '——————————————————————————————————————————————————————————————————————————————————————
        '锁定表格
        ExcelApp.ThisWorkbook.Worksheets("指标数据").Protect(Password:="wscjc")
        '重新打开excel自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
    End Sub
    Sub 清空时间计划表数据()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————    
        '屏蔽事件
        ExcelApp.Application.EnableEvents = False
        '屏蔽excel自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationManual
        '屏蔽屏幕更新
        ExcelApp.Application.ScreenUpdating = False
        '————————————————————————————————————————————————————————————————————————————————————————    
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("G205:G214").Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("O205:P214").Value = 0
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("B3:B4").Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("D3:D4").Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("J3:J4").Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("L3:L4").Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("B35:B36").Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("D35:D36").Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("J35:J36").Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("L35:L36").Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("B67:B68").Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("D67:D68").Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("J67:J68").Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("L67:L68").Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("B99:B100").Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("D99:D100").Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("J99:J100").Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("L99:L100").Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("B131:B132").Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("D131:D132").Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("J131:J132").Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("L131:L132").Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(172, 2).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(172, 4).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(172, 10).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(172, 12).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(174, 2).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(174, 4).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(174, 10).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(174, 12).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(176, 2).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(176, 4).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(176, 10).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(176, 12).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(178, 2).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(178, 4).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(178, 10).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(178, 12).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(180, 2).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(180, 4).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(180, 10).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(180, 12).Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("L314:L323").Value = 0
        '计算一次
        ExcelApp.Calculate()
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '投资各方出资比例重置回默认值
        ExcelApp.Worksheets("建设期时间计划表").Cells(163, 2).Value = 0
        ExcelApp.Worksheets("建设期时间计划表").Cells(164, 2).Value = 0
        ExcelApp.Worksheets("建设期时间计划表").Cells(165, 2).Value = 0
        ExcelApp.Worksheets("建设期时间计划表").Cells(166, 2).Value = 0
        ExcelApp.Worksheets("建设期时间计划表").Cells(167, 2).Value = 0
        ExcelApp.Worksheets("建设期时间计划表").Cells(163, 2).Value = 100
        '计算一次
        ExcelApp.Calculate()
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '将已经输入的投资金额和各种量清空
        Dim XZ1 = MsgBox("是否需要同时清空<估算表>中输入的逐年投资额、建设期可抵扣增值税额、各种年量等数据？", vbOKCancel)
        If XZ1 = vbOK Then
            '第一到第十次投资金额
            For i = 30 To 73 Step 43 '估算表行号
                For j = 3 To 11 Step 2 '估算表列号
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value > 0 Then
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value = 0
                    End If
                Next
            Next
            '建设期可抵扣增值税额
            For i = 38 To 81 Step 43 '估算表行号
                For j = 3 To 11 Step 2 '估算表列号
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value > 0 Then
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value = 0
                    End If
                Next
            Next
            '各种年量、前五次投资
            For i = 44 To 55 '估算表行号
                For j = 3 To 11 Step 2 '估算表列号
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value > 0 Then
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value = 0
                    End If
                Next
            Next
            '各种年量、后五次投资
            For i = 91 To 102 '估算表行号
                For j = 3 To 11 Step 2 '估算表列号
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value > 0 Then
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value = 0
                    End If
                Next
            Next
        End If
        '计算一次
        ExcelApp.Calculate()
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '建设期每年投产的月份数计算清空
        ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").unProtect(Password:="wscjc")
        '手动输入的建设期投产月份数清空
        For i = 127 To 136 '行号
            For j = 3 To 33 '列号
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i, j).Value = 0
            Next
        Next
        '仅建设期彻底结束当年的投产月份数清空
        For i = 142 To 151 '行号
            For j = 3 To 33 '列号
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i, j).Value = 0
            Next
        Next
        '计算一次
        ExcelApp.Calculate()
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '清空逐年投产月份（自动计算和手动计算）
        For i = 1 To 10
            For j = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(126 + i, j).Value = 0 '手动
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(141 + i, j).Value = 0 '自动
            Next
        Next
        For i = 3 To 33
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value = 0 '汇总
        Next
        '计算一次
        ExcelApp.Calculate()
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '项目计算年数改变后改变的计算用系数
        Call 项目计算年限变化后改变相关系数()
        '折旧年数改变后改变的计算用系数
        Call 折旧计算年限变化后改变相关系数()
        '无形资产摊销年限改变后改变的计算用系数
        Call 无形资产摊销年限变化后改变相关系数()
        '长期贷款还款年数改变后改变的计算用系数
        Call 长期贷款计算年限变化后改变相关系数()
        '年数总和法逐年折旧摊销系数
        Call 年数总和法逐年折旧摊销系数()
        '计算长期贷款是否处于宽限期
        Call 是否处于长期贷款宽限期()
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '第2到第10次投资年份输入发生变化时修改计算系数
        Call 第二至第五次投资年份改变后修改相关计算系数()
        Call 第六至第十次投资年份改变后修改相关计算系数()
        '建设期可抵扣增值税计算系数
        Call 建设期增值税抵扣系数()
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '计算一次负荷率
        Call 直接输入综合负荷率()
        Call 分投资逐次输入负荷率()
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        Call 将部分设置重置回默认状态()
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '计算一次
        ExcelApp.Calculate()
        '计算流动资金
        Call 流动资金相关计算()
        '重新计算投资回收期
        Call 投资回收期计算()
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '锁定表格
        ExcelApp.ThisWorkbook.Worksheets("估算表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Protect(Password:="wscjc")
        '开启事件触发
        ExcelApp.Application.EnableEvents = True
        '重新打开excel自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
        MsgBox("清空建设期时间计划完成！")
    End Sub
    Sub 确定建设期时间计划()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Call 计算前基本处理()
        '解锁建设期时间计划表
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").unProtect(Password:="wscjc")
        '定义局部变量
        Dim j1 As Integer = 0
        Dim j2 As Integer = 0
        Dim j3 As Integer = 0
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————     
        '添加一些报错功能
        '建设期年份总数不可以大于10
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(162, 7).Value > 10 Then
            MsgBox("建设期总年份数量不可以大于10，请重新输入建设期时间计划！")
            '重新打开屏幕更新
            ExcelApp.Application.ScreenUpdating = True
            Exit Sub
        End If
        '------------------------------------------------------------------------------------------------------------
        '第一次建设期
        '结束年份不可以小于开始年份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 2).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 2).Value Then
            MsgBox("第一次建设期结束年份不可以小于开始年份，程序将自动修改相关参数！")
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 2).Value + 1
        End If
        '结束年份等于开始年份，但是结束月份不可以小于开始月份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 2).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 4).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 4).Value Then
            MsgBox("第一次建设期结束年份等于开始年份时，结束月份不可以小于开始月份，程序将自动修改相关参数！")
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 4).Value + 1 <= 12 Then
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 4).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 4).Value + 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 4).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 4).Value - 1
            End If
        End If
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '第二次建设期
        '第二次开始年份不可以小于第一次结束年份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 10).Value <= ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 2).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 10).Value > 0 Then
            MsgBox("第二次建设期开始年份不可以小于等于第一次建设期的结束年份，程序将自动修改相关参数！")
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 2).Value + 1
        End If
        '结束年份不可以小于开始年份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 10).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 10).Value Then
            MsgBox("第二次建设期结束年份不可以小于开始年份，程序将自动修改相关参数！")
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 10).Value + 1
        End If
        '结束年份等于开始年份，但是结束月份不可以小于开始月份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 10).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 12).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 12).Value Then
            MsgBox("第二次建设期结束年份等于开始年份时，结束月份不可以小于开始月份，程序将自动修改相关参数！")
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 12).Value + 1 <= 12 Then
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 12).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 12).Value + 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 12).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 12).Value - 1
            End If
        End If
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '第三次建设期
        '第三次开始年份不可以小于第二次结束年份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 2).Value <= ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 10).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 2).Value > 0 Then
            MsgBox("第三次建设期开始年份不可以小于等于第二次建设期的结束年份，程序将自动修改相关参数！")
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(4, 10).Value + 1
        End If
        '结束年份不可以小于开始年份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 2).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 2).Value Then
            MsgBox("第三次建设期结束年份不可以小于开始年份，程序将自动修改相关参数！")
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 2).Value + 1
        End If
        '结束年份等于开始年份，但是结束月份不可以小于开始月份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 2).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 4).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 4).Value Then
            MsgBox("第三次建设期结束年份等于开始年份时，结束月份不可以小于开始月份，程序将自动修改相关参数！")
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 4).Value + 1 <= 12 Then
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 4).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 4).Value + 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 4).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 4).Value - 1
            End If
        End If
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '第四次建设期
        '第四次开始年份不可以小于第三次结束年份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 10).Value <= ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 2).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 10).Value > 0 Then
            MsgBox("第四次建设期开始年份不可以小于等于第三次建设期的结束年份，程序将自动修改相关参数！")
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 2).Value + 1
        End If
        '结束年份不可以小于开始年份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 10).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 10).Value Then
            MsgBox("第四次建设期结束年份不可以小于开始年份，程序将自动修改相关参数！")
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 10).Value + 1
        End If
        '结束年份等于开始年份，但是结束月份不可以小于开始月份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 10).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 12).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 12).Value Then
            MsgBox("第四次建设期结束年份等于开始年份时，结束月份不可以小于开始月份，程序将自动修改相关参数！")
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 12).Value + 1 <= 12 Then
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 12).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 12).Value + 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(35, 12).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 12).Value - 1
            End If
        End If
        '------------------------------------------------------------------------------------------------------------
        '第五次建设期
        '第五次开始年份不可以小于第四次结束年份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 2).Value <= ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 10).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 2).Value > 0 Then
            MsgBox("第五次建设期开始年份不可以小于等于第四次建设期的结束年份，程序将自动修改相关参数！")
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(36, 10).Value + 1
        End If
        '结束年份不可以小于开始年份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 2).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 2).Value Then
            MsgBox("第五次建设期结束年份不可以小于开始年份，程序将自动修改相关参数！")
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 2).Value + 1
        End If
        '结束年份等于开始年份，但是结束月份不可以小于开始月份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 2).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 4).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 4).Value Then
            MsgBox("第五次建设期结束年份等于开始年份时，结束月份不可以小于开始月份，程序将自动修改相关参数！")
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 4).Value + 1 <= 12 Then
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 4).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 4).Value + 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 4).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 4).Value - 1
            End If
        End If
        '------------------------------------------------------------------------------------------------------------
        '第六次建设期
        '第六次开始年份不可以小于第五次结束年份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 10).Value <= ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 2).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 10).Value > 0 Then
            MsgBox("第六次建设期开始年份不可以小于等于第五次建设期的结束年份，程序将自动修改相关参数！")
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 2).Value + 1
        End If
        '结束年份不可以小于开始年份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 10).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 10).Value Then
            MsgBox("第六次建设期结束年份不可以小于开始年份，程序将自动修改相关参数！")
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 10).Value + 1
        End If
        '结束年份等于开始年份，但是结束月份不可以小于开始月份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 10).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 12).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 12).Value Then
            MsgBox("第六次建设期结束年份等于开始年份时，结束月份不可以小于开始月份，程序将自动修改相关参数！")
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 12).Value + 1 <= 12 Then
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 12).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 12).Value + 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(67, 12).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 12).Value - 1
            End If
        End If
        '------------------------------------------------------------------------------------------------------------
        '第七次建设期
        '第七次开始年份不可以小于第六次结束年份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 2).Value <= ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 10).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 2).Value > 0 Then
            MsgBox("第七次建设期开始年份不可以小于等于第六次建设期的结束年份，程序将自动修改相关参数！")
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(68, 10).Value + 1
        End If
        '结束年份不可以小于开始年份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 2).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 2).Value Then
            MsgBox("第七次建设期结束年份不可以小于开始年份，程序将自动修改相关参数！")
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 2).Value + 1
        End If
        '结束年份等于开始年份，但是结束月份不可以小于开始月份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 2).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 4).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 4).Value Then
            MsgBox("第七次建设期结束年份等于开始年份时，结束月份不可以小于开始月份，程序将自动修改相关参数！")
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 4).Value + 1 <= 12 Then
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 4).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 4).Value + 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 4).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 4).Value - 1
            End If
        End If
        '------------------------------------------------------------------------------------------------------------
        '第八次建设期
        '第八次开始年份不可以小于第七次结束年份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 10).Value <= ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 2).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 10).Value > 0 Then
            MsgBox("第八次建设期开始年份不可以小于等于第七次建设期的结束年份，程序将自动修改相关参数！")
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 2).Value + 1
        End If
        '结束年份不可以小于开始年份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 10).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 10).Value Then
            MsgBox("第八次建设期结束年份不可以小于开始年份，程序将自动修改相关参数！")
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 10).Value + 1
        End If
        '结束年份等于开始年份，但是结束月份不可以小于开始月份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 10).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 12).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 12).Value Then
            MsgBox("第八次建设期结束年份等于开始年份时，结束月份不可以小于开始月份，程序将自动修改相关参数！")
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 12).Value + 1 <= 12 Then
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 12).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 12).Value + 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(99, 12).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 12).Value - 1
            End If
        End If
        '------------------------------------------------------------------------------------------------------------
        '第九次建设期
        '第九次开始年份不可以小于第八次结束年份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 2).Value <= ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 10).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 2).Value > 0 Then
            MsgBox("第九次建设期开始年份不可以小于等于第八次建设期的结束年份，程序将自动修改相关参数！")
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(100, 10).Value + 1
        End If
        '结束年份不可以小于开始年份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 2).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 2).Value Then
            MsgBox("第九次建设期结束年份不可以小于开始年份，程序将自动修改相关参数！")
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 2).Value + 1
        End If
        '结束年份等于开始年份，但是结束月份不可以小于开始月份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 2).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 2).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 4).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 4).Value Then
            MsgBox("第九次建设期结束年份等于开始年份时，结束月份不可以小于开始月份，程序将自动修改相关参数！")
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 4).Value + 1 <= 12 Then
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 4).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 4).Value + 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 4).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 4).Value - 1
            End If
        End If
        '------------------------------------------------------------------------------------------------------------
        '第十次建设期
        '第十次开始年份不可以小于第九次结束年份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 10).Value <= ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 2).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 10).Value > 0 Then
            MsgBox("第十次建设期开始年份不可以小于等于第九次建设期的结束年份，程序将自动修改相关参数！")
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 2).Value + 1
        End If
        '结束年份不可以小于开始年份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 10).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 10).Value Then
            MsgBox("第十次建设期结束年份不可以小于开始年份，程序将自动修改相关参数！")
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 10).Value + 1
        End If
        '结束年份等于开始年份，但是结束月份不可以小于开始月份
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 10).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 10).Value And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 12).Value < ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 12).Value Then
            MsgBox("第十次建设期结束年份等于开始年份时，结束月份不可以小于开始月份，程序将自动修改相关参数！")
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 12).Value + 1 <= 12 Then
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 12).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 12).Value + 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(131, 12).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(132, 12).Value - 1
            End If
        End If
        '计算一次工作簿
        ExcelApp.Calculate()
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '将之前已有的和确定建设期时间计划相关的计算结果全部重置回默认状态，防止出错
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("G205:G214").Value = 0
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("O205:P214").Value = 0
        '计算一次
        ExcelApp.Calculate()
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Range("L314:L323").Value = 0
        '计算一次
        ExcelApp.Calculate()
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '建设期每年投产的月份数计算清空
        ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").unProtect(Password:="wscjc")
        '手动输入的建设期投产月份数清空
        For i = 127 To 136 '行号
            For j = 3 To 33 '列号
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i, j).Value = 0
            Next
        Next
        '仅建设期彻底结束当年的投产月份数清空
        For i = 142 To 151 '行号
            For j = 3 To 33 '列号
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i, j).Value = 0
            Next
        Next
        '计算一次
        ExcelApp.Calculate()
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '清空逐年投产月份（自动计算和手动计算）
        For i = 1 To 10
            For j = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(126 + i, j).Value = 0 '手动
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(141 + i, j).Value = 0 '自动
            Next
        Next
        For i = 3 To 33
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value = 0 '汇总
        Next
        '计算一次
        ExcelApp.Calculate()
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '将投资发生年份从小到大排序并且踢出为0的项，同时归类对应的开始月份和结束月份
        For i = 205 To 304
            '投资发生年份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(i, 2).Value > 0 Then
                j1 = j1 + 1 '计数加1
                '记录下投资年份
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(204 + j1, 7).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(i, 2).Value
            End If
            '开始月份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(i, 10).Value > 0 Then
                j2 = j2 + 1
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(204 + j2, 15).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(i, 10).Value
            End If
            '结束月份
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(i, 11).Value > 0 Then
                j3 = j3 + 1
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(204 + j3, 16).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(i, 11).Value
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
        '计算建设期内，每次投资年份中的投产月份数
        For i = 314 To 323
            For j = 314 To 323
                If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(i, 4).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(i, 4).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(j, 11).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(j, 12).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(i, 7).Value
                End If
            Next
        Next
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
        '如果手动输入的建设期投产年份大于建设期最大年份，报错
        '求建设期年份序号最大值
        Dim JSQNFXHZDZ As Integer = 0
        For i = 205 To 214 '建设期年份序号
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(i, 7).Value > JSQNFXHZDZ Then
                JSQNFXHZDZ = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(i, 7).Value
            End If
        Next
        For i = 314 To 323 '手动输入的建设期投产年份序号
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(i, 4).Value > JSQNFXHZDZ Then
                MsgBox("手动输入的建设期投产月份所在的年份大于了建设期最大年份，请检查！")
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
        '将已有的自动计算的建设期投产月份和手动输入的建设期投产月份全部重置回0，防止出错
        '将是否处于建设期判断结果清空，防止出错
        '第一到第十次投资发生年份
        For i = 3 To 33  '投资计划与资金筹措表行号
            For j = 1 To 10 '投资次数序号
                '手动输入的投产月份
                If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(126 + j, i).Value > 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(126 + j, i).Value = 0
                End If
                '自动计算的投产月份
                If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(141 + j, i).Value > 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(141 + j, i).Value = 0
                End If
            Next
            '将建设期投产月份汇总结果清空
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value = 0
            '将是否是建设期的结果清空
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(162, i).Value = 0
        Next
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
        '判断哪些年份属于建设期，是建设期的设置为1，不是的是0
        For i = 205 To 214 '建设期时间计划表列号
            For j = 3 To 33 '投资计划与资金筹措表行号
                If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(161, j).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(i, 7).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(162, j).Value = 1
                End If
            Next
        Next
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
        '仅建设期彻底结束当年的投产月份数
        Dim JS2 As Integer = 0 '投资年份计数
        '第一到第十次投资发生年份
        For i = 3 To 21 Step 2
            JS2 = JS2 + 1
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(146, i).Value > 0 Then
                For j = 3 To 33 '投资计划与资金筹措表列号
                    If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(141, j).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(146, i).Value Then
                        ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(141 + JS2, j).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(166, i).Value
                    Else
                        ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(141 + JS2, j).Value = 0
                    End If
                Next
            Else
                For j = 3 To 33 '投资计划与资金筹措表列号
                    ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(141 + JS2, j).Value = 0
                Next
            End If
        Next
        '计算一次工作簿
        ExcelApp.Calculate()
        '建设期每年投产的月份数计算
        '手动输入的建设期投产月份数（不包括建设期结束当年，也就是建设期最后一年）
        Dim nfxh As Integer = 0
        Dim JS1 As Integer = 0
        Dim JS4 As Integer = 0
        For i = 314 To 323
            nfxh = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(i, 11).Value
            JS1 = JS1 + 1
            If nfxh > 0 Then
                JS4 = JS4 + 1
                For k = 3 To 33 '投资计划与资金筹措表列号
                    If nfxh = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(126, k).Value Then
                        ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(126 + JS4, k).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(i, 12).Value
                    Else
                        ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(126 + JS4, k).Value = 0
                    End If
                Next
            Else
                For k = 3 To 33 '投资计划与资金筹措表列号
                    If nfxh = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(126, k).Value Then
                        ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(126 + JS1, k).Value = 0
                    End If
                Next
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
        '检查建设期是否跨度超过2年，如果超过，弹出提醒
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(162, 7).Value > 2 Then
            MsgBox("建设期时间跨度超过2年，如果在建设期内有生产需要计提折旧和摊销，建议在估算表中将折旧摊销计算方式改为方法二，或者将建设期内投产月份设置为每年的1月。")
        End If
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '解锁表格
        ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").unProtect(Password:="wscjc")
        '将所得税征收比例重置回100%
        For i = 2 To 31
            ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(53, i + 3).Value = 1
        Next
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(44, 18).Value = "常规设置"
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
        '针对计算出的建设期生产月份数，汇总建设期生产月份数计算结果
        Dim CFNFXH(20) As Integer '用于储存自动计算出的建设期投产月份数与手动输入的建设期投产月份数重复的年份序号
        Dim JS5 As Integer = 0
        For i = 1 To 10 '投资计划与资金筹措表行号
            For j = 3 To 33 '投资计划与资金筹措表列号
                '处于建设期，手动的月份大于0
                If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(162, j).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(126 + i, j).Value > 0 Then '处于建设期
                    JS5 = JS5 + 1
                    CFNFXH(JS5) = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(141, j).Value
                    '计算出汇总的逐年投产月份数，保留自动计算的和手动计算的两种过程值，以手动输入的为主
                    ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, j).Value = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(138, j).Value
                    '处于建设期，手动的月份等于=0，自动的月份大于0
                ElseIf ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(162, j).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(126 + i, j).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(141 + i, j).Value > 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, j).Value = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(153, j).Value
                    '处于建设期，手动的月份等于=0，自动的月份=0
                ElseIf ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(162, j).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(126 + i, j).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(141 + i, j).Value = 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, j).Value = 0
                End If
            Next
        Next
        '计算一次工作簿
        ExcelApp.Calculate()
        '建设期内的投产月份合并（再次合并，防止遗漏）
        For i = 3 To 33 '投资计划与资金筹措表列号
            If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(162, i).Value = 1 Then '处于建设期
                If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(138, i).Value > 0 Then '以手动输入的结果为主
                    ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(138, i).Value '手动输入=0.自动计算大于0
                ElseIf ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(138, i).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(153, i).Value > 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(153, i).Value
                ElseIf ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(138, i).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(153, i).Value = 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value = 0
                End If
            End If
        Next
        '计算一次工作簿
        ExcelApp.Calculate()
        '不处于建设期的投产月份全部设置为12
        For i = 3 To 33 '投资计划与资金筹措表列号
            If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(162, i).Value = 0 Then '不处于建设期
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value = 12
            End If
        Next
        '计算一次工作簿
        ExcelApp.Calculate()
        Dim CFNFXH_1(JS5 - 1) As Integer
        For i = 0 To JS5 - 1
            CFNFXH_1(i) = CFNFXH(i + 1)
        Next
        '将自动计算出的建设期投产月份数与手动输入的建设期投产月份数重复的年份序号显示出来
        Dim XianShi_1 As String = Nothing
        For Each XXX In CFNFXH_1
            XianShi_1 = XianShi_1 & XXX.ToString & "   "
        Next
        If JS5 > 0 Then
            MsgBox("两种方法计算出的建设期生产月份数量存在冲突和重复，程序将以手动输入的建设期投产月份为依据进行计算！该建设期年份序号为：" & vbCrLf & XianShi_1)
        End If
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
        Dim MYTCYFNF(20) As Integer
        Dim JS6 As Integer = 0
        For i = 3 To 21 Step 2
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(146, i).Value > 1 Then '不包括第一年
                For j = 3 To 33 '投资计划与资金筹措表列号
                    If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(156, j).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(146, i).Value Then
                        '是建设期，年份大于1，且当年生产月份数为0，则报错
                        If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, j).Value = 0 Then
                            JS6 = JS6 + 1
                            MYTCYFNF(JS6) = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(156, j).Value
                        End If
                    End If
                Next
            End If
        Next
        Dim MYTCYFNF_1(JS6) As Integer
        For i = 0 To JS6 - 1
            MYTCYFNF_1(i) = MYTCYFNF(i + 1)
        Next
        '将建设期中存在没有生产月份数的年份序号显示出来
        Dim XianShi_2 As String = Nothing
        For Each XXX In MYTCYFNF_1
            If XXX > 0 Then
                XianShi_2 = XianShi_2 & XXX.ToString & "   "
            End If
        Next
        If JS6 > 0 Then
            MsgBox("建设期中存在没有生产月份数的年份，请检查是否输入有误！该建设期年份序号为：" & vbCrLf & XianShi_2)
        End If
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '弹窗显示逐年投产月份数量
        '计算一次工作簿
        ExcelApp.Calculate()
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        Dim ZNTCYFS(jsnx - 1) As Integer
        Dim ZNTCYFS_NFXH(jsnx - 1) As Integer '年份序号
        For i = 0 To jsnx - 1
            ZNTCYFS(i) = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 3).Value
            ZNTCYFS_NFXH(i) = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(156, i + 3).Value
        Next
        Dim XianShi_3 As String = Nothing
        For i = 0 To jsnx - 1
            XianShi_3 = XianShi_3 & ZNTCYFS(i).ToString & "(" & ZNTCYFS_NFXH(i).ToString & ")  "
        Next
        MsgBox("计算年限内逐年投产的月份数量为：" & vbCrLf & XianShi_3)
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '项目计算年数改变后改变的计算用系数
        Call 项目计算年限变化后改变相关系数()
        '折旧年数改变后改变的计算用系数
        Call 折旧计算年限变化后改变相关系数()
        '无形资产摊销年限改变后改变的计算用系数
        Call 无形资产摊销年限变化后改变相关系数()
        '长期贷款还款年数改变后改变的计算用系数
        Call 长期贷款计算年限变化后改变相关系数()
        '年数总和法逐年折旧摊销系数
        Call 年数总和法逐年折旧摊销系数()
        '计算长期贷款是否处于宽限期
        Call 是否处于长期贷款宽限期()
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————       
        '第2到第10次投资年份输入发生变化时修改计算系数
        Call 第二至第五次投资年份改变后修改相关计算系数()
        Call 第六至第十次投资年份改变后修改相关计算系数()
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '建设期可抵扣增值税计算系数
        Call 建设期增值税抵扣系数()
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '计算一次负荷率
        Call 直接输入综合负荷率()
        Call 分投资逐次输入负荷率()
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '将部分设置重置回默认状态
        Call 将部分设置重置回默认状态()
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
        '计算流动资金
        Call 流动资金相关计算()
        '重新计算投资回收期
        Call 投资回收期计算()
        '锁定表格
        ExcelApp.ThisWorkbook.Worksheets("估算表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Protect(Password:="wscjc")
        '重新打开excel自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
        MsgBox("确定建设期时间计划完成！")
    End Sub
    Sub 将部分设置重置回默认状态()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '解锁表格
        ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入税收表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").unProtect(Password:="wscjc")
        '————————————————————————————————————————————————————————————————————————————————————————        
        '在确认建设期时间计划和修改估算表的参数时，将部分设置重置回默认值（折旧摊销和长期贷款计算参数、建设期资金运用方式）
        '————————————————————————————————————————————————————————————————————————————————————————        
        '写入每次投资的折旧摊销和长期贷款计算参数（每次投资均相同，根据估算表中的参数进行设置）
        '计算一次工作簿
        ExcelApp.Calculate()
        '写入折旧摊销系数默认值
        For i = 4 To 13
            '固定资产折旧年限
            ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 22).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 7).Value
            '固定资产残值率
            ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 24).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(17, 5).Value
            '无形资产摊销年限
            ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 26).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 7).Value
        Next
        '写入长期贷款还款年限和宽限年限默认值
        For i = 23 To 32
            '长期贷款还款年限
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 39).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value
        Next
        For i = 34 To 43
            '长期贷款宽限年限
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 39).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value
        Next
        '记录折旧摊销和长期贷款计算模式
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(166, 7).Value = "相同"
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(167, 7).Value = "相同"
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '建设期资金运用方式重置回默认模式（占资本金比例，每次投资均相同）
        '计算一次工作簿
        ExcelApp.Calculate()
        '将资本金比例和建设期贷款利率设置为默认情况（每次投资均相同，占动态投资比例，根据估算表中的参数进行设置、增值税退税计算）
        Dim JSQZJYYFSSZ As New 建设期资金运用方式设置
        '读取资本金(动态)比例，建设期贷款利率
        Dim ZBJBL = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(14, 7).Value
        Dim DKLL = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(9, 9).Value
        Call JSQZJYYFSSZ.建设期默认资金运用模式(ZBJBL, DKLL)
        '将逐次设置资本金运用方式计数次数重置回0
        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(170, 11).Value = 0
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '将投资各方收益计算参数重置回默认值
        '根据在<建设期时间计划表>中输入的投资各方出资比例，将投资各方的出资比例、利润分配比例、资产处置比例设置为一样的值           
        '将输入写入Excel
        '投资方1
        ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(4, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(163, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(5, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(163, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(6, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(163, 2).Value / 100
        '投资方2
        ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(4, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(5, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(6, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 2).Value / 100
        '投资方3
        ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(4, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(5, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(6, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 2).Value / 100
        '投资方4
        ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(4, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(166, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(5, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(166, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(6, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(166, 2).Value / 100
        '投资方5
        ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(4, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(167, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(5, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(167, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(6, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(167, 2).Value / 100
        '设置投资方1~5现金流量表的隐藏和显示         
        '先将5个表格都彻底隐藏
        ExcelApp.Worksheets("投资方1现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("投资方2现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("投资方3现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("投资方4现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        ExcelApp.Worksheets("投资方5现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        '设置需要显示的表格
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(163, 2).Value > 0 Then
            ExcelApp.Worksheets("投资方1现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
        End If
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 2).Value > 0 Then
            ExcelApp.Worksheets("投资方2现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
        End If
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 2).Value > 0 Then
            ExcelApp.Worksheets("投资方3现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
        End If
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(166, 2).Value > 0 Then
            ExcelApp.Worksheets("投资方4现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
        End If
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(167, 2).Value > 0 Then
            ExcelApp.Worksheets("投资方5现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
        End If
        '写入计算模式
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(168, 7).Value = "相同"
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '将计算期末回收固定资产残值重置回默认状态
        '前15年
        For i = 5 To 19
            ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(1078, i).Value = 0
        Next
        '后16年
        For i = 4 To 19
            ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(1080, i).Value = 0
        Next
        '将设置状态写入表格
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(53, 18).Value = "期末回收残值"
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '将增值税退税计算重置回默认
        '计算一次工作簿
        ExcelApp.Calculate()
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        For i = 3 To 33
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(171, i).Value <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(172, i).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(13, 7).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(172, i).Value = 0
            End If
        Next
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(45, 18).Value = "常规设置"
        ExcelApp.Calculate()
        '设置表格中的收入行隐藏或者显示
        '如果增值税退税比例为0，则隐藏增值税退税收入，前15年表格
        If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(172, 34).Value = 0 Then
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(16).EntireRow.Hidden = True
        Else
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(16).EntireRow.Hidden = False
        End If
        '如果增值税退税比例为0，则隐藏增值税退税收入，后15年表格
        If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(172, 34).Value = 0 Then
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(39).EntireRow.Hidden = True
        Else
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(39).EntireRow.Hidden = False
        End If
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '所得税逐年征收比例重置回默认状态
        '检测目前的逐年所得税征收比例，如果在计算年限内，前一个是100%，当前是0，则将0变成100%
        For i = 1 To 31
            If ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(49, i + 3).Value <= jsnx Then
                If ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(53, i + 3 - 1).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(53, i + 3).Value = 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(53, i + 3).Value = 1
                End If
            Else
                ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(53, i + 3).Value = 0
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '各种年限系数的计算开始年份（补贴收入、销售收入和成本）
        Dim JSKSNF As Integer = 0
        For i = 3 To 33
            If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value > 0 Then
                JSKSNF = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(156, i).Value
                Exit For
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '补贴收入逐年计算系数（第2年到计算期最后一年，每年都是100%）
        '计算一次工作簿
        ExcelApp.Calculate()
        For i = 3 To 33
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= JSKSNF And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 1 '补贴收入1
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 1 '补贴收入2
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(165, i).Value = 1 '光伏补贴收入
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(165, i).Value = 0
            End If
        Next
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 18).Value = "保持每年100%"
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(43, 18).Value = "保持每年100%"
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '接入费计算模式（第2年到计算期最后一年，根据逐年达产率增加值计算）
        '计算一次工作簿
        ExcelApp.Calculate()
        '第1年
        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, 3).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, 3).Value
        '第2-15年
        For i = 4 To 17
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= 2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i - 1).Value)
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
            End If
        Next
        '第16年
        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, 18).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, 3).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, 17).Value)
        '第17-31年
        For i = 19 To 33
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= 2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15 - 1).Value)
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
            End If
        Next
        '将小于0的结果设置为0
        For i = 3 To 33
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value < 0 Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
            End If
        Next
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(41, 18).Value = "逐年达产率增加值"
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '部分销售收入和成本逐年计算系数
        For i = 3 To 33
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= JSKSNF And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 1 '充电桩收入
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(148, i).Value = 1 '常规设备修理费成本
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(17, 11).Value = "折算" Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(143, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12 '风电收入
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(145, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12 '购电容量费成本
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(146, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12 '城市管廊费成本
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(147, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12 '人员工资成本
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12 '充电桩收入
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(48, 18).Value = "逐年投产月份比例"
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(49, 18).Value = "逐年投产月份比例"
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(50, 18).Value = "逐年投产月份比例"
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(51, 18).Value = "逐年投产月份比例"
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(52, 18).Value = "逐年投产月份比例"
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(143, i).Value = 1 '风电收入
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(145, i).Value = 1 '购电容量费成本
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(146, i).Value = 1 '城市管廊费成本
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(147, i).Value = 1 '人员工资成本
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 1  '充电桩收入
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(48, 18).Value = "保持每年100%"
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(49, 18).Value = "保持每年100%"
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(50, 18).Value = "保持每年100%"
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(51, 18).Value = "保持每年100%"
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(52, 18).Value = "保持每年100%"
                End If
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(143, i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(145, i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(146, i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(147, i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(148, i).Value = 0
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '锁定表格
        ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Protect(Password:="wscjc")
    End Sub
    Sub 打开表格自动运行()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '启动迭代计算
        ExcelApp.Application.Iteration = True
        '打开自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '打开事件触发
        ExcelApp.Application.EnableEvents = True
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
        '锁定表格
        Call 锁定表格()
        '自保护程序
        Call 自保护程序()
        '验证Excel表格的更新时间
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 29).Value < 20191203 Then
            MsgBox（"Excel文件版本已过期，无法配合最新的计算程序使用，需要更新到最新版本才可以使用！本Excel文件仅可以查看已有的计算结果，不可能用于新的计算！"）
            '跳过投资各方收益表格操作，防止多次报错
            Exit Sub
        End If
        Call 投资各方收益率表格操作()
    End Sub
    Function Excel版本号验证()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim ZTJC As Integer
        '验证Excel表格的更新时间
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 29).Value < 20191203 Then
            MsgBox（"Excel文件版本已过期，无法配合最新的计算程序使用，需要更新到最新版本才可以使用！本Excel文件仅可以查看已有的计算结果，不可能用于新的计算！"）
            ZTJC = 1
        Else
            ZTJC = 0
        End If
        Call 自保护程序()
        '返回结果
        Return ZTJC
    End Function
    Sub 资本金税后收益率反算收入单价(FSLJDJSCSMax As Integer)
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '资本金税后收益率，资本金税后收益率，资本金税后收益率，
        '反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入
        '实例化计算进度显示窗体
        Dim Form1 As New 计算进度显示
        Dim SRBH As Integer '收入编号
        Dim SRDJ '收入单价
        Dim SRDJWT As Double = 0 '收入单价微调（用于修正计算结果的误差）
        Dim SRZL '收入总量
        Dim JSCSMax As Integer = FSLJDJSCSMax '计算次数最大值
        Dim WTJSBC '微调计算步长
        For i = 13 To 27 '收入&成本表中的有收入的项
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 7).Value > 0 Then '存在收入
                SRBH = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 27).Value
                SRDJ = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value
                SRZL = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 6).Value
                For j = 43 To 67
                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 27).Value = SRBH And SRBH > 0 And SRDJ > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 8).Value = "Y" Then '找到编号相同的收入
                        If SRDJ < 1 Then '根据收入单价情况确定不同的微调计算步长
                            WTJSBC = 0.0001
                        ElseIf SRDJ >= 1 And SRDJ < 10 Then
                            WTJSBC = 0.001
                        ElseIf SRDJ >= 10 And SRDJ < 100 Then
                            WTJSBC = 0.01
                        ElseIf SRDJ >= 100 And SRDJ < 1000 Then
                            WTJSBC = 0.1
                        ElseIf SRDJ >= 1000 And SRDJ < 10000 Then
                            WTJSBC = 1
                        Else
                            WTJSBC = 10
                        End If
                        '判断此时的内部收益率和设定的临界点内部收益率的大小关系
                        '如果此时的内部收益率大于设定的临界点内部收益率，则将目前的单价往下减
                        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value > ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                            For m = 1 To JSCSMax
                                '计算进度显示
                                Form1.Show()
                                Form1.Label1.Text = "正在进行" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算资本金临界点计算,已计算" & m & "次。"
                                Form1.TopMost = True
                                System.Windows.Forms.Application.DoEvents()
                                '反算临界点
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJ * (1 - m / JSCSMax) '每次改变原单价的千分之0.5
                                '计算流动资金
                                'Call 流动资金相关计算()
                                '设置跳出条件
                                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value <= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                    SRDJWT = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value, 4) '用于微调收入单价的初始数据
                                    'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.WorksheetFunction.RoundUp(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value * (1 + 1 / JSCSMax), 3) '将反算出来的单价记在表格中
                                    Exit For
                                End If
                            Next
                            '计算流动资金
                            Call 流动资金相关计算()
                            '排除收益率恰好等于设定的收益率的情况，进行微调操作
                            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                '加入微调修正，使得计算结果更加准确
                                For m = 1 To JSCSMax * 10 '此时的内部收益率小于设定的临界点内部收益率，则将目前的SRDJWT往上加
                                    '计算进度显示
                                    Form1.Show()
                                    Form1.Label1.Text = "正在最后微调" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算资本金临界点计算,已微调" & m & "次。"
                                    Form1.TopMost = True
                                    System.Windows.Forms.Application.DoEvents()
                                    '反算临界点
                                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJWT + WTJSBC * m '每次增加WTJSBC
                                    '计算流动资金
                                    Call 流动资金相关计算()
                                    '设置跳出条件
                                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value '将反算出来的单价记在表格中
                                        Exit For
                                    End If
                                Next
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = SRDJWT '将反算出来的单价记在表格中
                            End If
                            '返回初始单价
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJ
                            '计算流动资金
                            Call 流动资金相关计算()
                            '计算此时的单价比原值变化了百分之多少，从而求出反算总量
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 7).Value = ExcelApp.WorksheetFunction.RoundUp((ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value / SRDJ) * SRZL, 2)
                        End If
                        '如果此时的内部收益率小于设定的临界点内部收益率，则将目前的单价往上加
                        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                            For m = 1 To JSCSMax * 10
                                '计算进度显示
                                Form1.Show()
                                Form1.Label1.Text = "正在进行" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算资本金临界点计算,已计算" & m & "次。"
                                Form1.TopMost = True
                                System.Windows.Forms.Application.DoEvents()
                                '反算临界点
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJ * (1 + m / JSCSMax) '每次改变原单价的千分之0.5
                                '计算流动资金
                                'Call 流动资金相关计算()
                                '设置跳出条件
                                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                    SRDJWT = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value, 4) '用于微调收入单价的初始数据
                                    'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.WorksheetFunction.RoundUp(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value, 3) '将反算出来的单价记在表格中
                                    Exit For
                                End If
                            Next
                            '计算流动资金
                            Call 流动资金相关计算()
                            '排除收益率恰好等于设定的收益率的情况， 进行微调操作
                            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value > ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                '加入微调修正，使得计算结果更加准确
                                For m = 1 To JSCSMax '此时的内部收益率大于设定的临界点内部收益率，则将目前的SRDJWT往下减
                                    '计算进度显示
                                    Form1.Show()
                                    Form1.Label1.Text = "正在最后微调" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算资本金临界点计算,已微调" & m & "次。"
                                    Form1.TopMost = True
                                    System.Windows.Forms.Application.DoEvents()
                                    '反算临界点
                                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJWT - WTJSBC * m '每次减少WTJSBC
                                    '计算流动资金
                                    Call 流动资金相关计算()
                                    '设置跳出条件
                                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then '排除收益率恰好等于的情况
                                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value  '将反算出来的单价记在表格中
                                        Exit For
                                    ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value + WTJSBC
                                        Exit For
                                    End If
                                Next
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = SRDJWT '将反算出来的单价记在表格中
                            End If
                            '返回初始单价
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJ
                            '计算流动资金
                            Call 流动资金相关计算()
                            '计算此时的单价比原值变化了百分之多少，从而求出反算总量
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 7).Value = ExcelApp.WorksheetFunction.RoundUp((ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value / SRDJ) * SRZL, 2)
                        End If
                    End If
                Next
            End If
        Next
        Form1.Close()
    End Sub
    Sub 资本金税后收益率反算成本单价(FSLJDJSCSMax As Integer)
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '资本金税后收益率，资本金税后收益率，资本金税后收益率，
        '反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本
        '实例化计算进度显示窗体
        Dim Form1 As New 计算进度显示
        Dim CBBH As Integer '成本编号
        Dim CBDJ '成本单价
        Dim CBDJWT As Double = 0 '成本单价微调（用于修正计算结果的误差）
        Dim CBZL '成本总量
        Dim JSCSMax As Integer = FSLJDJSCSMax '计算次数最大值
        Dim WTJSBC '微调计算步长
        For i = 13 To 27 '收入&成本表中的有收入的项
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 14).Value > 0 Then '存在成本
                CBBH = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 28).Value
                CBDJ = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value
                CBZL = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 13).Value
                For j = 43 To 67
                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 28).Value = CBBH And CBBH > 0 And CBDJ > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 8).Value = "Y" Then '找到编号相同的收入
                        If CBDJ < 1 Then '根据成本单价情况确定不同的微调计算步长
                            WTJSBC = 0.0001
                        ElseIf CBDJ >= 1 And CBDJ < 10 Then
                            WTJSBC = 0.001
                        ElseIf CBDJ >= 10 And CBDJ < 100 Then
                            WTJSBC = 0.01
                        ElseIf CBDJ >= 100 And CBDJ < 1000 Then
                            WTJSBC = 0.1
                        ElseIf CBDJ >= 1000 And CBDJ < 10000 Then
                            WTJSBC = 1
                        Else
                            WTJSBC = 10
                        End If
                        '判断此时的内部收益率和设定的临界点内部收益率的大小关系
                        '如果此时的内部收益率小于设定的临界点内部收益率，则将目前的单价往下减
                        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                            For m = 1 To JSCSMax
                                '计算进度显示
                                Form1.Show()
                                Form1.Label1.Text = "正在进行" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算资本金临界点计算,已计算" & m & "次。"
                                Form1.TopMost = True
                                System.Windows.Forms.Application.DoEvents()
                                '反算临界点
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJ * (1 - m / JSCSMax)
                                '计算流动资金
                                'Call 流动资金相关计算()
                                '设置跳出条件
                                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                    CBDJWT = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value, 4) '用于微调成本单价的初始数据
                                    'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.WorksheetFunction.RoundUp(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value, 3) '将反算出来的单价记在表格中
                                    Exit For
                                End If
                            Next
                            '计算流动资金
                            Call 流动资金相关计算()
                            '排除收益率恰好等于设定的收益率的情况，进行微调操作
                            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value > ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                '加入微调修正，使得计算结果更加准确
                                For m = 1 To JSCSMax * 10 '此时的内部收益率大于设定的临界点内部收益率，则将目前的CBDJWT往上加
                                    '计算进度显示
                                    Form1.Show()
                                    Form1.Label1.Text = "正在最后微调" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算资本金临界点计算,已微调" & m & "次。"
                                    Form1.TopMost = True
                                    System.Windows.Forms.Application.DoEvents()
                                    '反算临界点
                                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJWT + WTJSBC * m '每次增加WTJSBC
                                    '计算流动资金
                                    Call 流动资金相关计算()
                                    '设置跳出条件
                                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then '排除收益率恰好等于的情况
                                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value  '将反算出来的单价记在表格中
                                        Exit For
                                    ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value - WTJSBC
                                        Exit For
                                    End If
                                Next
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = CBDJWT '将反算出来的单价记在表格中
                            End If
                            '返回初始单价
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJ
                            '计算流动资金
                            Call 流动资金相关计算()
                            '计算此时的单价比原值变化了百分之多少，从而求出反算总量
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 7).Value = ExcelApp.WorksheetFunction.RoundUp((ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value / CBDJ) * CBZL, 2)
                        End If
                        '如果此时的内部收益率大于设定的临界点内部收益率，则将目前的单价往上加
                        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                            For m = 1 To JSCSMax * 10
                                '计算进度显示
                                Form1.Show()
                                Form1.Label1.Text = "正在进行" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算资本金临界点计算,已计算" & m & "次。"
                                Form1.TopMost = True
                                System.Windows.Forms.Application.DoEvents()
                                '反算临界点
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJ * (1 + m / JSCSMax) '每次改变原单价的千分之0.5
                                '计算流动资金
                                'Call 流动资金相关计算()
                                '设置跳出条件
                                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value <= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                    CBDJWT = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value, 4) '用于微调成本单价的初始数据
                                    'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.WorksheetFunction.RoundUp(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value * (1 - 1 / JSCSMax), 3) '将反算出来的单价记在表格中
                                    Exit For
                                End If
                            Next
                            '计算流动资金
                            Call 流动资金相关计算()
                            '排除收益率恰好等于设定的收益率的情况，进行微调操作
                            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                '加入微调修正，使得计算结果更加准确
                                For m = 1 To JSCSMax '此时的内部收益率小于设定的临界点内部收益率，则将目前的CBDJWT往下减
                                    '计算进度显示
                                    Form1.Show()
                                    Form1.Label1.Text = "正在最后微调" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算资本金临界点计算,已微调" & m & "次。"
                                    Form1.TopMost = True
                                    System.Windows.Forms.Application.DoEvents()
                                    '反算临界点
                                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJWT - WTJSBC * m '每次减少WTJSBC
                                    '计算流动资金
                                    Call 流动资金相关计算()
                                    '设置跳出条件
                                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value '将反算出来的单价记在表格中
                                        Exit For
                                    End If
                                Next
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = CBDJWT '将反算出来的单价记在表格中
                            End If
                            '返回初始单价
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJ
                            '计算流动资金
                            Call 流动资金相关计算()
                            '计算此时的单价比原值变化了百分之多少，从而求出反算总量
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 7).Value = ExcelApp.WorksheetFunction.RoundUp((ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value / CBDJ) * CBZL, 2)
                        End If
                    End If
                Next
            End If
        Next
        Form1.Close()
    End Sub
    Sub 资本金税后收益率反算静态投资(FSLJDJSCSMax As Integer)
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '实例化计算进度显示窗体
        Dim Form1 As New 计算进度显示
        Dim TZHH(50) As Integer '数组，储存投资所在的行号
        Dim TZHHJS As Integer = 0 '投资行号计数
        Dim TZJECC(200) '数组，储存各种投资金额
        Dim JTTZWT(200) '数组，静态投资微调
        Dim JTTZZEWT As Double = 0 '静态投资总额微调
        Dim WTJSBC As Integer '微调计算步长
        Dim JSCSMax As Integer = FSLJDJSCSMax '计算次数最大值
        For i = 30 To 110 '检索估算表中投资编号，投资编号都是1表示
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, 27).Value = 1 Then
                TZHHJS = TZHHJS + 1
                TZHH(TZHHJS) = i '行号储存进数组，从下标1开始
            End If
        Next
        Dim TZJESLJS As Integer = 0 '投资金额数量计数
        '将各种投资金额储存在数组中
        For i = 1 To TZHHJS '行
            For j = 3 To 11 Step 2 '列
                TZJESLJS = TZJESLJS + 1
                TZJECC(TZJESLJS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(i), j).Value
            Next
        Next
        Dim TZmin = TZJECC(1）  '投资金额中的最小值，初始值为数组中的第一个数
        For i = 1 To 200
            If TZJECC(i) < TZmin And TZJECC(i) > 0 Then
                TZmin = TZJECC(i）
            End If
        Next
        WTJSBC = ExcelApp.WorksheetFunction.RoundUp(TZmin / (JSCSMax * 10), 0)
        Dim JS As Integer = 0
        '判断此时的内部收益率和设定的临界点内部收益率的大小关系
        '如果此时的内部收益率小于设定的临界点内部收益率，则将目前的单价往下减
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 8).Value = "Y" Then
            For i = 1 To JSCSMax
                '计算进度显示
                Form1.Show()
                Form1.Label1.Text = "正在进行静态投资反算资本金临界点计算,已计算" & i & "次。"
                Form1.TopMost = True
                System.Windows.Forms.Application.DoEvents()
                '改变分年静态投资
                '遍历所有的投资，每个值都变化
                For j = 1 To TZHHJS '行
                    For m = 3 To 11 Step 2 '列
                        JS = (j - 1) * 5 + Int(m / 2)
                        If TZJECC(JS) > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = TZJECC(JS) * (1 - i / JSCSMax)
                            '计算流动资金
                            'Call 流动资金相关计算()
                        End If
                    Next
                Next
                '设置跳出条件
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                    '将静态投资储存在JTTZWT数组
                    TZJESLJS = 0
                    For k = 1 To TZHHJS '行
                        For l = 3 To 11 Step 2 '列
                            TZJESLJS = TZJESLJS + 1
                            JTTZWT(TZJESLJS) = Math.Round(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(k), l).Value, 2)
                        Next
                    Next
                    JTTZZEWT = ExcelApp.WorksheetFunction.RoundUp(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value, 2) '将反算出来的总投资记在表格中
                    'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = ExcelApp.WorksheetFunction.RoundUp(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value, 2) '将反算出来的总投资记在表格中
                    Exit For
                End If
            Next
            '计算流动资金
            Call 流动资金相关计算()
            '排除收益率恰好等于设定的收益率的情况，进行微调操作
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value > ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                '加入微调修正，使得计算结果更加准确
                For i = 1 To JSCSMax * 10 '此时的内部收益率大于设定的临界点内部收益率，则将目前的JTTZWT往上加
                    '计算进度显示
                    Form1.Show()
                    Form1.Label1.Text = "正在最后微调静态投资反算资本金临界点计算,已微调" & i & "次。"
                    Form1.TopMost = True
                    System.Windows.Forms.Application.DoEvents()
                    '遍历所有的投资，每个值都变化
                    For j = 1 To TZHHJS '行
                        For m = 3 To 11 Step 2 '列
                            JS = (j - 1) * 5 + Int(m / 2)
                            If TZJECC(JS) > 0 Then
                                'ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = JTTZWT(JS) + WTJSBC * i '每次增加WTJSBC
                                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = JTTZWT(JS) * (1 + i / (JSCSMax * 25))
                                '计算流动资金
                                Call 流动资金相关计算()
                            End If
                        Next
                    Next
                    '设置跳出条件
                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then '排除收益率恰好等于的情况
                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value  '将反算出来的总投资记在表格中 
                        Exit For
                    ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                        'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value - WTJSBC
                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value * (1 - 1 / (JSCSMax * 25))
                        Exit For
                    End If
                Next
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = JTTZZEWT '将反算出来的总投资记在表格中
            End If
            '静态总投资重置回默认值
            JS = 0
            For j = 1 To TZHHJS '行
                For m = 3 To 11 Step 2 '列
                    JS = (j - 1) * 5 + Int(m / 2)
                    If TZJECC(JS) > 0 Then
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = TZJECC(JS)
                        '计算流动资金
                        Call 流动资金相关计算()
                    End If
                Next
            Next
        End If
        '如果此时的内部收益率大于设定的临界点内部收益率，则将目前的单价往上加
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value > ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 8).Value = "Y" Then
            For i = 1 To JSCSMax * 10
                '计算进度显示
                Form1.Show()
                Form1.Label1.Text = "正在进行静态投资反算资本金临界点计算,已计算" & i & "次。"
                Form1.TopMost = True
                System.Windows.Forms.Application.DoEvents()
                '改变分年静态投资
                '遍历所有的投资，每个值都变化
                For j = 1 To TZHHJS '行
                    For m = 3 To 11 Step 2 '列
                        JS = (j - 1) * 5 + Int(m / 2)
                        If TZJECC(JS) > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = TZJECC(JS) * (1 + i / JSCSMax)
                            '计算流动资金
                            'Call 流动资金相关计算()
                        End If
                    Next
                Next
                '设置跳出条件
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value <= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                    '将静态投资储存在JTTZWT数组
                    TZJESLJS = 0
                    For k = 1 To TZHHJS '行
                        For l = 3 To 11 Step 2 '列
                            TZJESLJS = TZJESLJS + 1
                            JTTZWT(TZJESLJS) = Math.Round(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(k), l).Value, 2)
                        Next
                    Next
                    JTTZZEWT = ExcelApp.WorksheetFunction.RoundUp(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value, 2) '将反算出来的总投资记在表格中
                    'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = ExcelApp.WorksheetFunction.RoundUp(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value * (1 - 1 / 2000), 2) '将反算出来的总投资记在表格中
                    Exit For
                End If
            Next
            '计算流动资金
            Call 流动资金相关计算()
            '排除收益率恰好等于设定的收益率的情况，进行微调操作
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                '加入微调修正，使得计算结果更加准确
                For i = 1 To JSCSMax '此时的内部收益率小于设定的临界点内部收益率，则将目前的JTTZWT往下减
                    '计算进度显示
                    Form1.Show()
                    Form1.Label1.Text = "正在最后微调静态投资反算资本金临界点计算,已微调" & i & "次。"
                    Form1.TopMost = True
                    System.Windows.Forms.Application.DoEvents()
                    '遍历所有的投资，每个值都变化
                    For j = 1 To TZHHJS '行
                        For m = 3 To 11 Step 2 '列
                            JS = (j - 1) * 5 + Int(m / 2)
                            If TZJECC(JS) > 0 Then
                                'ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = JTTZWT(JS) - WTJSBC * i '每次减少WTJSBC
                                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = JTTZWT(JS) * (1 - i / (JSCSMax * 25))
                                '计算流动资金
                                Call 流动资金相关计算()
                            End If
                        Next
                    Next
                    '设置跳出条件
                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value  '将反算出来的总投资记在表格中 
                        Exit For
                    End If
                Next
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = JTTZZEWT '将反算出来的总投资记在表格中
            End If
            '静态总投资重置回默认值
            JS = 0
            For j = 1 To TZHHJS '行
                For m = 3 To 11 Step 2 '列
                    JS = (j - 1) * 5 + Int(m / 2)
                    If TZJECC(JS) > 0 Then
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = TZJECC(JS)
                        '计算流动资金
                        Call 流动资金相关计算()
                    End If
                Next
            Next
        End If
        '将反算出的静态投资结果保留为2位小数
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 8).Value = "Y" Then
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value, 2)
        End If
        Form1.Close()
    End Sub
    Sub 全投资税后收益率反算收入单价(FSLJDJSCSMax As Integer)
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '全投资税后收益率，全投资税后收益率，全投资税后收益率，
        '反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入，反算收入
        '实例化计算进度显示窗体
        Dim Form1 As New 计算进度显示
        Dim SRBH As Integer '收入编号
        Dim SRDJ '收入单价
        Dim SRDJWT As Double = 0 '收入单价微调（用于修正计算结果的误差）
        Dim SRZL '收入总量
        Dim JSCSMax As Integer = FSLJDJSCSMax '计算次数最大值
        Dim WTJSBC '微调计算步长
        For i = 13 To 27 '收入&成本表中的有收入的项
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 7).Value > 0 Then '存在收入
                SRBH = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 27).Value
                SRDJ = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value
                SRZL = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 6).Value
                For j = 43 To 67
                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 27).Value = SRBH And SRBH > 0 And SRDJ > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 8).Value = "Y" Then '找到编号相同的收入
                        If SRDJ < 1 Then '根据收入单价情况确定不同的微调计算步长
                            WTJSBC = 0.0001
                        ElseIf SRDJ >= 1 And SRDJ < 10 Then
                            WTJSBC = 0.001
                        ElseIf SRDJ >= 10 And SRDJ < 100 Then
                            WTJSBC = 0.01
                        ElseIf SRDJ >= 100 And SRDJ < 1000 Then
                            WTJSBC = 0.1
                        ElseIf SRDJ >= 1000 And SRDJ < 10000 Then
                            WTJSBC = 1
                        Else
                            WTJSBC = 10
                        End If
                        '判断此时的内部收益率和设定的临界点内部收益率的大小关系
                        '如果此时的内部收益率大于设定的临界点内部收益率，则将目前的单价往下减
                        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value > ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                            For m = 1 To JSCSMax
                                '计算进度显示
                                Form1.Show()
                                Form1.Label1.Text = "正在进行" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算全投资临界点计算,已计算" & m & "次。"
                                Form1.TopMost = True
                                System.Windows.Forms.Application.DoEvents()
                                '反算临界点
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJ * (1 - m / JSCSMax) '每次改变原单价的千分之0.5
                                '计算流动资金
                                'Call 流动资金相关计算()
                                '设置跳出条件
                                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value <= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                    SRDJWT = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value, 4) '用于微调收入单价的初始数据
                                    'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.WorksheetFunction.RoundUp(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value * (1 + 1 / JSCSMax), 3) '将反算出来的单价记在表格中
                                    Exit For
                                End If
                            Next
                            '计算流动资金
                            Call 流动资金相关计算()
                            '排除收益率恰好等于设定的收益率的情况，进行微调操作
                            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                '加入微调修正，使得计算结果更加准确
                                For m = 1 To JSCSMax * 10 '此时的内部收益率小于设定的临界点内部收益率，则将目前的SRDJWT往上加
                                    '计算进度显示
                                    Form1.Show()
                                    Form1.Label1.Text = "正在最后微调" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算全投资临界点计算,已微调" & m & "次。"
                                    Form1.TopMost = True
                                    System.Windows.Forms.Application.DoEvents()
                                    '反算临界点
                                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJWT + WTJSBC * m '每次增加WTJSBC
                                    '计算流动资金
                                    Call 流动资金相关计算()
                                    '设置跳出条件
                                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value '将反算出来的单价记在表格中
                                        Exit For
                                    End If
                                Next
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = SRDJWT '将反算出来的单价记在表格中
                            End If
                            '返回初始单价
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJ
                            '计算流动资金
                            Call 流动资金相关计算()
                            '计算此时的单价比原值变化了百分之多少，从而求出反算总量
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 7).Value = ExcelApp.WorksheetFunction.RoundUp((ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value / SRDJ) * SRZL, 2)
                        End If
                        '如果此时的内部收益率小于设定的临界点内部收益率，则将目前的单价往上加
                        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                            For m = 1 To JSCSMax * 10
                                '计算进度显示
                                Form1.Show()
                                Form1.Label1.Text = "正在进行" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算全投资临界点计算,已计算" & m & "次。"
                                Form1.TopMost = True
                                System.Windows.Forms.Application.DoEvents()
                                '反算临界点
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJ * (1 + m / JSCSMax) '每次改变原单价的千分之0.5
                                '计算流动资金
                                'Call 流动资金相关计算()
                                '设置跳出条件
                                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                    SRDJWT = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value, 4) '用于微调收入单价的初始数据
                                    'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.WorksheetFunction.RoundUp(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value, 3) '将反算出来的单价记在表格中
                                    Exit For
                                End If
                            Next
                            '计算流动资金
                            Call 流动资金相关计算()
                            '排除收益率恰好等于设定的收益率的情况，进行微调操作
                            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value > ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                '加入微调修正，使得计算结果更加准确
                                For m = 1 To JSCSMax '此时的内部收益率大于设定的临界点内部收益率，则将目前的SRDJWT往下减
                                    '计算进度显示
                                    Form1.Show()
                                    Form1.Label1.Text = "正在最后微调" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算全投资临界点计算,已微调" & m & "次。"
                                    Form1.TopMost = True
                                    System.Windows.Forms.Application.DoEvents()
                                    '反算临界点
                                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJWT - WTJSBC * m '每次减小WTJSBC
                                    '计算流动资金
                                    Call 流动资金相关计算()
                                    '设置跳出条件
                                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then '排除收益率恰好等于的情况
                                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value  '将反算出来的单价记在表格中
                                        Exit For
                                    ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value + WTJSBC
                                        Exit For
                                    End If
                                Next
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = SRDJWT '将反算出来的单价记在表格中
                            End If
                            '返回初始单价
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJ
                            '计算流动资金
                            Call 流动资金相关计算()
                            '计算此时的单价比原值变化了百分之多少，从而求出反算总量
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 7).Value = ExcelApp.WorksheetFunction.RoundUp((ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value / SRDJ) * SRZL, 2)
                        End If
                    End If
                Next
            End If
        Next
        Form1.Close()
    End Sub
    Sub 全投资税后收益率反算成本单价(FSLJDJSCSMax As Integer)
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '全投资税后收益率，全投资税后收益率，全投资税后收益率，
        '反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本，反算成本
        '实例化计算进度显示窗体
        Dim Form1 As New 计算进度显示
        Dim CBBH As Integer '成本编号
        Dim CBDJ '成本单价
        Dim CBDJWT As Double = 0 '成本单价微调（用于修正计算结果的误差）
        Dim CBZL '成本总量
        Dim JSCSMax As Integer = FSLJDJSCSMax '计算次数最大值
        Dim WTJSBC '微调计算步长
        For i = 13 To 27 '收入&成本表中的有收入的项
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 14).Value > 0 Then '存在成本
                CBBH = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 28).Value
                CBDJ = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value
                CBZL = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 13).Value
                For j = 43 To 67
                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 28).Value = CBBH And CBBH > 0 And CBDJ > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 8).Value = "Y" Then '找到编号相同的收入
                        If CBDJ < 1 Then '根据成本单价情况确定不同的微调计算步长
                            WTJSBC = 0.0001
                        ElseIf CBDJ >= 1 And CBDJ < 10 Then
                            WTJSBC = 0.001
                        ElseIf CBDJ >= 10 And CBDJ < 100 Then
                            WTJSBC = 0.01
                        ElseIf CBDJ >= 100 And CBDJ < 1000 Then
                            WTJSBC = 0.1
                        ElseIf CBDJ >= 1000 And CBDJ < 10000 Then
                            WTJSBC = 1
                        Else
                            WTJSBC = 10
                        End If
                        '判断此时的内部收益率和设定的临界点内部收益率的大小关系
                        '如果此时的内部收益率小于设定的临界点内部收益率，则将目前的单价往下减
                        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                            For m = 1 To JSCSMax
                                '计算进度显示
                                Form1.Show()
                                Form1.Label1.Text = "正在进行" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算全投资临界点计算,已计算" & m & "次。"
                                Form1.TopMost = True
                                System.Windows.Forms.Application.DoEvents()
                                '反算临界点
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJ * (1 - m / JSCSMax) '每次改变原单价的千分之0.5
                                '计算流动资金
                                'Call 流动资金相关计算()
                                '设置跳出条件
                                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                    CBDJWT = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value, 4) '用于微调成本单价的初始数据
                                    'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.WorksheetFunction.RoundUp(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value, 3) '将反算出来的单价记在表格中
                                    Exit For
                                End If
                            Next
                            '计算流动资金
                            Call 流动资金相关计算()
                            '排除收益率恰好等于设定的收益率的情况，进行微调操作
                            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value > ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                '加入微调修正，使得计算结果更加准确
                                For m = 1 To JSCSMax * 10 '此时的内部收益率大于设定的临界点内部收益率，则将目前的CBDJWT往上加
                                    '计算进度显示
                                    Form1.Show()
                                    Form1.Label1.Text = "正在最后微调" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算全投资临界点计算,已微调" & m & "次。"
                                    Form1.TopMost = True
                                    System.Windows.Forms.Application.DoEvents()
                                    '反算临界点
                                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJWT + WTJSBC * m '每次增加WTJSBC
                                    '计算流动资金
                                    Call 流动资金相关计算()
                                    '设置跳出条件
                                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then '排除收益率恰好等于的情况
                                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value '将反算出来的单价记在表格中
                                        Exit For
                                    ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value - WTJSBC
                                        Exit For
                                    End If
                                Next
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = CBDJWT '将反算出来的单价记在表格中
                            End If
                            '返回初始单价
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJ
                            '计算流动资金
                            Call 流动资金相关计算()
                            '计算此时的单价比原值变化了百分之多少，从而求出反算总量
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 7).Value = ExcelApp.WorksheetFunction.RoundUp((ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value / CBDJ) * CBZL, 2)
                        End If
                        '如果此时的内部收益率大于设定的临界点内部收益率，则将目前的单价往上加
                        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                            For m = 1 To JSCSMax * 10
                                '计算进度显示
                                Form1.Show()
                                Form1.Label1.Text = "正在进行" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算全投资临界点计算,已计算" & m & "次。"
                                Form1.TopMost = True
                                System.Windows.Forms.Application.DoEvents()
                                '反算临界点
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJ * (1 + m / JSCSMax) '每次改变原单价的千分之0.5
                                '计算流动资金
                                'Call 流动资金相关计算()
                                '设置跳出条件
                                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value <= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                    CBDJWT = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value, 4) '用于微调成本单价的初始数据
                                    'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.WorksheetFunction.RoundUp(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value * (1 - 1 / JSCSMax), 3) '将反算出来的单价记在表格中
                                    Exit For
                                End If
                            Next
                            '计算流动资金
                            Call 流动资金相关计算()
                            '排除收益率恰好等于设定的收益率的情况，进行微调操作
                            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                '加入微调修正，使得计算结果更加准确
                                For m = 1 To JSCSMax '此时的内部收益率小于设定的临界点内部收益率，则将目前的CBDJWT往下减
                                    '计算进度显示
                                    Form1.Show()
                                    Form1.Label1.Text = "正在最后微调" & ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 2).Value & "反算全投资临界点计算,已微调" & m & "次。"
                                    Form1.TopMost = True
                                    System.Windows.Forms.Application.DoEvents()
                                    '反算临界点
                                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJWT - WTJSBC * m '每次减小WTJSBC
                                    '计算流动资金
                                    Call 流动资金相关计算()
                                    '设置跳出条件
                                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value '将反算出来的单价记在表格中
                                        Exit For
                                    End If
                                Next
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value = CBDJWT '将反算出来的单价记在表格中
                            End If
                            '返回初始单价
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJ
                            '计算流动资金
                            Call 流动资金相关计算()
                            '计算此时的单价比原值变化了百分之多少，从而求出反算总量
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 7).Value = ExcelApp.WorksheetFunction.RoundUp((ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 4).Value / CBDJ) * CBZL, 2)
                        End If
                    End If
                Next
            End If
        Next
        Form1.Close()
    End Sub
    Sub 全投资税后收益率反算静态投资(FSLJDJSCSMax As Integer)
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '实例化计算进度显示窗体
        Dim Form1 As New 计算进度显示
        Dim TZHH(50) As Integer '数组，储存投资所在的行号
        Dim TZHHJS As Integer = 0 '投资行号计数
        Dim TZJECC(200) '数组，储存各种投资金额
        Dim JTTZWT(200) '数组，静态投资微调
        Dim JTTZZEWT As Double = 0 '静态投资总额微调
        Dim WTJSBC As Integer '微调计算步长
        Dim JSCSMax As Integer = FSLJDJSCSMax '计算次数最大值
        For i = 30 To 110 '检索估算表中投资编号，投资编号都是1表示
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, 27).Value = 1 Then
                TZHHJS = TZHHJS + 1
                TZHH(TZHHJS) = i '行号储存进数组，从下标1开始
            End If
        Next
        Dim TZJESLJS As Integer = 0 '投资金额数量计数
        '将各种投资金额储存在数组中
        For i = 1 To TZHHJS '行
            For j = 3 To 11 Step 2 '列
                TZJESLJS = TZJESLJS + 1
                TZJECC(TZJESLJS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(i), j).Value
            Next
        Next
        Dim TZmin = TZJECC(1）  '投资金额中的最小值，初始值为数组中的第一个数
        For i = 1 To 200
            If TZJECC(i) < TZmin And TZJECC(i) > 0 Then
                TZmin = TZJECC(i）
            End If
        Next
        WTJSBC = ExcelApp.WorksheetFunction.RoundUp(TZmin / (JSCSMax * 10), 0)
        Dim JS As Integer = 0
        '判断此时的内部收益率和设定的临界点内部收益率的大小关系
        '如果此时的内部收益率小于设定的临界点内部收益率，则将目前的单价往下减
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 8).Value = "Y" Then
            For i = 1 To JSCSMax
                '计算进度显示
                Form1.Show()
                Form1.Label1.Text = "正在进行静态投资反算全投资临界点计算,已计算" & i & "次。"
                Form1.TopMost = True
                System.Windows.Forms.Application.DoEvents()
                '改变分年静态投资
                '遍历所有的投资，每个值都变化
                For j = 1 To TZHHJS '行
                    For m = 3 To 11 Step 2 '列
                        JS = (j - 1) * 5 + Int(m / 2)
                        If TZJECC(JS) > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = TZJECC(JS) * (1 - i / JSCSMax)
                            '计算流动资金
                            'Call 流动资金相关计算()
                        End If
                    Next
                Next
                '设置跳出条件
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                    '将静态投资储存在JTTZWT数组
                    TZJESLJS = 0
                    For k = 1 To TZHHJS '行
                        For l = 3 To 11 Step 2 '列
                            TZJESLJS = TZJESLJS + 1
                            JTTZWT(TZJESLJS) = Math.Round(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(k), l).Value, 2)
                        Next
                    Next
                    JTTZZEWT = ExcelApp.WorksheetFunction.RoundUp(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value, 2) '将反算出来的总投资记在表格中
                    'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = ExcelApp.WorksheetFunction.RoundUp(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value, 2) '将反算出来的总投资记在表格中
                    Exit For
                End If
            Next
            '计算流动资金
            Call 流动资金相关计算()
            '排除收益率恰好等于设定的收益率的情况，进行微调操作
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value > ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                '加入微调修正，使得计算结果更加准确
                For i = 1 To JSCSMax * 10 '此时的内部收益率大于设定的临界点内部收益率，则将目前的JTTZWT往上加
                    '计算进度显示
                    Form1.Show()
                    Form1.Label1.Text = "正在最后微调静态投资反算全投资临界点计算,已微调" & i & "次。"
                    Form1.TopMost = True
                    System.Windows.Forms.Application.DoEvents()
                    '遍历所有的投资，每个值都变化
                    For j = 1 To TZHHJS '行
                        For m = 3 To 11 Step 2 '列
                            JS = (j - 1) * 5 + Int(m / 2)
                            If TZJECC(JS) > 0 Then
                                'ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = JTTZWT(JS) + WTJSBC * i '每次增加WTJSBC
                                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = JTTZWT(JS) * (1 + i / (JSCSMax * 25))
                                '计算流动资金
                                Call 流动资金相关计算()
                            End If
                        Next
                    Next
                    '设置跳出条件
                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then '排除收益率恰好等于的情况
                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value '将反算出来的总投资记在表格中 
                        Exit For
                    ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                        'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value - WTJSBC
                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value * (1 - 1 / (JSCSMax * 25))
                        Exit For
                    End If
                Next
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = JTTZZEWT '将反算出来的总投资记在表格中
            End If
            '静态总投资重置回默认值
            JS = 0
            For j = 1 To TZHHJS '行
                For m = 3 To 11 Step 2 '列
                    JS = (j - 1) * 5 + Int(m / 2)
                    If TZJECC(JS) > 0 Then
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = TZJECC(JS)
                        '计算流动资金
                        Call 流动资金相关计算()
                    End If
                Next
            Next
        End If
        '如果此时的内部收益率大于设定的临界点内部收益率，则将目前的单价往上加
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value > ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 8).Value = "Y" Then
            For i = 1 To JSCSMax * 10
                '计算进度显示
                Form1.Show()
                Form1.Label1.Text = "正在进行静态投资反算全投资临界点计算,已计算" & i & "次。"
                Form1.TopMost = True
                System.Windows.Forms.Application.DoEvents()
                '改变分年静态投资
                '遍历所有的投资，每个值都变化
                For j = 1 To TZHHJS '行
                    For m = 3 To 11 Step 2 '列
                        JS = (j - 1) * 5 + Int(m / 2)
                        If TZJECC(JS) > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = TZJECC(JS) * (1 + i / JSCSMax)
                            '计算流动资金
                            'Call 流动资金相关计算()
                        End If
                    Next
                Next
                '设置跳出条件
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value <= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                    '将静态投资储存在JTTZWT数组
                    TZJESLJS = 0
                    For k = 1 To TZHHJS '行
                        For l = 3 To 11 Step 2 '列
                            TZJESLJS = TZJESLJS + 1
                            JTTZWT(TZJESLJS) = Math.Round(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(k), l).Value, 2)
                        Next
                    Next
                    JTTZZEWT = ExcelApp.WorksheetFunction.RoundUp(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value, 2) '将反算出来的总投资记在表格中
                    'ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = ExcelApp.WorksheetFunction.RoundUp(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value * (1 - 1 / 2000), 2) '将反算出来的总投资记在表格中
                    Exit For
                End If
            Next
            '计算流动资金
            Call 流动资金相关计算()
            '排除收益率恰好等于设定的收益率的情况，进行微调操作
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value < ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                '加入微调修正，使得计算结果更加准确
                For i = 1 To JSCSMax '此时的内部收益率小于设定的临界点内部收益率，则将目前的JTTZWT往下减
                    '计算进度显示
                    Form1.Show()
                    Form1.Label1.Text = "正在最后微调静态投资反算全投资临界点计算,已微调" & i & "次。"
                    Form1.TopMost = True
                    System.Windows.Forms.Application.DoEvents()
                    '遍历所有的投资，每个值都变化
                    For j = 1 To TZHHJS '行
                        For m = 3 To 11 Step 2 '列
                            JS = (j - 1) * 5 + Int(m / 2)
                            If TZJECC(JS) > 0 Then
                                'ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = JTTZWT(JS) - WTJSBC * i '每次减少WTJSBC
                                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = JTTZWT(JS) * (1 - i / (JSCSMax * 25))
                                '计算流动资金
                                Call 流动资金相关计算()
                            End If
                        Next
                    Next
                    '设置跳出条件
                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 22).Value >= ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 4).Value Then
                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value  '将反算出来的总投资记在表格中 
                        Exit For
                    End If
                Next
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = JTTZZEWT '将反算出来的总投资记在表格中
            End If
            '静态总投资重置回默认值
            JS = 0
            For j = 1 To TZHHJS '行
                For m = 3 To 11 Step 2 '列
                    JS = (j - 1) * 5 + Int(m / 2)
                    If TZJECC(JS) > 0 Then
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = TZJECC(JS)
                        '计算流动资金
                        Call 流动资金相关计算()
                    End If
                Next
            Next
        End If
        '将反算出的静态投资结果保留为2位小数
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 8).Value = "Y" Then
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value = Math.Round(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 7).Value, 2)
        End If
        Form1.Close()
    End Sub
    Sub 静态投资敏感性分析计算(MGXFXBHL As Double)
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '实例化计算进度显示窗体
        Dim Form1 As New 计算进度显示
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(7, 29).Value = 1 Then
            '计算进度提醒
            Form1.Show()
            Form1.Label1.Text = "正在读取静态投资敏感性分析基本数据。"
            Form1.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            '将总静态投资写入敏感性分析表格中
            ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(7, 8).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value
            Dim TZHH(50) As Integer '数组，储存投资所在的行号
            Dim TZHHJS As Integer = 0 '投资行号计数
            Dim TZJECC(200) '数组，储存各种投资金额
            For i = 30 To 110 '检索估算表中投资编号，投资编号都是1表示
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, 27).Value = 1 Then
                    TZHHJS = TZHHJS + 1
                    TZHH(TZHHJS) = i '行号储存进数组，从下标1开始
                End If
            Next
            Dim TZJESLJS As Integer = 0 '投资金额数量计数
            '将各种投资金额储存在数组中
            For i = 1 To TZHHJS '行
                For j = 3 To 11 Step 2 '列
                    TZJESLJS = TZJESLJS + 1
                    TZJECC(TZJESLJS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(i), j).Value
                Next
            Next
            Dim JS As Integer = 0
            '计算静态总投资敏感性分析
            For i = 1 To 5 '敏感性分析循环5次
                '计算进度提醒
                Form1.Show()
                Form1.Label1.Text = "正在计算静态投资敏感性分析，计算进度：(" & i & "/ 5)"
                Form1.TopMost = True
                System.Windows.Forms.Application.DoEvents()
                '遍历所有的投资，每个值都变化
                For j = 1 To TZHHJS '行
                    For m = 3 To 11 Step 2 '列
                        JS = (j - 1) * 5 + Int(m / 2)
                        If TZJECC(JS) > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = TZJECC(JS) * (（1 - 2 * MGXFXBHL） + (i - 1) * MGXFXBHL)
                        End If
                    Next
                Next
                '计算流动资金
                Call 流动资金相关计算()
                '计算投资回收期
                Call 投资回收期计算()
                '读取计算的结果，内部收益率和回收年限
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(6 + i, 11).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 19).Value
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(6 + i, 12).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(13, 19).Value
            Next
            '静态总投资重置回默认值
            Form1.Show()
            Form1.Label1.Text = "正在将静态投资敏感性基本数据重置回默认值。"
            Form1.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            '静态投资重置回默认值
            JS = 0
            For j = 1 To TZHHJS '行
                For m = 3 To 11 Step 2 '列
                    JS = (j - 1) * 5 + Int(m / 2)
                    If TZJECC(JS) > 0 Then
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = TZJECC(JS)
                    End If
                Next
            Next
            '计算流动资金
            Call 流动资金相关计算()
            '计算投资回收期
            Call 投资回收期计算()
        End If
        Form1.Close()
    End Sub
    Sub 收入敏感性分析计算(MGXFXBHL As Double)
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '实例化计算进度显示窗体
        Dim Form1 As New 计算进度显示
        Dim SRBH As Integer '收入编号
        Dim SRDJ '收入单价
        For i = 13 To 25 '收入&成本表中的有收入的项
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 7).Value > 0 Then '存在收入
                SRBH = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 27).Value
                SRDJ = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value
                For j = 6 To 150
                    If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(j, 27).Value = SRBH And SRBH > 0 And SRDJ > 0 Then '找到编号相同的收入
                        Dim BQ As String = ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(j, 18).Value '标签，用户计算进度显示
                        'Dim XZ1 = MsgBox("是否要进行" & BQ & "敏感性分析计算？", vbOKCancel)
                        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(j, 29).Value = 1 Then
                            ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(j, 8).Value = SRDJ '收入单价写入表格
                            For m = 1 To 5 '敏感性分析
                                '计算进度显示
                                Form1.Show()
                                Form1.Label1.Text = "正在计算" & BQ & "敏感性分析，计算进度：(" & m & "/ 5)"
                                Form1.TopMost = True
                                System.Windows.Forms.Application.DoEvents()
                                '敏感性分析
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJ * (（1 - 2 * MGXFXBHL） + (m - 1) * MGXFXBHL)
                                '计算流动资金
                                Call 流动资金相关计算()
                                '计算投资回收期
                                Call 投资回收期计算()
                                '读取收益率和回收期
                                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(j - 1 + m, 11).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 19).Value
                                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(j - 1 + m, 12).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(13, 19).Value
                            Next
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJ '收入单价初始值返回
                            '计算流动资金
                            Call 流动资金相关计算()
                            '计算投资回收期
                            Call 投资回收期计算()
                        End If
                    End If
                Next
            End If
        Next
        Form1.Close()
    End Sub
    Sub 成本敏感性分析计算(MGXFXBHL As Double)
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '实例化计算进度显示窗体
        Dim Form1 As New 计算进度显示
        Dim CBBH As Integer '成本编号
        Dim CBDJ '成本单价
        For i = 14 To 27 '收入&成本表中的有成本的项
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 14).Value > 0 Then '存在成本
                CBBH = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 28).Value
                CBDJ = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value
                For j = 6 To 150
                    If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(j, 28).Value = CBBH And CBBH > 0 And CBDJ > 0 Then '找到编号相同的成本
                        Dim BQ = ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(j, 18).Value '标签，用户计算进度显示
                        'Dim XZ1 = MsgBox("是否要进行" & BQ & "敏感性分析计算？", vbOKCancel)
                        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(j, 29).Value = 1 Then
                            ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(j, 8).Value = CBDJ '成本单价写入表格
                            For m = 1 To 5 '敏感性分析
                                '计算进度显示
                                Form1.Show()
                                Form1.Label1.Text = "正在计算" & BQ & "敏感性分析，计算进度：(" & m & "/ 5)"
                                Form1.TopMost = True
                                System.Windows.Forms.Application.DoEvents()
                                '敏感性分析
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJ * (（1 - 2 * MGXFXBHL） + (m - 1) * MGXFXBHL)
                                '计算流动资金
                                Call 流动资金相关计算()
                                '计算投资回收期
                                Call 投资回收期计算()
                                '读取收益率和回收期
                                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(j - 1 + m, 11).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 19).Value
                                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(j - 1 + m, 12).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(13, 19).Value
                            Next
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJ '成本单价返回初始值
                            '计算流动资金
                            Call 流动资金相关计算()
                            '计算投资回收期
                            Call 投资回收期计算()
                        End If
                    End If
                Next
            End If
        Next
        Form1.Close()
    End Sub
    Sub 年运行小时数敏感性分析(MGXFXBHL As Double)
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '实例化计算进度显示窗体
        Dim Form1 As New 计算进度显示
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(137, 29).Value = 1 Then
            '只有当输入了年运行小时数时，才计算
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(137, 8).Value > 0 Then
                '计算进度显示
                Form1.Show()
                Form1.Label1.Text = "正在读取年利用小时数敏感性分析基本数据。"
                Form1.TopMost = True
                System.Windows.Forms.Application.DoEvents()
                '基本数据读入
                Dim NLHH(50) As Integer '数组，储存各种年量所在的行号
                Dim NLHHJS As Integer = 0 '各种年量行号计数
                Dim NLCC(200) '数组，储存各种年量
                For i = 30 To 110 '检索估算表中各种年量编号，各种年量编号都是2表示
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, 28).Value = 2 Then
                        NLHHJS = NLHHJS + 1
                        NLHH(NLHHJS) = i '行号储存进数组，从下标1开始
                    End If
                Next
                Dim NLSLJS As Integer = 0 '各种年量数量计数
                '将各种年量储存在数组中
                For i = 1 To NLHHJS '行
                    For j = 3 To 11 Step 2 '列
                        NLSLJS = NLSLJS + 1
                        NLCC(NLSLJS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(NLHH(i), j).Value
                    Next
                Next
                Dim SRNL(20) '收入年量
                Dim CBNL（20） '成本年量
                For i = 13 To 25
                    SRNL(i - 12) = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 6).Value
                Next
                For i = 14 To 24
                    CBNL(i - 12) = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 13).Value
                Next
                Dim JS As Integer = 0
                '计算年运行小时数敏感性分析
                For i = 1 To 5 '敏感性分析循环5次
                    '计算进度显示
                    Form1.Show()
                    Form1.Label1.Text = "正在计算年运行小时数敏感性分析，计算进度：(" & i & "/ 5)"
                    Form1.TopMost = True
                    System.Windows.Forms.Application.DoEvents()
                    '遍历所有的各种年量，每个值都变化
                    For j = 1 To NLHHJS '行
                        For m = 3 To 11 Step 2 '列
                            JS = (j - 1) * 5 + Int(m / 2)
                            If NLCC(JS) > 0 Then
                                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(NLHH(j), m).Value = NLCC(JS) * (（1 - 2 * MGXFXBHL） + (i - 1) * MGXFXBHL)
                            End If
                        Next
                    Next
                    '遍历各种收入和成本的全年输入，每个值都变化
                    For j = 13 To 25 '行号
                        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 6).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 6).Value = SRNL(j - 12) * (（1 - 2 * MGXFXBHL） + (i - 1) * MGXFXBHL)
                        End If
                    Next
                    For j = 14 To 24 '行号
                        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 13).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 13).Value = CBNL(j - 12) * (（1 - 2 * MGXFXBHL） + (i - 1) * MGXFXBHL)
                        End If
                    Next
                    '计算流动资金
                    Call 流动资金相关计算()
                    '计算投资回收期
                    Call 投资回收期计算()
                    '读取计算的结果，内部收益率和回收年限
                    ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(136 + i, 11).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 19).Value
                    ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(136 + i, 12).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(13, 19).Value
                Next
                '计算进度显示
                Form1.Show()
                Form1.Label1.Text = "正将年利用小时数敏感性分析基本数据重置回默认值。"
                Form1.TopMost = True
                System.Windows.Forms.Application.DoEvents()
                '各种重置回默认值
                JS = 0
                For j = 1 To NLHHJS '行
                    For m = 3 To 11 Step 2 '列
                        JS = (j - 1) * 5 + Int(m / 2)
                        If NLCC(JS) > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(NLHH(j), m).Value = NLCC(JS)
                        End If
                    Next
                Next
                '收入成本量重置回默认值
                For j = 13 To 25 '行号
                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 6).Value > 0 Then
                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 6).Value = SRNL(j - 12)
                    End If
                Next
                For j = 14 To 24 '行号
                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 13).Value > 0 Then
                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 13).Value = CBNL(j - 12)
                    End If
                Next
                '计算流动资金
                Call 流动资金相关计算()
                '计算投资回收期
                Call 投资回收期计算()
            End If
        End If
        Form1.Close()
    End Sub
    Sub 绘制单因素敏感性分析图()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim Form10 As New 选择单因素敏感性分析内容
        If Form10.hzzxt.Checked = True Then
            '绘制单因素敏感性分析图
            ExcelApp.ThisWorkbook.Worksheets("指标数据").Shapes.AddChart.Name = "单因素敏感性分析图" '创建图表并重命名
            ExcelApp.ThisWorkbook.Worksheets("指标数据").Shapes("单因素敏感性分析图").Select
            ExcelApp.ActiveChart.ChartType = Excel.XlChartType.xlLineMarkers '选择图表类型
            ExcelApp.ActiveSheet.Shapes("单因素敏感性分析图").IncrementLeft(-415)
            ExcelApp.ActiveSheet.Shapes("单因素敏感性分析图").IncrementTop(340)
            ExcelApp.ThisWorkbook.Worksheets("指标数据").ChartObjects("单因素敏感性分析图").Activate
            ExcelApp.ActiveChart.ChartTitle.Text = "单因素敏感性分析图表"
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————        
            Dim js As Integer = 0
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(7, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$7"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$7:$K$11"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(12, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$12"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$12:$K$16"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(17, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$17"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$17:$K$21"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(22, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$22"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$22:$K$26"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(27, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$27"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$27:$K$31"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(32, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$32"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$32:$K$36"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(37, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$37"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$37:$K$41"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(42, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$42"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$42:$K$46"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(47, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$47"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$47:$K$51"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(52, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$52"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$52:$K$56"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(57, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$57"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$57:$K$61"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(62, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$62"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$62:$K$66"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(67, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$67"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$67:$K$71"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(72, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$72"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$72:$K$76"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(77, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$77"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$77:$K$81"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(82, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$82"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$82:$K$86"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(87, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$87"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$87:$K$91"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(92, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$92"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$92:$K$96"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(97, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$97"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$97:$K$101"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(102, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$102"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$102:$K$106"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(107, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$107"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$107:$K$111"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(112, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$112"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$112:$K$116"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(117, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$117"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$117:$K$121"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(122, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$122"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$122:$K$126"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(127, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$127"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$127:$K$131"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(132, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$132"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$132:$K$136"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(137, 8).Value > 0 Then
                js = js + 1
                ExcelApp.ActiveChart.SeriesCollection.NewSeries
                ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$137"
                ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$137:$K$141"
                ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
            End If
        End If
    End Sub
    Sub 设置敏感性分析图格式()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '求取计算出的内部收益率中的最小值和最大值
        Dim NBSYLmin = ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(7, 11).Value '内部收益率最小值
        Dim NBSYLmax = ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(7, 11).Value '内部收益率最大值
        For i = 8 To 141 '求最小值
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i, 11).Value < NBSYLmin And ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i, 11).Value <> Nothing Then
                NBSYLmin = ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i, 11).Value
            End If
        Next
        For i = 8 To 141 '求最大值
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i, 11).Value > NBSYLmax And ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i, 11).Value <> Nothing Then
                NBSYLmax = ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i, 11).Value
            End If
        Next
        ExcelApp.ThisWorkbook.Worksheets("指标数据").ChartObjects("单因素敏感性分析图").Activate
        'Y轴数字显示格式
        ExcelApp.ActiveChart.HasAxis(Excel.XlAxisType.xlValue) = True
        ExcelApp.ActiveChart.Axes(Excel.XlAxisType.xlValue).Select
        ExcelApp.Selection.TickLabels.NumberFormatLocal = "0%"
        ExcelApp.ActiveChart.Axes(Excel.XlAxisType.xlValue).MinimumScale = ExcelApp.WorksheetFunction.RoundDown(NBSYLmin - 0.005, 2) '设置折线图坐标下限
        ExcelApp.ActiveChart.Axes(Excel.XlAxisType.xlValue).MaximumScale = ExcelApp.WorksheetFunction.RoundUp(NBSYLmax + 0.005, 2) '设置折线图坐标上限
        '设置横坐标标题、纵坐标标题、图表总标题
        ExcelApp.ActiveSheet.ChartObjects("单因素敏感性分析图").Activate
        'ExcelApp.ActiveChart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionBottom '设置图例出现在底部
        ExcelApp.ActiveChart.SetElement(MsoChartElementType.msoElementChartTitleAboveChart)
        ExcelApp.Selection.Format.TextFrame2.TextRange.Characters.Text = "单因素敏感性分析折线图"
        With ExcelApp.Selection.Format.TextFrame2.TextRange.Characters(1, 11).ParagraphFormat
            .TextDirection = MsoTextDirection.msoTextDirectionLeftToRight
            .Alignment = MsoAlignCmd.msoAlignCenters
        End With
        With ExcelApp.Selection.Format.TextFrame2.TextRange.Characters(1, 11).Font
            .BaselineOffset = 0
            .Bold = MsoTriState.msoTrue
            .NameComplexScript = "+mn-cs"
            .NameFarEast = "+mn-ea"
            .Fill.Visible = MsoTriState.msoTrue
            .Fill.ForeColor.RGB = RGB(0, 0, 0)
            .Fill.Transparency = 0
            .Fill.Solid
            .Size = 18
            .Italic = MsoTriState.msoFalse
            .Kerning = 12
            .Name = "+mn-lt"
            .UnderlineStyle = MsoTextUnderlineType.msoNoUnderline
            .Strike = MsoTextStrike.msoNoStrike
        End With
        ExcelApp.ActiveChart.ChartTitle.Text = "单因素敏感性分析折线图"
        ExcelApp.Selection.Format.TextFrame2.TextRange.Characters.Text = "单因素敏感性分析折线图"
        With ExcelApp.Selection.Format.TextFrame2.TextRange.Characters(1, 11).ParagraphFormat
            .TextDirection = MsoTextDirection.msoTextDirectionLeftToRight
            .Alignment = MsoAlignCmd.msoAlignCenters
        End With
        With ExcelApp.Selection.Format.TextFrame2.TextRange.Characters(1, 11).Font
            .BaselineOffset = 0
            .Bold = MsoTriState.msoTrue
            .NameComplexScript = "+mn-cs"
            .NameFarEast = "+mn-ea"
            .Fill.Visible = MsoTriState.msoTrue
            .Fill.ForeColor.RGB = RGB(0, 0, 0)
            .Fill.Transparency = 0
            .Fill.Solid
            .Size = 18
            .Italic = MsoTriState.msoFalse
            .Kerning = 12
            .Name = "+mn-lt"
            .UnderlineStyle = MsoTextUnderlineType.msoNoUnderline
            .Strike = MsoTextStrike.msoNoStrike
        End With
        ExcelApp.ActiveChart.ChartArea.Select()
        ExcelApp.ActiveChart.SetElement(MsoChartElementType.msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
        ExcelApp.Selection.Format.TextFrame2.TextRange.Characters.Text = "单因素值变化率（%）"
        With ExcelApp.Selection.Format.TextFrame2.TextRange.Characters(1, 7).ParagraphFormat
            .TextDirection = MsoTextDirection.msoTextDirectionLeftToRight
            .Alignment = MsoAlignCmd.msoAlignCenters
        End With
        With ExcelApp.Selection.Format.TextFrame2.TextRange.Characters(1, 7).Font
            .BaselineOffset = 0
            .Bold = MsoTriState.msoTrue
            .NameComplexScript = "+mn-cs"
            .NameFarEast = "+mn-ea"
            .Fill.Visible = MsoTriState.msoTrue
            .Fill.ForeColor.RGB = RGB(0, 0, 0)
            .Fill.Transparency = 0
            .Fill.Solid
            .Size = 10
            .Italic = MsoTriState.msoFalse
            .Kerning = 12
            .Name = "+mn-lt"
            .UnderlineStyle = MsoTextUnderlineType.msoNoUnderline
            .Strike = MsoTextStrike.msoNoStrike
        End With
        ExcelApp.ActiveChart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = "单因素值变化率（%）"
        ExcelApp.ActiveChart.ChartArea.Select()
        ExcelApp.ActiveChart.SetElement(MsoChartElementType.msoElementPrimaryValueAxisTitleVertical)
        ExcelApp.Selection.Format.TextFrame2.TextRange.Characters.Text = "资本金"
        With ExcelApp.Selection.Format.TextFrame2.TextRange.Characters(1, 3).ParagraphFormat
            .TextDirection = MsoTextDirection.msoTextDirectionLeftToRight
            .Alignment = MsoAlignCmd.msoAlignCenters
        End With
        With ExcelApp.Selection.Format.TextFrame2.TextRange.Characters(1, 3).Font
            .BaselineOffset = 0
            .Bold = MsoTriState.msoTrue
            .NameComplexScript = "+mn-cs"
            .NameFarEast = "+mn-ea"
            .Fill.Visible = MsoTriState.msoTrue
            .Fill.ForeColor.RGB = RGB(0, 0, 0)
            .Fill.Transparency = 0
            .Fill.Solid
            .Size = 10
            .Italic = MsoTriState.msoFalse
            .Kerning = 12
            .Name = "+mn-lt"
            .UnderlineStyle = MsoTextUnderlineType.msoNoUnderline
            .Strike = MsoTextStrike.msoNoStrike
        End With
        ExcelApp.Selection.Format.TextFrame2.TextRange.Characters.Text = "资本金所得税后内部收益率（%）"
        With ExcelApp.Selection.Format.TextFrame2.TextRange.Characters(1, 12).ParagraphFormat
            .TextDirection = MsoTextDirection.msoTextDirectionLeftToRight
            .Alignment = MsoAlignCmd.msoAlignCenters
        End With
        With ExcelApp.Selection.Format.TextFrame2.TextRange.Characters(1, 12).Font
            .BaselineOffset = 0
            .Bold = MsoTriState.msoTrue
            .NameComplexScript = "+mn-cs"
            .NameFarEast = "+mn-ea"
            .Fill.Visible = MsoTriState.msoTrue
            .Fill.ForeColor.RGB = RGB(0, 0, 0)
            .Fill.Transparency = 0
            .Fill.Solid
            .Size = 10
            .Italic = MsoTriState.msoFalse
            .Kerning = 12
            .Name = "+mn-lt"
            .UnderlineStyle = MsoTextUnderlineType.msoNoUnderline
            .Strike = MsoTextStrike.msoNoStrike
        End With
        ExcelApp.ActiveChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = "资本金所得税后内部收益率（%）"
        ExcelApp.ActiveChart.ChartArea.Select()
        '移动图表至新的位置并放大
        ExcelApp.ActiveSheet.Shapes("单因素敏感性分析图").ScaleWidth(1.9, MsoTriState.msoFalse, MsoScaleFrom.msoScaleFromTopLeft)
        ExcelApp.ActiveSheet.Shapes("单因素敏感性分析图").ScaleHeight(1.8, MsoTriState.msoFalse, MsoScaleFrom.msoScaleFromTopLeft)
        ExcelApp.ActiveSheet.Shapes("单因素敏感性分析图").IncrementLeft(890)
        ExcelApp.ActiveSheet.Shapes("单因素敏感性分析图").IncrementTop(-80)
        '检测是否有多余的数据标签，有多余的就删除
        Dim js As Integer = 0
        For i = 7 To 137 Step 5 '统计应该有多少个标签
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(i, 8).Value > 0 Then
                js = js + 1
            End If
        Next
        ExcelApp.ThisWorkbook.Worksheets("指标数据").ChartObjects("单因素敏感性分析图").Activate
        For i = 1 To 30
            For j = 1 To 30
                ExcelApp.ActiveChart.SeriesCollection(js + j).Delete
            Next
        Next
    End Sub
    Sub 投资回收期计算()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '自有资金
        '资本金现金流量表
        For iii = 1 To 31 '年份1至31年
            '所得税前累计净现金流量
            If ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(25, 4 + iii).Value >= 0 Then
                '将所得税前累计净现金流量第一个大于等于0的年份记录在表格中
                ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(45, 22).Value = iii
                '将所得税前累计净现金流量第一个大于等于0的年份，之前一年的所得税前累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(49, 22).Value = ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(25, 4 + iii - 1).Value
                '将所得税前累计净现金流量第一个大于等于0的年份，当年的所得税前净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(52, 22).Value = ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(24, 4 + iii).Value
                Exit For
            End If
        Next
        For jjj = 1 To 31 '年份1至31年
            '所得税后累计净现金流量
            If ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(23, 4 + jjj).Value >= 0 Then
                '将所得税后累计净现金流量第一个大于等于0的年份记录在表格中
                ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(46, 22).Value = jjj
                '将所得税后累计净现金流量第一个大于等于0的年份，之前一年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(50, 22).Value = ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(23, 4 + jjj - 1).Value
                '将所得税后累计净现金流量第一个大于等于0的年份，当年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(53, 22).Value = ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(22, 4 + jjj).Value
                Exit For
            End If
        Next
        '全投资
        '项目投资现金流量表
        For iii = 1 To 31 '年份1至31年
            '所得税前累计净现金流量
            If ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(17, 4 + iii).Value >= 0 Then
                '将所得税前累计净现金流量第一个大于等于0的年份记录在表格中
                ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(41, 22).Value = iii
                '将所得税前累计净现金流量第一个大于等于0的年份，之前一年的所得税前累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(44, 22).Value = ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(17, 4 + iii - 1).Value
                '将所得税前累计净现金流量第一个大于等于0的年份，当年的所得税前净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(47, 22).Value = ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(16, 4 + iii).Value
                Exit For
            End If
        Next
        For jjj = 1 To 31 '年份1至31年
            '所得税后累计净现金流量
            If ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(20, 4 + jjj).Value >= 0 Then
                '将所得税后累计净现金流量第一个大于等于0的年份记录在表格中
                ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(42, 22).Value = jjj
                '将所得税后累计净现金流量第一个大于等于0的年份，之前一年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(45, 22).Value = ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(20, 4 + jjj - 1).Value
                '将所得税后累计净现金流量第一个大于等于0的年份，当年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(48, 22).Value = ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(19, 4 + jjj).Value
                Exit For
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '投资各方所得税后回收年限计算
        '投资方1        
        For iii = 1 To 31 '年份1至31年
            '所得税后累计净现金流量
            If ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(16, 4 + iii).Value >= 0 Then
                '将所得税后累计净现金流量第一个大于等于0的年份记录在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(30, 22).Value = iii
                '将所得税后累计净现金流量第一个大于等于0的年份，之前一年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(33, 22).Value = ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(16, 4 + iii - 1).Value
                '将所得税后累计净现金流量第一个大于等于0的年份，当年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(35, 22).Value = ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(15, 4 + iii).Value
                Exit For
            End If
        Next
        '投资方2        
        For iii = 1 To 31 '年份1至31年
            '所得税后累计净现金流量
            If ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(16, 4 + iii).Value >= 0 Then
                '将所得税后累计净现金流量第一个大于等于0的年份记录在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(30, 22).Value = iii
                '将所得税后累计净现金流量第一个大于等于0的年份，之前一年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(33, 22).Value = ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(16, 4 + iii - 1).Value
                '将所得税后累计净现金流量第一个大于等于0的年份，当年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(35, 22).Value = ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(15, 4 + iii).Value
                Exit For
            End If
        Next
        '投资方3        
        For iii = 1 To 31 '年份1至31年
            '所得税后累计净现金流量
            If ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(16, 4 + iii).Value >= 0 Then
                '将所得税后累计净现金流量第一个大于等于0的年份记录在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(30, 22).Value = iii
                '将所得税后累计净现金流量第一个大于等于0的年份，之前一年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(33, 22).Value = ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(16, 4 + iii - 1).Value
                '将所得税后累计净现金流量第一个大于等于0的年份，当年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(35, 22).Value = ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(15, 4 + iii).Value
                Exit For
            End If
        Next
        '投资方4        
        For iii = 1 To 31 '年份1至31年
            '所得税后累计净现金流量
            If ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(16, 4 + iii).Value >= 0 Then
                '将所得税后累计净现金流量第一个大于等于0的年份记录在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(30, 22).Value = iii
                '将所得税后累计净现金流量第一个大于等于0的年份，之前一年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(33, 22).Value = ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(16, 4 + iii - 1).Value
                '将所得税后累计净现金流量第一个大于等于0的年份，当年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(35, 22).Value = ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(15, 4 + iii).Value
                Exit For
            End If
        Next
        '投资方5       
        For iii = 1 To 31 '年份1至31年
            '所得税后累计净现金流量
            If ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(16, 4 + iii).Value >= 0 Then
                '将所得税后累计净现金流量第一个大于等于0的年份记录在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(30, 22).Value = iii
                '将所得税后累计净现金流量第一个大于等于0的年份，之前一年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(33, 22).Value = ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(16, 4 + iii - 1).Value
                '将所得税后累计净现金流量第一个大于等于0的年份，当年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(35, 22).Value = ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(15, 4 + iii).Value
                Exit For
            End If
        Next
    End Sub

    Sub 设置所得税减免和增值税退税包含内容()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '定义局部变量
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '屏蔽屏幕更新，防止屏闪
        ExcelApp.Application.ScreenUpdating = False
        '屏蔽事件
        ExcelApp.Application.EnableEvents = False
        Dim XZ = MsgBox("是否需要设置所得税减免和增值税退税包含内容?", vbOKCancel)
        Dim Form6 As New 设置所得税减免和增值税退税包含内容
        If XZ = vbOK Then
            Form6.ShowDialog() '窗口显示
            Form6.TopMost = True
            System.Windows.Forms.Application.DoEvents()
        End If
        '重新锁定表格
        ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Protect(Password:="wscjc")
        '计算流动资金
        Call 流动资金相关计算()
        '重新计算投资回收期
        Call 投资回收期计算()
        '打开事件
        ExcelApp.Application.EnableEvents = True
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Activate()
    End Sub

    Sub 设置部分销售收入和经营成本计算年限()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '定义局部变量
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '屏蔽屏幕更新，防止屏闪
        ExcelApp.Application.ScreenUpdating = False
        '屏蔽事件
        ExcelApp.Application.EnableEvents = False
        '手动计算，关闭excel的自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationManual
        Dim XZ = MsgBox("是否需要设置部分销售收入和经营成本计算年限?", vbOKCancel)
        Dim Form13 As New 设置部分销售收入和经营成本计算年限
        If XZ = vbOK Then
            Form13.ShowDialog() '窗口显示
            Form13.TopMost = True
            System.Windows.Forms.Application.DoEvents()
        End If
        '清空窗体中已有的数据
        Form13.ksnf1.Clear()
        Form13.ksnf2.Clear()
        Form13.ksnf3.Clear()
        Form13.jsnf1.Clear()
        Form13.jsnf2.Clear()
        Form13.jsnf3.Clear()
        '计算流动资金
        Call 流动资金相关计算()
        '重新计算投资回收期
        Call 投资回收期计算()
        '重新锁定表格
        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Protect(Password:="wscjc")
        '重新打开excel自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '打开事件
        ExcelApp.Application.EnableEvents = True
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Activate()
    End Sub
    Sub 设置接入费收入计算方式()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '定义局部变量
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '屏蔽屏幕更新，防止屏闪
        ExcelApp.Application.ScreenUpdating = False
        '屏蔽事件
        ExcelApp.Application.EnableEvents = False
        '手动计算，关闭excel的自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationManual
        Dim XZ = MsgBox("是否需要设置接入费收入计算方式?", vbOKCancel)
        Dim Form12 As New 设置接入费计算方式
        If XZ = vbOK Then
            Form12.ShowDialog() '窗口显示
            Form12.TopMost = True
            System.Windows.Forms.Application.DoEvents()
        End If
        '清空窗体中已有的数据
        Form12.ksnf1.Clear()
        Form12.ksnf2.Clear()
        Form12.ksnf3.Clear()
        Form12.jsnf1.Clear()
        Form12.jsnf2.Clear()
        Form12.jsnf3.Clear()
        Form12.CheckBox1.Checked = False
        '计算流动资金
        Call 流动资金相关计算()
        '重新计算投资回收期
        Call 投资回收期计算()
        '重新锁定表格
        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Protect(Password:="wscjc")
        '打开事件
        ExcelApp.Application.EnableEvents = True
        '重新打开excel自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Activate()
    End Sub

    Sub 设置补贴收入计算年限()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '定义局部变量
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '屏蔽屏幕更新，防止屏闪
        ExcelApp.Application.ScreenUpdating = False
        '屏蔽事件
        ExcelApp.Application.EnableEvents = False
        '手动计算，关闭excel的自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationManual
        Dim XZ = MsgBox("是否需要设置补贴收入的计算年限?", vbOKCancel)
        Dim Form7 As New 设置补贴收入计算年限
        If XZ = vbOK Then
            Form7.ShowDialog() '窗口显示
            Form7.TopMost = True
            System.Windows.Forms.Application.DoEvents()
        End If
        '清空窗体中已有的数据
        Form7.ksnf1.Clear()
        Form7.ksnf2.Clear()
        Form7.ksnf3.Clear()
        Form7.jsnf1.Clear()
        Form7.jsnf2.Clear()
        Form7.jsnf3.Clear()
        Form7.CheckBox1.Checked = False
        Form7.CheckBox2.Checked = False
        '计算流动资金
        Call 流动资金相关计算()
        '重新计算投资回收期
        Call 投资回收期计算()
        '重新锁定表格
        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Protect(Password:="wscjc")
        '打开事件
        ExcelApp.Application.EnableEvents = True
        '重新打开excel自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Activate()
    End Sub

    Sub 逐年衰减系数输入()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '定义局部变量
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '屏蔽屏幕更新，防止屏闪
        ExcelApp.Application.ScreenUpdating = False
        '屏蔽事件
        ExcelApp.Application.EnableEvents = False
        '手动计算，关闭excel的自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationManual
        Dim XZ = MsgBox("是否需要修改程序内置的默认光伏发电逐年衰减效率或者蓄电池逐年衰减效率?", vbOKCancel)
        Dim Form2 As New 逐年衰减系数设置
        If XZ = vbOK Then
            '读取输入的衰减开始年份和衰减率（%）
            Form2.ShowDialog() '窗口显示
            Form2.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            '计算过程在“逐年衰减系数设置.确定”
        End If
        '清空窗体中已有的数据
        Form2.TextBox1.Clear()
        Form2.TextBox2.Clear()
        Form2.TextBox3.Clear()
        Form2.TextBox4.Clear()
        Form2.TextBox5.Clear()
        Form2.TextBox6.Clear()
        Form2.TextBox7.Clear()
        Form2.TextBox8.Clear()
        Form2.TextBox9.Clear()
        Form2.TextBox10.Clear()
        Form2.RichTextBox1.Rtf = Nothing
        Form2.RichTextBox1.Clear()
        '计算流动资金
        Call 流动资金相关计算()
        '重新计算投资回收期
        Call 投资回收期计算()
        '重新锁定表格
        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Protect(Password:="wscjc")
        '打开事件
        ExcelApp.Application.EnableEvents = True
        '重新打开excel自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Activate()
    End Sub

    Sub 修理费率逐年变化设置()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '定义局部变量
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '屏蔽屏幕更新，防止屏闪
        ExcelApp.Application.ScreenUpdating = False
        '屏蔽事件
        ExcelApp.Application.EnableEvents = False
        '手动计算，关闭excel的自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationManual
        Dim XZ = MsgBox("是否需要设置逐年动态变化的修理费率?", vbOKCancel)
        Dim Form4 As New 修理费率逐年变化设置
        If XZ = vbOK Then
            '读取输入的修理费率变化率（%）
            Form4.ShowDialog() '窗口显示
            Form4.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            '计算过程在“修理费率逐年变化设置.确定”
        End If
        '清空窗体中已有的数据
        Form4.ksnf1.Clear()
        Form4.ksnf2.Clear()
        Form4.ksnf3.Clear()
        Form4.ksnf4.Clear()
        Form4.ksnf5.Clear()
        Form4.jsnf1.Clear()
        Form4.jsnf2.Clear()
        Form4.jsnf3.Clear()
        Form4.jsnf4.Clear()
        Form4.jsnf5.Clear()
        Form4.ksfl1.Clear()
        Form4.ksfl2.Clear()
        Form4.ksfl3.Clear()
        Form4.ksfl4.Clear()
        Form4.ksfl5.Clear()
        Form4.jsfl1.Clear()
        Form4.jsfl2.Clear()
        Form4.jsfl3.Clear()
        Form4.jsfl4.Clear()
        Form4.RichTextBox1.Rtf = Nothing
        '计算流动资金
        Call 流动资金相关计算()
        '重新计算投资回收期
        Call 投资回收期计算()
        '重新锁定表格
        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Protect(Password:="wscjc")
        '打开事件
        ExcelApp.Application.EnableEvents = True
        '重新打开excel自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Activate()
    End Sub

    Sub 建设期资金运用方式设置()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '定义局部变量
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '屏蔽屏幕更新，防止屏闪
        ExcelApp.Application.ScreenUpdating = False
        '屏蔽事件
        ExcelApp.Application.EnableEvents = False
        '手动计算，关闭excel的自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationManual
        Dim XZ = MsgBox("是否需要设置建设期资金运用方式?包括每次投资的资本金比例系数取值方式和建设期贷款利率。", vbOKCancel)
        Dim Form5 As New 建设期资金运用方式设置
        If XZ = vbOK Then
            Form5.ShowDialog() '窗口显示
            Form5.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            '计算过程在“建设期资金运用方式设置.确定”
        End If
        '清空已有内容
        Form5.zbjbl_a.Clear()
        Form5.dkll_a.Clear()
        Form5.zbjbl1.Clear()
        Form5.zbjbl2.Clear()
        Form5.zbjbl3.Clear()
        Form5.zbjbl4.Clear()
        Form5.zbjbl5.Clear()
        Form5.zbjbl6.Clear()
        Form5.zbjbl7.Clear()
        Form5.zbjbl8.Clear()
        Form5.zbjbl9.Clear()
        Form5.zbjbl10.Clear()
        Form5.dkll1.Clear()
        Form5.dkll2.Clear()
        Form5.dkll3.Clear()
        Form5.dkll4.Clear()
        Form5.dkll5.Clear()
        Form5.dkll6.Clear()
        Form5.dkll7.Clear()
        Form5.dkll8.Clear()
        Form5.dkll9.Clear()
        Form5.dkll10.Clear()
        '重新锁定表格
        ExcelApp.ThisWorkbook.Worksheets("估算表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Protect(Password:="wscjc")
        '激活表格
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Activate()
        '重新打开excel自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '打开事件
        ExcelApp.Application.EnableEvents = True
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
    End Sub

    Sub 建设期逐次设置折旧摊销计算方式()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '定义局部变量
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '屏蔽屏幕更新，防止屏闪
        ExcelApp.Application.ScreenUpdating = False
        '屏蔽事件
        ExcelApp.Application.EnableEvents = False
        '手动计算，关闭excel的自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationManual
        Dim XZ = MsgBox("是否需要每次投资设置不同的固定资产折旧和无形资产摊销计算方式？", vbOKCancel)
        Dim Form8 As New 每次投资设置不同的折旧摊销计算方式
        If XZ = vbOK Then
            Form8.ShowDialog() '窗口显示
            Form8.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            '计算过程在“每次投资设置不同的折旧摊销计算方式.确定”
        End If
        '清空已有的全部数据
        '固定资产折旧年限
        Form8.gdzczjnx1.Clear()
        Form8.gdzczjnx2.Clear()
        Form8.gdzczjnx3.Clear()
        Form8.gdzczjnx4.Clear()
        Form8.gdzczjnx5.Clear()
        Form8.gdzczjnx6.Clear()
        Form8.gdzczjnx7.Clear()
        Form8.gdzczjnx8.Clear()
        Form8.gdzczjnx9.Clear()
        Form8.gdzczjnx10.Clear()
        '固定资产残值率
        Form8.gdzcczl1.Clear()
        Form8.gdzcczl2.Clear()
        Form8.gdzcczl3.Clear()
        Form8.gdzcczl4.Clear()
        Form8.gdzcczl5.Clear()
        Form8.gdzcczl6.Clear()
        Form8.gdzcczl7.Clear()
        Form8.gdzcczl8.Clear()
        Form8.gdzcczl9.Clear()
        Form8.gdzcczl10.Clear()
        '无形资产摊销年限
        Form8.wxzctxnx1.Clear()
        Form8.wxzctxnx2.Clear()
        Form8.wxzctxnx3.Clear()
        Form8.wxzctxnx4.Clear()
        Form8.wxzctxnx5.Clear()
        Form8.wxzctxnx6.Clear()
        Form8.wxzctxnx7.Clear()
        Form8.wxzctxnx8.Clear()
        Form8.wxzctxnx9.Clear()
        Form8.wxzctxnx10.Clear()
        '重新锁定表格
        ExcelApp.ThisWorkbook.Worksheets("估算表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Protect(Password:="wscjc")
        '激活表格
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Activate()
        '重新打开excel自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '打开事件
        ExcelApp.Application.EnableEvents = True
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
    End Sub

    Sub 逐次设置长期贷款还款和宽限年限()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '定义局部变量
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '屏蔽屏幕更新，防止屏闪
        ExcelApp.Application.ScreenUpdating = False
        '屏蔽事件
        ExcelApp.Application.EnableEvents = False
        '手动计算，关闭excel的自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationManual
        Dim XZ = MsgBox("是否需要每次投资设置不同的长期贷款还款年限和宽限年限？", vbOKCancel)
        Dim Form9 As New 每次投资设置不同的长期贷款还款和宽限年限
        If XZ = vbOK Then
            Form9.ShowDialog() '窗口显示
            Form9.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            '计算过程在“每次投资设置不同的长期贷款还款年限和宽限年限.确定”
        End If
        '长期贷款还款年限
        Form9.cqdkhknx1.Clear()
        Form9.cqdkhknx2.Clear()
        Form9.cqdkhknx3.Clear()
        Form9.cqdkhknx4.Clear()
        Form9.cqdkhknx5.Clear()
        Form9.cqdkhknx6.Clear()
        Form9.cqdkhknx7.Clear()
        Form9.cqdkhknx8.Clear()
        Form9.cqdkhknx9.Clear()
        Form9.cqdkhknx10.Clear()
        '长期贷款宽限年限
        Form9.kxnx1.Clear()
        Form9.kxnx2.Clear()
        Form9.kxnx3.Clear()
        Form9.kxnx4.Clear()
        Form9.kxnx5.Clear()
        Form9.kxnx6.Clear()
        Form9.kxnx7.Clear()
        Form9.kxnx8.Clear()
        Form9.kxnx9.Clear()
        Form9.kxnx10.Clear()
        '重新锁定表格
        ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("估算表").protect(Password:="wscjc")
        '激活表格
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Activate()
        '重新打开excel自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '打开事件
        ExcelApp.Application.EnableEvents = True
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
    End Sub

    Sub 分别设置投资各方收益计算参数()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '定义局部变量
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '屏蔽屏幕更新，防止屏闪
        ExcelApp.Application.ScreenUpdating = False
        '屏蔽事件
        ExcelApp.Application.EnableEvents = False
        '手动计算，关闭excel的自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationManual
        Dim XZ = MsgBox("是否需要分别设置投资各方收益计算参数？", vbOKCancel)
        Dim Form14 As New 投资各方设置不同的出资比例资产处置比例和利润分配比例
        If XZ = vbOK Then
            Form14.ShowDialog() '窗口显示
            Form14.TopMost = True
            System.Windows.Forms.Application.DoEvents()
        End If
        '清空窗体内的数据
        '出资比例
        Form14.czbl1.Clear()
        Form14.czbl2.Clear()
        Form14.czbl3.Clear()
        Form14.czbl4.Clear()
        Form14.czbl5.Clear()
        '资产处置比例
        Form14.zcczbl1.Clear()
        Form14.zcczbl2.Clear()
        Form14.zcczbl3.Clear()
        Form14.zcczbl4.Clear()
        Form14.zcczbl5.Clear()
        '利润分配比例
        Form14.lrfpbl1.Clear()
        Form14.lrfpbl2.Clear()
        Form14.lrfpbl3.Clear()
        Form14.lrfpbl4.Clear()
        Form14.lrfpbl5.Clear()
        '重新锁定表格
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Protect(Password:="wscjc")
        '激活表格
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Activate()
        '重新打开excel自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '打开事件
        ExcelApp.Application.EnableEvents = True
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
    End Sub
    Sub 项目计算年限变化后改变相关系数()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '定义局部变量
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        Dim tznfzdz, zjdknx '投资年份最大值、折旧贷款年限中较大的值
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '折旧贷款年限中的最大值
        zjdknx = ExcelApp.WorksheetFunction.Max(ExcelApp.ThisWorkbook.Worksheets("估算表").Range("G6:G8").Value)
        '项目计算年数改变后改变的计算用系数
        '读取输入的10次投资发生年份的最大值
        '如果贷款计算方法是方法一，同时折旧和摊销的计算方法为方法一或者方法二，定义投资年份最大值等于1
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法一" And (ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法一" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法二") Then
            tznfzdz = 1
            '添加报错
            If tznfzdz + zjdknx > jsnx Then
                '报错（只在这里报错一次，折旧摊销贷款第二到第十次投资年份模块里就不写了，防止重复报错）
                MsgBox("输入的<项目计算年限、第二次到第十次项目投资年份、长期贷款年限、固定资产折旧年限、无形资产摊销年限>之中的某一项或者某几项参数不合理，请重新核对！！！程序将自动重置不合理的数值！！！")
                '固定资产折旧年限
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 7).Value > jsnx - tznfzdz Then
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 7).Value = jsnx - tznfzdz
                End If
                '长期贷款年限
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value > jsnx - tznfzdz Then
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value = jsnx - tznfzdz
                End If
                '无形资产摊销年限
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 7).Value > jsnx - tznfzdz Then
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 7).Value = jsnx - tznfzdz
                End If
            End If
        Else
            tznfzdz = ExcelApp.WorksheetFunction.Max(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 3).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value) '投资年份最大值
        End If

        '输入的数字必需为从2到31的整数，同时要大于等于项目投资年份最大值与折旧和贷款年限的较大值的和
        If jsnx >= 2 And jsnx <= 31 And jsnx >= tznfzdz + zjdknx And Int(jsnx) = jsnx Then
            '屏蔽屏幕更新，防止屏闪；手动计算，关闭excel的自动计算；解锁表格
            Call 计算前基本处理()
            '根据输入的计算年数，修改相关计算系数
            '前15年
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(111, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(112, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(112, i).Value = 0
                End If
            Next
            '后16年
            For i = 2 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(114, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(115, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(115, i).Value = 0
                End If
            Next
            '计算年限改变后，改变相应的逐年达产系数
            '前15年
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(102, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(5, i).Value
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value = 0
                End If
            Next
            '后16年
            For i = 2 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(105, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(8, i).Value
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i).Value = 0
                End If
            Next
            '计算期末回收固定资产余值和回收流动资金系数
            '前15年
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(135, i).Value = jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(136, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(136, i).Value = 0
                End If
            Next
            '后16年
            For i = 2 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(138, i).Value = jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(139, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(139, i).Value = 0
                End If
            Next
        ElseIf jsnx > 31 Then
            MsgBox("项目计算年限不可以大于31年，请重新输入！")
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value = 31
        ElseIf jsnx < 2 Then
            MsgBox("项目计算年限不可以小于2年，请重新输入！")
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value = 2
        Else
            '如果贷款的计算方法中为方法二，或者折旧和摊销的计算方法为方法三或者方法四，进行下列计算
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法二" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法三" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法四" Then
                '报错（只在这里报错一次，折旧摊销贷款第二到第十次投资年份模块里就不写了，防止重复报错）
                MsgBox("输入的<项目计算年限、第二次到第十次项目投资年份、长期贷款年限、固定资产折旧年限、无形资产摊销年限>之中的某一项或者某几项参数不合理，请重新核对！！！程序将自动重置不合理的数值！！！")
                '固定资产折旧年限
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 7).Value > jsnx - tznfzdz Then
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 7).Value = jsnx - tznfzdz
                End If
                '长期贷款年限
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value > jsnx - tznfzdz Then
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value = jsnx - tznfzdz
                End If
                '无形资产摊销年限
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 7).Value > jsnx - tznfzdz Then
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 7).Value = jsnx - tznfzdz
                End If
            End If
        End If
    End Sub
    Sub 折旧计算年限变化后改变相关系数()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '定义局部变量
        Dim gdzczjnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 7).Value '固定资产折旧年限
        Dim tznfzdz As Integer
        '读取输入的项目计算年限
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '读取输入的10次投资发生年份的最大值
        '如果贷款计算方法是方法一，同时折旧和摊销的计算方法为方法一或者方法二，定义投资年份最大值等于1
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法一" And (ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法一" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法二") Then
            tznfzdz = 1
        Else
            tznfzdz = ExcelApp.WorksheetFunction.Max(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 3).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value) '投资年份最大值
        End If
        '输入的数字必需为从1到30的整数，并且小于项目计算年限
        If gdzczjnx >= 1 And gdzczjnx <= 30 And gdzczjnx <= jsnx - tznfzdz And Int(gdzczjnx) = gdzczjnx Then
            '屏蔽屏幕更新，防止屏闪；手动计算，关闭excel的自动计算；解锁表格
            Call 计算前基本处理()
            '第1次投资发生后，折旧计算系数，默认第一次投资发生在第一年且不可以修改
            '根据输入的折旧计算年数，修改相关计算系数
            '前15年
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(119, i).Value <= gdzczjnx + 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(119, i).Value > 1 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(120, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(120, i).Value = 0
                End If
            Next
            '后16年
            For i = 2 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(122, i).Value <= gdzczjnx + 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(122, i).Value > 1 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(123, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(123, i).Value = 0
                End If
            Next
            '第2次投资发生后，折旧计算系数
            '前15年
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(175, i).Value <= (gdzczjnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(175, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(176, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(176, i).Value = 0
                End If
            Next
            '后16年
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(178, i).Value <= (gdzczjnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(178, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(179, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(179, i).Value = 0
                End If
            Next
            '第3次投资发生后，折旧计算系数
            '前15年
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(183, i).Value <= (gdzczjnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(183, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(184, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(184, i).Value = 0
                End If
            Next
            '后16年
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(186, i).Value <= (gdzczjnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(186, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(187, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(187, i).Value = 0
                End If
            Next
            '第4次投资发生后，折旧计算系数
            '前15年
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(191, i).Value <= (gdzczjnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(191, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(192, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(192, i).Value = 0
                End If
            Next
            '后16年
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(194, i).Value <= (gdzczjnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(194, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(195, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(195, i).Value = 0
                End If
            Next
            '第5次投资发生后，折旧计算系数
            '前15年
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(199, i).Value <= (gdzczjnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(199, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(200, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(200, i).Value = 0
                End If
            Next
            '后16年
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(202, i).Value <= (gdzczjnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(202, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(203, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(203, i).Value = 0
                End If
            Next
            '第6次投资发生后，折旧计算系数
            '前15年
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(279, i).Value <= (gdzczjnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(279, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(280, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(280, i).Value = 0
                End If
            Next
            '后16年
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(282, i).Value <= (gdzczjnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(282, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(283, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(283, i).Value = 0
                End If
            Next
            '第7次投资发生后，折旧计算系数
            '前15年
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(287, i).Value <= (gdzczjnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(287, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(288, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(288, i).Value = 0
                End If
            Next
            '后16年
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(290, i).Value <= (gdzczjnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(290, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(291, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(291, i).Value = 0
                End If
            Next
            '第8次投资发生后，折旧计算系数
            '前15年
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(295, i).Value <= (gdzczjnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(295, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(296, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(296, i).Value = 0
                End If
            Next
            '后16年
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(298, i).Value <= (gdzczjnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(298, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(299, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(299, i).Value = 0
                End If
            Next
            '第9次投资发生后，折旧计算系数
            '前15年
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(303, i).Value <= (gdzczjnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(303, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(304, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(304, i).Value = 0
                End If
            Next
            '后16年
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(306, i).Value <= (gdzczjnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(306, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(307, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(307, i).Value = 0
                End If
            Next
            '第10次投资发生后，折旧计算系数
            '前15年
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(311, i).Value <= (gdzczjnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(311, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(312, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(312, i).Value = 0
                End If
            Next
            '后16年
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(314, i).Value <= (gdzczjnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(314, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(315, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(315, i).Value = 0
                End If
            Next
        ElseIf gdzczjnx < 1 Then
            MsgBox("折旧计算年限不可以小于1年，请重新输入！")
            If jsnx >= 2 And jsnx <= 31 Then
                '只有在项目计算年限大于等于2年，小于等于31年时才进行操作，防止死循环
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 7).Value = 1 '重置回默认值
            End If
        Else
            '如果贷款的计算方法中为方法二，或者折旧和摊销的计算方法为方法三或者方法四，进行下列计算
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法二" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法三" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法四" Then
                '报错（全部在项目计算年限修改模块内报错，这里不再重复报错）
                'MsgBox("输入的<项目计算年限、第二次到第十次项目投资年份、长期贷款年限、固定资产折旧年限、无形资产摊销年限>之中的某一项或者某几项参数不合理，请重新核对！！！为防止计算异常，程序不会自动修改，请重新手动输入！！！")
                'If jsnx >= 2 And jsnx <= 31 Then
                '    '只有在项目计算年限大于等于2年，小于等于31年时才进行操作，防止死循环
                '    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 7).Value = jsnx - tznfzdz '重置回默认值
                'End If
            End If
        End If
    End Sub
    Sub 长期贷款计算年限变化后改变相关系数()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        'cqdkhknx 长期贷款还款年限
        '定义局部变量
        Dim tznfzdz
        '读取输入的项目计算年限
        Dim cqdkhknx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value '长期贷款年限
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '读取输入的10次投资发生年份的最大值
        '如果贷款计算方法是方法一，同时折旧和摊销的计算方法为方法一或者方法二，定义投资年份最大值等于1
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法一" And (ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法一" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法二") Then
            tznfzdz = 1
        Else
            tznfzdz = ExcelApp.WorksheetFunction.Max(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 3).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value) '投资年份最大值
        End If
        '输入的数字必需为从1到30的整数，并且小于项目计算年限
        If cqdkhknx >= 1 And cqdkhknx <= 30 And cqdkhknx <= jsnx - tznfzdz And Int(cqdkhknx) = cqdkhknx Then
            '屏蔽屏幕更新，防止屏闪；手动计算，关闭excel的自动计算；解锁表格
            Call 计算前基本处理()
            '根据输入的长期贷款还款计算年数，修改相关计算系数
            '第1次投资发生后，长期贷款计算系数，默认第一次投资发生在第一年且不可以修改
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(127, i).Value <= cqdkhknx + 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(127, i).Value > 1 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(128, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(128, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(130, i).Value <= cqdkhknx + 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(130, i).Value > 1 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(131, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(131, i).Value = 0
                End If
            Next
            '第2次投资发生后，长期贷款计算系数
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(207, i).Value <= (cqdkhknx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(207, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(208, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(208, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(210, i).Value <= (cqdkhknx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(210, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(211, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(211, i).Value = 0
                End If
            Next
            '第3次投资发生后，长期贷款计算系数
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(215, i).Value <= (cqdkhknx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(215, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(216, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(216, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(218, i).Value <= (cqdkhknx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(218, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(219, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(219, i).Value = 0
                End If
            Next
            '第4次投资发生后，长期贷款计算系数
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(223, i).Value <= (cqdkhknx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(223, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(224, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(224, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(226, i).Value <= (cqdkhknx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(226, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(227, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(227, i).Value = 0
                End If
            Next
            '第5次投资发生后，长期贷款计算系数
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(231, i).Value <= (cqdkhknx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(231, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(232, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(232, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(234, i).Value <= (cqdkhknx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(234, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(235, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(235, i).Value = 0
                End If
            Next
            '第6次投资发生后，长期贷款计算系数
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(319, i).Value <= (cqdkhknx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(319, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(320, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(320, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(322, i).Value <= (cqdkhknx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(322, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(323, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(323, i).Value = 0
                End If
            Next
            '第7次投资发生后，长期贷款计算系数
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(327, i).Value <= (cqdkhknx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(327, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(328, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(328, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(330, i).Value <= (cqdkhknx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(330, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(331, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(331, i).Value = 0
                End If
            Next
            '第8次投资发生后，长期贷款计算系数
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(335, i).Value <= (cqdkhknx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(335, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(336, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(336, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(338, i).Value <= (cqdkhknx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(338, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(339, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(339, i).Value = 0
                End If
            Next
            '第9次投资发生后，长期贷款计算系数
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(343, i).Value <= (cqdkhknx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(343, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(344, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(344, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(346, i).Value <= (cqdkhknx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(346, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(347, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(347, i).Value = 0
                End If
            Next
            '第10次投资发生后，长期贷款计算系数
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(351, i).Value <= (cqdkhknx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(351, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(352, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(352, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(354, i).Value <= (cqdkhknx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(354, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(355, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(355, i).Value = 0
                End If
            Next
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '每次贷款已经发生的年份数，和剩余的贷款年份数计算
            Dim JS1 As Integer = 0
            For i = 100 To 600 '遍历借款还本付息计划表的行号
                If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 52).Value = 8888 Then '贷款已经发生的年份数，标签是8888
                    For j = 28 To 71 Step 43 '估算表行号
                        For k = 3 To 11 Step 2 '估算表列号
                            JS1 = JS1 + 1
                            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(j, k).Value > 0 Then
                                For l = 3 To 33 '借款还本付息计划表列号
                                    If (ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, l).Value - ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(j, k).Value) >= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value And (ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, l).Value - ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(j, k).Value) <= （cqdkhknx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value - 1） Then
                                        ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i + JS1, l).Value = ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, l).Value - ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(j, k).Value - （ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value - 1）
                                    Else
                                        ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i + JS1, l).Value = 0
                                    End If
                                Next
                            Else
                                For l = 3 To 33 '借款还本付息计划表列号
                                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i + JS1, l).Value = 0
                                Next
                            End If
                        Next
                    Next
                End If
            Next
            Dim JS2 As Integer = 0
            For i = 100 To 600 '遍历借款还本付息计划表的行号
                If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 53).Value = 9999 Then '贷款剩余的年份数，标签是9999
                    For j = 28 To 71 Step 43 '估算表行号
                        For k = 3 To 11 Step 2 '估算表列号
                            JS2 = JS2 + 1
                            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(j, k).Value > 0 Then
                                For l = 3 To 33 '借款还本付息计划表列号
                                    If (ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, l).Value - ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(j, k).Value) >= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value And (ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, l).Value - ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(j, k).Value) <= （cqdkhknx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value - 1） Then
                                        ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i + JS2, l).Value = cqdkhknx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(j, k).Value - ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, l).Value + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value
                                    Else
                                        ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i + JS2, l).Value = 0
                                    End If
                                Next
                            Else
                                For l = 3 To 33 '借款还本付息计划表列号
                                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i + JS2, l).Value = 0
                                Next
                            End If
                        Next
                    Next
                End If
            Next
        ElseIf cqdkhknx < 1 Then
            MsgBox("长期贷款计算年限不可以小于1年，请重新输入！")
            If jsnx >= 2 And jsnx <= 31 Then
                '只有在项目计算年限大于等于2年，小于等于31年时才进行操作，防止死循环
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value = 1 '重置回默认值
            End If
        Else
            '如果贷款的计算方法中为方法二，或者折旧和摊销的计算方法为方法三或者方法四，进行下列计算
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法二" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法三" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法四" Then
                '报错（全部在项目计算年限修改模块内报错，这里不再重复报错）
                'MsgBox("输入的<项目计算年限、第二次到第十次项目投资年份、长期贷款年限、固定资产折旧年限、无形资产摊销年限>之中的某一项或者某几项参数不合理，请重新核对！！！为防止计算异常，程序不会自动修改，请重新手动输入！！！")
                'If jsnx >= 2 And jsnx <= 31 Then
                '    '只有在项目计算年限大于等于2年，小于等于31年时才进行操作，防止死循环
                '    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value = jsnx - tznfzdz '重置回默认值
                'End If
            End If
        End If
    End Sub
    Sub 无形资产摊销年限变化后改变相关系数()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '定义局部变量
        Dim tznfzdz
        Dim wxzctxnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 7).Value '无形资产摊销年限
        '读取输入的项目计算年限
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '读取输入的10次投资发生年份的最大值
        '如果贷款计算方法是方法一，同时折旧和摊销的计算方法为方法一或者方法二，定义投资年份最大值等于1
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法一" And (ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法一" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法二") Then
            tznfzdz = 1
        Else
            tznfzdz = ExcelApp.WorksheetFunction.Max(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 3).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value, ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value) '投资年份最大值
        End If
        '输入的数字必需为从1到30的整数，并且小于项目计算年限
        If wxzctxnx >= 1 And wxzctxnx <= 30 And wxzctxnx <= jsnx - tznfzdz And Int(wxzctxnx) = wxzctxnx Then
            '屏蔽屏幕更新，防止屏闪；手动计算，关闭excel的自动计算；解锁表格
            Call 计算前基本处理()
            '第1次投资发生后，无形资产摊销计算系数，默认第一次投资发生在第一年且不可以修改
            '根据输入的无形资产摊销计算年数，修改相关计算系数
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(359, i).Value <= wxzctxnx + 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(359, i).Value > 1 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(360, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(360, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(362, i).Value <= wxzctxnx + 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(362, i).Value > 1 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(363, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(363, i).Value = 0
                End If
            Next
            '第2次投资发生后，无形资产摊销计算系数
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(367, i).Value <= (wxzctxnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(367, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(368, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(368, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(370, i).Value <= (wxzctxnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(370, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(371, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(371, i).Value = 0
                End If
            Next
            '第3次投资发生后，无形资产摊销计算系数
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(375, i).Value <= (wxzctxnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(375, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(376, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(376, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(378, i).Value <= (wxzctxnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(378, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(379, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(379, i).Value = 0
                End If
            Next
            '第4次投资发生后，无形资产摊销计算系数
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(383, i).Value <= (wxzctxnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(383, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(384, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(384, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(386, i).Value <= (wxzctxnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(386, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(387, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(387, i).Value = 0
                End If
            Next
            '第5次投资发生后，无形资产摊销计算系数
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(391, i).Value <= (wxzctxnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(391, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(392, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(392, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(394, i).Value <= (wxzctxnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(394, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(395, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(395, i).Value = 0
                End If
            Next
            '第6次投资发生后，无形资产摊销计算系数
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(399, i).Value <= (wxzctxnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(399, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(400, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(400, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(402, i).Value <= (wxzctxnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(402, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(403, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(403, i).Value = 0
                End If
            Next
            '第7次投资发生后，无形资产摊销计算系数
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(407, i).Value <= (wxzctxnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(407, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(408, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(408, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(410, i).Value <= (wxzctxnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(410, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(411, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(411, i).Value = 0
                End If
            Next
            '第8次投资发生后，无形资产摊销计算系数
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(415, i).Value <= (wxzctxnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(415, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(416, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(416, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(418, i).Value <= (wxzctxnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(418, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(419, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(419, i).Value = 0
                End If
            Next
            '第9次投资发生后，无形资产摊销计算系数
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(423, i).Value <= (wxzctxnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(423, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(424, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(424, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(426, i).Value <= (wxzctxnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(426, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(427, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(427, i).Value = 0
                End If
            Next
            '第10次投资发生后，无形资产摊销计算系数
            For i = 3 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(431, i).Value <= (wxzctxnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(431, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(432, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(432, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(434, i).Value <= (wxzctxnx + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(434, i).Value > ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value <> 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(435, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(435, i).Value = 0
                End If
            Next
        ElseIf wxzctxnx < 1 Then
            MsgBox("无形资产摊销计算年限不可以小于1年，请重新输入！")
            If jsnx >= 2 And jsnx <= 31 Then
                '只有在项目计算年限大于等于2年，小于等于31年时才进行操作，防止死循环
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 7).Value = 1
            End If
        Else
            '如果贷款的计算方法中为方法二，或者折旧和摊销的计算方法为方法三或者方法四，进行下列计算
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法二" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法三" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法四" Then
                '报错（全部在项目计算年限修改模块内报错，这里不再重复报错）
                'MsgBox("输入的<项目计算年限、第二次到第十次项目投资年份、长期贷款年限、固定资产折旧年限、无形资产摊销年限>之中的某一项或者某几项参数不合理，请重新核对！！！为防止计算异常，程序不会自动修改，请重新手动输入！！！")
                'If jsnx >= 2 And jsnx <= 31 Then
                '    '只有在项目计算年限大于等于2年，小于等于31年时才进行操作，防止死循环
                '    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 7).Value = jsnx - tznfzdz '重置回默认值
                'End If
            End If
        End If
    End Sub
    Sub 第二至第五次投资年份改变后修改相关计算系数()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '定义局部变量
        Dim jsnx, zjdknx, tzzdnf '项目计算年限、折旧贷款年限中的较大值
        Dim tznf2, tznf3, tznf4, tznf5, tznf6, tznf7, tznf8, tznf9, tznf10 '投资年份2-10
        '读取本项目计算年限、以及折旧年限和贷款年限中的较大值
        jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '折旧贷款年限最大值
        zjdknx = ExcelApp.WorksheetFunction.Max(ExcelApp.ThisWorkbook.Worksheets("估算表").Range("G6:G8").Value)
        '允许输入的最大年份数
        tzzdnf = jsnx - zjdknx '投资最大年份
        '项目第二次投资年份发生变化时
        '输入的数字必需大于等于2，小于等于允许输入的最大年份的整数，且不为0
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value >= 2 And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value <= tzzdnf And Int(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value <> 0 Then
            '屏蔽屏幕更新，防止屏闪；手动计算，关闭excel的自动计算；解锁表格
            Call 计算前基本处理()
            '根据输入的项目第二次投资年份，修改相关计算系数
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(143, i).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(144, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(144, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(146, i).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(147, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(147, i).Value = 0
                End If
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value = 0 Then
            '屏蔽屏幕更新，防止屏闪；手动计算，关闭excel的自动计算；解锁表格
            Call 计算前基本处理()
            For i = 3 To 17
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(144, i).Value = 0
            Next
            For i = 2 To 17
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(147, i).Value = 0
            Next
            '如果贷款OR折旧摊销计算方法中有一个为方法二，进行下列计算
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法二" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法三" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法四" Then
            '报错（全部在项目计算年限修改模块内报错，这里不再重复报错）
            'MsgBox("输入的<项目计算年限、第二次到第十次项目投资年份、长期贷款年限、固定资产折旧年限、无形资产摊销年限>之中的某一项或者某几项参数不合理，请重新核对！！！为防止计算异常，程序不会自动修改，请重新手动输入！！！")
            'ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value = 2 '重置回默认值
        End If
        '项目第三次投资年份发生变化时
        '读取项目第二次投资年份
        tznf2 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value
        '输入的数字必需大于项目第二次投资年份，，且第二次投资年份不为0，小于等于允许输入的最大年份的整数，且不为0
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value > tznf2 And tznf2 <> 0 And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value <= tzzdnf And Int(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value <> 0 Then
            '屏蔽屏幕更新，防止屏闪；手动计算，关闭excel的自动计算；解锁表格
            Call 计算前基本处理()
            '根据输入的项目第3次投资年份，修改相关计算系数
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(151, i).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(152, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(152, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(154, i).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(155, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(155, i).Value = 0
                End If
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value = 0 Then
            '屏蔽屏幕更新，防止屏闪；手动计算，关闭excel的自动计算；解锁表格
            Call 计算前基本处理()
            For i = 3 To 17
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(152, i).Value = 0
            Next
            For i = 2 To 17
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(155, i).Value = 0
            Next
            '如果贷款OR折旧摊销计算方法中有一个为方法二，进行下列计算
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法二" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法三" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法四" Then
            '报错（全部在项目计算年限修改模块内报错，这里不再重复报错）
            'MsgBox("输入的<项目计算年限、第二次到第十次项目投资年份、长期贷款年限、固定资产折旧年限、无形资产摊销年限>之中的某一项或者某几项参数不合理，请重新核对！！！为防止计算异常，程序不会自动修改，请重新手动输入！！！")
            'If tznf2 <> 0 Then
            '    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value = tznf2 + 1 '重置回默认值
            'Else
            '    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value = 0
            'End If
        End If
        '项目第四次投资年份发生变化时
        '读取项目第三次投资年份
        tznf3 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value
        '输入的数字必需大于项目第三次投资年份，小于等于允许输入的最大年份的整数，且不为0
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value > tznf3 And tznf3 <> 0 And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value <= tzzdnf And Int(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value <> 0 Then
            '屏蔽屏幕更新，防止屏闪；手动计算，关闭excel的自动计算；解锁表格
            Call 计算前基本处理()
            '根据输入的项目第4次投资年份，修改相关计算系数
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(159, i).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(160, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(160, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(162, i).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(163, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(163, i).Value = 0
                End If
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value = 0 Then
            '屏蔽屏幕更新，防止屏闪；手动计算，关闭excel的自动计算；解锁表格
            Call 计算前基本处理()
            For i = 3 To 17
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(160, i).Value = 0
            Next
            For i = 2 To 17
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(163, i).Value = 0
            Next
            '如果贷款OR折旧摊销计算方法中有一个为方法二，进行下列计算
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法二" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法三" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法四" Then
            '报错（全部在项目计算年限修改模块内报错，这里不再重复报错）
            'MsgBox("输入的<项目计算年限、第二次到第十次项目投资年份、长期贷款年限、固定资产折旧年限、无形资产摊销年限>之中的某一项或者某几项参数不合理，请重新核对！！！为防止计算异常，程序不会自动修改，请重新手动输入！！！")
            'If tznf3 <> 0 Then
            '    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value = tznf3 + 1 '重置回默认值
            'Else
            '    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value = 0
            'End If
        End If
        '项目第五次投资年份发生变化时
        '读取项目第四次投资年份
        tznf4 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value
        '输入的数字必需大于项目第四次投资年份，小于等于允许输入的最大年份的整数，且不为0
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value > tznf4 And tznf4 <> 0 And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value <= tzzdnf And Int(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value <> 0 Then
            '屏蔽屏幕更新，防止屏闪；手动计算，关闭excel的自动计算；解锁表格
            Call 计算前基本处理()
            '根据输入的项目第5次投资年份，修改相关计算系数
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(167, i).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(168, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(168, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(170, i).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(171, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(171, i).Value = 0
                End If
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value = 0 Then
            '屏蔽屏幕更新，防止屏闪；手动计算，关闭excel的自动计算；解锁表格
            Call 计算前基本处理()
            For i = 3 To 17
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(168, i).Value = 0
            Next
            For i = 2 To 17
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(171, i).Value = 0
            Next
            '如果贷款OR折旧摊销计算方法中有一个为方法二，进行下列计算
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法二" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法三" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法四" Then
            '报错（全部在项目计算年限修改模块内报错，这里不再重复报错）
            'MsgBox("输入的<项目计算年限、第二次到第十次项目投资年份、长期贷款年限、固定资产折旧年限、无形资产摊销年限>之中的某一项或者某几项参数不合理，请重新核对！！！为防止计算异常，程序不会自动修改，请重新手动输入！！！")
            'If tznf4 <> 0 Then
            '    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value = tznf4 + 1 '重置回默认值
            'Else
            '    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value = 0
            'End If
        End If
    End Sub
    Sub 第六至第十次投资年份改变后修改相关计算系数()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '定义局部变量
        Dim jsnx, zjdknx, tzzdnf '项目计算年限、折旧贷款年限中的较大值
        Dim tznf2, tznf3, tznf4, tznf5, tznf6, tznf7, tznf8, tznf9, tznf10 '投资年份2-10
        '读取本项目计算年限、以及折旧年限和贷款年限中的较大值
        jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '折旧贷款年限最大值
        zjdknx = ExcelApp.WorksheetFunction.Max(ExcelApp.ThisWorkbook.Worksheets("估算表").Range("G6:G8").Value)
        '允许输入的最大年份数
        tzzdnf = jsnx - zjdknx '投资最大年份
        '项目第六次投资年份发生变化时
        '读取项目第五次投资年份
        tznf5 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value
        '输入的数字必需大于项目第5次投资年份，小于等于允许输入的最大年份的整数，且不为0
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value > tznf5 And tznf5 <> 0 And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value <= tzzdnf And Int(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value <> 0 Then
            '屏蔽屏幕更新，防止屏闪；手动计算，关闭excel的自动计算；解锁表格
            Call 计算前基本处理()
            '根据输入的项目第6次投资年份，修改相关计算系数
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(239, i).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(240, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(240, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(242, i).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(243, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(243, i).Value = 0
                End If
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value = 0 Then
            '屏蔽屏幕更新，防止屏闪；手动计算，关闭excel的自动计算；解锁表格
            Call 计算前基本处理()
            For i = 3 To 17
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(240, i).Value = 0
            Next
            For i = 2 To 17
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(243, i).Value = 0
            Next
            '如果贷款OR折旧摊销计算方法中有一个为方法二，进行下列计算
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法二" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法三" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法四" Then
            '报错（全部在项目计算年限修改模块内报错，这里不再重复报错）
            'MsgBox("输入的<项目计算年限、第二次到第十次项目投资年份、长期贷款年限、固定资产折旧年限、无形资产摊销年限>之中的某一项或者某几项参数不合理，请重新核对！！！为防止计算异常，程序不会自动修改，请重新手动输入！！！")
            'If tznf5 <> 0 Then
            '    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value = tznf5 + 1 '重置回默认值
            'Else
            '    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value = 0
            'End If
        End If
        '项目第7次投资年份发生变化时
        '读取项目第6次投资年份
        tznf6 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value
        '输入的数字必需大于项目第6次投资年份，小于等于允许输入的最大年份的整数，且不为0
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value > tznf6 And tznf6 <> 0 And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value <= tzzdnf And Int(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value <> 0 Then
            '屏蔽屏幕更新，防止屏闪；手动计算，关闭excel的自动计算；解锁表格
            Call 计算前基本处理()
            '根据输入的项目第6次投资年份，修改相关计算系数
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(247, i).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(248, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(248, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(250, i).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(251, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(251, i).Value = 0
                End If
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value = 0 Then
            '屏蔽屏幕更新，防止屏闪；手动计算，关闭excel的自动计算；解锁表格
            Call 计算前基本处理()
            For i = 3 To 17
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(248, i).Value = 0
            Next
            For i = 2 To 17
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(251, i).Value = 0
            Next
            '如果贷款OR折旧摊销计算方法中有一个为方法二，进行下列计算
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法二" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法三" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法四" Then
            '报错（全部在项目计算年限修改模块内报错，这里不再重复报错）
            'MsgBox("输入的<项目计算年限、第二次到第十次项目投资年份、长期贷款年限、固定资产折旧年限、无形资产摊销年限>之中的某一项或者某几项参数不合理，请重新核对！！！为防止计算异常，程序不会自动修改，请重新手动输入！！！")
            'If tznf6 <> 0 Then
            '    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value = tznf6 + 1 '重置回默认值
            'Else
            '    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value = 0
            'End If
        End If
        '项目第8次投资年份发生变化时
        '读取项目第7次投资年份
        tznf7 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value
        '输入的数字必需大于项目第7次投资年份，小于等于允许输入的最大年份的整数，且不为0
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value > tznf7 And tznf7 <> 0 And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value <= tzzdnf And Int(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value <> 0 Then
            '屏蔽屏幕更新，防止屏闪；手动计算，关闭excel的自动计算；解锁表格
            Call 计算前基本处理()
            '根据输入的项目第6次投资年份，修改相关计算系数
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(255, i).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(256, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(256, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(258, i).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(259, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(259, i).Value = 0
                End If
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value = 0 Then
            '屏蔽屏幕更新，防止屏闪；手动计算，关闭excel的自动计算；解锁表格
            Call 计算前基本处理()
            For i = 3 To 17
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(256, i).Value = 0
            Next
            For i = 2 To 17
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(259, i).Value = 0
            Next
            '如果贷款OR折旧摊销计算方法中有一个为方法二，进行下列计算
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法二" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法三" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法四" Then
            '报错（全部在项目计算年限修改模块内报错，这里不再重复报错）
            'MsgBox("输入的<项目计算年限、第二次到第十次项目投资年份、长期贷款年限、固定资产折旧年限、无形资产摊销年限>之中的某一项或者某几项参数不合理，请重新核对！！！为防止计算异常，程序不会自动修改，请重新手动输入！！！")
            'If tznf7 <> 0 Then
            '    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value = tznf7 + 1 '重置回默认值
            'Else
            '    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value = 0
            'End If
        End If
        '项目第9次投资年份发生变化时
        '读取项目第8次投资年份
        tznf8 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value
        '输入的数字必需大于项目第8次投资年份，小于等于允许输入的最大年份的整数，且不为0
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value > tznf8 And tznf8 <> 0 And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value <= tzzdnf And Int(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value <> 0 Then
            '屏蔽屏幕更新，防止屏闪；手动计算，关闭excel的自动计算；解锁表格
            Call 计算前基本处理()
            '根据输入的项目第6次投资年份，修改相关计算系数
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(263, i).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(264, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(264, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(266, i).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(267, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(267, i).Value = 0
                End If
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value = 0 Then
            '屏蔽屏幕更新，防止屏闪；手动计算，关闭excel的自动计算；解锁表格
            Call 计算前基本处理()
            For i = 3 To 17
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(264, i).Value = 0
            Next
            For i = 2 To 17
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(267, i).Value = 0
            Next
            '如果贷款OR折旧摊销计算方法中有一个为方法二，进行下列计算
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法二" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法三" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法四" Then
            '报错（全部在项目计算年限修改模块内报错，这里不再重复报错）
            'MsgBox("输入的<项目计算年限、第二次到第十次项目投资年份、长期贷款年限、固定资产折旧年限、无形资产摊销年限>之中的某一项或者某几项参数不合理，请重新核对！！！为防止计算异常，程序不会自动修改，请重新手动输入！！！")
            'If tznf8 <> 0 Then
            '    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value = tznf8 + 1 '重置回默认值
            'Else
            '    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value = 0
            'End If
        End If
        '项目第10次投资年份发生变化时
        '读取项目第9次投资年份
        tznf9 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value
        '输入的数字必需大于项目第9次投资年份，小于等于允许输入的最大年份的整数，且不为0
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value > tznf9 And tznf9 <> 0 And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value <= tzzdnf And Int(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value <> 0 Then
            '屏蔽屏幕更新，防止屏闪；手动计算，关闭excel的自动计算；解锁表格
            Call 计算前基本处理()
            '根据输入的项目第6次投资年份，修改相关计算系数
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(271, i).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(272, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(272, i).Value = 0
                End If
            Next
            For i = 2 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(274, i).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(275, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(275, i).Value = 0
                End If
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value = 0 Then
            '屏蔽屏幕更新，防止屏闪；手动计算，关闭excel的自动计算；解锁表格
            Call 计算前基本处理()
            For i = 3 To 17
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(272, i).Value = 0
            Next
            For i = 2 To 17
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(275, i).Value = 0
            Next
            '如果贷款OR折旧摊销计算方法中有一个为方法二，进行下列计算
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法二" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法三" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法四" Then
            '报错（全部在项目计算年限修改模块内报错，这里不再重复报错）
            'MsgBox("输入的<项目计算年限、第二次到第十次项目投资年份、长期贷款年限、固定资产折旧年限、无形资产摊销年限>之中的某一项或者某几项参数不合理，请重新核对！！！为防止计算异常，程序不会自动修改，请重新手动输入！！！")
            'If tznf9 <> 0 Then
            '    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value = tznf9 + 1 '重置回默认值
            'Else
            '    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value = 0
            'End If
        End If
    End Sub
    Sub 年数总和法逐年折旧摊销系数()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim XMJSNX As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        Dim GDZCZJNX As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 7).Value '固定资产折旧年限
        Dim WXZCTXNX As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 7).Value '无形资产摊销年限        
        Call 计算前基本处理()
        '————————————————————————————————————————————————————————————————————————————————————————        
        '折旧摊销计算方法四采用的逐年折旧率摊销率计算
        Dim JS1 As Integer = 0 '投资年份计数
        '第一到第十次投资发生年份
        For i = 28 To 71 Step 43 '估算表行号
            For j = 3 To 11 Step 2 '估算表列号
                JS1 = JS1 + 1
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value > 0 Then
                    '固定资产折旧逐年系数
                    Dim JS2 As Integer = 0
                    For k = 3 To 33 '投资计划与资金筹措表列号
                        If (ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(166, k).Value >= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(166, k).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value + GDZCZJNX) Then
                            JS2 = JS2 + 1
                            If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(167, 2 + JS2).Value > 0 Then '仅写入大于0的值
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(167 + JS1, k).Value = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(167, 2 + JS2).Value
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(167 + JS1, k).Value = Nothing
                            End If
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(167 + JS1, k).Value = Nothing
                        End If
                    Next
                    '无形资产摊销逐年系数
                    Dim JS3 As Integer = 0
                    For k = 3 To 33 '投资计划与资金筹措表列号
                        If (ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(181, k).Value >= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(181, k).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value + GDZCZJNX) Then
                            JS3 = JS3 + 1
                            If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(182, 2 + JS3).Value > 0 Then
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(182 + JS1, k).Value = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(182, 2 + JS3).Value
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(182 + JS1, k).Value = Nothing
                            End If
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(182 + JS1, k).Value = Nothing
                        End If
                    Next
                Else
                    For k = 3 To 33 '投资计划与资金筹措表列号
                        '固定资产折旧逐年系数
                        ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(167 + JS1, k).Value = Nothing
                        '无形资产摊销逐年系数
                        ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(182 + JS1, k).Value = Nothing
                    Next
                End If
            Next
        Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '折旧摊销计算方法二采用的逐年折旧率摊销率计算
        '查找第一个有投产月份数量的年份
        Dim DYGTCNF As Integer = 0 '第一个投产年份
        For i = 3 To 33 '投资计划与资金筹措表列号
            If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value > 0 Then
                'i代表第一个有投产月份数量的年份所在列号
                DYGTCNF = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(156, i).Value
                '记录下第一个有生产月份数的月份数量
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(195, 19).Value = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 1).Value
                Exit For '跳出循环
            End If
        Next
        '生成方法二采用的固定资产折旧率和无形资产摊销率
        Dim JS4 As Integer = 0
        Dim JS5 As Integer = 0
        For i = 3 To 33 '投资计划与资金筹措表列号
            '固定资产折旧率
            If (ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(196, i).Value >= DYGTCNF And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(196, i).Value <= DYGTCNF + GDZCZJNX) Then
                JS4 = JS4 + 1
                If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(167, 2 + JS4).Value > 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(197, i).Value = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(167, 2 + JS4).Value
                Else
                    ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(197, i).Value = Nothing
                End If
            Else
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(197, i).Value = Nothing
            End If
            '无形资产摊销率
            If (ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(196, i).Value >= DYGTCNF And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(196, i).Value <= DYGTCNF + WXZCTXNX) Then
                JS5 = JS5 + 1
                If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(182, 2 + JS5).Value > 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(198, i).Value = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(182, 2 + JS5).Value
                Else
                    ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(198, i).Value = Nothing
                End If
            Else
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(198, i).Value = Nothing
            End If
        Next
    End Sub
    Sub 建设期增值税抵扣系数()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim JSQZZSDKNS As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(9, 7).Value '建设期增值税抵扣年数
        Dim XMJSNX As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '输入的数字必需为从0到31的整数，同时要大于等于项目投资年份最大值与折旧和贷款年限的较大值的和
        If JSQZZSDKNS >= 0 And JSQZZSDKNS <= XMJSNX And Int(JSQZZSDKNS) = JSQZZSDKNS Then
            Call 计算前基本处理()
            Dim JS As Integer = 0 '投资年份计数
            '第一到第十次投资发生年份
            For i = 28 To 71 Step 43 '估算表行号
                For j = 3 To 11 Step 2 '估算表列号
                    JS = JS + 1
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value > 0 Then
                        For k = 3 To 33 '投资计划与资金筹措表列号
                            If (ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(111, k).Value >= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(111, k).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value + JSQZZSDKNS) Then
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(111 + JS, k).Value = 1
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(111 + JS, k).Value = 0
                            End If
                        Next
                    Else
                        For k = 3 To 33 '投资计划与资金筹措表列号
                            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(111 + JS, k).Value = 0
                        Next
                    End If
                Next
            Next
        ElseIf JSQZZSDKNS < 0 Then '如果输入的年份小于0
            Call 计算前基本处理()
            MsgBox("输入的建设期增值税抵扣年限不可以小于0，程序将自动修改年限为0！")
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(9, 7).Value = 0
        ElseIf JSQZZSDKNS > XMJSNX - 1 Then
            Call 计算前基本处理()
            MsgBox("输入的建设期增值税抵扣年限不可以大于项目计算年限，程序将自动修改年限！")
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(9, 7).Value = XMJSNX
        ElseIf JSQZZSDKNS >= 0 And JSQZZSDKNS <= XMJSNX And Int(JSQZZSDKNS) <> JSQZZSDKNS Then
            Call 计算前基本处理()
            MsgBox("输入的建设期增值税抵扣年限必需为0到项目计算年限减一之间的正整数，程序将自动修改年限！")
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(9, 7).Value = 2
        End If
    End Sub
    Sub 是否处于长期贷款宽限期()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim DKKXQNS As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value '贷款宽限期年数
        Dim XMJSNX As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '输入的数字必需为从1到30的整数，同时要小于等于项目计算年分数
        If DKKXQNS >= 1 And DKKXQNS <= XMJSNX And Int(DKKXQNS) = DKKXQNS Then
            '判断每个年份是否处于长期贷款宽限期
            Dim JS1 As Integer = 0
            For i = 100 To 600 '遍历借款还本付息计划表的行号
                If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 54).Value = 7777 Then '标签是7777
                    Call 计算前基本处理()
                    For j = 28 To 71 Step 43 '估算表行号
                        For k = 3 To 11 Step 2 '估算表列号
                            JS1 = JS1 + 1
                            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(j, k).Value > 0 Then
                                For l = 3 To 33 '借款还本付息计划表列号
                                    If (ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, l).Value - ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(j, k).Value) >= 0 And (ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, l).Value - ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(j, k).Value) <= （ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value - 1） Then
                                        ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i + JS1, l).Value = 1 '处于宽限期就赋值为1
                                    Else
                                        ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i + JS1, l).Value = 0
                                    End If
                                Next
                            Else
                                For l = 3 To 33 '借款还本付息计划表列号
                                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i + JS1, l).Value = 0
                                Next
                            End If
                        Next
                    Next
                End If
            Next
        ElseIf DKKXQNS < 1 Then '如果输入的年份小于0
            Call 计算前基本处理()
            MsgBox("输入的长期贷款宽限期年限不可以小于1，程序将自动修改年限为1！")
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value = 1
        ElseIf DKKXQNS > XMJSNX Then
            Call 计算前基本处理()
            MsgBox("输入的长期贷款宽限期年限不可以大于项目计算年限，程序将自动修改年限！")
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value = XMJSNX
        ElseIf Int(DKKXQNS) <> DKKXQNS Then
            Call 计算前基本处理()
            MsgBox("输入的长期贷款宽限期年限必需为1到项目计算年限之间的正整数，程序将自动修改年限！")
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value = 1
        End If
    End Sub
    Sub 隐藏收入税收表中收入为0的行()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '定义局部变量
        Dim js
        '屏蔽屏幕更新，防止屏闪
        ExcelApp.Application.ScreenUpdating = False
        '手动计算，关闭excel的自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationManual
        '解锁收入税收表
        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Unprotect(Password:="wscjc")
        '第1年到第15年
        For i = 1 To 13
            For j = 1 To 10
                '如果某一项收入为0，并且收入类型与输入的类型匹配，则隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 7).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 2).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(5 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(5 + j).EntireRow.Hidden = True
                End If
                '如果某一项收入大于0，并且收入类型与输入的类型匹配，则不隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 7).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 2).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(5 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(5 + j).EntireRow.Hidden = False
                End If
            Next
        Next
        For i = 1 To 13
            For j = 1 To 13
                '隐藏销项增值税为0的
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 7).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 2).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(64 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(64 + j).EntireRow.Hidden = True
                End If
                '增值税大于0，不隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 7).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 2).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(64 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(64 + j).EntireRow.Hidden = False
                End If
            Next
        Next
        '如果增值税退税比例为0，则隐藏增值税退税收入，前15年表格
        If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(172, 34).Value = 0 Then
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(16).EntireRow.Hidden = True
        Else
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(16).EntireRow.Hidden = False
        End If
        '隐藏补贴收入分项中为0的行，前15年表格
        For i = 1 To 13
            For j = 1 To 3
                '如果某一项收入为0，并且收入类型与输入的类型匹配，则隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 7).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 2).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(52 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(52 + j).EntireRow.Hidden = True
                End If
                '如果某一项收入大于0，并且收入类型与输入的类型匹配，则不隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 7).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 2).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(52 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(52 + j).EntireRow.Hidden = False
                End If
            Next
        Next
        '自动给收入税收表的条目编序号，前15年表格
        js = 0 '计数
        For i = 1 To 11
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(5 + i).Height > 0 Then
                js = js + 1
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(5 + i, 1).Value = 1 + js / 10
            End If
        Next
        '自动给销项增值税表编号，前15年表格
        js = 0 '计数
        For i = 1 To 13
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(64 + i).Height > 0 Then
                js = js + 1
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(64 + i, 1).Value = js
            End If
        Next
        '自动给补贴收入分项计算表编号，前15年表格
        js = 0 '计数
        For i = 1 To 3
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(52 + i).Height > 0 Then
                js = js + 1
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(52 + i, 1).Value = js
            End If
        Next
        '第16年到第30年
        For i = 1 To 13
            For j = 1 To 10
                '如果某一项收入为0，并且收入类型与输入的类型匹配，则隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 7).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 2).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(28 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(28 + j).EntireRow.Hidden = True
                End If
                '如果某一项收入大于0，并且收入类型与输入的类型匹配，则隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 7).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 2).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(28 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(28 + j).EntireRow.Hidden = False
                End If
            Next
        Next
        For i = 1 To 13
            For j = 1 To 13
                '隐藏销项增值税为0的
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 7).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 2).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(79 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(79 + j).EntireRow.Hidden = True
                End If
                '增值税大于0，不隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 7).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 2).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(79 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(79 + j).EntireRow.Hidden = False
                End If
            Next
        Next
        '如果增值税退税比例为0，则隐藏增值税退税收入，后15年表格
        If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(172, 34).Value = 0 Then
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(39).EntireRow.Hidden = True
        Else
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(39).EntireRow.Hidden = False
        End If
        '隐藏补贴收入分项中为0的行，后15年表格
        For i = 1 To 13
            For j = 1 To 3
                '如果某一项收入为0，并且收入类型与输入的类型匹配，则隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 7).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 2).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(58 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(58 + j).EntireRow.Hidden = True
                End If
                '如果某一项收入大于0，并且收入类型与输入的类型匹配，则隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 7).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 2).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(58 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(58 + j).EntireRow.Hidden = False
                End If
            Next
        Next
        '自动给收入税收表的条目编序号，后15年表格
        js = 0 '计数
        For i = 1 To 11
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(28 + i).Height > 0 Then
                js = js + 1
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(28 + i, 1).Value = 1 + js / 10
            End If
        Next
        '自动给销项增值税表编号，后15年表格
        js = 0 '计数
        For i = 1 To 13
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(79 + i).Height > 0 Then
                js = js + 1
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(79 + i, 1).Value = js
            End If
        Next
        '自动给补贴收入分项计算表编号，后15年表格
        js = 0 '计数
        For i = 1 To 3
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(58 + i).Height > 0 Then
                js = js + 1
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(58 + i, 1).Value = js
            End If
        Next
        '锁定收入税收表
        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Protect(Password:="wscjc")
        '重新打开excel自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
    End Sub
    Sub 隐藏总成本表中成本为0的行()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '定义局部变量
        Dim js
        '屏蔽屏幕更新，防止屏闪
        ExcelApp.Application.ScreenUpdating = False
        '手动计算，关闭excel的自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationManual
        '解锁总成本表
        ExcelApp.ThisWorkbook.Worksheets("总成本表").Unprotect(Password:="wscjc")
        '第1年到第15年
        For i = 1 To 15
            For j = 1 To 15
                '如果某一项成本为0，并且成本类型与输入的类型匹配，则隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 14).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 9).Value = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(4 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("总成本表").Rows(4 + j).EntireRow.Hidden = True
                End If
                '如果某一项成本大于0，并且成本类型与输入的类型匹配，则隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 14).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 9).Value = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(4 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("总成本表").Rows(4 + j).EntireRow.Hidden = False
                End If
            Next
        Next
        '自动给总成本表的条目编序号
        js = 0 '计数
        For i = 1 To 19
            If ExcelApp.ThisWorkbook.Worksheets("总成本表").Rows(4 + i).Height > 0 Then
                js = js + 1
                ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(4 + i, 1).Value = js
            End If
        Next
        '第16年到第30年
        For i = 1 To 15
            For j = 1 To 15
                '如果某一项成本为0，并且成本类型与输入的类型匹配，则隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 14).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 9).Value = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(36 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("总成本表").Rows(36 + j).EntireRow.Hidden = True
                End If
                '如果某一项成本大于0，并且成本类型与输入的类型匹配，则隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 14).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 9).Value = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(36 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("总成本表").Rows(36 + j).EntireRow.Hidden = False
                End If
            Next
        Next
        '自动给总成本表的条目编序号
        js = 0 '计数
        For i = 1 To 19
            If ExcelApp.ThisWorkbook.Worksheets("总成本表").Rows(36 + i).Height > 0 Then
                js = js + 1
                ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(36 + i, 1).Value = js
            End If
        Next
        '锁定总成本表
        ExcelApp.ThisWorkbook.Worksheets("总成本表").Protect(Password:="wscjc")
        '重新打开excel自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
    End Sub
    Sub 逐年判断是否有收入以及所得税减免计算系数()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim DYGSRNF As Integer = 0 '第一个有收入的年份
        Dim SDSMC As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(11, 7).Value '所得税免除年分数
        Dim SDSJZ As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(12, 7).Value '所得税减征年分数
        Dim XMJSNX As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        If SDSMC + SDSJZ > 0 Then
            '屏蔽屏幕更新，防止屏闪
            ExcelApp.Application.ScreenUpdating = False
            '手动计算，关闭excel的自动计算
            ExcelApp.Application.Calculation = XlCalculation.xlCalculationManual
            '解锁表格
            ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").unProtect(Password:="wscjc")
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").unProtect(Password:="wscjc")
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").unProtect(Password:="wscjc")
            ExcelApp.Calculate() '计算一次
            '输入的所得税免除和减征年数必需为整数，大于等于0且相加小于等于项目计算年数
            If SDSMC >= 0 And SDSJZ >= 0 And Int(SDSMC) = SDSMC And Int(SDSJZ) = SDSJZ And SDSMC + SDSJZ <= XMJSNX Then
                '逐年判断该年是否有收入
                For i = 5 To 35
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(5, i).Value > 0 Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(154, i - 2).Value = 1
                    Else
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(154, i - 2).Value = 0
                    End If
                Next
                '查找第一个有收入的年份序号
                For i = 5 To 35
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(5, i).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(5, i - 1).Value = 0 Then
                        DYGSRNF = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(4, i).Value
                        Exit For
                    End If
                Next
                '所得税免除
                For i = 4 To 34
                    If ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(49, i).Value >= DYGSRNF And ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(49, i).Value <= DYGSRNF + SDSMC - 1 Then
                        ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(51, i).Value = 1
                    Else
                        ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(51, i).Value = 0
                    End If
                Next
                '所得税减征
                For i = 4 To 34
                    If ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(49, i).Value >= DYGSRNF + SDSMC And ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(49, i).Value <= DYGSRNF + SDSMC + SDSJZ - 1 Then
                        ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(52, i).Value = 1
                    Else
                        ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(52, i).Value = 0
                    End If
                Next
                '写入所得税征收比例
                For i = 4 To 34
                    '所得税免除
                    If ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(51, i).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(52, i).Value = 0 Then
                        ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(53, i).Value = 0
                        '所得税减征收
                    ElseIf ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(51, i).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(52, i).Value = 1 Then
                        ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(53, i).Value = 1 - ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(16, 5).Value
                    Else
                        ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(53, i).Value = 1
                    End If
                Next
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(44, 18).Value = "常规设置"
            ElseIf SDSMC < 0 Then
                MsgBox("输入的所得税免除年限不可以小于0，程序将自动修改年限为0！")
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(11, 7).Value = 0
            ElseIf SDSJZ < 0 Then
                MsgBox("输入的所得税减征年限不可以小于0，程序将自动修改年限为0！")
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(12, 7).Value = 0
            ElseIf Int(SDSMC) <> SDSMC Then
                MsgBox("输入的所得税免除年限必须为整数，程序将自动修改年限为0！")
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(11, 7).Value = 0
            ElseIf Int(SDSJZ) <> SDSJZ Then
                MsgBox("输入的所得税减征年限必须为整数，程序将自动修改年限为0！")
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(12, 7).Value = 0
            ElseIf SDSMC + SDSJZ > XMJSNX Then
                MsgBox("输入的所得税免除年限和所得税减征年限之和必须为小于等于项目计算总年限，程序将自动修改年限！")
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(11, 7).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(12, 7).Value = 0
            End If
            '锁定收入税收表
            ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Protect(Password:="wscjc")
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Protect(Password:="wscjc")
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Protect(Password:="wscjc")
            '重新打开excel自动计算
            ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
            '重新打开屏幕更新
            ExcelApp.Application.ScreenUpdating = True
        End If
    End Sub
    Sub 流动资金相关计算()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '解锁表格
        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("资产负债表").unProtect(Password:="wscjc")
        '————————————————————————————————————————————————————————————————————————————————————————
        '第一年的应收账款、原材料、燃料和动力、现金的和
        Dim DYN = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(49, 3).Value + ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(50, 3).Value + ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(51, 3).Value + ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(52, 3).Value
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '手动计算，关闭excel的自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationManual
        '————————————————————————————————————————————————————————————————————————————————————————
        '如果计算期第一年应收账款、原材料、燃料和动力、现金的和大于0，但是计算期第一年的负荷率等于0，则报错
        If DYN > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(5, 3).Value = 0 Then
            MsgBox("计算期第一年设置的负荷率等于0，但是计算期第一年存在流动资金计算，请检查。")
        End If
        '计算一次Excel
        ExcelApp.Calculate()
        If DYN > 0 Then
            '如果第一年有收入和成本，则流动资金的计算年份与项目计算年份一一对应，流动资金计算年限等于项目计算年限
            '流动资金当期增加额计算系数,
            For i = 1 To 31
                If i <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(48, i + 2).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(48, i + 2).Value = 0
                End If
            Next
            '计算一次Excel
            ExcelApp.Calculate()
            '计算逐年应收账款、原材料、燃料和动力、现金
            For i = 1 To 31
                If i <= jsnx Then
                    If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 2).Value = 0 Then
                        '应收账款
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(6, i + 5).Value = 0
                        '原材料
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(8, i + 5).Value = 0
                        '燃料和动力
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(9, i + 5).Value = 0
                        '现金
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(10, i + 5).Value = 0
                    Else
                        '应收账款
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(6, i + 5).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(49, i + 2).Value / ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 2).Value
                        '原材料
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(8, i + 5).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(50, i + 2).Value / ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 2).Value
                        '燃料和动力
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(9, i + 5).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(51, i + 2).Value / ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 2).Value
                        '现金
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(10, i + 5).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(52, i + 2).Value / ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 2).Value
                    End If

                Else
                    '应收账款
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(6, i + 5).Value = 0
                    '原材料
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(8, i + 5).Value = 0
                    '燃料和动力
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(9, i + 5).Value = 0
                    '现金
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(10, i + 5).Value = 0
                End If
            Next
            '计算一次Excel
            ExcelApp.Calculate()
            '计算流动资金逐年贷款金额、付息金额
            '前15年
            For i = 1 To 15
                If i <= jsnx Then
                    '流动资金贷款金额
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(14, i + 4).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(15, i + 5).Value
                    '流动资金付息金额
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(17, i + 4).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(17, i + 5).Value
                Else
                    '流动资金贷款金额
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(14, i + 4).Value = 0
                    '流动资金付息金额
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(17, i + 4).Value = 0
                End If
            Next
            '后16年
            For i = 16 To 31
                If i <= jsnx Then
                    '流动资金贷款金额
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(44, i - 12).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(33, i - 12).Value
                    '流动资金付息金额
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(47, i - 12).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(35, i - 12).Value
                Else
                    '流动资金贷款金额
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(44, i - 12).Value = 0
                    '流动资金付息金额
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(47, i - 12).Value = 0
                End If
            Next
            '计算一次Excel
            ExcelApp.Calculate()
            '计算期末回收流动资金
            '前15年
            If jsnx <= 15 Then
                '流动资金总额
                ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(31, 20).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(13, jsnx + 5).Value
                '自有流动资金
                ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(32, 20).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(14, jsnx + 5).Value
            Else
                '后16年
                '流动资金总额
                ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(31, 20).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(31, jsnx - 12).Value
                '自有流动资金
                ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(32, 20).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(32, jsnx - 12).Value
            End If
            '计算一次Excel
            ExcelApp.Calculate()
            '计算期末最后一年减去自有流动资金
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("资产负债表").Cells(64, i).Value = jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("资产负债表").Cells(66, i).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(32, 20).Value
                Else
                    ExcelApp.ThisWorkbook.Worksheets("资产负债表").Cells(66, i).Value = 0
                End If
            Next
            '计算一次Excel
            ExcelApp.Calculate()
        Else
            '如果第一年没有收入和成本，则流动资金的计算年份与项目计算年份不是一一对应，流动资金计算年份比项目计算年份提前一年，流动资金计算年限等于项目计算年限-1
            '流动资金当期增加额计算系数,
            For i = 1 To 31
                If i <= jsnx - 1 Then
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(48, i + 2).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(48, i + 2).Value = 0
                End If
            Next
            '计算一次Excel
            ExcelApp.Calculate()
            '计算逐年应收账款、原材料、燃料和动力、现金
            For i = 1 To 31
                If i <= jsnx - 1 Then
                    If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 3).Value = 0 Then
                        '应收账款
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(6, i + 5).Value = 0
                        '原材料
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(8, i + 5).Value = 0
                        '燃料和动力
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(9, i + 5).Value = 0
                        '现金
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(10, i + 5).Value = 0
                    Else
                        '应收账款
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(6, i + 5).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(49, i + 3).Value / ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 3).Value
                        '原材料
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(8, i + 5).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(50, i + 3).Value / ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 3).Value
                        '燃料和动力
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(9, i + 5).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(51, i + 3).Value / ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 3).Value
                        '现金
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(10, i + 5).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(52, i + 3).Value / ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 3).Value
                    End If

                Else
                    '应收账款
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(6, i + 5).Value = 0
                    '原材料
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(8, i + 5).Value = 0
                    '燃料和动力
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(9, i + 5).Value = 0
                    '现金
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(10, i + 5).Value = 0
                End If
            Next
            '计算一次Excel
            ExcelApp.Calculate()
            '计算流动资金逐年贷款金额、付息金额
            '第1年
            '流动资金贷款金额
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(14, 5).Value = 0
            '流动资金付息金额
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(17, 5).Value = 0
            '第2-15年
            For i = 2 To 15
                If i <= jsnx Then '流动资金贷款计算到计算年限，而不是计算年限减一
                    '流动资金贷款金额
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(14, i + 4).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(15, i + 4).Value
                    '流动资金付息金额
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(17, i + 4).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(17, i + 4).Value
                Else
                    '流动资金贷款金额
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(14, i + 4).Value = 0
                    '流动资金付息金额
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(17, i + 4).Value = 0
                End If
            Next
            '第16年
            '流动资金贷款金额
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(44, 4).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(15, 20).Value
            '流动资金付息金额
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(47, 4).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(17, 20).Value
            '第17-31年
            For i = 17 To 31
                If i <= jsnx Then '流动资金贷款计算到计算年限，而不是计算年限减一
                    '流动资金贷款金额
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(44, i - 12).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(33, i - 13).Value
                    '流动资金付息金额
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(47, i - 12).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(35, i - 13).Value
                Else
                    '流动资金贷款金额
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(44, i - 12).Value = 0
                    '流动资金付息金额
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(47, i - 12).Value = 0
                End If
            Next
            '计算一次Excel
            ExcelApp.Calculate()
            '计算期末回收流动资金
            '前15年
            If jsnx - 1 <= 15 Then
                '流动资金总额
                ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(31, 20).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(13, jsnx - 1 + 5).Value
                '自有流动资金
                ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(32, 20).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(14, jsnx - 1 + 5).Value
            Else
                '后16年
                '流动资金总额
                ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(31, 20).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(31, jsnx - 1 - 12).Value
                '自有流动资金
                ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(32, 20).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(32, jsnx - 1 - 12).Value
            End If
            '计算一次Excel
            ExcelApp.Calculate()
            '计算期末最后一年不减去自有流动资金
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("资产负债表").Cells(66, i).Value = 0
            Next
            '计算一次Excel
            ExcelApp.Calculate()
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '锁定表格
        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("资产负债表").Protect(Password:="wscjc")
        '重新打开excel自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
    End Sub
    Sub 进入维护模式()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '屏蔽屏幕更新，防止屏闪
        ExcelApp.Application.ScreenUpdating = False
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '读取计数
        Dim Form3 As New 进入维护模式
        If (ExcelApp.Worksheets("建设期时间计划表").Cells(2, 26).Value < 3 And ExcelApp.Worksheets("建设期时间计划表").Cells(2, 26).Value >= 0) Then '最多只可以连续错3次。单元格Z2
            Form3.ShowDialog() '窗口显示
            Form3.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            Form3.TextBox1.Text = Nothing '清空已有的内容
        Else
            MsgBox("已超过最大尝试次数，不可以再尝试输入密码！")
            '锁定表格
            Call 锁定表格()
        End If
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
    End Sub
    Sub 计算前基本处理()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '屏蔽屏幕更新，防止屏闪
        ExcelApp.Application.ScreenUpdating = False
        '手动计算，关闭excel的自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationManual
        '解锁表格
        ExcelApp.ThisWorkbook.Worksheets("估算表").Unprotect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Unprotect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Unprotect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Unprotect(Password:="wscjc")
    End Sub
    Sub 投资各方收益率表格操作()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '屏蔽屏幕更新，防止屏闪
        ExcelApp.Application.ScreenUpdating = False
        '读取输入的投资各方出资比例
        Dim TZGFCZBL(10) As Double '数组，储存投资各方出资比例
        Dim TZFJS As Integer = 0 '投资各方出资比例大于0的数量计数
        For i = 0 To 4
            TZGFCZBL(i) = ExcelApp.Worksheets("建设期时间计划表").Cells(163 + i, 2).Value
            If TZGFCZBL(i) > 0 Then
                TZFJS = TZFJS + 1
            End If
        Next
        '针对读取的投资各方出资比例添加报错功能
        '不能隔行输入
        If TZGFCZBL(0) = 0 And (TZGFCZBL(1) > 0 Or TZGFCZBL(2) > 0 Or TZGFCZBL(3) > 0 Or TZGFCZBL(4) > 0) Then
            MsgBox("投资各方出资比例不可以隔行输入，程序将自动重置回默认值！")
            '屏蔽事件
            ExcelApp.Application.EnableEvents = False
            '投资各方出资比例重置回默认值
            ExcelApp.Worksheets("建设期时间计划表").Cells(163, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(164, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(165, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(166, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(167, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(163, 2).Value = 100
            '开启事件触发
            ExcelApp.Application.EnableEvents = True
            GoTo aaaaa
        End If
        If TZGFCZBL(1) = 0 And (TZGFCZBL(2) > 0 Or TZGFCZBL(3) > 0 Or TZGFCZBL(4) > 0) Then
            MsgBox("投资各方出资比例不可以隔行输入，程序将自动重置回默认值！")
            '屏蔽事件
            ExcelApp.Application.EnableEvents = False
            '投资各方出资比例重置回默认值
            ExcelApp.Worksheets("建设期时间计划表").Cells(163, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(164, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(165, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(166, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(167, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(163, 2).Value = 100
            '开启事件触发
            ExcelApp.Application.EnableEvents = True
            GoTo aaaaa
        End If
        If TZGFCZBL(2) = 0 And (TZGFCZBL(3) > 0 Or TZGFCZBL(4) > 0) Then
            MsgBox("投资各方出资比例不可以隔行输入，程序将自动重置回默认值！")
            '屏蔽事件
            ExcelApp.Application.EnableEvents = False
            '投资各方出资比例重置回默认值
            ExcelApp.Worksheets("建设期时间计划表").Cells(163, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(164, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(165, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(166, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(167, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(163, 2).Value = 100
            '开启事件触发
            ExcelApp.Application.EnableEvents = True
            GoTo aaaaa
        End If
        If TZGFCZBL(3) = 0 And TZGFCZBL(4) > 0 Then
            MsgBox("投资各方出资比例不可以隔行输入，程序将自动重置回默认值！")
            '屏蔽事件
            ExcelApp.Application.EnableEvents = False
            '投资各方出资比例重置回默认值
            ExcelApp.Worksheets("建设期时间计划表").Cells(163, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(164, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(165, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(166, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(167, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(163, 2).Value = 100
            '开启事件触发
            ExcelApp.Application.EnableEvents = True
            GoTo aaaaa
        End If
        '根据输入的情况，自动补全后面一个出资比例，保证总出资比例为100%
        If ExcelApp.Worksheets("建设期时间计划表").Cells(168, 2).Value <= 100 Then '如果总和小于等于100
            If TZFJS = 1 Then
                '屏蔽事件
                ExcelApp.Application.EnableEvents = False
                ExcelApp.Worksheets("建设期时间计划表").Cells(164, 2).Value = 100 - TZGFCZBL(0)
                '开启事件触发
                ExcelApp.Application.EnableEvents = True
            End If
            If TZFJS = 2 Then
                '屏蔽事件
                ExcelApp.Application.EnableEvents = False
                ExcelApp.Worksheets("建设期时间计划表").Cells(165, 2).Value = 100 - TZGFCZBL(1) - TZGFCZBL(0)
                '开启事件触发
                ExcelApp.Application.EnableEvents = True
            End If
            If TZFJS = 3 Then
                '屏蔽事件
                ExcelApp.Application.EnableEvents = False
                ExcelApp.Worksheets("建设期时间计划表").Cells(166, 2).Value = 100 - TZGFCZBL(2) - TZGFCZBL(1) - TZGFCZBL(0)
                '开启事件触发
                ExcelApp.Application.EnableEvents = True
            End If
            If TZFJS = 4 Then
                '屏蔽事件
                ExcelApp.Application.EnableEvents = False
                ExcelApp.Worksheets("建设期时间计划表").Cells(167, 2).Value = 100 - TZGFCZBL(3) - TZGFCZBL(2) - TZGFCZBL(1) - TZGFCZBL(0)
                '开启事件触发
                ExcelApp.Application.EnableEvents = True
            End If
        End If
        '总和不能大于100
        If ExcelApp.Worksheets("建设期时间计划表").Cells(168, 2).Value > 100 Then
            'MsgBox("投资各方出资比例总和不可以大于100%，程序将自动重置回默认值！")
            '屏蔽事件
            ExcelApp.Application.EnableEvents = False
            '投资各方出资比例设置为：第一的不变，第二的变成100减去第一年，第三到第五为0            
            ExcelApp.Worksheets("建设期时间计划表").Cells(164, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(165, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(166, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(167, 2).Value = 0
            ExcelApp.Worksheets("建设期时间计划表").Cells(164, 2).Value = 100 - ExcelApp.Worksheets("建设期时间计划表").Cells(163, 2).Value
            '开启事件触发
            ExcelApp.Application.EnableEvents = True
            GoTo aaaaa
        End If
        '如果输入到了投资方5输入比例，但是总和小于100
        If TZGFCZBL(4) > 0 And ExcelApp.Worksheets("建设期时间计划表").Cells(168, 2).Value < 100 Then
            'MsgBox("投资各方出资比例总和不可以小于100%，程序将自动重置回默认值！")
            '屏蔽事件
            ExcelApp.Application.EnableEvents = False
            '投资各方出资比例设置为：第一到第四不变，第五等于100减去前四个之和
            ExcelApp.Worksheets("建设期时间计划表").Cells(167, 2).Value = 100 - ExcelApp.Worksheets("建设期时间计划表").Cells(163, 2).Value - ExcelApp.Worksheets("建设期时间计划表").Cells(164, 2).Value - ExcelApp.Worksheets("建设期时间计划表").Cells(165, 2).Value - ExcelApp.Worksheets("建设期时间计划表").Cells(166, 2).Value
            '开启事件触发
            ExcelApp.Application.EnableEvents = True
            GoTo aaaaa
        End If
aaaaa：
        '再重新读取一次各方投资比例
        For i = 0 To 4
            TZGFCZBL(i) = ExcelApp.Worksheets("建设期时间计划表").Cells(163 + i, 2).Value
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '将输入的投资各方比例写入表格
        '解锁表格
        ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Unprotect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Unprotect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Unprotect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Unprotect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Unprotect(Password:="wscjc")
        '投资方1
        ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(4, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(163, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(5, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(163, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(6, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(163, 2).Value / 100
        '投资方2
        ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(4, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(5, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(6, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 2).Value / 100
        '投资方3
        ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(4, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(5, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(6, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 2).Value / 100
        '投资方4
        ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(4, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(166, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(5, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(166, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(6, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(166, 2).Value / 100
        '投资方5
        ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(4, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(167, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(5, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(167, 2).Value / 100
        ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(6, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(167, 2).Value / 100
        '————————————————————————————————————————————————————————————————————————————————————————
        '根据输入的投资各方出资比例，自动显示或者隐藏相关工作表
        If TZGFCZBL(0) > 0 Then
            '显示工作表
            ExcelApp.Worksheets("投资方1现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
        Else
            '彻底隐藏工作表
            ExcelApp.Worksheets("投资方1现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        End If
        If TZGFCZBL(1) > 0 Then
            '显示工作表
            ExcelApp.Worksheets("投资方2现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
        Else
            '彻底隐藏工作表
            ExcelApp.Worksheets("投资方2现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        End If
        If TZGFCZBL(2) > 0 Then
            '显示工作表
            ExcelApp.Worksheets("投资方3现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
        Else
            '彻底隐藏工作表
            ExcelApp.Worksheets("投资方3现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        End If
        If TZGFCZBL(3) > 0 Then
            '显示工作表
            ExcelApp.Worksheets("投资方4现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
        Else
            '彻底隐藏工作表
            ExcelApp.Worksheets("投资方4现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        End If
        If TZGFCZBL(4) > 0 Then
            '显示工作表
            ExcelApp.Worksheets("投资方5现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
        Else
            '彻底隐藏工作表
            ExcelApp.Worksheets("投资方5现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '写入计算模式
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(168, 7).Value = "相同"
        '————————————————————————————————————————————————————————————————————————————————————————
        '锁定表格
        ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Protect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Protect(Password:="wscjc")
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算流动资金
        Call 流动资金相关计算()
        '计算回收期
        Call 投资回收期计算()
        '开启屏幕更新
        ExcelApp.Application.ScreenUpdating = True
    End Sub

    Sub 解锁表格()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim ZTJC As Integer = Excel版本号验证()
        If ZTJC = 1 Then
            Call 锁定表格()
            ZTJC = 0
            Exit Sub
        End If
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
    Sub 锁定表格()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
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
    Sub 自保护程序()
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————  
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
                'Call 获取本地服务器版本信息并验证()
                Call 获取本机MAC地址并验证()
                Call 网络时间和本地时间交替验证()
            End If
        Catch ex As Exception
            '发生任何异常，则进行自保护验证
            'Call 获取本地服务器版本信息并验证()
            Call 获取本机MAC地址并验证()
            Call 网络时间和本地时间交替验证()
            Exit Sub
        End Try
    End Sub
    Sub 获取本地服务器版本信息并验证()
        'On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————  
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
    Sub 获取本机MAC地址并验证()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
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
            'MAC地址白名单(随便写一个)
            If MAC = "45:40:C1:8F:AD:4A" Then
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
    Sub 网络时间和本地时间交替验证()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
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
                If GetDate >= #01/01/2020# Then '月/日/年，验证网络时间
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
            '进程暂停一段时间（2000分钟）
            Threading.Thread.Sleep(120000000)
            Dim Local_Time = Date.Now '获取系统本地时间
            Dim Dead_Time = Convert.ToDateTime("2020/01/01 01:00:00") '设定程序有效期截止时间
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
    Sub 获取系统时间并验证()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        Dim Local_Time = Date.Now '获取系统本地时间
        Dim Dead_Time = Convert.ToDateTime("2020/01/01 01:00:00") '设定程序有效期截止时间
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
    Sub 程序联网验证()
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————
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
                If GetDate >= #01/01/2020# Then '月/日/年
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
    '————————————————————————————————————————————————————————————————————————————————————————
    '————————————————————————————————————————————————————————————————————————————————————————        
    'Public ZTJC '状态监测
    'Form1是计算进度显示，在每个有需要的Sub内单独实例化
    'Public Shared Form2 As New 逐年衰减系数设置
    'Public Shared Form3 As New 进入维护模式
    'Public Shared Form4 As New 修理费率逐年变化设置
    'Public Shared Form5 As New 建设期资金运用方式设置
    'Public Shared Form6 As New 设置所得税减免和增值税退税包含内容
    'Public Shared Form7 As New 设置补贴收入计算年限
    'Public Shared Form8 As New 每次投资设置不同的折旧摊销计算方式
    'Public Shared Form9 As New 每次投资设置不同的长期贷款还款和宽限年限
    'Public Shared Form10 As New 选择单因素敏感性分析内容
    'Public Shared Form11 As New 经济评价表格导出
    'Public Shared Form12 As New 设置接入费计算方式
    'Public Shared Form13 As New 设置部分销售收入和经营成本计算年限
    'Public Shared Form14 As New 投资各方设置不同的出资比例资产处置比例和利润分配比例

End Class
