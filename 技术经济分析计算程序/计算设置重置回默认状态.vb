Imports Microsoft.Office.Interop
Module 计算设置重置回默认状态
    Sub 计算设置重置回默认状态_part1(ExcelApp As Object)
        On Error Resume Next
        '<确定建设期时间计划>、<清空建设期时间计划>，彻底重置回初始状态
        '————————————————————————————————————————————————————————————————————————————————————————

        '建设期资金运用模式
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 7).Value = "动态"
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 7).Value = "相同"
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '记录折旧摊销和长期贷款计算模式
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(166, 7).Value = "相同"
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(167, 7).Value = "相同"
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '将投资各方收益计算计算模式
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(168, 7).Value = "相同"
        '————————————————————————————————————————————————————————————————————————————————————————
        '接入费计算模式
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(41, 18).Value = "逐年达产率增加值"
        '————————————————————————————————————————————————————————————————————————————————————————
        '补贴收入计算模式
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 18).Value = "保持每年100%"
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(43, 18).Value = "保持每年100%"
        '————————————————————————————————————————————————————————————————————————————————————————
        '所得税免除和减征
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(44, 18).Value = "常规设置"
        '增值税退税收入
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(45, 18).Value = "常规设置"
        '————————————————————————————————————————————————————————————————————————————————————————
        '光伏逐年衰减系数
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(46, 18).Value = "常规设置"
        '蓄电池逐年衰减系数
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(47, 18).Value = "常规设置"
        '风力发电逐年系数	
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(48, 18).Value = "常规设置"
        '————————————————————————————————————————————————————————————————————————————————————————
        '部分收入成本计算模式
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(9, 11).Value = "折算" Then
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(49, 18).Value = "逐年投产月份比例"
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(50, 18).Value = "逐年投产月份比例"
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(51, 18).Value = "逐年投产月份比例"
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(52, 18).Value = "逐年投产月份比例"
        Else
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(49, 18).Value = "保持每年100%"
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(50, 18).Value = "保持每年100%"
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(51, 18).Value = "保持每年100%"
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(52, 18).Value = "保持每年100%"
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '将计算期末回收固定资产残值重置回默认状态
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(53, 18).Value = "期末回收残值"
        '————————————————————————————————————————————————————————————————————————————————————————
        '设备修理费
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(54, 18).Value = "常规设置"
        '材料费其它费
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(55, 18).Value = "常规设置"
        '————————————————————————————————————————————————————————————————————————————————————————
        '光伏、风电修理费计算基数设置
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(56, 18).Value = "投资额百分比(%)"
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(57, 18).Value = "投资额百分比(%)"
    End Sub
    Sub 计算设置重置回默认状态_part2(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '<确定建设期时间计划>、<清空建设期时间计划>，<确定估算表参数设置>涉及到的内容
        '————————————————————————————————————————————————————————————————————————————————————————        
        '建设期资金运用方式
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 7).Value = "相同" Then
            '资本金比例
            For i = 1 To 10
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 4, 22).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 5).Value
            Next
            '—————————————————————————————————————————————————————————————————————————————————————————
            '建设期贷款计息次数
            Dim jsqdkjxcs = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 9).Value
            '计算复利贷款利率
            Dim dkll = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 9).Value
            Dim dkll_fl As Double = (1 + dkll / jsqdkjxcs) ^ jsqdkjxcs - 1
            '常规设备
            For i = 12 To 21
                '长期贷款年利率
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 39).Value = dkll_fl
            Next
            '其它设备
            For i = 5 To 14
                '长期贷款利率
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 42).Value = dkll_fl
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 45).Value = dkll_fl
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 48).Value = dkll_fl
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 51).Value = dkll_fl
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 54).Value = dkll_fl
            Next
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————
        '折旧摊销计算方式
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(166, 7).Value = "相同" Then
            '常规设备写入各种系数默认值
            For i = 4 To 13
                '固定资产折旧年限
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 22).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 7).Value
                '固定资产残值率
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 24).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 5).Value
                '无形资产摊销年限
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 26).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 7).Value
            Next
            '其它设备
            For i = 4 To 13
                '固定资产折旧年限
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 29).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 7).Value
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 32).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 7).Value
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 35).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 7).Value
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 38).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 7).Value
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 41).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 7).Value
            Next
            For i = 15 To 24
                '固定资产残值率
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 29).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 5).Value
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 32).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 5).Value
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 35).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 5).Value
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 38).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 5).Value
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 41).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 5).Value
            Next
            For i = 26 To 35
                '无形资产摊销年限
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 29).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 7).Value
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 32).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 7).Value
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 35).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 7).Value
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 38).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 7).Value
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 41).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 7).Value
            Next
        End If
        '————————————————————————————————————————————————————————————————————————————————————————       
        '长期贷款计算方式
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(167, 7).Value = "相同" Then
            '常规设备
            For i = 23 To 32
                '长期贷款还款年限
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 39).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value
            Next
            For i = 34 To 43
                '长期贷款宽限年限
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 39).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value
            Next
            '其它设备
            For i = 16 To 25
                '长期贷款年限
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 42).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 45).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 48).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 51).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 54).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value
            Next
            For i = 27 To 36
                '宽限年限
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 42).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 45).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 48).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 51).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 54).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value
            Next
        End If
        '———————————————————————————————————————————————————————————————————————————————————————————
        '投资各方收益方式
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(168, 7).Value = "相同" Then
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
        End If
        '————————————————————————————————————————————————————————————————————————————————————————      
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
    Sub 计算设置重置回默认状态_part3(ExcelApp As Object)
        On Error Resume Next
        '与计算年限相关的内容
        '————————————————————————————————————————————————————————————————————————————————————————
        '读取项目计算年限
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '————————————————————————————————————————————————————————————————————————————————————————       
        '读取逐年投产月份数
        Dim tcyfs_list(31) As Integer
        For i = 3 To 33
            tcyfs_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value
        Next
        '————————————————————————————————————————————————————————————————————————————————————————       
        '接入费计算模式
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(41, 18).Value = "逐年达产率增加值" Then
            '接入费计算模式（第2年到计算期最后一年，根据逐年达产率增加值计算）
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
        ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(41, 18).Value = "保持每年100%" Then
            For i = 3 To 33
                If i - 2 <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1
                End If
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(41, 18).Value = "逐年达产率" Then
            '前15年
            For i = 3 To 17
                If i - 2 <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value
                End If
            Next
            '16-31年
            For i = 18 To 33
                If i - 2 <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 16).Value
                End If
            Next
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '补贴收入逐年计算系数（第2年到计算期最后一年，每年都是100%）
        For i = 3 To 33
            If i - 2 >= 2 And i - 2 <= jsnx Then
                '补贴收入1
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 18).Value = "保持每年100%" Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 1
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 18).Value = "逐年投产月份比例" Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 1 * tcyfs_list(i - 2) / 12
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 18).Value = "逐年达产率" Then
                    '前15年
                    If i - 2 <= 15 Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value
                        '16-31年
                    ElseIf i - 2 >= 16 Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 16).Value
                    End If
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 18).Value = "逐年达产率增加值" Then
                    '第1年
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, 3).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, 3).Value
                    '第2-15年
                    If i - 2 >= 2 And i - 2 <= 15 Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i - 1).Value)
                    End If
                    '第16年
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, 18).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, 3).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, 17).Value)
                    '第17-31年
                    If i - 2 >= 17 Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15 - 1).Value)
                    End If
                End If
                '补贴收入2
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(43, 18).Value = "保持每年100%" Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 1
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(43, 18).Value = "逐年投产月份比例" Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 1 * tcyfs_list(i - 2) / 12
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(43, 18).Value = "逐年达产率" Then
                    '前15年
                    If i - 2 <= 15 Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value
                        '16-31年
                    ElseIf i - 2 >= 16 Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 16).Value
                    End If
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(43, 18).Value = "逐年达产率增加值" Then
                    '第1年
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, 3).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, 3).Value
                    '第2-15年
                    If i - 2 >= 2 And i - 2 <= 15 Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i - 1).Value)
                    End If
                    '第16年
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, 18).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, 3).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, 17).Value)
                    '第17-31年
                    If i - 2 >= 17 Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15 - 1).Value)
                    End If
                End If
                '光伏补贴收入
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(165, i).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(165, i).Value = 0
            End If
        Next
        '—————————————————————————————————————————————————————————————————————————————————————
        '所得税减免
        '此处不再需要
        '—————————————————————————————————————————————————————————————————————————————————————
        '增值税退税
        '此处不再需要
        '—————————————————————————————————————————————————————————————————————————————————————
        '光伏逐年衰减系数
        '此处不再需要
        '—————————————————————————————————————————————————————————————————
        '蓄电池逐年衰减系数
        '此处不再需要
        '—————————————————————————————————————————————————————————————————
        '风力发电逐年系数	
        '此处不再需要
        '—————————————————————————————————————————————————————————————————
        '购电容量费成本
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(49, 18).Value = "逐年投产月份比例" Then
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(145, i).Value = 1 * tcyfs_list(i - 2) / 12
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(49, 18).Value = "保持每年100%" Then
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(145, i).Value = 1
            Next
        End If
        '—————————————————————————————————————————————————————————————————
        '城市管廊成本
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(50, 18).Value = "逐年投产月份比例" Then
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(146, i).Value = 1 * tcyfs_list(i - 2) / 12
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(50, 18).Value = "保持每年100%" Then
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(146, i).Value = 1
            Next
        End If
        '—————————————————————————————————————————————————————————————————
        '人员工资
        '如果是逐年递增情况，不进行修改
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(51, 18).Value = "逐年投产月份比例" Then
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(147, i).Value = 1 * tcyfs_list(i - 2) / 12
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(51, 18).Value = "保持每年100%" Then
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(147, i).Value = 1
            Next
        End If
        '—————————————————————————————————————————————————————————————————
        '充电桩收入
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(52, 18).Value = "逐年投产月份比例" Then
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 1 * tcyfs_list(i - 2) / 12
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(52, 18).Value = "保持每年100%" Then
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 1
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(52, 18).Value = "逐年达产率" Then
            For i = 3 To 33
                '前15年
                If i - 2 <= 15 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value
                    '16-31年
                ElseIf i - 2 >= 16 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 16).Value
                End If
            Next
        End If
        '——————————————————————————————————————————————————————————————————————————————————————
        '修理费率重置回默认值
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(54, 18).Value = "常规设置" Then
            '修理费率默认值
            Dim xlfl_mr = 默认逐年修理费率(ExcelApp)
            '常规设备修理费率（%）
            Dim xlfl_cg_mr_list = xlfl_mr(0)
            '燃机修理费率（%）
            Dim xlfl_rj_mr_list = xlfl_mr(1)
            '蓄电池修理费率（%）
            Dim xlfl_xdc_mr_list = xlfl_mr(2)
            '光伏设备修理费率（%）
            Dim xlfl_gf_mr_list = xlfl_mr(3)
            '暖通设备修理费率（%）
            Dim xlfl_nt_mr_list = xlfl_mr(4)
            '风电设备修理费率（%）
            Dim xlfl_fd_mr_list = xlfl_mr(5)
            '写入Excel
            For i = 3 To 33
                If i - 2 <= jsnx Then
                    '常规设备修理费率（%）
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value = xlfl_cg_mr_list(i - 2)
                    '燃机修理费率（%）
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(154, i).Value = xlfl_rj_mr_list(i - 2)
                    '蓄电池修理费率（%）
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(155, i).Value = xlfl_xdc_mr_list(i - 2)
                    '光伏设备修理费率（%）
                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(56, 18).Value = "投资额百分比(%)" Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(156, i).Value = xlfl_gf_mr_list(i - 2)
                    End If
                    '暖通设备修理费率（%）
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(157, i).Value = xlfl_nt_mr_list(i - 2)
                    '风电设备修理费率（%）
                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(57, 18).Value = "投资额百分比(%)" Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value = xlfl_fd_mr_list(i - 2)
                    End If
                Else
                    '常规设备修理费率（%）
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value = 0
                    '燃机修理费率（%）
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(154, i).Value = 0
                    '蓄电池修理费率（%）
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(155, i).Value = 0
                    '光伏设备修理费率（%）
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(156, i).Value = 0
                    '暖通设备修理费率（%）
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(157, i).Value = 0
                    '风电设备修理费率（%）
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value = 0
                End If
            Next
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '其它设备计算基数扣除默认设置
            Dim kcje_mr = 设备修理费计算基数扣除默认设置(ExcelApp)
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(103, 3).Value = kcje_mr(0)
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(103, 4).Value = kcje_mr(1)
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(103, 5).Value = kcje_mr(2)
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(103, 6).Value = kcje_mr(3)
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(103, 7).Value = kcje_mr(4)
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(103, 8).Value = kcje_mr(5)
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(104, 3).Value = kcje_mr(6)
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(104, 4).Value = kcje_mr(7)
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(104, 5).Value = kcje_mr(8)
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(104, 6).Value = kcje_mr(9)
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(104, 7).Value = kcje_mr(10)
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(104, 8).Value = kcje_mr(11)
        End If
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '材料费率、其它费率重置回默认值
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(55, 18).Value = "常规设置" Then
            '读取材料费其它费计算默认值
            Dim clqtfl = 默认逐年材料费率其它费率(ExcelApp)
            '逐年材料费默认值
            Dim clfl_mr_rj_list = clqtfl(0)
            Dim clfl_mr_rm_list = clqtfl(1)
            Dim clfl_mr_ljfd_list = clqtfl(2)
            Dim clfl_mr_glgr_list = clqtfl(3)
            Dim clfl_mr_gf_list = clqtfl(4)
            Dim clfl_mr_fd_list = clqtfl(5)
            '逐年其它费默认值
            Dim qtfl_mr_rj_list = clqtfl(6)
            Dim qtfl_mr_rm_list = clqtfl(7)
            Dim qtfl_mr_ljfd_list = clqtfl(8)
            Dim qtfl_mr_glgr_list = clqtfl(9)
            Dim qtfl_mr_gf_list = clqtfl(10)
            Dim qtfl_mr_fd_list = clqtfl(11)
            '写入Excel
            For i = 3 To 33 '列
                If i - 2 <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(182, i).Value = clfl_mr_rj_list(i - 2)
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(183, i).Value = clfl_mr_rm_list(i - 2)
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(184, i).Value = clfl_mr_ljfd_list(i - 2)
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(185, i).Value = clfl_mr_glgr_list(i - 2)
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(186, i).Value = clfl_mr_gf_list(i - 2)
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(187, i).Value = clfl_mr_fd_list(i - 2)
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(188, i).Value = qtfl_mr_rj_list(i - 2)
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(189, i).Value = qtfl_mr_rm_list(i - 2)
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(190, i).Value = qtfl_mr_ljfd_list(i - 2)
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(191, i).Value = qtfl_mr_glgr_list(i - 2)
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(192, i).Value = qtfl_mr_gf_list(i - 2)
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(193, i).Value = qtfl_mr_fd_list(i - 2)
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(182, i).Value = 0
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(183, i).Value = 0
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(184, i).Value = 0
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(185, i).Value = 0
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(186, i).Value = 0
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(187, i).Value = 0
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(188, i).Value = 0
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(189, i).Value = 0
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(190, i).Value = 0
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(191, i).Value = 0
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(192, i).Value = 0
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(193, i).Value = 0
                End If
            Next
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            Dim kcje_mr = 材料费其它费计算基数扣除默认设置(ExcelApp)
            '设备材料费其它费计算基数扣除计算模式写入Excel，供其它程序调用
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(106, 3).Value = kcje_mr(0)
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(106, 4).Value = kcje_mr(1)
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(106, 5).Value = kcje_mr(2)
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(106, 6).Value = kcje_mr(3)
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(106, 7).Value = kcje_mr(4)
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(106, 8).Value = kcje_mr(5)
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(106, 9).Value = kcje_mr(6)
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(107, 3).Value = kcje_mr(7)
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(107, 4).Value = kcje_mr(8)
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(107, 5).Value = kcje_mr(9)
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(107, 6).Value = kcje_mr(10)
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(107, 7).Value = kcje_mr(11)
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(107, 8).Value = kcje_mr(12)
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(107, 9).Value = kcje_mr(13)
        End If
        '————————————————————————————————————————————————————————————————————————————————————————      
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
End Module
