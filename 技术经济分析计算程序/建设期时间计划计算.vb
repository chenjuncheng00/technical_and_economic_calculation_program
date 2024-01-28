Module 建设期时间计划计算
    Sub 建设期时间计划计算(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '读取输入数据
        '读取输入的建设期时间计划数据
        Dim jsqsjjhsj = 读取建设期时间计划数据(ExcelApp)
        '开始年份列表
        Dim year_start_list = jsqsjjhsj(0)
        '结束年份列表
        Dim year_end_list = jsqsjjhsj(1)
        '开始月份列表
        Dim month_start_list = jsqsjjhsj(2)
        '结束月份列表
        Dim month_end_list = jsqsjjhsj(3)
        '手动输入的年份
        Dim year_shuru_list = jsqsjjhsj(4)
        '手动输入的月份
        Dim month_shuru_list = jsqsjjhsj(5)
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------
        '计算建设期时间计划
        Dim ans_jsq = 建设期时间计划计算_main(year_start_list, year_end_list, month_start_list, month_end_list, year_shuru_list, month_shuru_list)
        '读取计算结果
        Dim ans_jsq_index_list = ans_jsq(0)
        Dim ans_year_list = ans_jsq(1)
        Dim ans_month_list = ans_jsq(2)
        Dim ans_month_start_list = ans_jsq(3)
        Dim ans_month_end_list = ans_jsq(4)
        '计算结果写入Excel
        '写入<建设期时间计划表>
        Dim js As Integer = 0
        For i = 1 To 31
            '如果处于建设期
            If ans_jsq_index_list(i) = 1 Then
                '投资年份序号
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(205 + js, 7).Value = ans_year_list(i)
                '投产开始月份序号
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(205 + js, 15).Value = ans_month_start_list(i)
                '投产结束月份序号
                ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(205 + js, 16).Value = ans_month_end_list(i)
                '计数+1
                js += 1
            End If
        Next
        '写入<投资计划与资金筹措表>
        For i = 1 To 31
            'ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 2).Value = ans_month_list(i)
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(162, i + 2).Value = ans_jsq_index_list(i)
        Next
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '计算一次Excel
        ExcelApp.Calculate()
    End Sub
    Function 建设期时间计划计算_main(year_start_list As Array, year_end_list As Array, month_start_list As Array,
                                     month_end_list As Array, year_shuru_list As Array, month_shuru_list As Array)
        'year_start_list：建设期开始的年份，列表，长度10
        'year_end_list：建设期结束的年份，列表，长度10
        'month_start_list：建设期开始的月份，列表，长度10
        'month_end_list：建设期结束的月份，列表，长度10
        'year_shuru_list：手动输入的建设期年份，列表，长度10
        'month_shuru_list：手动输入的建设期开始投产的月份(建设期于这个月份之前的一个月结束)，列表，长度10

        '储存最终结果的31长度的列表(建设期年份序号，投产的月份数，建设期开始月份序号，建设期结束月份序号)
        Dim ans_year_list(31) As Integer
        Dim ans_month_list(31) As Integer
        Dim ans_month_start_list(31) As Integer
        Dim ans_month_end_list(31) As Integer
        '储存处于建设期的年份标记：(1：处于建设期，0：不处于建设期)
        Dim ans_jsq_index_list(31) As Integer
        '计算自动计算的结果
        Dim ans_zd = 建设期时间计划计算_自动计算(year_start_list, year_end_list, month_start_list, month_end_list)
        Dim ans_year_list_zd = ans_zd(0)
        Dim ans_month_list_zd = ans_zd(1)
        Dim ans_month_start_list_zd = ans_zd(2)
        Dim ans_month_end_list_zd = ans_zd(3)
        '计算手动输入的结果
        Dim ans_sd = 建设期时间计划计算_手动输入(year_start_list, year_shuru_list, month_shuru_list)
        Dim ans_year_list_sd = ans_sd(0)
        Dim ans_month_list_sd = ans_sd(1)
        Dim ans_month_start_list_sd = ans_sd(2)
        Dim ans_month_end_list_sd = ans_sd(3)
        '判断两个结果的投产月份数是否一致，如果不一致，则用手动计算的结果为准
        For i = 1 To 31
            If ans_year_list_zd(i) <> 0 Or ans_year_list_sd(i) <> 0 Then
                '建设期年份序号
                If ans_year_list_zd(i) = 0 And ans_year_list_sd(i) <> 0 Then
                    ans_year_list(i) = ans_year_list_sd(i)
                Else
                    ans_year_list(i) = ans_year_list_zd(i)
                End If
                '1：处于建设期，0：不处于建设期
                ans_jsq_index_list(i) = 1
                '比较两个计算的投产月份数结果
                If ans_month_list_zd(i) <> ans_month_list_sd(i) And ans_year_list_sd(i) <> 0 Then
                    '如果不相同，用手动输入的计算结果，代替自动输入的计算结果
                    ans_month_list(i) = ans_month_list_sd(i)
                    ans_month_start_list(i) = ans_month_start_list_sd(i)
                    ans_month_end_list(i) = ans_month_start_list_sd(i) + ans_month_list_sd(i) - 1
                Else
                    ans_month_list(i) = ans_month_list_zd(i)
                    ans_month_start_list(i) = ans_month_start_list_zd(i)
                    ans_month_end_list(i) = ans_month_end_list_zd(i)
                End If
            Else
                ans_jsq_index_list(i) = 0 ' 1：处于建设期，0：不处于建设期    
                '不处于建设期，投产月份数设置为12
                ans_month_list(i) = 12
                ans_month_start_list(i) = 0
                ans_month_end_list(i) = 0
            End If
        Next
        '结果列表
        Dim ans(4)
        '返回结果
        ans(0) = ans_jsq_index_list
        ans(1) = ans_year_list
        ans(2) = ans_month_list
        ans(3) = ans_month_start_list
        ans(4) = ans_month_end_list
        '返回结果
        Return ans
    End Function
    Function 建设期时间计划计算_自动计算(year_start_list As Array, year_end_list As Array, month_start_list As Array,
                                         month_end_list As Array)
        'year_start_list：建设期开始的年份，列表，长度10
        'year_end_list：建设期结束的年份，列表，长度10
        'month_start_list：建设期开始的月份，列表，长度10
        'month_end_list：建设期结束的月份，列表，长度10

        '储存最终结果的31长度的列表(建设期年份序号，投产的月份数，建设期开始月份序号，建设期结束月份序号)
        Dim ans_year_list(31) As Integer
        Dim ans_month_list(31) As Integer
        Dim ans_month_start_list(31) As Integer
        Dim ans_month_end_list(31) As Integer

        '项目起始年份
        Dim year_0 As Integer = year_start_list(1)
        For i = 1 To 10
            Dim year_start As Integer = year_start_list(i)
            Dim year_end As Integer = year_end_list(i)
            Dim month_start As Integer = month_start_list(i)
            Dim month_end As Integer = month_end_list(i)
            If year_start <> 0 And year_end <> 0 And month_end <> 0 Then
                '建设期年份序号和投产月份数
                Dim jsq_list = 建设期年份序号和投产月份数_自动计算(year_0, year_start, year_end, month_start, month_end)
                Dim ans_year = jsq_list(0)
                Dim ans_month = jsq_list(1)
                Dim ans_month_start = jsq_list(2)
                Dim ans_month_end = jsq_list(3)
                '记录年份序号的开始和结束的序号
                Dim index_start As Integer = year_start - year_0 + 1
                Dim index_end As Integer = year_end - year_0 + 1
                '记录进31长度的列表
                Dim js As Integer = 1 '计数
                For j = index_start To index_end
                    ans_year_list(j) = ans_year(js)
                    ans_month_list(j) = ans_month(js)
                    ans_month_start_list(j) = ans_month_start(js)
                    ans_month_end_list(j) = ans_month_end(js)
                    js += 1
                Next
            End If
        Next
        '结果列表
        Dim ans(3)
        '返回结果
        ans(0) = ans_year_list
        ans(1) = ans_month_list
        ans(2) = ans_month_start_list
        ans(3) = ans_month_end_list
        '返回结果
        Return ans
    End Function
    Function 建设期年份序号和投产月份数_自动计算(year_0 As Integer, year_start As Integer, year_end As Integer,
                                                 month_start As Integer, month_end As Integer)
        'year_0：整个项目建设期开始的第一个年份
        'year_start：建设期开始的年份
        'year_end：建设期结束的年份
        'month_start：建设期开始的月份
        'month_end：建设期结束的月份

        '横跨的年份数量
        Dim year_num As Integer = year_end - year_start + 1
        '年份序号列表
        Dim year_list(year_num) As Integer
        '投产月份数列表
        Dim month_list(year_num) As Integer
        '开始的月份列表
        Dim month_start_list(year_num) As Integer
        '结束的月份列表
        Dim month_end_list(year_num) As Integer
        '计数
        Dim js As Integer = 0
        For i = 1 To year_num
            '年份序号列表
            year_list(i) = year_start - year_0 + 1 + js
            '计算建设期开始月份序号和结束月份序号
            If i = 1 And year_num > 1 Then
                '第一年
                month_start_list(i) = month_start
                month_end_list(i) = 12
            ElseIf i > 1 And i < year_num And year_num > 1 Then
                '中间年份
                month_start_list(i) = 1
                month_end_list(i) = 12
            ElseIf i = year_num And year_num > 1 Then
                '最后一年
                month_start_list(i) = 1
                month_end_list(i) = month_end
            Else
                '开始和结束的月份序号列表
                month_start_list(i) = month_start
                month_end_list(i) = month_end
            End If
            '计算投产月份数列表
            '如果是投产计划的最后一年，则计算投产月份数
            If i = year_num Then
                '投产月份数量
                month_list(i) = 12 - month_end
            Else
                month_list(i) = 0
            End If
            '投产月份数列表
            js += 1
        Next
        '结果列表
        Dim ans(3)
        '返回结果
        ans(0) = year_list
        ans(1) = month_list
        ans(2) = month_start_list
        ans(3) = month_end_list
        '返回结果
        Return ans
    End Function
    Function 建设期时间计划计算_手动输入(year_start_list As Array, year_shuru_list As Array, month_shuru_list As Array)
        'year_start：建设期开始的年份，长度10
        'year_shuru_list：手动输入的建设期年份，列表，长度10
        'month_shuru_list：手动输入的建设期开始投产的月份(建设期于这个月份之前的一个月结束)，列表，长度10

        '储存最终结果的31长度的列表(建设期年份序号，投产的月份数，建设期开始月份序号，建设期结束月份序号)
        Dim ans_year_list(31) As Integer
        Dim ans_month_list(31) As Integer
        Dim ans_month_start_list(31) As Integer
        Dim ans_month_end_list(31) As Integer

        '项目起始年份
        Dim year_0 As Integer = year_start_list(1)
        For i = 1 To 10
            '建设期年份序号和投产月份数
            Dim year_shuru As Integer = year_shuru_list(i)
            Dim month_shuru As Integer = month_shuru_list(i)
            If year_shuru <> 0 And month_shuru <> 0 Then
                Dim ans_jsq = 建设期年份序号和投产月份数_手动输入(year_0, year_shuru, month_shuru)
                Dim year_index As Integer = ans_jsq(0)
                Dim month_sc As Integer = ans_jsq(1)
                '记录进31长度的列表
                ans_year_list(year_index) = year_index
                ans_month_list(year_index) = month_sc
                ans_month_start_list(year_index) = 1
                ans_month_end_list(year_index) = month_shuru
            End If
        Next
        '结果列表
        Dim ans(3)
        '返回结果
        ans(0) = ans_year_list
        ans(1) = ans_month_list
        ans(2) = ans_month_start_list
        ans(3) = ans_month_end_list
        '返回结果
        Return ans
    End Function
    Function 建设期年份序号和投产月份数_手动输入(year_0 As Integer, year_shuru As Integer, month_shuru As Integer)
        'year_0：整个项目建设期开始的第一个年份
        'year_shuru：手动输入的建设期年份
        'month_shuru：手动输入的建设期开始投产的月份(建设期于这个月份之前的一个月结束)

        Dim year_index As Integer = year_shuru - year_0 + 1
        Dim month_sc As Integer = 12 - month_shuru + 1
        '结果列表
        Dim ans(1)
        '返回结果
        ans(0) = year_index
        ans(1) = month_sc
        '返回结果
        Return ans
    End Function
    Function 逐年投产月份数_10次投资(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '读取输入数据
        '项目计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '读取输入的建设期时间计划数据
        Dim jsqsjjhsj = 读取建设期时间计划数据(ExcelApp)
        '开始年份列表
        Dim year_start_list = jsqsjjhsj(0)
        '结束年份列表
        Dim year_end_list = jsqsjjhsj(1)
        '开始月份列表
        Dim month_start_list = jsqsjjhsj(2)
        '结束月份列表
        Dim month_end_list = jsqsjjhsj(3)
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------
        '计算自动计算的结果
        Dim ans_zd = 建设期时间计划计算_自动计算(year_start_list, year_end_list, month_start_list, month_end_list)
        Dim ans_year_list = ans_zd(0)
        Dim ans_month_list = ans_zd(1)
        Dim ans_month_start_list = ans_zd(2)
        Dim ans_month_end_list = ans_zd(3)
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------
        Dim js As Integer = 1
        '10次投资，计算每次投资的建设年份编号
        Dim year_10_list(10) As Integer
        '10次投资，每个建设期年份的投产月份数
        Dim tcyf_list(10) As Integer
        '10次投资，计算每次投资的逐年投产月份数
        Dim month_10_list(10)
        For i = 1 To 31
            If ans_year_list(i) > 0 Then
                Dim tmp_month(31) As Integer
                For j = 1 To 31
                    If j < ans_year_list(i) Then
                        tmp_month(j) = 0
                    ElseIf j = ans_year_list(i) Then
                        tmp_month(j) = ans_month_list(i)
                    ElseIf j > ans_year_list(i) And j <= jsnx Then
                        tmp_month(j) = 12
                    Else
                        tmp_month(j) = 0
                    End If
                Next
                month_10_list(js) = tmp_month
                year_10_list(js) = ans_year_list(i)
                tcyf_list(js) = ans_month_list(i)
                js = js + 1
            End If
        Next
        For i = js + 1 To 10
            Dim tmp_month(31) As Integer
            For j = 1 To 31
                tmp_month(j) = 0
            Next
            month_10_list(i) = tmp_month
            year_10_list(js) = 0
        Next
        '返回结果
        Dim ans(2)
        ans(0) = year_10_list
        ans(1) = month_10_list
        ans(2) = tcyf_list
        Return ans
    End Function
    Function 计算项目运营年限(jsnx As Integer, month_10_list As Array)
        'jsnx：计算年限
        'month_10_list：10次投资，计算每次投资的逐年投产月份数，列表，长度10
        '————————————————————————————————————————————————————————————————————————————————————————       
        Dim yynx As Integer
        If month_10_list(1)(1) > 0 Then
            yynx = jsnx
        Else
            yynx = jsnx - 1
        End If
        Return yynx
    End Function
End Module
