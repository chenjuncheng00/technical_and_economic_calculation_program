Module 逐年达产率计算
    Sub 逐年达产率计算_main(ExcelApp As Object)
        '读取输入数据
        Dim GSBSJ = 读取估算表数据(ExcelApp)
        '投资年份
        Dim tznf_list = GSBSJ(0)
        '项目计算年限
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '获取10次投资,每一次投资的投产月份数
        Dim ans_month = 逐年投产月份数_10次投资(ExcelApp)
        Dim month_10_list = ans_month(1)
        Dim tcyf_list = ans_month(2)
        '逐年达产率计算
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 19).Value = "直接输入" Then
            Call 直接输入综合达产率(ExcelApp, jsnx, month_10_list)
        Else
            Call 分投资逐次输入达产率(ExcelApp, jsnx, month_10_list, tznf_list, tcyf_list)
        End If
    End Sub

    Sub 直接输入综合达产率(ExcelApp As Object, jsnx As Integer, month_10_list As Array)
        'jsnx：计算年限
        'month_10_list：10次投资，计算每次投资的逐年投产月份数，列表，长度10
        '————————————————————————————————————————————————————————————————————————————————————————        
        Call 逐年综合达产率_写入EXCEL(ExcelApp, jsnx, month_10_list)
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '此处不要计算：接入费逐年达产率、补贴收入逐年达产率、固定收入成本逐年达产率（购电容量费成本、城市管廊成本、人员工资、充电桩收入）
        '此处不要计算：光伏逐年达产率、风电逐年达产率、蓄电池逐年达产率
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '计算一次Excel
        ExcelApp.Calculate()
    End Sub
    Sub 分投资逐次输入达产率(ExcelApp As Object, jsnx As Integer, month_10_list As Array, tznf_list As Array, tcyf_list As Array)
        'jsnx：计算年限
        'month_10_list：10次投资，计算每次投资的逐年投产月份数，列表，长度10
        'tznf_list：10次投资的年份序号，列表，长度10
        'tcyf_list：10次投资的投产月份数，列表，长度10
        'zs_set：数值是否需要按照当年的（投产月份数/12）进行折算
        '————————————————————————————————————————————————————————————————————————————————————————        
        '10次投资综合达产率计算
        Call 十次投资综合达产率计算(ExcelApp, jsnx, month_10_list, tznf_list, tcyf_list, True)
        '————————————————————————————————————————————————————————————————————————————————————————        
        Call 逐年综合达产率_写入EXCEL(ExcelApp, jsnx, month_10_list)
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '此处不要计算：接入费逐年达产率、补贴收入逐年达产率、固定收入成本逐年达产率（购电容量费成本、城市管廊成本、人员工资、充电桩收入）
        '此处不要计算：光伏逐年达产率、风电逐年达产率、蓄电池逐年达产率
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '计算一次Excel
        ExcelApp.Calculate()
    End Sub
    Sub 逐年综合达产率_写入EXCEL(ExcelApp As Object, jsnx As Integer, month_10_list As Array)
        'jsnx：计算年限
        'month_10_list：10次投资，计算每次投资的逐年投产月份数，列表，长度10
        '————————————————————————————————————————————————————————————————————————————————————————       
        '“收入&成本输入”103行和106行的数据是用来计算收入成本的最终数据
        For i = 3 To 17
            '综合负荷率，前15年
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(102, i).Value <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(5, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value = 0
            End If
        Next
        For i = 2 To 17
            '综合负荷率，后16年        
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(105, i).Value <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(8, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i).Value = 0
            End If
        Next
        '检测建设期，将没有投产月份数量的年份负荷率设置为0
        For i = 3 To 17 '投资计划与资金筹措表列号
            '前15年
            If month_10_list(1)(i - 2) = 0 Then '如果当年投产的月份数=0                
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value = 0 '负荷率设置为0
            End If
        Next
        For i = 2 To 17 '投资计划与资金筹措表列号           
            '后16年
            If month_10_list(1)(i + 14) = 0 Then '如果当年投产的月份数=0                
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i).Value = 0 '负荷率设置为0
            End If
        Next
    End Sub
    Sub 接入费逐年达产率计算(ExcelApp As Object, jsnx As Integer, month_10_list As Array, qtnf_list As List(Of Integer))
        'jsnx：计算年限
        'month_10_list：10次投资，计算每次投资的逐年投产月份数，列表，长度10
        'qtnf_list：其它年份列表，根本列表中的年份序号，将对应年份的负荷率设置为0
        '————————————————————————————————————————————————————————————————————————————————————————       
        '读取基础负荷率
        Dim fhl_base = 读取逐年负荷率(ExcelApp)
        '接入费计算模式
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(41, 18).Value = "逐年达产率增加值" Then
            For i = 1 To 31
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i + 2).Value = fhl_base(i) - fhl_base(i - 1)
            Next
            '将小于0的结果设置为0
            For i = 1 To 31
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i + 2).Value < 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i + 2).Value = 0
                End If
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(41, 18).Value = "保持每年100%" Then
            For i = 1 To 31
                If i <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i + 2).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i + 2).Value = 0
                End If
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(41, 18).Value = "逐年达产率" Then
            For i = 1 To 31
                If i <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i + 2).Value = fhl_base(i)
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i + 2).Value = 0
                End If
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(41, 18).Value = "逐年投产月份比例" Then
            For i = 1 To 31
                If i <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i + 2).Value = 1 * month_10_list(1)(i) / 12
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i + 2).Value = 0
                End If
            Next
        End If
        For Each qtnf In qtnf_list
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, qtnf + 2).Value = 0
        Next
    End Sub
    Sub 补贴收入逐年达产率计算_main(ExcelApp As Object, jsnx As Integer, month_10_list As Array)
        'jsnx：计算年限
        'month_10_list：10次投资，计算每次投资的逐年投产月份数，列表，长度10
        '————————————————————————————————————————————————————————————————————————————————————————       
        '生成一个空的List
        Dim qtnf_list As New List(Of Integer)
        '读取基础负荷率
        Dim fhl_base = 读取逐年负荷率(ExcelApp)
        '补贴收入1
        Call 补贴收入1逐年达产率计算(ExcelApp, jsnx, month_10_list, fhl_base, qtnf_list)
        '补贴收入2
        Call 补贴收入2逐年达产率计算(ExcelApp, jsnx, month_10_list, fhl_base, qtnf_list)
        '光伏补贴收入
        For i = 1 To 31
            If i >= 1 And i <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(165, i + 2).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(165, i + 2).Value = 0
            End If
        Next
    End Sub
    Sub 补贴收入1逐年达产率计算(ExcelApp As Object, jsnx As Integer, month_10_list As Array, fhl_base As Array, qtnf_list As List(Of Integer))
        'jsnx：计算年限
        'month_10_list：10次投资，计算每次投资的逐年投产月份数，列表，长度10
        'fhl_base：逐年负荷率，列表，长度31
        'jsms：计算模式（0：保持每年100%，1：逐年投产月份比例，2：逐年达产率，3：逐年达产率增加值）
        'qtnf_list：其它年份列表，根本列表中的年份序号，将对应年份的负荷率设置为0
        '————————————————————————————————————————————————————————————————————————————————————————       
        For i = 1 To 31
            If i >= 1 And i <= jsnx Then
                '补贴收入1
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 18).Value = "保持每年100%" Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i + 2).Value = 1
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 18).Value = "逐年投产月份比例" Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i + 2).Value = 1 * month_10_list(1)(i) / 12
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 18).Value = "逐年达产率" Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i + 2).Value = fhl_base(i)
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 18).Value = "逐年达产率增加值" Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i + 2).Value = fhl_base(i) - fhl_base(i - 1)
                End If
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i + 2).Value = 0
            End If
        Next
        For Each qtnf In qtnf_list
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, qtnf + 2).Value = 0
        Next
    End Sub
    Sub 补贴收入2逐年达产率计算(ExcelApp As Object, jsnx As Integer, month_10_list As Array, fhl_base As Array, qtnf_list As List(Of Integer))
        'jsnx：计算年限
        'month_10_list：10次投资，计算每次投资的逐年投产月份数，列表，长度10
        'fhl_base：逐年负荷率，列表，长度31
        'jsms：计算模式（0：保持每年100%，1：逐年投产月份比例，2：逐年达产率，3：逐年达产率增加值）
        'qtnf_list：其它年份列表，根本列表中的年份序号，将对应年份的负荷率设置为0
        '————————————————————————————————————————————————————————————————————————————————————————       
        For i = 1 To 31
            If i >= 1 And i <= jsnx Then
                '补贴收入2
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(43, 18).Value = "保持每年100%" Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i + 2).Value = 1
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(43, 18).Value = "逐年投产月份比例" Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i + 2).Value = 1 * month_10_list(1)(i) / 12
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(43, 18).Value = "逐年达产率" Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i + 2).Value = fhl_base(i)
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(43, 18).Value = "逐年达产率增加值" Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i + 2).Value = fhl_base(i) - fhl_base(i - 1)
                End If
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i + 2).Value = 0
            End If
        Next
        For Each qtnf In qtnf_list
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, qtnf + 2).Value = 0
        Next
    End Sub
    Sub 固定收入成本逐年达产率计算(ExcelApp As Object, jsnx As Integer, month_10_list As Array)
        'jsnx：计算年限
        'month_10_list：10次投资，计算每次投资的逐年投产月份数，列表，长度10
        '————————————————————————————————————————————————————————————————————————————————————————       
        '生成一个空的List
        Dim qtnf_list As New List(Of Integer)
        '读取基础负荷率
        Dim fhl_base = 读取逐年负荷率(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————       
        '————————————————————————————————————————————————————————————————————————————————————————       
        '购电容量费成本
        Call 购电容量费成本逐年达产率计算(ExcelApp, jsnx, month_10_list, qtnf_list)
        '—————————————————————————————————————————————————————————————————
        '城市管廊成本
        Call 城市管廊成本逐年达产率计算(ExcelApp, jsnx, month_10_list, qtnf_list)
        '—————————————————————————————————————————————————————————————————
        '人员工资
        Call 人员工资成本逐年达产率计算(ExcelApp, jsnx, month_10_list, 0, qtnf_list)
        '—————————————————————————————————————————————————————————————————
        '充电桩收入
        Call 充电桩收入逐年达产率计算(ExcelApp, jsnx, month_10_list, fhl_base, qtnf_list)
    End Sub
    Sub 购电容量费成本逐年达产率计算(ExcelApp As Object, jsnx As Integer, month_10_list As Array, qtnf_list As List(Of Integer))
        'jsnx：计算年限
        'month_10_list：10次投资，计算每次投资的逐年投产月份数，列表，长度10
        'fhl_base：逐年负荷率，列表，长度31
        'jsms：计算模式（0：逐年投产月份比例，1：保持每年100%）
        'qtnf_list：其它年份列表，根本列表中的年份序号，将对应年份的负荷率设置为0
        '————————————————————————————————————————————————————————————————————————————————————————       
        '购电容量费成本
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(49, 18).Value = "逐年投产月份比例" Then
            For i = 3 To 33
                If i - 2 <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(145, i).Value = 1 * month_10_list(1)(i - 2) / 12
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(145, i).Value = 0
                End If
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(49, 18).Value = "保持每年100%" Then
            For i = 3 To 33
                If i - 2 <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(145, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(145, i).Value = 0
                End If
            Next
        End If
        For Each qtnf In qtnf_list
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(145, qtnf + 2).Value = 0
        Next
    End Sub
    Sub 城市管廊成本逐年达产率计算(ExcelApp As Object, jsnx As Integer, month_10_list As Array, qtnf_list As List(Of Integer))
        'jsnx：计算年限
        'month_10_list：10次投资，计算每次投资的逐年投产月份数，列表，长度10
        'fhl_base：逐年负荷率，列表，长度31
        'jsms：计算模式（0：逐年投产月份比例，1：保持每年100%）
        'qtnf_list：其它年份列表，根本列表中的年份序号，将对应年份的负荷率设置为0
        '————————————————————————————————————————————————————————————————————————————————————————       
        '城市管廊成本
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(50, 18).Value = "逐年投产月份比例" Then
            For i = 3 To 33
                If i - 2 <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(146, i).Value = 1 * month_10_list(1)(i - 2) / 12
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(146, i).Value = 0
                End If
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(50, 18).Value = "保持每年100%" Then
            For i = 3 To 33
                If i - 2 <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(146, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(146, i).Value = 0
                End If
            Next
        End If
        For Each qtnf In qtnf_list
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(146, qtnf + 2).Value = 0
        Next
    End Sub
    Sub 人员工资成本逐年达产率计算(ExcelApp As Object, jsnx As Integer, month_10_list As Array, ZNDZBL As Double, qtnf_list As List(Of Integer))
        'jsnx：计算年限
        'month_10_list：10次投资，计算每次投资的逐年投产月份数，列表，长度10
        'fhl_base：逐年负荷率，列表，长度31
        'jsms：计算模式（0：逐年投产月份比例，1：保持每年100%，2：逐年递增）
        'ZNDZBL：人员工资成本逐年递增比例
        'qtnf_list：其它年份列表，根本列表中的年份序号，将对应年份的负荷率设置为0
        '————————————————————————————————————————————————————————————————————————————————————————       
        '人员工资
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(51, 18).Value = "逐年投产月份比例" Then
            For i = 3 To 33
                If i - 2 <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(147, i).Value = 1 * month_10_list(1)(i - 2) / 12
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(147, i).Value = 0
                End If
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(51, 18).Value = "保持每年100%" Then
            For i = 3 To 33
                If i - 2 <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(147, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(147, i).Value = 0
                End If
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(51, 18).Value = "逐年递增" Then
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(147, i).Value = 1 * (1 + ZNDZBL * (i - 3))
            Next
        End If
        For Each qtnf In qtnf_list
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(147, qtnf + 2).Value = 0
        Next
    End Sub
    Sub 充电桩收入逐年达产率计算(ExcelApp As Object, jsnx As Integer, month_10_list As Array, fhl_base As Array, qtnf_list As List(Of Integer))
        'jsnx：计算年限
        'month_10_list：10次投资，计算每次投资的逐年投产月份数，列表，长度10
        'fhl_base：逐年负荷率，列表，长度31
        'jsms：计算模式（0：逐年投产月份比例，1：保持每年100%，2：逐年达产率）
        'qtnf_list：其它年份列表，根本列表中的年份序号，将对应年份的负荷率设置为0
        '————————————————————————————————————————————————————————————————————————————————————————       
        '充电桩收入
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(52, 18).Value = "逐年投产月份比例" Then
            For i = 3 To 33
                If i - 2 <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 1 * month_10_list(1)(i - 2) / 12
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 0
                End If
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(52, 18).Value = "保持每年100%" Then
            For i = 3 To 33
                If i - 2 <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 0
                End If
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(52, 18).Value = "逐年达产率" Then
            For i = 3 To 33
                If i - 2 <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = fhl_base(i - 2)
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 0
                End If
            Next
        End If
        For Each qtnf In qtnf_list
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, qtnf + 2).Value = 0
        Next
    End Sub
    Sub 光伏逐年综合达产率计算(ExcelApp As Object, fhl_base As Array, zs_set As Boolean)
        'fhl_base：逐年负荷率，列表，长度31
        'zs_set：数值是否需要按照当年的（投产月份数/12）进行折算
        '————————————————————————————————————————————————————————————————————————————————————————       
        '项目计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '读取输入数据
        Dim GSBSJ = 读取估算表数据(ExcelApp)
        '投资年份
        Dim tznf_list = GSBSJ(0)
        '10次投资的装机功率
        Dim gfzjgl = GSBSJ(17)
        '获取10次投资,每一次投资的投产月份数
        Dim ans_month = 逐年投产月份数_10次投资(ExcelApp)
        Dim month_10_list = ans_month(1)
        Dim tcyf_list = ans_month(2)
        '计算项目运营年限
        Dim yynx As Integer = 计算项目运营年限(jsnx, month_10_list)
        '————————————————————————————————————————————————————————————————————————————————————————       
        '计算逐年综合负荷率
        Dim fhl_zh_gf = 逐年综合达产率计算基础功能(tznf_list, tcyf_list, gfzjgl, yynx, zs_set, fhl_base)
        '————————————————————————————————————————————————————————————————————————————————————————       
        '结果写入Excel
        For i = 1 To 31
            If i <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(100, 2 + i).Value = fhl_zh_gf(i)
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(100, 2 + i).Value = 0
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
    Sub 风电逐年综合达产率计算(ExcelApp As Object, fhl_base As Array, zs_set As Boolean)
        'fhl_base：逐年负荷率，列表，长度31
        'zs_set：数值是否需要按照当年的（投产月份数/12）进行折算
        '————————————————————————————————————————————————————————————————————————————————————————        
        '项目计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '读取输入数据
        Dim GSBSJ = 读取估算表数据(ExcelApp)
        '投资年份
        Dim tznf_list = GSBSJ(0)
        '10次投资的装机功率
        Dim fdzjgl = GSBSJ(19)
        '获取10次投资,每一次投资的投产月份数
        Dim ans_month = 逐年投产月份数_10次投资(ExcelApp)
        Dim month_10_list = ans_month(1)
        Dim tcyf_list = ans_month(2)
        '计算项目运营年限
        Dim yynx As Integer = 计算项目运营年限(jsnx, month_10_list)
        '————————————————————————————————————————————————————————————————————————————————————————       
        '计算逐年综合负荷率
        Dim fhl_zh_fd = 逐年综合达产率计算基础功能(tznf_list, tcyf_list, fdzjgl, yynx, zs_set, fhl_base)
        '————————————————————————————————————————————————————————————————————————————————————————       
        '结果写入Excel
        For i = 1 To 31
            If i <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(178, 2 + i).Value = fhl_zh_fd(i)
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(178, 2 + i).Value = 0
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
    Sub 蓄电池逐年综合达产率计算(ExcelApp As Object, fhl_base As Array, jsms As String, zs_set As Boolean)
        'fhl_base：逐年负荷率，列表，长度31
        'jsms：计算模式，“供电”、“购电”
        'zs_set：数值是否需要按照当年的（投产月份数/12）进行折算
        '————————————————————————————————————————————————————————————————————————————————————————        
        '项目计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '读取输入数据
        Dim GSBSJ = 读取估算表数据(ExcelApp)
        '投资年份
        Dim tznf_list = GSBSJ(0)
        '10次投资的装机功率
        Dim xdczjgl = GSBSJ(13)
        '获取10次投资,每一次投资的投产月份数
        Dim ans_month = 逐年投产月份数_10次投资(ExcelApp)
        Dim month_10_list = ans_month(1)
        Dim tcyf_list = ans_month(2)
        '计算项目运营年限
        Dim yynx As Integer = 计算项目运营年限(jsnx, month_10_list)
        '————————————————————————————————————————————————————————————————————————————————————————        
        '计算逐年综合负荷率
        Dim fhl_zh_xdc = 逐年综合达产率计算基础功能(tznf_list, tcyf_list, xdczjgl, yynx, zs_set, fhl_base)
        '————————————————————————————————————————————————————————————————————————————————————————        
        '结果写入Excel
        For i = 1 To 31
            If i <= jsnx Then
                If jsms = "供电" Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(96, 2 + i).Value = fhl_zh_xdc(i)
                End If
                If jsms = "购电" Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(97, 2 + i).Value = fhl_zh_xdc(i)
                End If
                If jsms = "供电和购电" Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(96, 2 + i).Value = fhl_zh_xdc(i)
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(97, 2 + i).Value = fhl_zh_xdc(i)
                End If
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(96, 2 + i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(97, 2 + i).Value = 0
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
    Sub 十次投资综合达产率计算(ExcelApp As Object, jsnx As Integer, month_10_list As Array, tznf_list As Array, tcyf_list As Array, zs_set As Boolean)
        'jsnx：计算年限
        'month_10_list：10次投资，计算每次投资的逐年投产月份数，列表，长度10
        'tznf_list：10次投资的年份序号，列表，长度10
        'tcyf_list：10次投资的投产月份数，列表，长度10
        'zs_set：数值是否需要按照当年的（投产月份数/12）进行折算
        '————————————————————————————————————————————————————————————————————————————————————————        
        '计算项目运营年限
        Dim yynx As Integer = 计算项目运营年限(jsnx, month_10_list)
        '读取1-10次的燃料费额
        Dim rlf(10) As Double
        For i = 1 To 10
            rlf(i) = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(119, 5 + i).Value
        Next
        '读取基础负荷率
        Dim fhl_base = 读取逐年负荷率(ExcelApp)
        '计算逐年综合负荷率
        Dim ans_fhl_zh = 逐年综合达产率计算基础功能(tznf_list, tcyf_list, rlf, yynx, zs_set, fhl_base)
        '结果写入Excel
        '前15年
        For i = 1 To 15
            If i <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(123, 2 + i).Value = ans_fhl_zh(i)
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(123, 2 + i).Value = 0
            End If
        Next
        '16—31年
        For i = 16 To 31
            If i <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(125, i - 14).Value = ans_fhl_zh(i)
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(125, i - 14).Value = 0
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
    Function 逐年综合达产率计算基础功能(tznf_list As Array, tcyf_list As Array, value_10_list As Array, yynx As Integer, zs_set As Boolean, fhl_base As Array)
        'tznf_list：10次投资的年份序号，列表，长度10
        'tcyf_list：10次投资的投产月份数，列表，长度10
        'value_10_list：10次投资的各种金额数值，列表，长度10
        'yynx：运营年限
        'zs_set：数值是否需要按照当年的（投产月份数/12）进行折算
        'fhl_base：负荷率(达产率)计算基础值，列表，长度31
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算Value_10_list的和
        Dim value_sum As Double = 0
        For i = 1 To 10
            value_sum += value_10_list(i)
        Next
        Dim ans_value = 计算逐年总金额_10次投资(tznf_list, tcyf_list, value_10_list, yynx, 1, zs_set, fhl_base, 1)
        '逐年综合负荷率计算结果
        Dim ans_fhl(31) As Double
        For i = 1 To 31
            If value_sum > 0 Then
                ans_fhl(i) = ans_value(i) / value_sum
            Else
                ans_fhl(i) = 0
            End If
        Next
        '返回结果
        Return ans_fhl
    End Function

    Function 默认逐年达产率()
        '风电逐年达产率
        Dim fhl_fd(31) As Double
        '光伏逐年达产率
        Dim fhl_gf(31) As Double
        '蓄电池逐年衰减率
        Dim fhl_xdc(31) As Double
        '风电
        For i = 1 To 20
            fhl_fd(i) = 1
        Next
        '光伏
        For i = 1 To 25
            fhl_gf(i) = 0.98 - 0.0055 * i
        Next
        '蓄电池
        For i = 1 To 10
            fhl_xdc(i) = 1.02 - 0.01 * i
        Next
        '返回结果
        Dim ans(2)
        ans(0) = fhl_fd
        ans(1) = fhl_gf
        ans(2) = fhl_xdc
        Return ans
    End Function
End Module
