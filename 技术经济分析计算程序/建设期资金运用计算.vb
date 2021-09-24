Module 建设期资金运用计算
    Sub 建设期资金运用计算(ExcelApp As Object, zbj_model As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '要先进行<建设期时间计划计算>，再运行本程序
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置

        '从估算表读取计算所需数据
        Dim GSBSJ = 读取估算表数据(ExcelApp)
        '10次的静态投资，列表，长度10
        Dim jttz = GSBSJ(1)
        '10次投资年份序号列表，列表，长度10
        Dim tznf_list = GSBSJ(0)
        '将静态投资转换成31年的列表
        Dim jttz_list = 基础计算功能_10_to_31(tznf_list, jttz)
        '10次的贷款年利率，列表，长度10
        Dim dkll(10) As Double
        For i = 1 To 10
            dkll(i) = ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i + 11, 39).Value
        Next
        '将贷款利率转换成31年的列表
        Dim dkll_list = 基础计算功能_10_to_31(tznf_list, dkll)
        '10次投资的资本金比例，列表，长度10
        Dim zbjbl(10) As Double
        For i = 1 To 10
            zbjbl(i) = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 4, 22).Value
        Next
        '将资本金比例转换成31年的列表
        Dim zbjbl_list = 基础计算功能_10_to_31(tznf_list, zbjbl)
        '读取手动输入的建设期贷款利息金额，列表，长度10
        Dim dklx_shuru(10) As Double
        For i = 1 To 5
            dklx_shuru(i) = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(162 + i, 12).Value
        Next
        For i = 6 To 10
            dklx_shuru(i) = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(157 + i, 16).Value
        Next
        '转为长度31的列表
        Dim dklx_shuru_list = 基础计算功能_10_to_31(tznf_list, dklx_shuru)
        '读取资本金计算模式、建设期贷款利息计算模式
        Dim dklx_model As Integer
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 11).Value = "自动计算" Then
            dklx_model = 0
        Else
            dklx_model = 1
        End If
        '从<建设期时间计划计算>的计算结果中读取数据
        'jsq_index_list：是否处于建设期标记，1：处于建设期，0：不处于建设期，列表，长度31
        Dim jsq_index_list(31) As Integer
        For i = 1 To 31
            jsq_index_list(i) = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(162, i + 2).Value
        Next
        'year_jsq：逐年建设年份序号，列表，长度10
        'month_start_jsq：逐年建设开始月份序号，列表，长度10
        'month_end_jsq：逐年建设结束月份序号，列表，长度10
        Dim year_jsq(10) As Integer
        Dim month_start_jsq(10) As Integer
        Dim month_end_jsq(10) As Integer
        For i = 1 To 10
            year_jsq(i) = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(204 + i, 7).Value
            month_start_jsq(i) = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(204 + i, 15).Value
            month_end_jsq(i) = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(204 + i, 16).Value
        Next
        '将长度10的列表转为31长度
        'year_list：逐年建设年份序号，列表，长度31
        'month_start_list：逐年建设开始月份序号，列表，长度31
        'month_end_list：逐年建设结束月份序号，列表，长度31
        Dim year_list = 基础计算功能_10_to_31(tznf_list, year_jsq)
        Dim month_start_list = 基础计算功能_10_to_31(tznf_list, month_start_jsq)
        Dim month_end_list = 基础计算功能_10_to_31(tznf_list, month_end_jsq)
        '进行建设期资金运用计算
        Dim ans_jsqzjyy = 建设期资金运用计算_main(zbj_model, dklx_model, jttz_list, jsq_index_list, dkll_list, year_list,
                                                  month_start_list, month_end_list, zbjbl_list, dklx_shuru_list)

        '动态投资，列表，长度31
        Dim ans_dttz_list = ans_jsqzjyy(0)
        '资本金，列表，长度31
        Dim ans_zbj_list = ans_jsqzjyy(1)
        '建设期贷款金额，列表，长度31
        Dim ans_dkje_list = ans_jsqzjyy(2)
        '建设期贷款利息，列表，长度31
        Dim ans_dklx_list = ans_jsqzjyy(3)
        '将列表转为长度10
        Dim dttz_list = 基础计算功能_31_to_10(tznf_list, ans_dttz_list)
        Dim zbj_list = 基础计算功能_31_to_10(tznf_list, ans_zbj_list)
        Dim dkje_list = 基础计算功能_31_to_10(tznf_list, ans_dkje_list)
        Dim dklx_list = 基础计算功能_31_to_10(tznf_list, ans_dklx_list)
        '数据写入Excel，写入<估算表>
        '前5次
        For i = 1 To 5
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(31, 2 * i + 1).Value = dttz_list(i)
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(32, 2 * i + 1).Value = zbj_list(i)
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(33, 2 * i + 1).Value = dkje_list(i)
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(34, 2 * i + 1).Value = dklx_list(i)
        Next
        '6-10次
        For i = 6 To 10
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(74, 2 * i - 9).Value = dttz_list(i)
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(75, 2 * i - 9).Value = zbj_list(i)
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(76, 2 * i - 9).Value = dkje_list(i)
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(77, 2 * i - 9).Value = dklx_list(i)
        Next
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '计算一次Excel
        ExcelApp.Calculate()
    End Sub
    Function 建设期资金运用计算_main(zbj_model As Integer, dklx_model As Integer, jttz_list As Array,
                                     jsq_index_list As Array, dkll_list As Array, year_list As Array,
                                     month_start_list As Array, month_end_list As Array, zbjbl_list As Array,
                                     dklx_shuru_list As Array)
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础
        'dklx_model：建设期贷款利息计算模式：0：自动计算，1：手动输入
        'jttz_list：静态投资金额，列表，长度31
        'jsq_index_list：是否处于建设期标记，1：处于建设期，0：不处于建设期，列表，长度31
        'dkll_list：逐年贷款利率，列表，长度31
        'year_list：逐年建设年份序号，列表，长度31
        'month_start_list：逐年建设开始月份序号，列表，长度31
        'month_end_list：逐年建设结束月份序号，列表，长度31
        'zbjbl_list：逐年资本金比例，列表，长度31
        'dklx_shuru_list：手动输入的建设期贷款利息，列表，长度31

        '列表，用于储存计算结果，列表长度31
        Dim ans_dttz_list(31) As Double '动态投资，列表，长度31
        Dim ans_zbj_list(31) As Double '资本金，列表，长度31
        Dim ans_dkje_list(31) As Double '建设期贷款金额，列表，长度31
        Dim ans_dklx_list(31) As Double '建设期贷款利息，列表，长度31

        '最大迭代次数：200
        For i = 1 To 200
            '计算建设期贷款利息
            ans_dklx_list = 建设期贷款利息计算_main(jsq_index_list, ans_dkje_list, dkll_list, year_list,
                                                    month_start_list, month_end_list, dklx_model, dklx_shuru_list)
            '遍历31年列表
            For j = 1 To 31
                '计算资本金
                If zbj_model = 0 Then
                    ans_zbj_list(j) = ans_dttz_list(j) * zbjbl_list(j)
                Else
                    ans_zbj_list(j) = jttz_list(j) * zbjbl_list(j)
                End If
                '计算建设期贷款金额
                ans_dkje_list(j) = ans_dttz_list(j) - ans_zbj_list(j)
                '计算动态投资
                ans_dttz_list(j) = jttz_list(j) + ans_dklx_list(j)
            Next
        Next
        Dim ans(3)
        ans(0) = ans_dttz_list
        ans(1) = ans_zbj_list
        ans(2) = ans_dkje_list
        ans(3) = ans_dklx_list
        '返回结果
        Return ans
    End Function
    Function 建设期贷款利息计算_main(jsq_index_list As Array, dkje_list As Array, dkll_list As Array, year_list As Array,
                                     month_start_list As Array, month_end_list As Array, dklx_model As Integer,
                                     dklx_shuru_list As Array)
        'jsq_index_list：是否处于建设期标记，1：处于建设期，0：不处于建设期，列表，长度31
        'dkje_list：逐年贷款金额，列表，长度31
        'dkll_list：逐年贷款利率，列表，长度31
        'year_list：逐年建设年份序号，列表，长度31
        'month_start_list：逐年建设开始月份序号，列表，长度31
        'month_end_list：逐年建设结束月份序号，列表，长度31
        'dklx_model：建设期贷款利息计算模式：0：自动计算，1：手动输入
        'dklx_shuru_list：手动输入的建设期贷款利息，列表，长度31

        '列表，用于储存计算结果，列表长度31
        Dim ans_dklx_list(31) As Double
        '如果是自动计算模式
        If dklx_model = 0 Then
            '转入下一年计算的贷款本息之和，初始值是0，之后每一年计算过程中会修改
            Dim dkbx_zr As Double = 0
            Dim dkbx_zr_list(31) As Double
            '单独算一下第1年
            If jsq_index_list(1) = 1 Then
                ans_dklx_list(1) = 建设期贷款利息计算_开工年份(dkje_list(1), dkll_list(1), month_start_list(1), month_end_list(1))
                dkbx_zr += ans_dklx_list(1) + dkje_list(1)
                dkbx_zr_list(1) = dkbx_zr
            End If
            '循环到第2-30年，防止循环越界
            For i = 2 To 30
                '判断是否处于建设期
                If jsq_index_list(i) = 1 Then
                    '判断是属于哪一种建设期贷款利息计算方式
                    If year_list(i) = year_list(i - 1) + 1 And year_list(i) = year_list(i + 1) - 1 And month_end_list(i - 1) = 12 And
                       month_start_list(i) = 1 And month_end_list(i) = 12 And month_start_list(i + 1) = 1 Then
                        '属于建设年份
                        ans_dklx_list(i) = 建设期贷款利息计算_建设年份(dkje_list(i), dkbx_zr, dkll_list(i))
                    ElseIf year_list(i) <> year_list(i - 1) + 1 Or
                           (year_list(i) = year_list(i - 1) + 1 And (month_end_list(i - 1) <> 12 Or month_start_list(i) <> 1)) Then
                        '属于开工年份
                        ans_dklx_list(i) = 建设期贷款利息计算_开工年份(dkje_list(i), dkll_list(i), month_start_list(i), month_end_list(i))
                    Else
                        '属于投产年份
                        ans_dklx_list(i) = 建设期贷款利息计算_投产年份(dkje_list(i), dkbx_zr, dkll_list(i), month_end_list(i))
                    End If
                    '修改转入下一年的本息之和
                    dkbx_zr += ans_dklx_list(i) + dkje_list(i)
                Else
                    '如果不属于建设期，则将已经累计的转入下一年的本息之和清零
                    dkbx_zr = 0
                End If
                dkbx_zr_list(i) = dkbx_zr
            Next
            '单独计算第31年，第31年肯定属于投产年份
            If jsq_index_list(31) = 1 Then
                ans_dklx_list(31) = 建设期贷款利息计算_投产年份(dkje_list(31), dkbx_zr, dkll_list(31), month_end_list(31))
                dkbx_zr += ans_dklx_list(31) + dkje_list(31)
                dkbx_zr_list(31) = dkbx_zr
            End If
        Else
            '如果是手动输入模式
            ans_dklx_list = dklx_shuru_list
        End If
        '返回结果
        Return ans_dklx_list
    End Function

    Function 建设期贷款利息计算_开工年份(dkje As Double, dkll As Double, month_start As Integer, month_end As Integer)
        'dkje：建设期贷款金额
        'dkll：建设期贷款利率，复利
        'month_start：建设期开始月份序号，建设期包含开始当月
        'month_end ：建设期结束月份序号，建设期包含结束当月

        '计算建设期贷款利息
        Dim ans_dklx = (dkje / 2) * dkll * (12 - month_start + 1) / 12 - (dkje / 2) * dkll * (12 - month_end) / 12
        ' 返回结果
        Return ans_dklx
    End Function
    Function 建设期贷款利息计算_建设年份(dkje As Double, dkbx_zr As Double, dkll As Double)
        'dkje：建设期贷款金额
        'dkbx_zr：上一年转入的贷款本息总和
        'dkll：建设期贷款利率，复利

        '计算建设期贷款利息
        Dim ans_dklx = (dkbx_zr + dkje / 2) * dkll
        ' 返回结果
        Return ans_dklx
    End Function
    Function 建设期贷款利息计算_投产年份(dkje As Double, dkbx_zr As Double, dkll As Double, month_end As Integer)
        'dkje：建设期贷款金额
        'dkbx_zr：上一年转入的贷款本息总和
        'dkll：建设期贷款利率，复利
        'month_end ：建设期结束月份序号，建设期包含结束当月

        '计算建设期贷款利息
        Dim ans_dklx = (dkbx_zr + dkje / 2) * dkll * (month_end / 12)
        ' 返回结果
        Return ans_dklx
    End Function
End Module
