Module 逐个建设年份投资数据计算
    Function 计算基数累加修正计算_10次投资(tznf_list As Array, tcyf_list As Array, value_10_list As Array, yynx As Integer, kcbl As Double, zs_set As Boolean)
        '本方法用于修理费、材料费其它费的特殊计算功能
        'tznf_list：10次投资的年份序号，列表，长度10
        'tcyf_list：10次投资的投产月份数，列表，长度10
        'value_10_list：10次投资的各种金额数值(静态投资，建设期利息，可抵扣增值税，资本金，建设期贷款)，列表，长度10
        'yynx：运营年限
        'kcbl：超过运营年限后的扣除比例
        'zs_set：数值是否需要按照当年的（投产月份数/12）进行折算
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '计算每一年的累计数值，列表，长度31
        Dim value_lj_list = 数据累加合并_10次投资(tznf_list, tcyf_list, value_10_list, zs_set)
        '根据十次投资的时间、运营年限和扣除比例，计算逐年扣除的金额
        Dim value_kc_list(31) As Double '逐年扣除的金额
        For i = 1 To 10 '10次投资
            If tznf_list(i) > 0 Then
                For j = 1 To 31 '31年
                    If j > tznf_list(i) + yynx And tcyf_list(i) = 0 Then
                        '建设当年没有投产月份，则从后一年开始计算
                        value_kc_list(j) += value_10_list(i) * kcbl
                    ElseIf j >= yynx + tznf_list(i) And tcyf_list(i) > 0 Then
                        '建设当年有投产月份，则从当年开始计算
                        value_kc_list(j) += value_10_list(i) * kcbl
                    Else
                        value_kc_list(j) += 0
                    End If
                Next
            End If
        Next
        '扣除后剩余的金额，应该实际上用于最终计算的数值
        Dim value_result_list(31) As Double
        For i = 1 To 31
            value_result_list(i) = value_lj_list(i) - value_kc_list(i)
        Next
        '返回结果
        Dim ans(2)
        ans(0) = value_result_list
        ans(1) = value_lj_list
        ans(2) = value_kc_list
        Return ans
    End Function
    Function 数据累加合并_10次投资(tznf_list As Array, tcyf_list As Array, value_10_list As Array, zs_set As Boolean)
        'tznf_list：10次投资的年份序号，列表，长度10
        'tcyf_list：10次投资的投产月份数，列表，长度10
        'value_10_list：10次投资的各种数值，列表，长度10
        'zs_set：数值是否需要按照当年的（投产月份数/12）进行折算
        '———————————————————————————————————————————————————————————————————————————————————————— 
        'tznf_list转换数据
        Dim new_tznf_list As Integer() = Array数据转换为Integer(tznf_list, 10)
        '获取长度31的数值列表
        Dim value_31_list = 基础计算功能_10_to_31(tznf_list, value_10_list)
        '31长度的列表累加值
        Dim value_lj_list(31) As Double
        '当年的数值累计值等于，当年的数值和之前数值的总和（当年的数值需要考虑投产月份数的比例）
        For i = 1 To 31
            '投产年份之前的数值直接累加
            For j = 0 To i - 1
                value_lj_list(i) += value_31_list(j)
            Next
            '投产年份当年的数值需要判断是否按照（投产月份数/12）进行折算
            Dim tmp_value_lj As Double = 0
            If new_tznf_list.Contains(i) Then
                Dim index_nf As Integer = Array.IndexOf(tznf_list, i)
                '如果是建设年份
                If zs_set = False And tcyf_list(index_nf) > 0 Then
                    '如果不按照月份数比例折算，并且这一年有投产月份数，则直接全部累加
                    tmp_value_lj = value_31_list(i)
                Else
                    tmp_value_lj = value_31_list(i) * (tcyf_list(index_nf) / 12)
                End If
            Else
                '如果不是建设期，这一年的数据等于0，不需要处理
            End If
            value_lj_list(i) += tmp_value_lj
        Next
        '返回结果
        Return value_lj_list
    End Function
    Function 计算逐年总金额_10次投资(tznf_list As Array, tcyf_list As Array, value_10_list As Array, yynx As Integer, kcbl As Double, zs_set As Boolean, fl_list As Array, dividend As Double)
        'tznf_list：10次投资的年份序号，列表，长度10
        'tcyf_list：10次投资的投产月份数，列表，长度10
        'value_10_list：10次投资的各种金额数值(静态投资，建设期利息，可抵扣增值税，资本金，建设期贷款)，列表，长度10
        'yynx：运营年限
        'kcbl：超过运营年限后的扣除比例
        'zs_set：数值是否需要按照当年的（投产月份数/12）进行折算
        'fl_list：逐年费率，列表，长度31
        'dividend：被除数，逐年费率乘以逐年基数后，除以被除数得到最终金额
        '———————————————————————————————————————————————————————————————————————————————————————— 
        Dim ans_value_list = 生成逐年计算基数_10次投资(tznf_list, tcyf_list, value_10_list, yynx, kcbl, zs_set)
        Dim ans_fl_list = 生成逐年费率_10次投资(tznf_list, fl_list)
        '使用10次投资的逐年计算基数和费率，计算逐年金额，再将10次的逐年金额累加得到最终金额
        Dim result_list(31) As Double
        For i = 1 To 31
            Dim value_sum As Double = 0
            For j = 1 To 10
                Dim tmp_value As Double = ans_value_list(j)(i) * ans_fl_list(j)(i) / dividend
                value_sum += tmp_value
            Next
            result_list(i) = value_sum
        Next
        Return result_list
    End Function
    Function 生成逐年计算基数_10次投资(tznf_list As Array, tcyf_list As Array, value_10_list As Array, yynx As Integer, kcbl As Double, zs_set As Boolean)
        'tznf_list：10次投资的年份序号，列表，长度10
        'tcyf_list：10次投资的投产月份数，列表，长度10
        'value_10_list：10次投资的各种金额数值(静态投资，建设期利息，可抵扣增值税，资本金，建设期贷款)，列表，长度10
        'yynx：运营年限
        'kcbl：超过运营年限后的扣除比例
        'zs_set：数值是否需要按照当年的（投产月份数/12）进行折算
        '———————————————————————————————————————————————————————————————————————————————————————— 
        Dim ans_value_list(10)
        For i = 1 To 10
            Dim tmp_list(31) As Double
            If tznf_list(i) > 0 Then
                For j = 1 To 31
                    If j < tznf_list(i) Then
                        tmp_list(j) = 0
                    ElseIf j = tznf_list(i) Then
                        If zs_set = False And tcyf_list(i) > 0 Then
                            '如果不按照月份数比例折算，并且这一年有投产月份数，则直接等于
                            tmp_list(j) = value_10_list(i)
                        Else
                            tmp_list(j) = value_10_list(i) * (tcyf_list(i) / 12)
                        End If
                    ElseIf j > yynx + tznf_list(i) And tcyf_list(i) = 0 Then
                        '建设当年没有投产月份，则从后一年开始计算
                        tmp_list(j) = value_10_list(i) * (1 - kcbl)
                    ElseIf j >= yynx + tznf_list(i) And tcyf_list(i) > 0 Then
                        '建设当年有投产月份，则从当年开始计算
                        tmp_list(j) = value_10_list(i) * (1 - kcbl)
                    Else
                        tmp_list(j) = value_10_list(i)
                    End If
                Next
            Else
                For j = 1 To 31
                    tmp_list(j) = 0
                Next
            End If
            ans_value_list(i) = tmp_list
        Next
        Return ans_value_list
    End Function
    Function 生成逐年费率_10次投资(tznf_list As Array, fl_list As Array)
        'tznf_list：10次投资的年份序号，列表，长度10
        'fl_list：逐年费率，列表，长度31
        '———————————————————————————————————————————————————————————————————————————————————————— 
        Dim ans_fl_list(10)
        For i = 1 To 10
            Dim tmp_list(31) As Double
            If tznf_list(i) > 0 Then
                For j = 1 To 31
                    If j < tznf_list(i) Then
                        tmp_list(j) = 0
                    Else
                        tmp_list(j) = fl_list(j - (tznf_list(i) - tznf_list(1)))
                    End If
                Next
            Else
                For j = 1 To 31
                    tmp_list(j) = 0
                Next
            End If
            ans_fl_list(i) = tmp_list
        Next
        Return ans_fl_list
    End Function
End Module
