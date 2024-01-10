Module 增值税和所得税相关计算
    Sub 增值税相关计算(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '逐年实际增值税额及可抵扣增值税计算
        '————————————————————————————————————————————————————————————————————————————————————————        
        '项目计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '建设期可抵扣增值税的最大抵扣年份数
        Dim kdk_nfs As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(9, 7).Value
        '判断是否进行可抵扣增值税计算
        Dim dkjs As Boolean
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(12, 11).Value = "计算" Then
            dkjs = True
        Else
            dkjs = False
        End If
        '从估算表读取计算所需数据
        Dim GSBSJ = 读取估算表数据(ExcelApp)
        '从<估算表>读取逐年可抵扣增值税，列表，长度10
        Dim kdk = GSBSJ(9)
        '10次投资年份序号列表，列表，长度10
        Dim tznf_list = GSBSJ(0)
        '转换成长度31的列表，列表，长度31
        Dim kdk_list = 基础计算功能_10_to_31(tznf_list, kdk)
        '读取逐年销项增值税金额，进项增值税金额，列表，长度31
        Dim zzs_xx_list(31) As Double
        Dim zzs_jx_list(31) As Double
        '前15年
        For i = 1 To 15
            zzs_xx_list(i) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(21, i + 4).Value
            zzs_jx_list(i) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(22, i + 4).Value
        Next
        '16-31年
        For i = 16 To 31
            zzs_xx_list(i) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(44, i - 12).Value
            zzs_jx_list(i) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(45, i - 12).Value
        Next
        '计算可抵扣增值税
        Dim ans_kdk = 可抵扣增值税计算_main(kdk_list, zzs_xx_list, zzs_jx_list, kdk_nfs, jsnx, dkjs)
        Dim ans_ydk_list = ans_kdk(0) '已经被抵扣掉的增值税金额，写入Excel时需要取负值
        Dim ans_zzs_ydk_list = ans_kdk(1) '扣除已经抵扣掉的增值税金额后，剩余的增值税金额
        Dim ans_wdk_list = ans_kdk(2) '未抵扣(逐年未被抵扣掉的建设期可抵扣增值税)
        '结果写入Excel
        '前15年
        For i = 1 To 15
            If i <= jsnx Then
                '写入已经被抵扣的增值税金额
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(9, i + 4).Value = -ans_ydk_list(i)
                '写入剩余的增值税金额
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(20, i + 4).Value = ans_zzs_ydk_list(i)
                '写入逐年未被抵扣掉的建设期可抵扣增值税
                ExcelApp.ThisWorkbook.Worksheets("资产负债表").Cells(12, i + 4).Value = ans_wdk_list(i)
            Else
                '写入已经被抵扣的增值税金额
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(9, i + 4).Value = 0
                '写入剩余的增值税金额
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(20, i + 4).Value = 0
                '写入逐年未被抵扣掉的建设期可抵扣增值税
                ExcelApp.ThisWorkbook.Worksheets("资产负债表").Cells(12, i + 4).Value = 0
            End If
        Next
        '16-31年
        For i = 16 To 31
            If i <= jsnx Then
                '写入已经被抵扣的增值税金额
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(30, i - 12).Value = -ans_ydk_list(i)
                '写入剩余的增值税金额
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(43, i - 12).Value = ans_zzs_ydk_list(i)
                '写入逐年未被抵扣掉的建设期可抵扣增值税
                ExcelApp.ThisWorkbook.Worksheets("资产负债表").Cells(41, i - 12).Value = ans_wdk_list(i)
            Else
                '写入已经被抵扣的增值税金额
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(30, i - 12).Value = 0
                '写入剩余的增值税金额
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(43, i - 12).Value = 0
                '写入逐年未被抵扣掉的建设期可抵扣增值税
                ExcelApp.ThisWorkbook.Worksheets("资产负债表").Cells(41, i - 12).Value = 0
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
    Sub 所得税相关计算(ExcelApp As Object, sdsl_model As Integer)
        On Error Resume Next

        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        '————————————————————————————————————————————————————————————————————————————————————————        
        '逐年实际所得税额计算
        '————————————————————————————————————————————————————————————————————————————————————————    
        '项目计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '所得税免除年份数(年)
        Dim sdsmc_nfs As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(11, 7).Value
        '所得税减征年份数(年)
        Dim sdsjz_nfs As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(12, 7).Value
        '所得税减少征收比例(%)
        Dim sdsjz_bl As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(13, 7).Value
        '所得税率(%)
        Dim sdsl As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(13, 7).Value
        '读取逐年收入金额，列表，长度31
        Dim znsr_list(31) As Double
        '逐年应纳所得税额，列表，长度31
        Dim ynsdse_list(31) As Double
        '前15年
        For i = 1 To 15
            znsr_list(i) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(5, i + 4).Value
            ynsdse_list(i) = ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(11, i + 4).Value
        Next
        '16-31年
        For i = 16 To 31
            znsr_list(i) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(28, i - 12).Value
            ynsdse_list(i) = ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(31, i - 12).Value
        Next
        '采用默认值，则按照实际情况计算所得税减免和逐年所得税
        If sdsl_model = 0 Then
            '计算所得税
            Dim ans_sds = 所得税减免计算_main(znsr_list, ynsdse_list, sdsmc_nfs, sdsjz_nfs, sdsjz_bl, sdsl)
            '逐年所得税，列表，长度31
            Dim ans_sds_sl_list = ans_sds(0)
            Dim ans_sds_je_list = ans_sds(1)
            '逐年所得税金额写入Excel
            '前15年
            For i = 1 To 15
                If i <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(12, i + 4).Value = ans_sds_je_list(i)
                Else
                    ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(12, i + 4).Value = 0
                End If
            Next
            '16-31年
            For i = 16 To 31
                If i <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(32, i - 12).Value = ans_sds_je_list(i)
                Else
                    ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(32, i - 12).Value = 0
                End If
            Next
            '逐年所得税率写入Excel
            For i = 3 To 33
                If i - 2 <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(49, i).Value = ans_sds_sl_list(i - 2)
                Else
                    ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(49, i).Value = 0
                End If
            Next
        Else
            '读取Excel中的逐年所得税率
            Dim sdsl_zn_list(31) As Double
            For i = 3 To 33
                sdsl_zn_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(49, i).Value
            Next
            '计算逐年所得税金额并写入Excel
            '前15年
            For i = 1 To 15
                If i <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(12, i + 4).Value = ynsdse_list(i) * sdsl_zn_list(i)
                Else
                    ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(12, i + 4).Value = 0
                End If
            Next
            '16-31年
            For i = 16 To 31
                If i <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(32, i - 12).Value = ynsdse_list(i) * sdsl_zn_list(i)
                Else
                    ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(32, i - 12).Value = 0
                End If
            Next
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
    Function 可抵扣增值税计算_main(kdk_list As Array, zzs_xx_list As Array, zzs_jx_list As Array, kdk_nfs As Integer, jsnx As Integer, dkjs As Boolean)
        'kdk_list：建设期逐年可抵扣增值税总额，列表，长度31
        'zzs_xx_list：逐年销项增值税金额，列表，长度31
        'zzs_jx_list：逐年进项增值税金额，列表，长度31
        'kdk_nfs：建设期可抵扣增值税的最大抵扣年份数
        'jsnx：项目计算年限
        'dkjs：是否计算建设期可抵扣增值税

        '逐年增值税金额(未考虑抵扣的情况)
        Dim zzs_wdk(31) As Double
        For i = 1 To 31
            If zzs_xx_list(i) - zzs_jx_list(i) < 0 Then
                zzs_wdk(i) = 0
            Else
                zzs_wdk(i) = zzs_xx_list(i) - zzs_jx_list(i)
            End If
        Next
        '逐年建设期可抵扣增值税累计值，考虑可抵扣年限的限制
        Dim kdk_lj_list(31) As Double
        For i = 1 To 31
            For j = Math.Max(i - jsnx, 0) To i
                kdk_lj_list(i) += kdk_list(j)
            Next
        Next
        '根据是否进行可抵扣增值税计算，分别计算
        '计算结果，逐年可抵扣增值税、已抵扣增值税、实际增值税金额
        Dim ans_ydk_list(31) As Double '已抵扣(已经被抵扣掉的增值税金额)
        Dim ans_zzs_ydk_list(31) As Double '增值税(扣除已经抵扣掉的增值税金额后，剩余的增值税金额)
        Dim ans_wdk_list(31) As Double '未抵扣(逐年未被抵扣掉的建设期可抵扣增值税)
        If dkjs = True Then
            For i = 1 To 31
                '计算已抵扣(已经被抵扣掉的增值税金额)
                If i <= kdk_nfs Then
                    If kdk_lj_list(i) >= zzs_wdk(i) Then
                        ans_ydk_list(i) = zzs_wdk(i)
                    Else
                        ans_ydk_list(i) = kdk_lj_list(i)
                    End If
                Else
                    ans_ydk_list(i) = 0
                End If
                '未抵扣(逐年未被抵扣掉的建设期可抵扣增值税)
                ans_wdk_list(i) = kdk_lj_list(i) - ans_ydk_list(i)
                '计算剩余的增值税(扣除已经抵扣掉的增值税金额后，剩余的增值税金额)
                ans_zzs_ydk_list(i) = zzs_wdk(i) - ans_ydk_list(i)
                '将i及之后年份的逐年累计可抵扣增值税减去已经抵扣的金额
                For j = i To 31
                    kdk_lj_list(j) -= ans_ydk_list(i)
                Next
            Next
        Else
            For i = 1 To 31
                ans_ydk_list(i) = 0
                ans_wdk_list(i) = kdk_lj_list(i) - ans_ydk_list(i)
                ans_zzs_ydk_list(i) = zzs_wdk(i) - ans_ydk_list(i)
            Next
        End If
        Dim ans(2)
        ans(0) = ans_ydk_list
        ans(1) = ans_zzs_ydk_list
        ans(2) = ans_wdk_list
        Return ans
    End Function
    Function 所得税减免计算_main(znsr_list As Array, ynsdse_list As Array, sdsmc_nfs As Integer,
                                 sdsjz_nfs As Integer, sdsjz_bl As Double, sdsl As Double)
        'znsr_list：逐年收入金额，列表，长度31
        'ynsdse_list：逐年应纳所得税额，列表，长度31
        'sdsmc_nfs：所得税免除的年份数量
        'sdsjz_nfs：所得税减征的年份数量
        'sdsjz_bl：所得税减征的百分比
        'sdsl：所得税率(初始值)

        '结果列表，逐年实际所得税率，列表，长度31
        Dim ans_sds_sl_list(31) As Double
        '结果列表，逐年实际所得税纳税金额，列表，长度31
        Dim ans_sds_je_list(31) As Double
        '判断第一个有收入的年份序号
        Dim nfxh_sr_No1 As Integer
        For i = 1 To 31
            If znsr_list(i) > 0 Then
                nfxh_sr_No1 = i
                Exit For
            End If
        Next
        '先进行所得税免除计算
        For i = 1 To 31
            If i >= nfxh_sr_No1 And i <= nfxh_sr_No1 + sdsmc_nfs - 1 Then
                '全部免除
                ans_sds_sl_list(i) = 0
                ans_sds_je_list(i) = 0
            ElseIf i >= nfxh_sr_No1 + sdsmc_nfs And i <= nfxh_sr_No1 + sdsmc_nfs + sdsjz_nfs - 1 Then
                '按照比例减征所得税
                ans_sds_sl_list(i) = sdsl * (1 - sdsjz_bl)
                ans_sds_je_list(i) = ynsdse_list(i) * ans_sds_sl_list(i)
            ElseIf i < nfxh_sr_No1 Then
                '没有收入，所以没有所得税
                ans_sds_sl_list(i) = 0
                ans_sds_je_list(i) = 0
            Else
                '足额缴纳所得税
                ans_sds_sl_list(i) = sdsl
                ans_sds_je_list(i) = ynsdse_list(i) * ans_sds_sl_list(i)
            End If
        Next
        '返回结果
        Dim ans(1)
        ans(0) = ans_sds_sl_list
        ans(1) = ans_sds_je_list
        Return ans
    End Function

End Module
