Module 流动资金计算
    Sub 建设期流动资金计算(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '从估算表读取计算所需数据
        Dim GSBSJ = 读取估算表数据(ExcelApp)
        '10次投资年份序号列表，列表，长度10
        Dim tznf_list = GSBSJ(0)
        '读取流动资金当期增加额，列表，长度31
        Dim ans_ldzjdqzje_list(31) As Double
        '前15年
        For i = 1 To 15
            ans_ldzjdqzje_list(i) = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(16, i + 5).Value
        Next
        '16-31年
        For i = 16 To 31
            ans_ldzjdqzje_list(i) = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(34, i - 12).Value
        Next
        '转换为长度10的列表
        Dim ldzjdqzje_list = 基础计算功能_31_to_10(tznf_list, ans_ldzjdqzje_list)
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '数据写入Excel，写入<估算表>
        '前5次
        For i = 1 To 5
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(35, 2 * i + 1).Value = ldzjdqzje_list(i)
        Next
        '6-10次
        For i = 6 To 10
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(78, 2 * i - 9).Value = ldzjdqzje_list(i)
        Next
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '数据写入<投资计划与资金筹措表>
        '前15年
        For i = 1 To 15
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(10, i + 4).Value = ans_ldzjdqzje_list(i)
        Next
        '16-31年
        For i = 16 To 31
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(31, i - 12).Value = ans_ldzjdqzje_list(i)
        Next
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '计算一次Excel
        ExcelApp.Calculate()
    End Sub
    Sub 流动资金相关计算(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '第一年的应收账款、原材料、燃料和动力、现金的和
        Dim DYN As Double = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(49, 3).Value + ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(50, 3).Value + ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(51, 3).Value + ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(52, 3).Value
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '————————————————————————————————————————————————————————————————————————————————————————
        '如果计算期第一年应收账款、原材料、燃料和动力、现金的和大于0，同时计算期第一年的负荷率等于0，同时计算期第一年收入为0，则报错
        If DYN > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(5, 3).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(5, 5).Value = 0 Then
            'MsgBox("计算期第一年设置的负荷率等于0，计算期第一年收入等于0，但是计算期第一年存在流动资金计算，请检查。")
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
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
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算逐年应收账款、原材料、燃料和动力、现金
            '前15年
            For i = 1 To 15
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
            '16-31年
            For i = 16 To 31
                If i <= jsnx Then
                    If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 2).Value = 0 Then
                        '应收账款
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(24, i - 12).Value = 0
                        '原材料
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(26, i - 12).Value = 0
                        '燃料和动力
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(27, i - 12).Value = 0
                        '现金
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(28, i - 12).Value = 0
                    Else
                        '应收账款
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(24, i - 12).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(49, i + 2).Value / ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 2).Value
                        '原材料
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(26, i - 12).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(50, i + 2).Value / ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 2).Value
                        '燃料和动力
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(27, i - 12).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(51, i + 2).Value / ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 2).Value
                        '现金
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(28, i - 12).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(52, i + 2).Value / ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 2).Value
                    End If
                Else
                    '应收账款
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(24, i - 12).Value = 0
                    '原材料
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(26, i - 12).Value = 0
                    '燃料和动力
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(27, i - 12).Value = 0
                    '现金
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(28, i - 12).Value = 0
                End If
            Next
            '计算一次Excel
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————
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
            '————————————————————————————————————————————————————————————————————————————————————————
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
            '————————————————————————————————————————————————————————————————————————————————————————
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
            '————————————————————————————————————————————————————————————————————————————————————————
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
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算逐年应收账款、原材料、燃料和动力、现金
            '前15年
            For i = 1 To 15
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
            '16-31年
            For i = 16 To 31
                If i <= jsnx - 1 Then
                    If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 3).Value = 0 Then
                        '应收账款
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(24, i - 12).Value = 0
                        '原材料
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(26, i - 12).Value = 0
                        '燃料和动力
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(27, i - 12).Value = 0
                        '现金
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(28, i - 12).Value = 0
                    Else
                        '应收账款
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(24, i - 12).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(49, i + 3).Value / ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 3).Value
                        '原材料
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(26, i - 12).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(50, i + 3).Value / ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 3).Value
                        '燃料和动力
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(27, i - 12).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(51, i + 3).Value / ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 3).Value
                        '现金
                        ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(28, i - 12).Value = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(52, i + 3).Value / ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 3).Value
                    End If
                Else
                    '应收账款
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(24, i - 12).Value = 0
                    '原材料
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(26, i - 12).Value = 0
                    '燃料和动力
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(27, i - 12).Value = 0
                    '现金
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(28, i - 12).Value = 0
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
            '计算期末偿还流动资金本金
            '前15年
            For i = 1 To 15
                If i = jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(16, i + 4).Value = ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(14, i + 4).Value
                Else
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(16, i + 4).Value = 0
                End If
            Next
            '16-31年
            For i = 16 To 31
                If i = jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(46, i - 12).Value = ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(44, i - 12).Value
                Else
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(46, i - 12).Value = 0
                End If
            Next
            '计算一次Excel
            ExcelApp.Calculate()
        End If
        '计算一次Excel
        ExcelApp.Calculate()
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '建设期流动资金计算
        Call 建设期流动资金计算(ExcelApp)
    End Sub
End Module
