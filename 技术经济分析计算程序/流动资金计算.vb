Module 流动资金计算
    Sub 流动资金相关计算(ExcelApp As Object, xlfl_cg_model As Integer, xlfl_qt_model As Integer, kcje_xlf_model As Integer, hscz As Boolean, zbj_model As Integer,
                         clfl_qtfl_model As Integer, kcje_clf_qtf_model As Integer, ldzj_model As Integer, kcje_ldzj_model As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————   
        'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'hscy：计算期末，是否回收资产残值
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'ldzj_model：流动资金的计算方式，0：使用默认值，1：从Excel中读取已有的值，预留功能
        'kcje_ldzj_model：流动资金计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值，预留功能
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '分项逐年流动资金计算
        Dim ans_ldzj = 分项逐年流动资金计算(ExcelApp, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model)
        Dim ldzj_all_list = ans_ldzj(0)
        Dim zyldzj_all_list = ans_ldzj(1)
        Dim xj_all_list = ans_ldzj(2)
        Dim yszk_cg = ans_ldzj(3)
        Dim ycl_cg = ans_ldzj(4)
        Dim rldl_cg = ans_ldzj(5)
        '————————————————————————————————————————————————————————————————————————————————————————
        '第一年的应收账款、原材料、燃料和动力、现金的和
        Dim DYN As Double = yszk_cg(1) + ycl_cg(1) + rldl_cg(1) + xj_all_list(1)
        If DYN > 0 Then
            '如果第一年有收入和成本，则流动资金的计算年份与项目计算年份一一对应，流动资金计算年限等于项目计算年限
            '流动资金当期增加额计算系数
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
            '逐年数据写入Excel
            '前15年
            For i = 1 To 15
                If i > jsnx Or ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 2).Value = 0 Then
                    '应收账款
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(6, i + 5).Value = 0
                    '原材料
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(8, i + 5).Value = 0
                    '燃料和动力
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(9, i + 5).Value = 0
                    '现金
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(10, i + 5).Value = 0
                    '流动资金总额
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(13, i + 5).Value = 0
                    '自有流动资金
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(14, i + 5).Value = 0
                Else
                    '应收账款
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(6, i + 5).Value = yszk_cg(i)
                    '原材料
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(8, i + 5).Value = ycl_cg(i)
                    '燃料和动力
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(9, i + 5).Value = rldl_cg(i)
                    '现金
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(10, i + 5).Value = xj_all_list(i)
                    '流动资金总额
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(13, i + 5).Value = ldzj_all_list(i)
                    '自有流动资金
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(14, i + 5).Value = zyldzj_all_list(i)
                End If
            Next
            '16-31年
            For i = 16 To 31
                If i > jsnx Or ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 2).Value = 0 Then
                    '应收账款
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(24, i - 12).Value = 0
                    '原材料
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(26, i - 12).Value = 0
                    '燃料和动力
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(27, i - 12).Value = 0
                    '现金
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(28, i - 12).Value = 0
                    '流动资金总额
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(31, i - 12).Value = 0
                    '自有流动资金
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(32, i - 12).Value = 0
                Else
                    '应收账款
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(24, i - 12).Value = yszk_cg(i)
                    '原材料
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(26, i - 12).Value = ycl_cg(i)
                    '燃料和动力
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(27, i - 12).Value = rldl_cg(i)
                    '现金
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(28, i - 12).Value = xj_all_list(i)
                    '流动资金总额
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(31, i - 12).Value = ldzj_all_list(i)
                    '自有流动资金
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(32, i - 12).Value = zyldzj_all_list(i)
                End If
            Next
            '计算一次Excel
            ExcelApp.Calculate()
            '———————————————————————————————————————————————————————————————————————————————————————— 
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
            '逐年数据写入Excel
            '前15年
            For i = 1 To 15
                If i > jsnx - 1 Or ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 3).Value = 0 Then
                    '应收账款
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(6, i + 5).Value = 0
                    '原材料
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(8, i + 5).Value = 0
                    '燃料和动力
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(9, i + 5).Value = 0
                    '现金
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(10, i + 5).Value = 0
                    '流动资金总额
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(13, i + 5).Value = 0
                    '自有流动资金
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(14, i + 5).Value = 0
                Else
                    '应收账款
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(6, i + 5).Value = yszk_cg(i + 1)
                    '原材料
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(8, i + 5).Value = ycl_cg(i + 1)
                    '燃料和动力
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(9, i + 5).Value = rldl_cg(i + 1)
                    '现金
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(10, i + 5).Value = xj_all_list(i + 1)
                    '流动资金总额
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(13, i + 5).Value = ldzj_all_list(i + 1)
                    '自有流动资金
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(14, i + 5).Value = zyldzj_all_list(i + 1)
                End If
            Next
            '16-31年
            For i = 16 To 31
                If i > jsnx - 1 Or ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 3).Value = 0 Then
                    '应收账款
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(24, i - 12).Value = 0
                    '原材料
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(26, i - 12).Value = 0
                    '燃料和动力
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(27, i - 12).Value = 0
                    '现金
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(28, i - 12).Value = 0
                    '流动资金总额
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(31, i - 12).Value = 0
                    '自有流动资金
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(32, i - 12).Value = 0
                Else
                    '应收账款
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(24, i - 12).Value = yszk_cg(i + 1)
                    '原材料
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(26, i - 12).Value = ycl_cg(i + 1)
                    '燃料和动力
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(27, i - 12).Value = rldl_cg(i + 1)
                    '现金
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(28, i - 12).Value = xj_all_list(i + 1)
                    '流动资金总额
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(31, i - 12).Value = ldzj_all_list(i + 1)
                    '自有流动资金
                    ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(32, i - 12).Value = zyldzj_all_list(i + 1)
                End If
            Next
            '计算一次Excel
            ExcelApp.Calculate()
            '———————————————————————————————————————————————————————————————————————————————————————— 
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
            '———————————————————————————————————————————————————————————————————————————————————————— 
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
            '———————————————————————————————————————————————————————————————————————————————————————— 
            '计算期末最后一年不减去自有流动资金
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("资产负债表").Cells(66, i).Value = 0
            Next
            '计算一次Excel
            ExcelApp.Calculate()
            '———————————————————————————————————————————————————————————————————————————————————————— 
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
            '———————————————————————————————————————————————————————————————————————————————————————— 
        End If
        '计算一次Excel
        ExcelApp.Calculate()
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '建设期流动资金计算
        Call 建设期流动资金计算(ExcelApp)
    End Sub
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
    Function 分项逐年流动资金计算(ExcelApp As Object, xlfl_cg_model As Integer, xlfl_qt_model As Integer, kcje_xlf_model As Integer, hscz As Boolean, zbj_model As Integer,
                                  clfl_qtfl_model As Integer, kcje_clf_qtf_model As Integer, ldzj_model As Integer, kcje_ldzj_model As Integer)
        'On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————   
        'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'hscy：计算期末，是否回收资产残值
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'ldzj_model：流动资金的计算方式，0：使用默认值，1：从Excel中读取已有的值，预留功能
        'kcje_ldzj_model：流动资金计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值，预留功能
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim GSBSJ = 读取估算表数据(ExcelApp)
        '投资年份，10次投资的情况
        Dim tznf_list = GSBSJ(0)
        '燃机总发电量(万kWh)，10次投资的情况
        Dim rjfdl = GSBSJ(11)
        Dim rjfdl_list = 基础计算功能_10_to_31(tznf_list, rjfdl)
        '蓄电池总装机功率(kW)，10次投资的情况
        Dim xdczjgl = GSBSJ(13)
        Dim xdczjgl_list = 基础计算功能_10_to_31(tznf_list, xdczjgl)
        '供冷供热总量(万kWh)，10次投资的情况
        Dim glgrl = GSBSJ(15)
        Dim glgrl_list = 基础计算功能_10_to_31(tznf_list, glgrl)
        '光伏总装机功率(kW)，10次投资的情况
        Dim gfzjgl = GSBSJ(17)
        Dim gfzjgl_list = 基础计算功能_10_to_31(tznf_list, gfzjgl)
        '风电总装机功率(kW)，10次投资的情况
        Dim fdzjgl = GSBSJ(19)
        Dim fdzjgl_list = 基础计算功能_10_to_31(tznf_list, fdzjgl)
        '流动资金费率，默认：从EXCEL读取，逐年固定费率
        Dim ans_fl_mr = 默认逐年流动资金费率(ExcelApp)
        Dim ldzjfl_gf_list = ans_fl_mr(0)
        Dim ldzjfl_fd_list = ans_fl_mr(1)
        Dim ldzjfl_xdc_list = ans_fl_mr(2)
        Dim ldzjfl_rj_list = ans_fl_mr(3)
        Dim ldzjfl_nt_list = ans_fl_mr(4)
        '计算基数扣除默认设置
        Dim yynx_rj As Integer
        Dim yynx_xdc As Integer
        Dim yynx_nt As Integer
        Dim yynx_gf As Integer
        Dim yynx_fd As Integer
        Dim kcbl_rj As Double
        Dim kcbl_xdc As Double
        Dim kcbl_nt As Double
        Dim kcbl_gf As Double
        Dim kcbl_fd As Double
        If kcje_ldzj_model = 0 Then
            Dim kcje_mr = 流动资金计算基数扣除默认设置(ExcelApp)
            yynx_rj = kcje_mr(0)
            yynx_xdc = kcje_mr(1)
            yynx_nt = kcje_mr(2)
            yynx_gf = kcje_mr(3)
            yynx_fd = kcje_mr(4)
            kcbl_rj = kcje_mr(5)
            kcbl_xdc = kcje_mr(6)
            kcbl_nt = kcje_mr(7)
            kcbl_gf = kcje_mr(8)
            kcbl_fd = kcje_mr(9)
        Else
            yynx_rj = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(56, 3).Value
            yynx_xdc = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(56, 4).Value
            yynx_nt = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(56, 5).Value
            yynx_gf = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(56, 6).Value
            yynx_fd = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(56, 7).Value
            kcbl_rj = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(57, 3).Value
            kcbl_xdc = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(57, 4).Value
            kcbl_nt = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(57, 5).Value
            kcbl_gf = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(57, 6).Value
            kcbl_fd = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(57, 7).Value
        End If
        '根据扣除设置，修改逐年计算基数
        '燃机
        Dim ans_ldzjfl_rj = 计算费率修正计算基础功能(tznf_list, rjfdl, rjfdl_list, ldzjfl_rj_list, yynx_rj, kcbl_rj)
        ldzjfl_rj_list = ans_ldzjfl_rj(0)
        '暖通
        Dim ans_ldzjfl_nt = 计算费率修正计算基础功能(tznf_list, glgrl, glgrl_list, ldzjfl_nt_list, yynx_nt, kcbl_nt)
        ldzjfl_nt_list = ans_ldzjfl_nt(0)
        '光伏
        Dim ans_ldzjfl_gf = 计算费率修正计算基础功能(tznf_list, gfzjgl, gfzjgl_list, ldzjfl_gf_list, yynx_gf, kcbl_gf)
        ldzjfl_gf_list = ans_ldzjfl_gf(0)
        '风电
        Dim ans_ldzjfl_fd = 计算费率修正计算基础功能(tznf_list, fdzjgl, fdzjgl_list, ldzjfl_fd_list, yynx_fd, kcbl_fd)
        ldzjfl_fd_list = ans_ldzjfl_fd(0)
        '蓄电池
        Dim ans_ldzjfl_xdc = 计算费率修正计算基础功能(tznf_list, xdczjgl, xdczjgl_list, ldzjfl_xdc_list, yynx_xdc, kcbl_xdc)
        ldzjfl_xdc_list = ans_ldzjfl_xdc(0)
        '————————————————————————————————————————————————————————————————————————————————————————
        '常规设备流动资金计算
        Dim ans_cg = 逐年流动资金计算_常规设备(ExcelApp, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model)
        Dim ldzj_cg = ans_cg(0)
        Dim zyldzj_cg = ans_cg(1)
        Dim yszk_cg = ans_cg(2)
        Dim ycl_cg = ans_cg(3)
        Dim rldl_cg = ans_cg(4)
        Dim xj_cg = ans_cg(5)
        '光伏风电蓄电池流动资金计算
        Dim ans_xny = 逐年流动资金计算_光伏风电蓄电池(ExcelApp, gfzjgl_list, fdzjgl_list, xdczjgl_list, ldzjfl_gf_list, ldzjfl_fd_list, ldzjfl_xdc_list)
        Dim ldzj_gf = ans_xny(0)
        Dim ldzj_fd = ans_xny(1)
        Dim ldzj_xdc = ans_xny(2)
        Dim zyldzj_gf = ans_xny(3)
        Dim zyldzj_fd = ans_xny(4)
        Dim zyldzj_xdc = ans_xny(5)
        '燃机和暖通流动资金计算
        Dim ans_rjnt = 逐年流动资金计算_燃机和暖通(ExcelApp, rjfdl_list, glgrl_list, ldzjfl_rj_list, ldzjfl_nt_list)
        Dim ldzj_rj = ans_rjnt(0)
        Dim ldzj_nt = ans_rjnt(1)
        Dim zyldzj_rj = ans_rjnt(2)
        Dim zyldzj_nt = ans_rjnt(3)
        '————————————————————————————————————————————————————————————————————————————————————————
        '数据汇总，列表长度32
        Dim ldzj_all_list(32) As Double
        Dim zyldzj_all_list(32) As Double
        Dim xj_all_list(32) As Double '光伏风电蓄电池的流动资金，全部加到“现金”这一项里
        For i = 1 To 31
            ldzj_all_list(i) = ldzj_cg(i) + ldzj_gf(i) + ldzj_fd(i) + ldzj_xdc(i) + ldzj_rj(i) + ldzj_nt(i)
            zyldzj_all_list(i) = zyldzj_cg(i) + zyldzj_gf(i) + zyldzj_fd(i) + zyldzj_xdc(i) + zyldzj_rj(i) + zyldzj_nt(i)
            xj_all_list(i) = xj_cg(i) + ldzj_gf(i) + ldzj_fd(i) + ldzj_xdc(i) + ldzj_rj(i) + ldzj_nt(i)
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '返回结果
        Dim ans(5)
        ans(0) = ldzj_all_list
        ans(1) = zyldzj_all_list
        ans(2) = xj_all_list
        ans(3) = yszk_cg
        ans(4) = ycl_cg
        ans(5) = rldl_cg
        Return ans
    End Function
    Function 逐年流动资金计算_常规设备(ExcelApp As Object, xlfl_cg_model As Integer, xlfl_qt_model As Integer, kcje_xlf_model As Integer, hscz As Boolean, zbj_model As Integer,
                                       clfl_qtfl_model As Integer, kcje_clf_qtf_model As Integer)
        'On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————   
        'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'hscy：计算期末，是否回收资产残值
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        '————————————————————————————————————————————————————————————————————————————————————————
        '分项逐年修理费，列表长度31
        Dim ans_xlf = 分项逐年修理费计算(ExcelApp, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model)
        Dim xlf_rj = ans_xlf(2)
        Dim xlf_nt = ans_xlf(3)
        Dim xlf_gf = ans_xlf(4)
        Dim xlf_fd = ans_xlf(5)
        Dim xlf_xdc = ans_xlf(6)
        '修理费进项税率
        Dim xlf_jxsl As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(140, 22).Value
        '————————————————————————————————————————————————————————————————————————————————————————
        '分项逐年保险费，列表长度31
        Dim ans_bxf = 分项逐年保险费金额计算(ExcelApp, hscz, zbj_model)
        Dim bxf_cg = ans_bxf(0)
        Dim bxf_gf = ans_bxf(1)
        Dim bxf_fd = ans_bxf(2)
        Dim bxf_xdc = ans_bxf(3)
        Dim bxf_rj = ans_bxf(4)
        Dim bxf_nt = ans_bxf(5)
        '保险费进项税率
        Dim bxf_jxsl As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(210, 22).Value
        '————————————————————————————————————————————————————————————————————————————————————————
        '分项逐年材料费其它费，列表长度31
        Dim ans_clfqtf = 分项逐年材料费其它费计算(ExcelApp, clfl_qtfl_model, kcje_clf_qtf_model)
        Dim clf_rj = ans_clfqtf(2)
        Dim clf_nt = ans_clfqtf(5)
        Dim clf_gf = ans_clfqtf(6)
        Dim clf_fd = ans_clfqtf(7)
        Dim clf_xdc = ans_clfqtf(8)
        Dim qtf_rj = ans_clfqtf(9)
        Dim qtf_nt = ans_clfqtf(12)
        Dim qtf_gf = ans_clfqtf(13)
        Dim qtf_fd = ans_clfqtf(14)
        Dim qtf_xdc = ans_clfqtf(15)
        '材料费其它费进项税率
        Dim clf_jxsl As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(235, 22).Value
        Dim qtf_jxsl As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(236, 22).Value
        '————————————————————————————————————————————————————————————————————————————————————————   
        '判断是否为纯新能源项目（投资中仅有：光伏+风电+蓄电池）
        Dim xnytz_all As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(46, 1).Value + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(50, 1).Value + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(53, 1).Value
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '流动资金常规计算模式设置，0：采用常规计算模型；1：采用新能源计算模式
        Dim trq_ldzj_cg As Integer = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 33).Value
        Dim mt_ldzj_cg As Integer = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 33).Value
        Dim gd_ldzj_cg As Integer = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 33).Value
        Dim wgzq_ldzj_cg As Integer = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 33).Value
        Dim gdrlf_ldzj_cg As Integer = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 33).Value
        Dim bs_ldzj_cg As Integer = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 33).Value
        Dim shs_ldzj_cg As Integer = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 33).Value
        Dim ssh_ldzj_cg As Integer = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 33).Value
        Dim ns_ldzj_cg As Integer = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 33).Value
        Dim csglf_ldzj_cg As Integer = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 33).Value
        Dim rygz_ldzj_cg As Integer = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 33).Value
        Dim rjtz_ldzj_cg As Integer = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(30, 33).Value
        Dim nttz_ldzj_cg As Integer = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(32, 33).Value
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '计算逐年应收账款、原材料、燃料和动力、现金（去除光伏、风电、蓄电池的相关金额），列表长度32
        Dim yszk_list(32) As Double
        Dim ycl_list(32) As Double
        Dim rldl_list(32) As Double
        Dim xj_list(32) As Double
        If xnytz_all = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value Then
            '如果是纯新能源项目，则数据都是0
            For i = 1 To 32
                yszk_list(i) = 0
                ycl_list(i) = 0
                rldl_list(i) = 0
                xj_list(i) = 0
            Next
        Else
            '如果不是纯新能源项目，则正常计算
            '前15年
            For i = 1 To 15
                '应收账款 = 含税经营成本 - 含税蓄电池购电成本 - 含税光伏风电蓄电池修理费 - 含税光伏风电蓄电池保险费 - 含税光伏风电蓄电池材料费 - 含税光伏风电蓄电池其它费
                yszk_list(i) = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(21, i + 4).Value + ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(22, i + 4).Value -
                               ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(53, i + 4).Value - ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(54, i + 4).Value -
                               (xlf_gf(i) + xlf_fd(i) + xlf_xdc(i)) * (1 + xlf_jxsl) - (bxf_gf(i) + bxf_fd(i) + bxf_xdc(i)) * (1 + bxf_jxsl) -
                               (clf_gf(i) + clf_fd(i) + clf_xdc(i)) * (1 + clf_jxsl) - (qtf_gf(i) + qtf_fd(i) + qtf_xdc(i)) * (1 + qtf_jxsl) -
                               trq_ldzj_cg * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(5, i + 4).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(6, i + 4).Value) -
                               mt_ldzj_cg * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(17, i + 4).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(18, i + 4).Value) -
                               gd_ldzj_cg * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(29, i + 4).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(30, i + 4).Value) -
                               wgzq_ldzj_cg * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(41, i + 4).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(42, i + 4).Value) -
                               gdrlf_ldzj_cg * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(65, i + 4).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(66, i + 4).Value) -
                               bs_ldzj_cg * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(77, i + 4).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(78, i + 4).Value) -
                               shs_ldzj_cg * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(89, i + 4).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(90, i + 4).Value) -
                               ssh_ldzj_cg * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(101, i + 4).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(102, i + 4).Value) -
                               ns_ldzj_cg * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(113, i + 4).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(114, i + 4).Value) -
                               csglf_ldzj_cg * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(125, i + 4).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(126, i + 4).Value) -
                               rygz_ldzj_cg * ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(199, i + 4).Value -
                               rjtz_ldzj_cg * (xlf_rj(i) * (1 + xlf_jxsl) + bxf_rj(i) * (1 + bxf_jxsl) + clf_rj(i) * (1 + clf_jxsl) + qtf_rj(i) * (1 + qtf_jxsl)) -
                               nttz_ldzj_cg * (xlf_nt(i) * (1 + xlf_jxsl) + bxf_nt(i) * (1 + bxf_jxsl) + clf_nt(i) * (1 + clf_jxsl) + qtf_nt(i) * (1 + qtf_jxsl))
                '原材料 = 含税所有的原材料成本 - 含税光伏风电蓄电池材料费
                ycl_list(i) = (1 - bs_ldzj_cg) * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(77, i + 4).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(78, i + 4).Value) +
                              (1 - shs_ldzj_cg) * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(89, i + 4).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(90, i + 4).Value) +
                              (1 - ssh_ldzj_cg) * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(101, i + 4).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(102, i + 4).Value) +
                              (1 - ns_ldzj_cg) * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(113, i + 4).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(114, i + 4).Value) +
                              (1 - csglf_ldzj_cg) * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(125, i + 4).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(126, i + 4).Value) +
                              ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(225, i + 4).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(226, i + 4).Value -
                              (clf_gf(i) + clf_fd(i) + clf_xdc(i) + rjtz_ldzj_cg * clf_rj(i) + nttz_ldzj_cg * clf_nt(i)) * (1 + clf_jxsl)
                '燃料和动力
                rldl_list(i) = (1 - trq_ldzj_cg) * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(5, i + 4).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(6, i + 4).Value) +
                               (1 - mt_ldzj_cg) * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(17, i + 4).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(18, i + 4).Value) +
                               (1 - gd_ldzj_cg) * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(29, i + 4).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(30, i + 4).Value) +
                               (1 - wgzq_ldzj_cg) * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(41, i + 4).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(42, i + 4).Value) +
                               (1 - gdrlf_ldzj_cg) * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(65, i + 4).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(66, i + 4).Value)
                '现金 = 工资 + 含税保险费 + 含税其它费 - 含税光伏风电蓄电池保险费 - 含税光伏风电蓄电池其它费
                xj_list(i) = (1 - rygz_ldzj_cg) * ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(199, i + 4).Value +
                             ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(211, i + 4).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(212, i + 4).Value +
                             ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(227, i + 4).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(228, i + 4).Value -
                             (bxf_gf(i) + bxf_fd(i) + bxf_xdc(i) + rjtz_ldzj_cg * bxf_rj(i) + nttz_ldzj_cg * bxf_nt(i)) * (1 + bxf_jxsl) -
                             (qtf_gf(i) + qtf_fd(i) + qtf_xdc(i) + rjtz_ldzj_cg * qtf_rj(i) + nttz_ldzj_cg * qtf_nt(i)) * (1 + qtf_jxsl)
            Next
            '第16~31年
            For i = 16 To 31
                '应收账款 = 含税经营成本 - 含税蓄电池购电成本 - 含税光伏风电蓄电池修理费 - 含税光伏风电蓄电池保险费 - 含税光伏风电蓄电池材料费 - 含税光伏风电蓄电池其它费
                yszk_list(i) = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(53, i - 12).Value + ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(45, i - 12).Value -
                               ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(57, i - 12).Value - ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(58, i - 12).Value -
                               (xlf_gf(i) + xlf_fd(i) + xlf_xdc(i)) * (1 + xlf_jxsl) - (bxf_gf(i) + bxf_fd(i) + bxf_xdc(i)) * (1 + bxf_jxsl) -
                               (clf_gf(i) + clf_fd(i) + clf_xdc(i)) * (1 + clf_jxsl) - (qtf_gf(i) + qtf_fd(i) + qtf_xdc(i)) * (1 + qtf_jxsl) -
                               trq_ldzj_cg * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(9, i - 12).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(10, i - 12).Value) -
                               mt_ldzj_cg * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(21, i - 12).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(22, i - 12).Value) -
                               gd_ldzj_cg * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(33, i - 12).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(34, i - 12).Value) -
                               wgzq_ldzj_cg * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(45, i - 12).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(46, i - 12).Value) -
                               gdrlf_ldzj_cg * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(69, i - 12).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(70, i - 12).Value) -
                               bs_ldzj_cg * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(81, i - 12).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(82, i - 12).Value) -
                               shs_ldzj_cg * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(93, i - 12).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(94, i - 12).Value) -
                               ssh_ldzj_cg * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(105, i - 12).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(106, i - 12).Value) -
                               ns_ldzj_cg * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(117, i - 12).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(118, i - 12).Value) -
                               csglf_ldzj_cg * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(129, i - 12).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(130, i - 12).Value) -
                               rygz_ldzj_cg * ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(203, i - 12).Value -
                               rjtz_ldzj_cg * (xlf_rj(i) * (1 + xlf_jxsl) + bxf_rj(i) * (1 + bxf_jxsl) + clf_rj(i) * (1 + clf_jxsl) + qtf_rj(i) * (1 + qtf_jxsl)) -
                               nttz_ldzj_cg * (xlf_nt(i) * (1 + xlf_jxsl) + bxf_nt(i) * (1 + bxf_jxsl) + clf_nt(i) * (1 + clf_jxsl) + qtf_nt(i) * (1 + qtf_jxsl))
                '原材料 = 含税所有的原材料成本 - 含税光伏风电蓄电池材料费
                ycl_list(i) = (1 - bs_ldzj_cg) * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(81, i - 12).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(82, i - 12).Value) +
                              (1 - shs_ldzj_cg) * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(93, i - 12).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(94, i - 12).Value) +
                              (1 - ssh_ldzj_cg) * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(105, i - 12).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(106, i - 12).Value) +
                              (1 - ns_ldzj_cg) * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(117, i - 12).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(118, i - 12).Value) +
                              (1 - csglf_ldzj_cg) * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(129, i - 12).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(130, i - 12).Value) +
                              ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(233, i - 12).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(234, i - 12).Value -
                              (clf_gf(i) + clf_fd(i) + clf_xdc(i) + rjtz_ldzj_cg * clf_rj(i) + nttz_ldzj_cg * clf_nt(i)) * (1 + clf_jxsl)
                '燃料和动力
                rldl_list(i) = (1 - trq_ldzj_cg) * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(9, i - 12).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(10, i - 12).Value) +
                               (1 - mt_ldzj_cg) * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(21, i - 12).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(22, i - 12).Value) +
                               (1 - gd_ldzj_cg) * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(33, i - 12).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(34, i - 12).Value) +
                               (1 - wgzq_ldzj_cg) * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(45, i - 12).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(46, i - 12).Value) +
                               (1 - gdrlf_ldzj_cg) * (ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(69, i - 12).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(70, i - 12).Value)
                '现金 = 工资 + 含税保险费 + 含税其它费 - 含税光伏风电蓄电池保险费 - 含税光伏风电蓄电池其它费
                xj_list(i) = (1 - rygz_ldzj_cg) * ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(203, i - 12).Value +
                             ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(215, i - 12).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(216, i - 12).Value +
                             ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(235, i - 12).Value + ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(236, i - 12).Value -
                             (bxf_gf(i) + bxf_fd(i) + bxf_xdc(i) + rjtz_ldzj_cg * bxf_rj(i) + nttz_ldzj_cg * bxf_nt(i)) * (1 + bxf_jxsl) -
                             (qtf_gf(i) + qtf_fd(i) + qtf_xdc(i) + rjtz_ldzj_cg * qtf_rj(i) + nttz_ldzj_cg * qtf_nt(i)) * (1 + qtf_jxsl)
            Next
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '金额除以周转次数
        For i = 1 To 31
            yszk_list(i) = yszk_list(i) / ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 2).Value
            ycl_list(i) = ycl_list(i) / ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 2).Value
            rldl_list(i) = rldl_list(i) / ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 2).Value
            xj_list(i) = xj_list(i) / ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 2).Value
        Next
        '————————————————————————————————————————————————————————————————————————————————————————   
        '自有流动资金比例
        Dim zyldzjbl As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 5).Value
        Dim ldzj_cg(32) As Double '流动资金总额，列表长度32
        Dim zyldzj_cg(32) As Double '自有流动资金金额，列表长度32
        '流动资金 = 流动资产 - 流动负债 = 应收账款 + 原材 + 燃料和动力 + 现金 - （原材 + 燃料和动力） = 应收账款 + 现金
        For i = 1 To 31
            ldzj_cg(i) = yszk_list(i) + xj_list(i)
            zyldzj_cg(i) = zyldzjbl * ldzj_cg(i)
        Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '返回结果
        Dim ans(5)
        ans(0) = ldzj_cg
        ans(1) = zyldzj_cg
        ans(2) = yszk_list
        ans(3) = ycl_list
        ans(4) = rldl_list
        ans(5) = xj_list
        Return ans
    End Function
    Function 逐年流动资金计算_光伏风电蓄电池(ExcelApp As Object, gfzjgl_list As Array, fdzjgl_list As Array, xdczjgl_list As Array, ldzjfl_gf_list As Array,
                                             ldzjfl_fd_list As Array, ldzjfl_xdc_list As Array)
        '光伏风电蓄电池流动资金直接按照费率计算，不用除以周转次数
        'gfzjgl_list，fdzjgl_list， xdczjgl_list： 与下面的量一一对应
        '光伏总装机功率(kW)，10次投资的情况，列表，长度31
        '风电总装机功率(kW)，10次投资的情况，列表，长度31
        '蓄电池总装机功率(kW)，10次投资的情况，列表，长度31
        'ldzjfl_gf_lis，ldzjfl_fd_list，ldzjfl_xdc_list：与下面的流动资金费率一一对应
        '光伏流动资金费率， 元 / kW，列表，长度31
        '风电流动资金费率， 元 / kW，列表，长度31
        '蓄电池流动资金费率， 元 / kW，列表，长度31
        '————————————————————————————————————————————————————————————————————————————————————————  
        '自有流动资金比例
        Dim zyldzjbl_gf As Double = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(5, 23).Value
        Dim zyldzjbl_fd As Double = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(7, 23).Value
        Dim zyldzjbl_xdc As Double = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(9, 23).Value
        '风电、光伏、蓄电池，装机kW的逐年累计值计算
        Dim gfzjgl_lj As Double = 0
        Dim fdzjgl_lj As Double = 0
        Dim xdczjgl_lj As Double = 0
        '风电、光伏、蓄电池，逐年流动资金金额，列表长度32
        Dim ldzj_gf(32) As Double '流动资金总额
        Dim ldzj_fd(32) As Double '流动资金总额
        Dim ldzj_xdc(32) As Double '流动资金总额
        Dim zyldzj_gf(32) As Double '自有流动资金金额
        Dim zyldzj_fd(32) As Double '自有流动资金金额
        Dim zyldzj_xdc(32) As Double '自有流动资金金额
        For i = 1 To 31
            '计算前一年累积量，用前一年的累积量进行计算
            gfzjgl_lj += gfzjgl_list(i - 1)
            fdzjgl_lj += fdzjgl_list(i - 1)
            xdczjgl_lj += xdczjgl_list(i - 1)
            '计算流动资金金额
            ldzj_gf(i) = gfzjgl_lj * ldzjfl_gf_list(i) / 10000
            ldzj_fd(i) = fdzjgl_lj * ldzjfl_fd_list(i) / 10000
            ldzj_xdc(i) = xdczjgl_lj * ldzjfl_xdc_list(i) / 10000
            '计算自有流动资金金额
            zyldzj_gf(i) = ldzj_gf(i) * zyldzjbl_gf
            zyldzj_fd(i) = ldzj_fd(i) * zyldzjbl_fd
            zyldzj_xdc(i) = ldzj_xdc(i) * zyldzjbl_xdc
        Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '返回结果
        Dim ans(5)
        ans(0) = ldzj_gf
        ans(1) = ldzj_fd
        ans(2) = ldzj_xdc
        ans(3) = zyldzj_gf
        ans(4) = zyldzj_fd
        ans(5) = zyldzj_xdc
        Return ans
    End Function
    Function 逐年流动资金计算_燃机和暖通(ExcelApp As Object, rjfdl_list As Array, glgrl_list As Array, ldzjfl_rj_list As Array, ldzjfl_nt_list As Array)
        '燃机和暖通流动资金直接按照费率计算，不用除以周转次数
        'rjfdl_list, glgrl_list：与下面的量一一对应
        '燃机总发电量(万kWh)，10次投资的情况，列表，长度31
        '供冷供热总量(万kWh)，10次投资的情况，列表，长度31
        'ldzjfl_rj_lis，ldzjfl_nt_list：与下面的流动资金费率一一对应
        '燃机流动资金费率， 元 / MWh，列表，长度31
        '暖通流动资金费率， 元 / MWh，列表，长度31
        '————————————————————————————————————————————————————————————————————————————————————————  
        '是否将燃机或者暖通流动资金使用“新能源”计算模式取计算
        Dim rj_mode As Integer = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(30, 33).Value
        Dim nt_mode As Integer = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(32, 33).Value
        '————————————————————————————————————————————————————————————————————————————————————————  
        '自有流动资金比例
        Dim zyldzjbl_rj As Double = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(11, 23).Value
        Dim zyldzjbl_nt As Double = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(13, 23).Value
        '燃机和暖通总量累计值
        Dim rjfdl_lj As Double = 0
        Dim glgrl_lj As Double = 0
        '燃机和暖通，逐年流动资金金额，列表长度32
        Dim ldzj_rj(32) As Double '流动资金总额
        Dim ldzj_nt(32) As Double '流动资金总额
        Dim zyldzj_rj(32) As Double '自有流动资金金额
        Dim zyldzj_nt(32) As Double '自有流动资金金额
        For i = 1 To 31
            '计算前一年累积量，用前一年的累积量进行计算
            rjfdl_lj += rjfdl_list(i - 1)
            glgrl_lj += glgrl_list(i - 1)
            '计算流动资金金额
            ldzj_rj(i) = rj_mode * rjfdl_lj * ldzjfl_rj_list(i) / 1000
            ldzj_nt(i) = nt_mode * glgrl_lj * ldzjfl_nt_list(i) / 1000
            '计算自有流动资金金额
            zyldzj_rj(i) = rj_mode * ldzj_rj(i) * zyldzjbl_rj
            zyldzj_nt(i) = nt_mode * ldzj_nt(i) * zyldzjbl_nt
        Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '返回结果
        Dim ans(3)
        ans(0) = ldzj_rj
        ans(1) = ldzj_nt
        ans(2) = zyldzj_rj
        ans(3) = zyldzj_nt
        Return ans
    End Function
    Function 默认逐年流动资金费率(ExcelApp As Object)
        '读取Excel中输入的默认流动资金费率
        '光伏、风电、蓄电池流动资金直接根据费率进行计算
        Dim ldzjfl_gf As Double = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(4, 23).Value
        Dim ldzjfl_fd As Double = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(6, 23).Value
        Dim ldzjfl_xdc As Double = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(8, 23).Value
        '燃机和暖通流动资金费率
        Dim ldzjfl_rj As Double = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(10, 23).Value
        Dim ldzjfl_nt As Double = ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(12, 23).Value
        '逐年费率
        Dim ldzjfl_gf_list(31) As Double
        Dim ldzjfl_fd_list(31) As Double
        Dim ldzjfl_xdc_list(31) As Double
        Dim ldzjfl_rj_list(31) As Double
        Dim ldzjfl_nt_list(31) As Double
        For i = 1 To 31
            ldzjfl_gf_list(i) = ldzjfl_gf
            ldzjfl_fd_list(i) = ldzjfl_fd
            ldzjfl_xdc_list(i) = ldzjfl_xdc
            ldzjfl_rj_list(i) = ldzjfl_rj
            ldzjfl_nt_list(i) = ldzjfl_nt
        Next
        '返回结果
        Dim ans(4)
        ans(0) = ldzjfl_gf_list
        ans(1) = ldzjfl_fd_list
        ans(2) = ldzjfl_xdc_list
        ans(3) = ldzjfl_rj_list
        ans(4) = ldzjfl_nt_list
        Return ans
    End Function
    Function 流动资金计算基数扣除默认设置(ExcelApp As Object)
        'yynx_xdc：蓄电池每次投资运营年限数量（年）
        'yynx_gf：光伏每次投资运营年限数量（年）
        'yynx_fd：风电每次投资运营年限数量（年）
        'kcbl_xdc：蓄电池每次投资运营年限结束后，上次投资计算材料费其它费的扣除比例（%）
        'kcbl_gf：光伏每次投资运营年限结束后，上次投资计算材料费其它费的扣除比例（%）
        'kcbl_fd：风电每次投资运营年限结束后，上次投资计算材料费其它费的扣除比例（%）

        '读取计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        Dim yynx_rj As Integer = jsnx - 1
        Dim yynx_xdc As Integer = 10
        Dim yynx_nt As Integer = jsnx - 1
        Dim yynx_gf As Integer = jsnx - 1
        Dim yynx_fd As Integer = jsnx - 1
        Dim kcbl_rj As Double = 0
        Dim kcbl_xdc As Double = 1
        Dim kcbl_nt As Double = 0
        Dim kcbl_gf As Double = 0
        Dim kcbl_fd As Double = 0

        '返回结果
        Dim ans(9)
        ans(0) = yynx_rj
        ans(1) = yynx_xdc
        ans(2) = yynx_nt
        ans(3) = yynx_gf
        ans(4) = yynx_fd
        ans(5) = kcbl_rj
        ans(6) = kcbl_xdc
        ans(7) = kcbl_nt
        ans(8) = kcbl_gf
        ans(9) = kcbl_fd
        Return ans
    End Function

End Module
