Module 计算投资收益率和回收期
    Sub 投资收益率和投资回收期计算(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        ExcelApp.Calculate()
        '自有资金
        '资本金现金流量表
        '列表，所得税前净现值
        Dim jxz_zbj_sq(31) As Double
        '列表，所得税前累计净现值
        Dim jxz_zbj_sq_lj(31) As Double
        '前15年
        For i = 1 To 15
            jxz_zbj_sq(i) = ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(24, 4 + i).Value
            jxz_zbj_sq_lj(i) = ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(25, 4 + i).Value
        Next
        '16-31年 
        For i = 16 To 31
            jxz_zbj_sq(i) = ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(50, i - 12).Value
            jxz_zbj_sq_lj(i) = ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(51, i - 12).Value
        Next
        For i = 1 To 31 '年份1至31年
            '所得税前累计净现金流量
            If jxz_zbj_sq_lj(i) >= 0 Then
                '将所得税前累计净现金流量第一个大于等于0的年份记录在表格中
                ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(45, 22).Value = i
                '将所得税前累计净现金流量第一个大于等于0的年份，之前一年的所得税前累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(49, 22).Value = jxz_zbj_sq_lj(i - 1)
                '将所得税前累计净现金流量第一个大于等于0的年份，当年的所得税前净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(52, 22).Value = jxz_zbj_sq(i)
                Exit For
            End If
        Next
        '列表，所得税后净现值
        Dim jxz_zbj_sh(31) As Double
        '列表，所得税后累计净现值
        Dim jxz_zbj_sh_lj(31) As Double
        '前15年
        For i = 1 To 15
            jxz_zbj_sh(i) = ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(22, 4 + i).Value
            jxz_zbj_sh_lj(i) = ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(23, 4 + i).Value
        Next
        '16-31年 
        For i = 16 To 31
            jxz_zbj_sh(i) = ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(48, i - 12).Value
            jxz_zbj_sh_lj(i) = ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(49, i - 12).Value
        Next
        For i = 1 To 31 '年份1至31年
            '所得税后累计净现金流量
            If jxz_zbj_sh_lj(i) >= 0 Then
                '将所得税后累计净现金流量第一个大于等于0的年份记录在表格中
                ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(46, 22).Value = i
                '将所得税后累计净现金流量第一个大于等于0的年份，之前一年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(50, 22).Value = jxz_zbj_sh_lj(i - 1)
                '将所得税后累计净现金流量第一个大于等于0的年份，当年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(53, 22).Value = jxz_zbj_sh(i)
                Exit For
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '全投资
        '列表，所得税前净现值
        Dim jxz_qtz_sq(31) As Double
        '列表，所得税前累计净现值
        Dim jxz_qtz_sq_lj(31) As Double
        '前15年
        For i = 1 To 15
            jxz_qtz_sq(i) = ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(16, 4 + i).Value
            jxz_qtz_sq_lj(i) = ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(17, 4 + i).Value
        Next
        '16-31年 
        For i = 16 To 31
            jxz_qtz_sq(i) = ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(37, i - 12).Value
            jxz_qtz_sq_lj(i) = ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(38, i - 12).Value
        Next
        '项目投资现金流量表
        For i = 1 To 31 '年份1至31年
            '所得税前累计净现金流量
            If jxz_qtz_sq_lj(i) >= 0 Then
                '将所得税前累计净现金流量第一个大于等于0的年份记录在表格中
                ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(41, 22).Value = i
                '将所得税前累计净现金流量第一个大于等于0的年份，之前一年的所得税前累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(44, 22).Value = jxz_qtz_sq_lj(i - 1)
                '将所得税前累计净现金流量第一个大于等于0的年份，当年的所得税前净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(47, 22).Value = jxz_qtz_sq(i)
                Exit For
            End If
        Next
        '列表，所得税后净现值
        Dim jxz_qtz_sh(31) As Double
        '列表，所得税后累计净现值
        Dim jxz_qtz_sh_lj(31) As Double
        '前15年
        For i = 1 To 15
            jxz_qtz_sh(i) = ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(19, 4 + i).Value
            jxz_qtz_sh_lj(i) = ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(20, 4 + i).Value
        Next
        '16-31年 
        For i = 16 To 31
            jxz_qtz_sh(i) = ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(40, i - 12).Value
            jxz_qtz_sh_lj(i) = ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(41, i - 12).Value
        Next
        For i = 1 To 31 '年份1至31年
            '所得税后累计净现金流量
            If jxz_qtz_sh_lj(i) >= 0 Then
                '将所得税后累计净现金流量第一个大于等于0的年份记录在表格中
                ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(42, 22).Value = i
                '将所得税后累计净现金流量第一个大于等于0的年份，之前一年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(45, 22).Value = jxz_qtz_sh_lj(i - 1)
                '将所得税后累计净现金流量第一个大于等于0的年份，当年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(48, 22).Value = jxz_qtz_sh(i)
                Exit For
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '投资各方所得税后回收年限计算
        '投资方1
        '列表，所得税后净现值
        Dim jxz_tzf1_sh(31) As Double
        '列表，所得税后累计净现值
        Dim jxz_tzf1_sh_lj(31) As Double
        '前15年
        For i = 1 To 15
            jxz_tzf1_sh(i) = ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(15, 4 + i).Value
            jxz_tzf1_sh_lj(i) = ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(16, 4 + i).Value
        Next
        '16-31年 
        For i = 16 To 31
            jxz_tzf1_sh(i) = ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(32, i - 12).Value
            jxz_tzf1_sh_lj(i) = ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(33, i - 12).Value
        Next
        For i = 1 To 31 '年份1至31年
            '所得税后累计净现金流量
            If jxz_tzf1_sh_lj(i) >= 0 Then
                '将所得税后累计净现金流量第一个大于等于0的年份记录在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(30, 22).Value = i
                '将所得税后累计净现金流量第一个大于等于0的年份，之前一年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(33, 22).Value = jxz_tzf1_sh_lj(i - 1)
                '将所得税后累计净现金流量第一个大于等于0的年份，当年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(35, 22).Value = jxz_tzf1_sh(i)
                Exit For
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '投资方2        
        '列表，所得税后净现值
        Dim jxz_tzf2_sh(31) As Double
        '列表，所得税后累计净现值
        Dim jxz_tzf2_sh_lj(31) As Double
        '前15年
        For i = 1 To 15
            jxz_tzf2_sh(i) = ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(15, 4 + i).Value
            jxz_tzf2_sh_lj(i) = ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(16, 4 + i).Value
        Next
        '16-31年 
        For i = 16 To 31
            jxz_tzf2_sh(i) = ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(32, i - 12).Value
            jxz_tzf2_sh_lj(i) = ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(33, i - 12).Value
        Next
        For i = 1 To 31 '年份1至31年
            '所得税后累计净现金流量
            If jxz_tzf2_sh_lj(i) >= 0 Then
                '将所得税后累计净现金流量第一个大于等于0的年份记录在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(30, 22).Value = i
                '将所得税后累计净现金流量第一个大于等于0的年份，之前一年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(33, 22).Value = jxz_tzf2_sh_lj(i - 1)
                '将所得税后累计净现金流量第一个大于等于0的年份，当年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(35, 22).Value = jxz_tzf2_sh(i)
                Exit For
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '投资方3        
        '列表，所得税后净现值
        Dim jxz_tzf3_sh(31) As Double
        '列表，所得税后累计净现值
        Dim jxz_tzf3_sh_lj(31) As Double
        '前15年
        For i = 1 To 15
            jxz_tzf3_sh(i) = ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(15, 4 + i).Value
            jxz_tzf3_sh_lj(i) = ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(16, 4 + i).Value
        Next
        '16-31年 
        For i = 16 To 31
            jxz_tzf3_sh(i) = ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(32, i - 12).Value
            jxz_tzf3_sh_lj(i) = ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(33, i - 12).Value
        Next
        For i = 1 To 31 '年份1至31年
            '所得税后累计净现金流量
            If jxz_tzf3_sh_lj(i) >= 0 Then
                '将所得税后累计净现金流量第一个大于等于0的年份记录在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(30, 22).Value = i
                '将所得税后累计净现金流量第一个大于等于0的年份，之前一年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(33, 22).Value = jxz_tzf3_sh_lj(i - 1)
                '将所得税后累计净现金流量第一个大于等于0的年份，当年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(35, 22).Value = jxz_tzf3_sh(i)
                Exit For
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '投资方4        
        '列表，所得税后净现值
        Dim jxz_tzf4_sh(31) As Double
        '列表，所得税后累计净现值
        Dim jxz_tzf4_sh_lj(31) As Double
        '前15年
        For i = 1 To 15
            jxz_tzf4_sh(i) = ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(15, 4 + i).Value
            jxz_tzf4_sh_lj(i) = ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(16, 4 + i).Value
        Next
        '16-31年 
        For i = 16 To 31
            jxz_tzf4_sh(i) = ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(32, i - 12).Value
            jxz_tzf4_sh_lj(i) = ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(33, i - 12).Value
        Next
        For i = 1 To 31 '年份1至31年
            '所得税后累计净现金流量
            If jxz_tzf4_sh_lj(i) >= 0 Then
                '将所得税后累计净现金流量第一个大于等于0的年份记录在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(30, 22).Value = i
                '将所得税后累计净现金流量第一个大于等于0的年份，之前一年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(33, 22).Value = jxz_tzf4_sh_lj(i - 1)
                '将所得税后累计净现金流量第一个大于等于0的年份，当年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(35, 22).Value = jxz_tzf4_sh(i)
                Exit For
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '投资方5       
        '列表，所得税后净现值
        Dim jxz_tzf5_sh(31) As Double
        '列表，所得税后累计净现值
        Dim jxz_tzf5_sh_lj(31) As Double
        '前15年
        For i = 1 To 15
            jxz_tzf5_sh(i) = ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(15, 4 + i).Value
            jxz_tzf5_sh_lj(i) = ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(16, 4 + i).Value
        Next
        '16-31年 
        For i = 16 To 31
            jxz_tzf5_sh(i) = ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(32, i - 12).Value
            jxz_tzf5_sh_lj(i) = ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(33, i - 12).Value
        Next
        For i = 1 To 31 '年份1至31年
            '所得税后累计净现金流量
            If jxz_tzf5_sh_lj(i) >= 0 Then
                '将所得税后累计净现金流量第一个大于等于0的年份记录在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(30, 22).Value = i
                '将所得税后累计净现金流量第一个大于等于0的年份，之前一年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(33, 22).Value = jxz_tzf5_sh_lj(i - 1)
                '将所得税后累计净现金流量第一个大于等于0的年份，当年的所得税后累计净现金流量储存在表格中
                ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(35, 22).Value = jxz_tzf5_sh(i)
                Exit For
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        '————————————————————————————————————————————————————————————————————————————————————————              
        '针对投资回收年限添加报错功能：如果累计净现金流量出现了先从负到正然后又变成负的情况
        '列表，储存所得税后累计净现值
        Dim ljjxz_sq_zbj(31) As Double '资本金所得税前累计净现值
        Dim ljjxz_sq_qtz(31) As Double '全投资所得税前累计净现值
        Dim ljjxz_sh_zbj(31) As Double '资本金所得税后累计净现值
        Dim ljjxz_sh_qtz(31) As Double '全投资所得税后累计净现值
        '前15年
        For i = 1 To 15
            '资本金所得税后累计净现值
            ljjxz_sh_zbj(i) = ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(23, 4 + i).Value
            '资本金所得税前累计净现值
            ljjxz_sq_zbj(i) = ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(25, 4 + i).Value
            '全投资所得税后累计净现值
            ljjxz_sh_qtz(i) = ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(20, 4 + i).Value
            '全投资所得税前累计净现值
            ljjxz_sq_qtz(i) = ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(17, 4 + i).Value
        Next
        '16-31年
        For i = 16 To 31
            '资本金所得税后累计净现值
            ljjxz_sh_zbj(i) = ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(49, i - 12).Value
            '资本金所得税前累计净现值
            ljjxz_sq_zbj(i) = ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(51, i - 12).Value
            '全投资所得税后累计净现值
            ljjxz_sh_qtz(i) = ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(41, i - 12).Value
            '全投资所得税前累计净现值
            ljjxz_sq_qtz(i) = ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(38, i - 12).Value
        Next
        '资本金所得税后状态检测
        Dim ZT_ZBJSDSH As Integer = 0
        '全投资所得税后状态检测
        Dim ZT_QTZSDSH As Integer = 0
        '资本金所得税前状态检测
        Dim ZT_ZBJSDSQ As Integer = 0
        '全投资所得税前状态检测
        Dim ZT_QTZSDSQ As Integer = 0
        '查找前一年为负，但是当年为正；或者当一年为正，后一年为负的情况
        For i = 1 To 30
            '资本金所得税后累计净现值
            If (ljjxz_sh_zbj(i - 1) < 0 And ljjxz_sh_zbj(i) > 0) Or (ljjxz_sh_zbj(i) > 0 And ljjxz_sh_zbj(i + 1) < 0) Then
                ZT_ZBJSDSH += 1
            End If
            '资本金所得税前累计净现值
            If (ljjxz_sq_zbj(i - 1) < 0 And ljjxz_sq_zbj(i) > 0) Or (ljjxz_sq_zbj(i) > 0 And ljjxz_sq_zbj(i + 1) < 0) Then
                ZT_ZBJSDSQ += 1
            End If
            '全投资所得税后累计净现值
            If (ljjxz_sh_qtz(i - 1) < 0 And ljjxz_sh_qtz(i) > 0) Or (ljjxz_sh_qtz(i) > 0 And ljjxz_sh_qtz(i + 1) < 0) Then
                ZT_QTZSDSH += 1
            End If
            '全投资所得税前累计净现值
            If (ljjxz_sq_qtz(i - 1) < 0 And ljjxz_sq_qtz(i) > 0) Or (ljjxz_sq_qtz(i) > 0 And ljjxz_sq_qtz(i + 1) < 0) Then
                ZT_QTZSDSQ += 1
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————  
        '判断结果写入Excel
        '写入结果
        If ZT_QTZSDSH > 1 Or ZT_QTZSDSQ > 1 Or ZT_ZBJSDSH > 1 Or ZT_ZBJSDSQ > 1 Then
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(14, 3).Value = "不正确"
        Else
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(14, 3).Value = "正确"
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '内部收益率计算
        '读取基本收益率
        Dim guess_zbj As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 3).Value
        Dim guess_qtz As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 3).Value
        Dim guess_tzf As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 3).Value
        '资本金所得税前
        '读取逐年现金流
        Dim xjl_zbjsdsq_list(31) As Double
        '前15年
        For i = 1 To 15
            xjl_zbjsdsq_list(i) = ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(24, i + 4).Value
        Next
        '16-31年
        For i = 16 To 31
            xjl_zbjsdsq_list(i) = ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(50, i - 12).Value
        Next
        '计算收益率
        Dim syl_zbjsdsq As Double = IRR(xjl_zbjsdsq_list, guess_zbj)
        '写入Excel
        ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(52, 4).Value = syl_zbjsdsq
        '————————————————————————————————————————————————————————————————————————————————————————
        '资本金所得税后
        '读取逐年现金流
        Dim xjl_zbjsdsh_list(31) As Double
        '前15年
        For i = 1 To 15
            xjl_zbjsdsh_list(i) = ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(22, i + 4).Value
        Next
        '16-31年
        For i = 16 To 31
            xjl_zbjsdsh_list(i) = ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(48, i - 12).Value
        Next
        '计算收益率
        Dim syl_zbjsdsh As Double = IRR(xjl_zbjsdsh_list, guess_zbj)
        '写入Excel
        ExcelApp.ThisWorkbook.Worksheets("资本金现金流量表").Cells(53, 4).Value = syl_zbjsdsh
        '————————————————————————————————————————————————————————————————————————————————————————
        '全投资所得税前
        '读取逐年现金流
        Dim xjl_qtzsdsq_list(31) As Double
        '前15年
        For i = 1 To 15
            xjl_qtzsdsq_list(i) = ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(16, i + 4).Value
        Next
        '16-31年
        For i = 16 To 31
            xjl_qtzsdsq_list(i) = ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(37, i - 12).Value
        Next
        '计算收益率
        Dim syl_qtzsdsq As Double = IRR(xjl_qtzsdsq_list, guess_qtz)
        '写入Excel
        ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(42, 4).Value = syl_qtzsdsq
        '————————————————————————————————————————————————————————————————————————————————————————
        '全投资所得税后
        '读取逐年现金流
        Dim xjl_qtzsdsh_list(31) As Double
        '前15年
        For i = 1 To 15
            xjl_qtzsdsh_list(i) = ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(19, i + 4).Value
        Next
        '16-31年
        For i = 16 To 31
            xjl_qtzsdsh_list(i) = ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(40, i - 12).Value
        Next
        '计算收益率
        Dim syl_qtzsdsh As Double = IRR(xjl_qtzsdsh_list, guess_qtz)
        '写入Excel
        ExcelApp.ThisWorkbook.Worksheets("项目投资现金流量表").Cells(43, 4).Value = syl_qtzsdsh
        '————————————————————————————————————————————————————————————————————————————————————————
        '投资方1
        '读取逐年现金流
        Dim xjl_tzf1_list(31) As Double
        '前15年
        For i = 1 To 15
            xjl_tzf1_list(i) = ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(15, i + 4).Value
        Next
        '16-31年
        For i = 16 To 31
            xjl_tzf1_list(i) = ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(32, i - 12).Value
        Next
        '计算收益率
        Dim syl_tzf1 As Double = IRR(xjl_tzf1_list, guess_tzf)
        '写入Excel
        ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(35, 4).Value = syl_tzf1
        '————————————————————————————————————————————————————————————————————————————————————————
        '投资方2
        '读取逐年现金流
        Dim xjl_tzf2_list(31) As Double
        '前15年
        For i = 1 To 15
            xjl_tzf2_list(i) = ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(15, i + 4).Value
        Next
        '16-31年
        For i = 16 To 31
            xjl_tzf2_list(i) = ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(32, i - 12).Value
        Next
        '计算收益率
        Dim syl_tzf2 As Double = IRR(xjl_tzf2_list, guess_tzf)
        '写入Excel
        ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(35, 4).Value = syl_tzf2
        '————————————————————————————————————————————————————————————————————————————————————————
        '投资方3
        '读取逐年现金流
        Dim xjl_tzf3_list(31) As Double
        '前15年
        For i = 1 To 15
            xjl_tzf3_list(i) = ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(15, i + 4).Value
        Next
        '16-31年
        For i = 16 To 31
            xjl_tzf3_list(i) = ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(32, i - 12).Value
        Next
        '计算收益率
        Dim syl_tzf3 As Double = IRR(xjl_tzf3_list, guess_tzf)
        '写入Excel
        ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(35, 4).Value = syl_tzf3
        '————————————————————————————————————————————————————————————————————————————————————————
        '投资方4
        '读取逐年现金流
        Dim xjl_tzf4_list(31) As Double
        '前15年
        For i = 1 To 15
            xjl_tzf4_list(i) = ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(15, i + 4).Value
        Next
        '16-31年
        For i = 16 To 31
            xjl_tzf4_list(i) = ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(32, i - 12).Value
        Next
        '计算收益率
        Dim syl_tzf4 As Double = IRR(xjl_tzf4_list, guess_tzf)
        '写入Excel
        ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(35, 4).Value = syl_tzf4
        '————————————————————————————————————————————————————————————————————————————————————————
        '投资方5
        '读取逐年现金流
        Dim xjl_tzf5_list(31) As Double
        '前15年
        For i = 1 To 15
            xjl_tzf5_list(i) = ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(15, i + 4).Value
        Next
        '16-31年
        For i = 16 To 31
            xjl_tzf5_list(i) = ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(32, i - 12).Value
        Next
        '计算收益率
        Dim syl_tzf5 As Double = IRR(xjl_tzf5_list, guess_tzf)
        '写入Excel
        ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(35, 4).Value = syl_tzf5
    End Sub
End Module
