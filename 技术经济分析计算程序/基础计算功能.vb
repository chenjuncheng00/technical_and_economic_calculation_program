Module 基础计算功能
    Function 基础计算功能_10_to_31(tznf_list As Array, je_list As Array)
        '将<估算表>中的10次投资情况生成成31年的列表
        '将长度10的列表转为长度31的列表
        'tznf_list：10次投资的年份序号，列表，长度10
        'je_list：10次投资的各种金额数值(静态投资，建设期利息，可抵扣增值税，资本金，建设期贷款)，列表，长度10

        '储存计算结果，长度为31的列表
        Dim ans(31) As Double
        '10次投资发生的年份序号
        For i = 1 To 10
            If tznf_list(i) > 0 Then
                ans(tznf_list(i)) = je_list(i)
            End If
        Next
        '返回结果
        Return ans
    End Function
    Function 基础计算功能_31_to_10(tznf_list As Array, je_list As Array)
        '将长度31的列表转为长度10的列表
        'tznf_list：10次投资的年份序号，列表，长度10
        'je_list：31年各种金额数值(静态投资，建设期利息，可抵扣增值税，资本金，建设期贷款)，列表，长度31

        '储存计算结果，长度为10的列表
        Dim ans(10) As Double
        '遍历10次投资的年份序号
        For i = 1 To 10
            If tznf_list(i) > 0 Then
                ans(i) = je_list(tznf_list(i))
            End If
        Next
        '返回结果
        Return ans
    End Function

    Sub 投资回收期计算(ExcelApp As Object)
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
    End Sub
    Sub 计算前基本处理(ExcelApp As Object, jbcl_model As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '启动迭代计算
        ExcelApp.Application.Iteration = True
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        '屏蔽屏幕更新
        ExcelApp.Application.ScreenUpdating = False
        '屏蔽事件
        ExcelApp.Application.EnableEvents = False
        '不一定每次都需要关闭Excel自动计算
        If jbcl_model = 1 Then
            '手动计算，关闭excel的自动计算
            ExcelApp.Application.Calculation = XlCalculation.xlCalculationManual
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
        '解锁表格
        Call 解锁表格(ExcelApp)
    End Sub
    Sub 计算后基本处理(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
        '打开事件
        ExcelApp.Application.EnableEvents = True
        '打开excel的自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '增值税相关计算
        Call 增值税相关计算(ExcelApp)
        '流动资金计算
        Call 流动资金相关计算(ExcelApp)
        '投资回收期计算
        Call 投资回收期计算(ExcelApp)
        '锁定表格
        Call 锁定表格(ExcelApp)
    End Sub
    Sub 计算还款利润(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '每次计算完折旧摊销和长期贷款后，都要计算一次本SUB
        '当折旧费不够还贷款本金时，用当年利润还

        '读取项目总的计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '长期贷款还本金额
        Dim znhb(31) As Double
        '固定资产折旧+无形资产摊销金额
        Dim zjtx(31) As Double
        '前15年
        For i = 1 To 15
            znhb(i) = ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(9, i + 4).Value
            zjtx(i) = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(22, i + 4).Value
        Next
        '16-31年
        For i = 16 To 31
            znhb(i) = ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(39, i - 12).Value
            zjtx(i) = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(54, i - 12).Value
        Next
        '计算逐年还款利润
        Dim hklr(31) As Double
        For i = 1 To 31
            If znhb(i) - zjtx(i) > 0 And i <= jsnx Then
                hklr(i) = znhb(i) - zjtx(i)
            Else
                hklr(i) = 0
            End If
        Next
        '结果写入Excel
        '前15年
        For i = 1 To 15
            ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(17, i + 4).Value = hklr(i)
        Next
        '16-31年
        For i = 16 To 31
            ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(37, i - 12).Value = hklr(i)
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
    Function Excel版本号验证(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim ZTJC As Integer
        '验证Excel表格的更新时间
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(3, 29).Value < 20210903 Then
            MsgBox（"Excel文件版本已过期，无法配合最新的计算程序使用，需要更新到最新版本才可以使用！本Excel文件仅可以查看已有的计算结果，不可能用于新的计算！"）
            ZTJC = 1
        Else
            ZTJC = 0
        End If
        Call 自保护程序(ExcelApp)
        '返回结果
        Return ZTJC
    End Function
End Module
