Module 修理费计算
    Sub 修理费计算(ExcelApp As Object, xlfl_cg_model As Integer, xlfl_qt_model As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        '————————————————————————————————————————————————————————————————————————————————————————
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
        '————————————————————————————————————————————————————————————————————————————————————————
        '采用的规设备修理费系数
        '常规设备修理费率（%）
        Dim xlfl_cg_list(31) As Double
        '燃机修理费率（%）
        Dim xlfl_rj_list(31) As Double
        '蓄电池修理费率（%）
        Dim xlfl_xdc_list(31) As Double
        '光伏设备修理费率（%）
        Dim xlfl_gf_list(31) As Double
        '暖通设备修理费率（%）
        Dim xlfl_nt_list(31) As Double
        '风电设备修理费率（%）
        Dim xlfl_fd_list(31) As Double
        '实际修理费率
        If xlfl_cg_model = 0 Then
            '常规设备修理费率（%）
            xlfl_cg_list = xlfl_cg_mr_list
        Else
            For i = 3 To 33
                '常规设备修理费率（%）
                xlfl_cg_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value
            Next
        End If
        '实际修理费率
        If xlfl_qt_model = 0 Then
            '燃机修理费率（%）
            xlfl_rj_list = xlfl_rj_mr_list
            '蓄电池修理费率（%）
            xlfl_xdc_list = xlfl_xdc_mr_list
            '光伏设备修理费率（%）
            xlfl_gf_list = xlfl_gf_mr_list
            '暖通设备修理费率（%）
            xlfl_nt_list = xlfl_nt_mr_list
            '风电设备修理费率（%）
            xlfl_fd_list = xlfl_fd_mr_list
        Else
            For i = 3 To 33
                '燃机修理费率（%）
                xlfl_rj_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(154, i).Value
                '蓄电池修理费率（%）
                xlfl_xdc_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(155, i).Value
                '光伏设备修理费率（%）
                xlfl_gf_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(156, i).Value
                '暖通设备修理费率（%）
                xlfl_nt_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(157, i).Value
                '风电设备修理费率（%）
                xlfl_fd_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value
            Next
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '无形资产比例（%）
        Dim wxzcbl_list = xlfl_mr(6)
        '————————————————————————————————————————————————————————————————————————————————————————
        '读取输入数据
        Dim GSBSJ = 读取估算表数据(ExcelApp)
        '项目计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '修理费率
        Dim xlfl As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(9, 5).Value
        '10次投资的数据，用于计算10次投资的固定资产原值
        Dim gdzcyz(10) As Double
        Dim tznf_list = GSBSJ(0)
        Dim dttz_list = GSBSJ(2)
        Dim kdkzzs_list = GSBSJ(9)
        '无形资产所占比例
        Dim wxzcbl As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 5).Value
        '读取开始年份和资产原值，前5年
        For i = 1 To 5
            gdzcyz(i) = (dttz_list(i) - kdkzzs_list(i)) * (1 - wxzcbl)
        Next
        '读取开始年份和资产原值，后5年
        For i = 6 To 10
            gdzcyz(i) = (dttz_list(i) - kdkzzs_list(i)) * (1 - wxzcbl)
        Next
        'jttz_list：逐年静态投资，列表
        Dim jttz = GSBSJ(1)
        Dim jttz_list = 基础计算功能_10_to_31(tznf_list, jttz)
        'gdzcyz_list：逐年固定资产原值，列表
        Dim gdzcyz_list = 基础计算功能_10_to_31(tznf_list, gdzcyz)
        'jsqdklx_list：逐年建设期贷款利息，列表
        Dim jsqdklx = GSBSJ(5)
        Dim jsqdklx_list = 基础计算功能_10_to_31(tznf_list, jsqdklx)
        '燃机总投资(万元)
        Dim rjtz = GSBSJ(10)
        Dim rjtz_list = 基础计算功能_10_to_31(tznf_list, rjtz)
        '蓄电池总投资(万元)
        Dim xdctz = GSBSJ(12)
        Dim xdctz_list = 基础计算功能_10_to_31(tznf_list, xdctz)
        '暖通总投资(万元)
        Dim nttz = GSBSJ(14)
        Dim nttz_list = 基础计算功能_10_to_31(tznf_list, nttz)
        '光伏总投资(万元)
        Dim gftz = GSBSJ(16)
        Dim gftz_list = 基础计算功能_10_to_31(tznf_list, gftz)
        '风电总投资(万元)
        Dim fdtz = GSBSJ(18)
        Dim fdtz_list = 基础计算功能_10_to_31(tznf_list, fdtz)
        'qttz_list：除了常规设备以外的其它设备的逐年投资额，列表
        Dim qttz_list(31) As Double
        For i = 1 To 31
            qttz_list(i) = rjtz_list(i) + xdctz_list(i) + nttz_list(i) + gftz_list(i) + fdtz_list(i)
        Next
        '每年投产月份数
        Dim tcyfs_list(31) As Integer
        '读取每年投产月份数
        For i = 1 To 31
            tcyfs_list(i) = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 2).Value
        Next
        '计算常规设备的逐年修理费率
        Dim ans_cgsb
        '计算其它设备的修理费
        Dim ans_qt(31) As Double
        '修理费总额
        Dim ans_znxlf(31) As Double
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 11).Value = "方法一" Then
            ans_cgsb = 设备修理费计算_常规方法(xlfl_cg_list, gdzcyz_list, jsqdklx_list, qttz_list, tcyfs_list, wxzcbl_list, True)
            '其它投资累计额
            Dim rjtz_lj = 0
            Dim xdctz_lj = 0
            Dim gftz_lj = 0
            Dim nttz_lj = 0
            Dim fdtz_lj = 0
            For i = 1 To 31
                '前一年其它投资累计额
                rjtz_lj += rjtz_list(i - 1)
                xdctz_lj += xdctz_list(i - 1)
                gftz_lj += gftz_list(i - 1)
                nttz_lj += nttz_list(i - 1)
                fdtz_lj += fdtz_list(i - 1)
                ans_qt(i) = (rjtz_lj * xlfl_rj_list(i) + xdctz_lj * xlfl_xdc_list(i) + gftz_lj * xlfl_gf_list(i) + nttz_lj * xlfl_nt_list(i) + fdtz_lj * xlfl_fd_list(i)) * (tcyfs_list(i) / 12)
                ans_znxlf(i) = ans_cgsb(i) + ans_qt(i)
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 11).Value = "方法二" Then
            ans_cgsb = 设备修理费计算_常规方法(xlfl_cg_list, gdzcyz_list, jsqdklx_list, qttz_list, tcyfs_list, wxzcbl_list, False)
            '其它投资累计额
            Dim rjtz_lj = 0
            Dim xdctz_lj = 0
            Dim gftz_lj = 0
            Dim nttz_lj = 0
            Dim fdtz_lj = 0
            For i = 1 To 31
                '前一年其它投资累计额
                rjtz_lj += rjtz_list(i - 1)
                xdctz_lj += xdctz_list(i - 1)
                gftz_lj += gftz_list(i - 1)
                nttz_lj += nttz_list(i - 1)
                fdtz_lj += fdtz_list(i - 1)
                ans_qt(i) = rjtz_lj * xlfl_rj_list(i) + xdctz_lj * xlfl_xdc_list(i) + gftz_lj * xlfl_gf_list(i) + nttz_lj * xlfl_nt_list(i) + fdtz_lj * xlfl_fd_list(i)
                ans_znxlf(i) = ans_cgsb(i) + ans_qt(i)
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 11).Value = "方法三" Then
            ans_cgsb = 设备修理费计算_博微方法(xlfl, jttz_list, gdzcyz_list, jsqdklx_list, qttz_list, tcyfs_list, wxzcbl, False)
            '其它投资累计额
            Dim rjtz_lj = 0
            Dim xdctz_lj = 0
            Dim gftz_lj = 0
            Dim nttz_lj = 0
            Dim fdtz_lj = 0
            For i = 1 To 31
                '前一年其它投资累计额
                rjtz_lj += rjtz_list(i - 1)
                xdctz_lj += xdctz_list(i - 1)
                gftz_lj += gftz_list(i - 1)
                nttz_lj += nttz_list(i - 1)
                fdtz_lj += fdtz_list(i - 1)
                ans_qt(i) = rjtz_lj * xlfl_rj_list(i) + xdctz_lj * xlfl_xdc_list(i) + gftz_lj * xlfl_gf_list(i) + nttz_lj * xlfl_nt_list(i) + fdtz_lj * xlfl_fd_list(i)
                ans_znxlf(i) = ans_cgsb(i) + ans_qt(i)
            Next
        End If
        '结果写入Excel
        '前15年
        For i = 1 To 15
            If i <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(137, 4 + i).Value = ans_znxlf(i)
            Else
                ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(137, 4 + i).Value = 0
            End If
        Next
        '16—31年
        For i = 16 To 31
            If i <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(141, i - 12).Value = ans_znxlf(i)
            Else
                ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(141, i - 12).Value = 0
            End If
        Next
        '写入设备修理费系数，供其它程序调用
        For i = 3 To 33
            If i - 2 <= jsnx Then
                '常规设备修理费率（%）
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value = xlfl_cg_list(i - 2)
                '燃机修理费率（%）
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(154, i).Value = xlfl_rj_list(i - 2)
                '蓄电池修理费率（%）
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(155, i).Value = xlfl_xdc_list(i - 2)
                '光伏设备修理费率（%）
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(156, i).Value = xlfl_gf_list(i - 2)
                '暖通设备修理费率（%）
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(157, i).Value = xlfl_nt_list(i - 2)
                '风电设备修理费率（%）
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value = xlfl_fd_list(i - 2)
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
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
    Function 设备修理费计算_常规方法(xlfl_list As Array, gdzcyz_list As Array, jsqdklx_list As Array, qttz_list As Array,
                                     tcyfs_list As Array, wxzcbl_list As Array, sfzs As Boolean)
        'xlfl_list：逐年修理费率，列表，长度31
        'gdzcyz_list：逐年固定资产原值，列表，长度31
        'jsqdklx_list：逐年建设期贷款利息，列表，长度31
        'qttz_list：除了常规设备以外的其它设备的逐年投资额，列表，长度31
        'tcyfs_list：逐年投产的月份数，列表，长度31
        'wxzcbl_list：无形资产占总资产比例，列表，长度31
        'sfzs：修理费是否按照(逐年投产的月份数/12)进行折算

        '计算累计固定资产原值和累计建设期贷款利息
        Dim gdzcyz_lj As Double = 0
        Dim jsqdklx_lj As Double = 0
        '计算累计其它部分总投资额
        Dim qttz_lj As Double = 0
        '计算出的逐年修理费金额，列表
        Dim ans_znxlf_list(31) As Double
        For i = 1 To 31
            gdzcyz_lj += gdzcyz_list(i - 1)
            jsqdklx_lj += jsqdklx_list(i - 1)
            qttz_lj += qttz_list(i - 1)
            If sfzs = True Then
                ans_znxlf_list(i) = (gdzcyz_lj - jsqdklx_lj * (1 - wxzcbl_list(i)) - qttz_lj) * xlfl_list(i) *
                                    (tcyfs_list(i) / 12)
            Else
                ans_znxlf_list(i) = (gdzcyz_lj - jsqdklx_lj * (1 - wxzcbl_list(i)) - qttz_lj) * xlfl_list(i)
            End If
        Next
        '返回结果
        Return ans_znxlf_list
    End Function
    Function 设备修理费计算_博微方法(xlfl As Double, jttz_list As Array, gdzcyz_list As Array, jsqdklx_list As Array,
                                     qttz_list As Array, tcyfs_list As Array, wxzcbl As Double, sfzs As Boolean)
        'xlfl：修理费率
        'jttz_list：逐年静态投资金额，列表，长度31
        'gdzcyz_list：逐年固定资产原值，列表，长度31
        'jsqdklx_list：逐年建设期贷款利息，列表，长度31
        'qttz_list：除了常规设备以外的其它设备的逐年投资额，列表，长度31
        'tcyfs_list：逐年投产的月份数，列表，长度31
        'wxzcbl：无形资产占总资产比例
        'sfzs：修理费是否按照(逐年投产的月份数/12)进行折算

        '计算累计固定资产原值和累计建设期贷款利息
        Dim gdzcyz_lj As Double = 0
        Dim jsqdklx_lj As Double = 0
        '计算累计其它部分总投资额
        Dim qttz_lj As Double = 0
        '计算累计静态投资总额
        Dim jttz_lj As Double = 0
        For i = 1 To 31
            gdzcyz_lj += gdzcyz_list(i)
            jsqdklx_lj += jsqdklx_list(i)
            qttz_lj += qttz_list(i)
            jttz_lj += jttz_list(i)
        Next
        '计算修理费最大值
        Dim xlf_max As Double = (gdzcyz_lj - jsqdklx_lj * (1 - wxzcbl) - qttz_lj) * xlfl
        '计算出的逐年修理费金额，列表
        '根据前一年累计出资比例折算修理费
        Dim ans_znxlf_list(31) As Double
        '前一年的累计静态投资总额
        Dim jttz_lj_1 As Double = 0
        For i = 1 To 31
            jttz_lj_1 += jttz_list(i - 1)
            If sfzs = True Then
                ans_znxlf_list(i) = (jttz_lj_1 / jttz_lj) * xlf_max * (tcyfs_list(i) / 12)
            Else
                ans_znxlf_list(i) = (jttz_lj_1 / jttz_lj) * xlf_max
            End If
        Next
        '返回结果
        Return ans_znxlf_list
    End Function
    Function 默认逐年修理费率(ExcelApp As Object)
        '读取Excel中输入的默认修理费
        '常规设备修理费率（%）
        Dim xlf_set_cg As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(135, 22).Value
        '燃机修理费率（%）
        Dim xlf_set_rj As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(136, 22).Value
        '蓄电池修理费率（%）
        Dim xlf_set_xdc As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(137, 22).Value
        '光伏设备修理费率（%）
        Dim xlf_set_gf As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(138, 22).Value
        '暖通设备修理费率（%）
        Dim xlf_set_nt As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(139, 22).Value
        '无形资产比例（%）
        Dim wxzcbl_set As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 5).Value
        Dim wxzcbl_list(31) As Double
        For i = 1 To 31
            wxzcbl_list(i) = wxzcbl_set
        Next
        '前5项逐年修理费率默认值，每年都一样
        Dim xlfl_cg_list(31)
        Dim xlfl_rj_list(31)
        Dim xlfl_xdc_list(31)
        Dim xlfl_gf_list(31)
        Dim xlfl_nt_list(31)
        '风电设备修理费率（%），逐年变化
        Dim xlfl_fd_list(31)
        For i = 1 To 31
            xlfl_cg_list(i) = xlf_set_cg
            xlfl_rj_list(i) = xlf_set_rj
            xlfl_xdc_list(i) = xlf_set_xdc
            xlfl_gf_list(i) = xlf_set_gf
            xlfl_nt_list(i) = xlf_set_nt
            '风电设备修理费率
            If i <= 6 Then
                '1-6：0.5%
                xlfl_fd_list(i) = 0.5 / 100
            ElseIf i >= 7 And i <= 11 Then
                '7-11：1%
                xlfl_fd_list(i) = 1 / 100
            ElseIf i >= 12 And i <= 16 Then
                '12-16：1.5%
                xlfl_fd_list(i) = 1.5 / 100
            ElseIf i >= 17 And i <= 21 Then
                '17-21：2%
                xlfl_fd_list(i) = 2 / 100
            ElseIf i >= 22 And i <= 26 Then
                '22-26：2.5%
                xlfl_fd_list(i) = 2.5 / 100
            End If
        Next
        '返回结果
        Dim ans(6)
        ans(0) = xlfl_cg_list
        ans(1) = xlfl_rj_list
        ans(2) = xlfl_xdc_list
        ans(3) = xlfl_gf_list
        ans(4) = xlfl_nt_list
        ans(5) = xlfl_fd_list
        ans(6) = wxzcbl_list
        Return ans
    End Function
End Module
