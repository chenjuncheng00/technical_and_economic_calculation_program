Module 修理费计算
    Sub 修理费计算(ExcelApp As Object, xlfl_cg_model As Integer, xlfl_qt_model As Integer, kcje_xlf_model As Integer)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        '————————————————————————————————————————————————————————————————————————————————————————
        '项目计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '计算逐年分项修理费金额
        Dim ans_xlf = 分项逐年修理费计算(ExcelApp, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model)
        Dim ans_znxlf = ans_xlf(0)
        '————————————————————————————————————————————————————————————————————————————————————————
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
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub

    Function 分项逐年修理费计算(ExcelApp As Object, xlfl_cg_model As Integer, xlfl_qt_model As Integer, kcje_xlf_model As Integer)
        'On Error Resume Next
        '只计算出分项逐年修理费的金额数值，不写入EXCEL
        '————————————————————————————————————————————————————————————————————————————————————————        
        'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        '————————————————————————————————————————————————————————————————————————————————————————
        '修理费率默认值
        Dim xlfl_mr = 默认逐年修理费率(ExcelApp)
        '常规设备修理费率（%）
        Dim xlfl_cg_mr_list = xlfl_mr(0)
        '燃机修理费率（%）
        Dim xlfl_rj_mr_list = xlfl_mr(1)
        '蓄电池修理费率（%）/(元/kW)
        Dim xlfl_xdc_mr_list = xlfl_mr(2)
        '光伏设备修理费率（%）/(元/kW)
        Dim xlfl_gf_mr_list = xlfl_mr(3)
        '暖通设备修理费率（%）
        Dim xlfl_nt_mr_list = xlfl_mr(4)
        '风电设备修理费率（%）/(元/kW)
        Dim xlfl_fd_mr_list = xlfl_mr(5)
        '————————————————————————————————————————————————————————————————————————————————————————
        '采用的设备修理费系数
        '常规设备修理费率（%）
        Dim xlfl_cg_list(31) As Double
        '燃机修理费率（%）
        Dim xlfl_rj_list(31) As Double
        '蓄电池修理费率（%）/(元/kW)
        Dim xlfl_xdc_list(31) As Double
        '光伏设备修理费率（%）/(元/kW)
        Dim xlfl_gf_list(31) As Double
        '暖通设备修理费率（%）
        Dim xlfl_nt_list(31) As Double
        '风电设备修理费率（%）/(元/kW)
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
            '蓄电池修理费率（%）/(元/kW)
            xlfl_xdc_list = xlfl_xdc_mr_list
            '光伏设备修理费率（%）/(元/kW)
            xlfl_gf_list = xlfl_gf_mr_list
            '暖通设备修理费率（%）
            xlfl_nt_list = xlfl_nt_mr_list
            '风电设备修理费率（%）/(元/kW)
            xlfl_fd_list = xlfl_fd_mr_list
        Else
            For i = 3 To 33
                '燃机修理费率（%）
                xlfl_rj_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(154, i).Value
                '蓄电池修理费率（%）/(元/kW)
                xlfl_xdc_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(155, i).Value
                '光伏设备修理费率（%）/(元/kW)
                xlfl_gf_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(156, i).Value
                '暖通设备修理费率（%）
                xlfl_nt_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(157, i).Value
                '风电设备修理费率（%）/(元/kW)
                xlfl_fd_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value
            Next
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '无形资产比例（%）
        Dim wxzcbl_list = xlfl_mr(6)
        '————————————————————————————————————————————————————————————————————————————————————————
        '其它设备计算基数扣除计算模式
        Dim yynx_rj As Integer
        Dim yynx_xdc As Integer
        Dim yynx_nt As Integer
        Dim yynx_gf As Integer
        Dim yynx_fd As Integer
        Dim yynx_cg As Integer
        Dim kcbl_rj As Double
        Dim kcbl_xdc As Double
        Dim kcbl_nt As Double
        Dim kcbl_gf As Double
        Dim kcbl_fd As Double
        Dim kcbl_cg As Double
        '读取默认值
        Dim kcje_mr = 设备修理费计算基数扣除默认设置(ExcelApp)
        If kcje_xlf_model = 0 Then
            yynx_rj = kcje_mr(0)
            yynx_xdc = kcje_mr(1)
            yynx_nt = kcje_mr(2)
            yynx_gf = kcje_mr(3)
            yynx_fd = kcje_mr(4)
            yynx_cg = kcje_mr(5)
            kcbl_rj = kcje_mr(6)
            kcbl_xdc = kcje_mr(7)
            kcbl_nt = kcje_mr(8)
            kcbl_gf = kcje_mr(9)
            kcbl_fd = kcje_mr(10)
            kcbl_cg = kcje_mr(11)
        Else
            yynx_rj = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(103, 3).Value
            yynx_xdc = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(103, 4).Value
            yynx_nt = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(103, 5).Value
            yynx_gf = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(103, 6).Value
            yynx_fd = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(103, 7).Value
            yynx_cg = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(103, 8).Value
            kcbl_rj = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(104, 3).Value
            kcbl_xdc = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(104, 4).Value
            kcbl_nt = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(104, 5).Value
            kcbl_gf = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(104, 6).Value
            kcbl_fd = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(104, 7).Value
            kcbl_cg = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(104, 8).Value
        End If
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
        '————————————————————————————————————————————————————————————————————————————————————————
        '每年投产月份数
        Dim tcyfs_list(31) As Integer
        '读取每年投产月份数
        For i = 1 To 31
            tcyfs_list(i) = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 2).Value
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
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
        '蓄电池装机(kW)
        Dim xdczj = GSBSJ(13)
        Dim xdczj_list = 基础计算功能_10_to_31(tznf_list, xdczj)
        '暖通总投资(万元)
        Dim nttz = GSBSJ(14)
        Dim nttz_list = 基础计算功能_10_to_31(tznf_list, nttz)
        '光伏总投资(万元)
        Dim gftz = GSBSJ(16)
        Dim gftz_list = 基础计算功能_10_to_31(tznf_list, gftz)
        '光伏装机功率(kW)
        Dim gfzj = GSBSJ(17)
        Dim gfzj_list = 基础计算功能_10_to_31(tznf_list, gfzj)
        '风电总投资(万元)
        Dim fdtz = GSBSJ(18)
        Dim fdtz_list = 基础计算功能_10_to_31(tznf_list, fdtz)
        '风电装机功率(kW)
        Dim fdzj = GSBSJ(19)
        Dim fdzj_list = 基础计算功能_10_to_31(tznf_list, fdzj)
        '分项投资之和不可以超过总投资
        For i = 1 To 10
            If rjtz(i) + xdctz(i) + nttz(i) + gftz(i) + fdtz(i) > jttz(i) Then
                MsgBox("每一年的分项投资之和不可以超过当年的总静态投资金额，修理费计算终止！")
                Exit Function
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '10次投资其它投资金额占当年静态投资金额的比例，长度10的列表
        Dim qttzbl(10) As Double
        For i = 1 To 10
            If jttz(i) > 0 Then
                qttzbl(i) = (rjtz(i) + xdctz(i) + nttz(i) + gftz(i) + fdtz(i)) / jttz(i)
            Else
                qttzbl(i) = 0
            End If
        Next
        '其它投资比例转为31年的列表
        Dim qttzbl_list = 基础计算功能_10_to_31(tznf_list, qttzbl)
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算常规设备的逐年修理费计算基数：相当于计算常规设备投资，按比例折算
        Dim cgtz(10) As Double '10次投资
        '计算常除了常规设备以外的其它设备的逐年投资额，按比例折算
        Dim qttz(10) As Double
        For i = 1 To 10
            cgtz(i) = (gdzcyz(i) - jsqdklx(i) * (1 - wxzcbl)) * (1 - qttzbl(i))
            qttz(i) = (gdzcyz(i) - jsqdklx(i) * (1 - wxzcbl)) * qttzbl(i)
        Next
        Dim cgtz_list = 基础计算功能_10_to_31(tznf_list, cgtz)
        Dim qttz_list = 基础计算功能_10_to_31(tznf_list, qttz)
        '————————————————————————————————————————————————————————————————————————————————————————
        '根据输入的运营年限和扣除比例，计算实际的逐年修理费率、逐年投资额度
        '保持计算用的基数不变，但是折算逐年计算费率
        '燃机
        Dim ans_xlf_rj = 计算费率修正计算基础功能(tznf_list, rjtz, rjtz_list, xlfl_rj_list, yynx_rj, kcbl_rj)
        xlfl_rj_list = ans_xlf_rj(0)
        Dim rjtz_lj_list = ans_xlf_rj(2)
        '暖通
        Dim ans_xlf_nt = 计算费率修正计算基础功能(tznf_list, nttz, nttz_list, xlfl_nt_list, yynx_nt, kcbl_nt)
        xlfl_nt_list = ans_xlf_nt(0)
        Dim nttz_lj_list = ans_xlf_nt(2)
        '光伏
        Dim ans_xlf_gf
        Dim gftz_lj_list
        Dim gfzj_lj_list
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(56, 18).Value = "投资额百分比(%)" Then
            ans_xlf_gf = 计算费率修正计算基础功能(tznf_list, gftz, gftz_list, xlfl_gf_list, yynx_gf, kcbl_gf)
            xlfl_gf_list = ans_xlf_gf(0)
            gftz_lj_list = ans_xlf_gf(2)
        ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(56, 18).Value = "装机功率(元/kW)" Then
            ans_xlf_gf = 计算费率修正计算基础功能(tznf_list, gfzj, gfzj_list, xlfl_gf_list, yynx_gf, kcbl_gf)
            xlfl_gf_list = ans_xlf_gf(0)
            gfzj_lj_list = ans_xlf_gf(2)
        End If
        '风电
        Dim ans_xlf_fd
        Dim fdtz_lj_list
        Dim fdzj_lj_list
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(57, 18).Value = "投资额百分比(%)" Then
            ans_xlf_fd = 计算费率修正计算基础功能(tznf_list, fdtz, fdtz_list, xlfl_fd_list, yynx_fd, kcbl_fd)
            xlfl_fd_list = ans_xlf_fd(0)
            fdtz_lj_list = ans_xlf_fd(2)
        ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(57, 18).Value = "装机功率(元/kW)" Then
            ans_xlf_fd = 计算费率修正计算基础功能(tznf_list, fdzj, fdzj_list, xlfl_fd_list, yynx_fd, kcbl_fd)
            xlfl_fd_list = ans_xlf_fd(0)
            fdzj_lj_list = ans_xlf_fd(2)
        End If
        '蓄电池
        Dim ans_xlf_xdc
        Dim xdctz_lj_list
        Dim xdczj_lj_list
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(58, 18).Value = "投资额百分比(%)" Then
            ans_xlf_xdc = 计算费率修正计算基础功能(tznf_list, xdctz, xdctz_list, xlfl_xdc_list, yynx_xdc, kcbl_xdc)
            xlfl_xdc_list = ans_xlf_xdc(0)
            xdctz_lj_list = ans_xlf_xdc(2)
        ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(58, 18).Value = "装机功率(元/kW)" Then
            ans_xlf_xdc = 计算费率修正计算基础功能(tznf_list, xdczj, xdczj_list, xlfl_xdc_list, yynx_xdc, kcbl_xdc)
            xlfl_xdc_list = ans_xlf_xdc(0)
            xdczj_lj_list = ans_xlf_xdc(2)
        End If
        '常规设备
        Dim ans_xlf_cg = 计算费率修正计算基础功能(tznf_list, cgtz, cgtz_list, xlfl_cg_list, yynx_cg, kcbl_cg)
        xlfl_cg_list = ans_xlf_cg(0)
        Dim cgtz_lj_list = ans_xlf_cg(2)
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算常规设备的逐年修理费率
        Dim ans_cgsb
        '燃机、暖通、光伏、风电、蓄电池
        Dim ans_rj(31) As Double
        Dim ans_nt(31) As Double
        Dim ans_gf(31) As Double
        Dim ans_fd(31) As Double
        Dim ans_xdc(31) As Double
        '修理费总额
        Dim ans_znxlf(31) As Double
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 11).Value = "方法一" Then
            '常规设备
            ans_cgsb = 设备修理费计算_常规方法(xlfl_cg_list, gdzcyz_list, jsqdklx_list, qttz_list, tcyfs_list, wxzcbl_list, True)
            '燃机、暖通
            For i = 1 To 31
                ans_rj(i) = rjtz_lj_list(i) * xlfl_rj_list(i) * (tcyfs_list(i) / 12)
                ans_nt(i) = nttz_lj_list(i) * xlfl_nt_list(i) * (tcyfs_list(i) / 12)
            Next
            '光伏
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(56, 18).Value = "投资额百分比(%)" Then
                For i = 1 To 31
                    ans_gf(i) = gftz_lj_list(i) * xlfl_gf_list(i) * (tcyfs_list(i) / 12)
                Next
            ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(56, 18).Value = "装机功率(元/kW)" Then
                For i = 1 To 31
                    ans_gf(i) = gfzj_lj_list(i) * xlfl_gf_list(i) * (tcyfs_list(i) / 12) / 10000
                Next
            End If
            '风电
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(57, 18).Value = "投资额百分比(%)" Then
                For i = 1 To 31
                    ans_fd(i) = fdtz_lj_list(i) * xlfl_fd_list(i) * (tcyfs_list(i) / 12)
                Next
            ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(57, 18).Value = "装机功率(元/kW)" Then
                For i = 1 To 31
                    ans_fd(i) = fdzj_lj_list(i) * xlfl_fd_list(i) * (tcyfs_list(i) / 12) / 10000
                Next
            End If
            '蓄电池
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(58, 18).Value = "投资额百分比(%)" Then
                For i = 1 To 31
                    ans_xdc(i) = xdctz_lj_list(i) * xlfl_xdc_list(i) * (tcyfs_list(i) / 12)
                Next
            ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(58, 18).Value = "装机功率(元/kW)" Then
                For i = 1 To 31
                    ans_xdc(i) = xdczj_lj_list(i) * xlfl_xdc_list(i) * (tcyfs_list(i) / 12) / 10000
                Next
            End If
            '设备修理费汇总
            For i = 1 To 31
                ans_znxlf(i) = ans_cgsb(i) + ans_rj(i) + ans_nt(i) + ans_gf(i) + ans_fd(i) + ans_xdc(i)
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 11).Value = "方法二" Then
            '常规设备
            ans_cgsb = 设备修理费计算_常规方法(xlfl_cg_list, gdzcyz_list, jsqdklx_list, qttz_list, tcyfs_list, wxzcbl_list, False)
            '燃机、暖通
            For i = 1 To 31
                ans_rj(i) = rjtz_lj_list(i) * xlfl_rj_list(i)
                ans_nt(i) = nttz_lj_list(i) * xlfl_nt_list(i)
            Next
            '光伏
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(56, 18).Value = "投资额百分比(%)" Then
                For i = 1 To 31
                    ans_gf(i) = gftz_lj_list(i) * xlfl_gf_list(i)
                Next
            ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(56, 18).Value = "装机功率(元/kW)" Then
                For i = 1 To 31
                    ans_gf(i) = gfzj_lj_list(i) * xlfl_gf_list(i) / 10000
                Next
            End If
            '风电
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(57, 18).Value = "投资额百分比(%)" Then
                For i = 1 To 31
                    ans_fd(i) = fdtz_lj_list(i) * xlfl_fd_list(i)
                Next
            ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(57, 18).Value = "装机功率(元/kW)" Then
                For i = 1 To 31
                    ans_fd(i) = fdzj_lj_list(i) * xlfl_fd_list(i) / 10000
                Next
            End If
            '蓄电池
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(58, 18).Value = "投资额百分比(%)" Then
                For i = 1 To 31
                    ans_xdc(i) = xdctz_lj_list(i) * xlfl_xdc_list(i)
                Next
            ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(58, 18).Value = "装机功率(元/kW)" Then
                For i = 1 To 31
                    ans_xdc(i) = xdczj_lj_list(i) * xlfl_xdc_list(i) / 10000
                Next
            End If
            '设备修理费汇总
            For i = 1 To 31
                ans_znxlf(i) = ans_cgsb(i) + ans_rj(i) + ans_nt(i) + ans_gf(i) + ans_fd(i) + ans_xdc(i)
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 11).Value = "方法三" Then
            '常规设备
            ans_cgsb = 设备修理费计算_博微方法(xlfl, jttz_list, gdzcyz_list, jsqdklx_list, qttz_list, tcyfs_list, wxzcbl, False)
            '燃机、暖通
            For i = 1 To 31
                ans_rj(i) = rjtz_lj_list(i) * xlfl_rj_list(i)
                ans_nt(i) = nttz_lj_list(i) * xlfl_nt_list(i)
            Next
            '光伏
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(56, 18).Value = "投资额百分比(%)" Then
                For i = 1 To 31
                    ans_gf(i) = gftz_lj_list(i) * xlfl_gf_list(i)
                Next
            ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(56, 18).Value = "装机功率(元/kW)" Then
                For i = 1 To 31
                    ans_gf(i) = gfzj_lj_list(i) * xlfl_gf_list(i) / 10000
                Next
            End If
            '风电
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(57, 18).Value = "投资额百分比(%)" Then
                For i = 1 To 31
                    ans_fd(i) = fdtz_lj_list(i) * xlfl_fd_list(i)
                Next
            ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(57, 18).Value = "装机功率(元/kW)" Then
                For i = 1 To 31
                    ans_fd(i) = fdzj_lj_list(i) * xlfl_fd_list(i) / 10000
                Next
            End If
            '蓄电池
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(58, 18).Value = "投资额百分比(%)" Then
                For i = 1 To 31
                    ans_xdc(i) = xdctz_lj_list(i) * xlfl_xdc_list(i)
                Next
            ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(58, 18).Value = "装机功率(元/kW)" Then
                For i = 1 To 31
                    ans_xdc(i) = xdczj_lj_list(i) * xlfl_xdc_list(i) / 10000
                Next
            End If
            '设备修理费汇总
            For i = 1 To 31
                ans_znxlf(i) = ans_cgsb(i) + ans_rj(i) + ans_nt(i) + ans_gf(i) + ans_fd(i) + ans_xdc(i)
            Next
        End If
        '返回结果
        Dim ans(6)
        ans(0) = ans_znxlf
        ans(1) = ans_cgsb
        ans(2) = ans_rj
        ans(3) = ans_nt
        ans(4) = ans_gf
        ans(5) = ans_fd
        ans(6) = ans_xdc
        Return ans
    End Function

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
    Function 设备修理费计算基数扣除默认设置(ExcelApp As Object)
        'yynx_rj：燃机每次投资运营年限数量（年）
        'yynx_xdc：蓄电池每次投资运营年限数量（年）
        'yynx_nt：暖通每次投资运营年限数量（年）
        'yynx_gf：光伏每次投资运营年限数量（年）
        'yynx_fd：风电每次投资运营年限数量（年）
        'yynx_cg：常规设备每次投资运营年限数量（年）
        'kcbl_rj：燃机每次投资运营年限结束后，上次投资计算修理费的扣除比例（%）
        'kcbl_xdc：蓄电池每次投资运营年限结束后，上次投资计算修理费的扣除比例（%）
        'kcbl_nt：暖通每次投资运营年限结束后，上次投资计算修理费的扣除比例（%）
        'kcbl_gf：光伏每次投资运营年限结束后，上次投资计算修理费的扣除比例（%）
        'kcbl_fd：风电每次投资运营年限结束后，上次投资计算修理费的扣除比例（%）
        'kcbl_cg：常规设备每次投资运营年限结束后，上次投资计算修理费的扣除比例（%）

        '读取计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        Dim yynx_rj As Integer = jsnx - 1
        Dim yynx_xdc As Integer = 10
        Dim yynx_nt As Integer = jsnx - 1
        Dim yynx_gf As Integer = 25
        Dim yynx_fd As Integer = 20
        Dim yynx_cg As Integer = jsnx - 1
        Dim kcbl_rj As Double = 1
        Dim kcbl_xdc As Double = 1
        Dim kcbl_nt As Double = 1
        Dim kcbl_gf As Double = 1
        Dim kcbl_fd As Double = 1
        Dim kcbl_cg As Double = 1

        Dim ans(11)
        ans(0) = yynx_rj
        ans(1) = yynx_xdc
        ans(2) = yynx_nt
        ans(3) = yynx_gf
        ans(4) = yynx_fd
        ans(5) = yynx_cg
        ans(6) = kcbl_rj
        ans(7) = kcbl_xdc
        ans(8) = kcbl_nt
        ans(9) = kcbl_gf
        ans(10) = kcbl_fd
        ans(11) = kcbl_cg
        Return ans
    End Function
End Module
