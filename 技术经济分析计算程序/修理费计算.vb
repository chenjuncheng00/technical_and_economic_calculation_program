Module 修理费计算
    Sub 修理费计算(ExcelApp As Object, xlfl_cg_model As Integer, xlfl_qt_model As Integer, kcje_xlf_model As Integer)
        'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        '————————————————————————————————————————————————————————————————————————————————————————
        '项目计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '计算逐年分项修理费金额
        Dim ans_xlf = 分项逐年修理费计算(ExcelApp, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, jsnx)
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

    Function 分项逐年修理费计算(ExcelApp As Object, xlfl_cg_model As Integer, xlfl_qt_model As Integer, kcje_xlf_model As Integer, jsnx As Integer)
        '只计算出分项逐年修理费的金额数值，不写入EXCEL
        '————————————————————————————————————————————————————————————————————————————————————————        
        'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'jsnx：项目计算年限
        '————————————————————————————————————————————————————————————————————————————————————————
        '修理费率
        Dim xlfl As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(9, 5).Value
        '无形资产所占比例
        Dim wxzcbl As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 5).Value
        '计算10次投资，每次建设年份的投产月份数
        Dim ans_month = 逐年投产月份数_10次投资(ExcelApp)
        Dim month_10_list = ans_month(1)
        Dim tcyf_list = ans_month(2)
        '读取输入数据
        Dim GSBSJ = 读取估算表数据(ExcelApp)
        Dim tznf_list = GSBSJ(0)
        Dim dttz_list = GSBSJ(2)
        Dim kdkzzs_list = GSBSJ(9)
        'jttz_list：逐年静态投资，列表
        Dim jttz = GSBSJ(1)
        'jsqdklx_list：逐年建设期贷款利息，列表
        Dim jsqdklx = GSBSJ(5)
        '燃机总投资(万元)
        Dim rjtz = GSBSJ(10)
        '蓄电池总投资(万元)
        Dim xdctz = GSBSJ(12)
        '蓄电池装机(kW)
        Dim xdczj = GSBSJ(13)
        '暖通总投资(万元)
        Dim nttz = GSBSJ(14)
        '光伏总投资(万元)
        Dim gftz = GSBSJ(16)
        '光伏装机功率(kW)
        Dim gfzj = GSBSJ(17)
        '风电总投资(万元)
        Dim fdtz = GSBSJ(18)
        '风电装机功率(kW)
        Dim fdzj = GSBSJ(19)
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
        Dim kcje_mr = 设备修理费计算基数扣除默认设置(jsnx, month_10_list)
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
        '10次投资的数据，用于计算10次投资的固定资产原值和无形资产原值
        Dim gdzcyz(10) As Double
        Dim wxzcyz(10) As Double
        '读取开始年份和资产原值
        For i = 1 To 10
            gdzcyz(i) = (dttz_list(i) - kdkzzs_list(i)) * (1 - wxzcbl)
            wxzcyz(i) = (dttz_list(i) - kdkzzs_list(i)) * wxzcbl
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '10次投资其它投资金额占当年静态投资金额的比例，长度10的列表
        Dim qttzbl(10) As Double
        Dim jttz_cg(10) As Double
        For i = 1 To 10
            If jttz(i) > 0 Then
                qttzbl(i) = (rjtz(i) + xdctz(i) + nttz(i) + gftz(i) + fdtz(i)) / jttz(i)
            Else
                qttzbl(i) = 0
            End If
            jttz_cg(i) = jttz(i) * (1 - qttzbl(i))
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算常规设备的逐年修理费计算基数：相当于计算常规设备投资，按比例折算
        Dim value_base_cg(10) As Double '10次投资
        '计算除了常规设备以外的其它设备的逐年投资额，按比例折算
        Dim value_base_qt(10) As Double
        For i = 1 To 10
            value_base_cg(i) = (gdzcyz(i) - jsqdklx(i) * (1 - wxzcbl)) * (1 - qttzbl(i))
            value_base_qt(i) = (gdzcyz(i) - jsqdklx(i) * (1 - wxzcbl)) * qttzbl(i)
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '建设年份需要按照投产月份数的比例是否需要进行折算
        Dim zs_set As Boolean
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 11).Value = "方法一" Then
            zs_set = True
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 11).Value = "方法二" Then
            zs_set = False
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 11).Value = "方法三" Then
            zs_set = False
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算常规设备的逐年修理费率
        Dim ans_cgsb As Double()
        '燃机、暖通、光伏、风电、蓄电池
        Dim ans_rj(31) As Double
        Dim ans_nt(31) As Double
        Dim ans_gf(31) As Double
        Dim ans_fd(31) As Double
        Dim ans_xdc(31) As Double
        '修理费总额
        Dim ans_znxlf(31) As Double
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 11).Value = "方法一" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 11).Value = "方法二" Then
            '常规设备
            ans_cgsb = 设备修理费计算_常规方法(tznf_list, tcyf_list, xlfl_cg_list, value_base_cg, yynx_cg, kcbl_cg, zs_set)
            '燃机、暖通
            ans_rj = 计算逐年总金额_10次投资(tznf_list, tcyf_list, rjtz, yynx_rj, kcbl_rj, zs_set, xlfl_rj_list, 1)
            ans_nt = 计算逐年总金额_10次投资(tznf_list, tcyf_list, nttz, yynx_nt, kcbl_nt, zs_set, xlfl_nt_list, 1)
            '光伏、风电、蓄电池
            Dim ans_xny = 逐年修理费计算_光伏风电蓄电池(ExcelApp, tznf_list, tcyf_list, gftz, fdtz, xdctz, gfzj, fdzj, xdczj, xlfl_gf_list, xlfl_fd_list,
                                                        xlfl_xdc_list, yynx_gf, yynx_fd, yynx_xdc, kcbl_gf, kcbl_fd, kcbl_xdc, zs_set)
            ans_gf = ans_xny(0)
            ans_fd = ans_xny(1)
            ans_xdc = ans_xny(2)
            '设备修理费汇总
            For i = 1 To 31
                ans_znxlf(i) = ans_cgsb(i) + ans_rj(i) + ans_nt(i) + ans_gf(i) + ans_fd(i) + ans_xdc(i)
            Next
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 11).Value = "方法三" Then
            '常规设备
            ans_cgsb = 设备修理费计算_博微方法(tznf_list, tcyf_list, xlfl_cg_list, value_base_cg, yynx_cg, kcbl_cg, jttz_cg, zs_set)
            '燃机、暖通
            ans_rj = 计算逐年总金额_10次投资(tznf_list, tcyf_list, rjtz, yynx_rj, kcbl_rj, zs_set, xlfl_rj_list, 1)
            ans_nt = 计算逐年总金额_10次投资(tznf_list, tcyf_list, nttz, yynx_nt, kcbl_nt, zs_set, xlfl_nt_list, 1)
            '光伏、风电、蓄电池
            Dim ans_xny = 逐年修理费计算_光伏风电蓄电池(ExcelApp, tznf_list, tcyf_list, gftz, fdtz, xdctz, gfzj, fdzj, xdczj, xlfl_gf_list, xlfl_fd_list,
                                                        xlfl_xdc_list, yynx_gf, yynx_fd, yynx_xdc, kcbl_gf, kcbl_fd, kcbl_xdc, zs_set)
            ans_gf = ans_xny(0)
            ans_fd = ans_xny(1)
            ans_xdc = ans_xny(2)
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

    Function 设备修理费计算_常规方法(tznf_list As Array, tcyf_list As Array, xlfl_cg_list As Array, value_base_cg As Array, yynx_cg As Integer, kcbl_cg As Double, zs_set As Boolean)
        'tznf_list：10次投资的年份序号，列表，长度10
        'tcyf_list：10次投资的投产月份数，列表，长度10
        'xlfl_cg_list：逐年常规设备修理费率，列表，长度31
        'value_base_cg：10次投资的常规设备修理费计算基数，列表，长度10
        'yynx：运营年限
        'kcbl：超过运营年限后的扣除比例
        'zs_set：数值是否需要按照当年的（投产月份数/12）进行折算

        '计算出的逐年修理费金额，列表
        Dim ans_znxlf_list As Double() = 计算逐年总金额_10次投资(tznf_list, tcyf_list, value_base_cg, yynx_cg, kcbl_cg, zs_set, xlfl_cg_list, 1)
        '返回结果
        Return ans_znxlf_list
    End Function
    Function 设备修理费计算_博微方法(tznf_list As Array, tcyf_list As Array, xlfl_cg_list As Array, value_base_cg As Array, yynx_cg As Integer, kcbl_cg As Double, jttz_cg As Array, zs_set As Boolean)
        'tznf_list：10次投资的年份序号，列表，长度10
        'tcyf_list：10次投资的投产月份数，列表，长度10
        'xlfl_cg_list：逐年常规设备修理费率，列表，长度31
        'value_base_cg：10次投资的常规设备修理费计算基数，列表，长度10
        'yynx：运营年限
        'kcbl：超过运营年限后的扣除比例
        'zs_set：数值是否需要按照当年的（投产月份数/12）进行折算
        'jttz_cg：10次投资的常规设备静态投资，列表，长度10

        '用常规计算模式计算逐年修理费，并且找到其中的最大值
        Dim xlf_31_list As Double() = 设备修理费计算_常规方法(tznf_list, tcyf_list, xlfl_cg_list, value_base_cg, yynx_cg, kcbl_cg, zs_set)
        Dim xlf_max As Double = xlf_31_list.Max
        'jttz数据转换
        Dim new_jttz As Double() = Array数据转换为Double(jttz_cg, 10)
        '计算静态投资最大值
        Dim jttz_max As Double = new_jttz.Max
        '静态投资修正转换
        Dim jttz_list = 计算基数累加修正计算_10次投资(tznf_list, tcyf_list, jttz_cg, yynx_cg, kcbl_cg, True)(0)
        '计算出的逐年修理费金额，列表
        '根据累计出资比例折算修理费
        Dim ans_znxlf_list(31) As Double
        For i = 1 To 31
            If xlf_max > 0 Then
                ans_znxlf_list(i) = (jttz_list(i) / jttz_max) * xlf_max
            Else
                ans_znxlf_list(i) = 0
            End If
        Next
        '返回结果
        Return ans_znxlf_list
    End Function
    Function 逐年修理费计算_光伏风电蓄电池(ExcelApp As Object, tznf_list As Array, tcyf_list As Array, gftz As Array, fdtz As Array, xdctz As Array,
                                           gfzjgl As Array, fdzjgl As Array, xdczjgl As Array, xlfl_gf_list As Array, xlfl_fd_list As Array,
                                           xlfl_xdc_list As Array, yynx_gf As Integer, yynx_fd As Integer, yynx_xdc As Integer, kcbl_gf As Double,
                                           kcbl_fd As Double, kcbl_xdc As Double, zs_set As Boolean)
        'tznf_list：10次投资的年份序号，列表，长度10
        'tcyf_list：10次投资的投产月份数，列表，长度10
        'gftz，fdtz， xdctz： 与下面的量一一对应
        '光伏总投资(万元)，10次投资的情况，列表，长度10
        '风电总投资(万元)，10次投资的情况，列表，长度10
        '蓄电池总投资(万元)，10次投资的情况，列表，长度10
        'gfzjgl，fdzjgl， xdczjgl： 与下面的量一一对应
        '光伏总装机功率(kW)，10次投资的情况，列表，长度10
        '风电总装机功率(kW)，10次投资的情况，列表，长度10
        '蓄电池总装机功率(kW)，10次投资的情况，列表，长度10
        'xlfl_gf_lis，xlfl_fd_list，xlfl_xdc_list：与下面的流动资金费率一一对应
        '光伏修理费率， 元/kW OR %，列表，长度31
        '风电修理费率， 元/kW OR %，列表，长度31
        '蓄电池修理费率， 元/kW OR %，列表，长度31
        'yynx：运营年限
        'kcbl：超过运营年限后的扣除比例
        'zs_set：数值是否需要按照当年的（投产月份数/12）进行折算
        '————————————————————————————————————————————————————————————————————————————————————————  
        '光伏、风电、蓄电池
        Dim ans_gf(31) As Double
        Dim ans_fd(31) As Double
        Dim ans_xdc(31) As Double
        '光伏
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(56, 18).Value = "投资额百分比(%)" Then
            ans_gf = 计算逐年总金额_10次投资(tznf_list, tcyf_list, gftz, yynx_gf, kcbl_gf, zs_set, xlfl_gf_list, 1)
        ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(56, 18).Value = "装机功率(元/kW)" Then
            ans_gf = 计算逐年总金额_10次投资(tznf_list, tcyf_list, gfzjgl, yynx_gf, kcbl_gf, zs_set, xlfl_gf_list, 10000)
        End If
        '风电
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(57, 18).Value = "投资额百分比(%)" Then
            ans_fd = 计算逐年总金额_10次投资(tznf_list, tcyf_list, fdtz, yynx_fd, kcbl_fd, zs_set, xlfl_fd_list, 1)
        ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(57, 18).Value = "装机功率(元/kW)" Then
            ans_fd = 计算逐年总金额_10次投资(tznf_list, tcyf_list, fdzjgl, yynx_fd, kcbl_fd, zs_set, xlfl_fd_list, 10000)
        End If
        '蓄电池
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(58, 18).Value = "投资额百分比(%)" Then
            ans_xdc = 计算逐年总金额_10次投资(tznf_list, tcyf_list, xdctz, yynx_xdc, kcbl_xdc, zs_set, xlfl_xdc_list, 1)
        ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(58, 18).Value = "装机功率(元/kW)" Then
            ans_gf = 计算逐年总金额_10次投资(tznf_list, tcyf_list, xdczjgl, yynx_xdc, kcbl_xdc, zs_set, xlfl_xdc_list, 10000)
        End If
        '————————————————————————————————————————————————————————————————————————————————————————  
        Dim ans(2)
        ans(0) = ans_gf
        ans(1) = ans_fd
        ans(2) = ans_xdc
        Return ans
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
        Dim xlfl_cg_list(31) As Double
        Dim xlfl_rj_list(31) As Double
        Dim xlfl_xdc_list(31) As Double
        Dim xlfl_gf_list(31) As Double
        Dim xlfl_nt_list(31) As Double
        '风电设备修理费率（%），逐年变化
        Dim xlfl_fd_list(31) As Double
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
    Function 设备修理费计算基数扣除默认设置(jsnx As Integer, month_10_list As Array)
        'jsnx：计算年限
        'month_10_list：10次投资，计算每次投资的逐年投产月份数，列表，长度10
        '————————————————————————————————————————————————————————————————————————————————————————       
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

        Dim yynx = 计算项目运营年限(jsnx, month_10_list)
        Dim yynx_rj As Integer = yynx
        Dim yynx_xdc As Integer = 10
        Dim yynx_nt As Integer = yynx
        Dim yynx_gf As Integer = 25
        Dim yynx_fd As Integer = 20
        Dim yynx_cg As Integer = yynx
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
