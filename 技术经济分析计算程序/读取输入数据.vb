Module 读取输入数据
    Function 读取估算表数据(ExcelApp As Object)
        '投资年份
        Dim tznf_list(10) As Integer
        '建设投资静态投资(万元)
        Dim jttz_list(10) As Double
        '建设投资动态投资(万元)
        Dim dttz_list(10) As Double
        '建设投资资本金(万元)
        Dim tzzbj_list(10) As Double
        '建设投资贷款金额(万元)
        Dim dkje_list(10) As Double
        '建设期贷款利息(万元)
        Dim dklx_list(10) As Double
        '流动资金(万元)
        Dim ldzj_list(10) As Double
        '铺底流动资金(万元)
        Dim pdldzj_list(10) As Double
        '流动资金贷款(万元)
        Dim ldzjdk_list(10) As Double
        '可抵扣增值税总额(万元)
        Dim kdkzzs_list(10) As Double

        '燃机总投资(万元)
        Dim rjtz_list(10) As Double
        '燃机总发电量(万kWh)
        Dim rjfdl_list(10) As Double
        '蓄电池总投资(万元)
        Dim xdctz_list(10) As Double
        '蓄电池总装机功率(kW)
        Dim xdczjgl_list(10) As Double
        '暖通总投资(万元)
        Dim nttz_list(10) As Double
        '供冷供热总量(万kWh)
        Dim glgrl_list(10) As Double
        '光伏总投资(万元)
        Dim gftz_list(10) As Double
        '光伏总装机功率(kW)
        Dim gfzjgl_list(10) As Double
        '燃煤机组总发电量(万kWh)
        Dim rmfdl_list(10) As Double
        '风电总投资(万元)
        Dim fdtz_list(10) As Double
        '风电总装机功率(kW)
        Dim fdzjgl_list(10) As Double
        '垃圾发电总发电量(万kWh)
        Dim ljfdl_list(10) As Double

        '读取数据
        '1-5次
        For i = 1 To 5
            tznf_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 2 * i + 1).Value
            jttz_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 2 * i + 1).Value
            dttz_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(31, 2 * i + 1).Value
            tzzbj_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(32, 2 * i + 1).Value
            dkje_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(33, 2 * i + 1).Value
            dklx_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(34, 2 * i + 1).Value
            ldzj_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(35, 2 * i + 1).Value
            pdldzj_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(36, 2 * i + 1).Value
            ldzjdk_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(37, 2 * i + 1).Value
            kdkzzs_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(38, 2 * i + 1).Value
            rjtz_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(44, 2 * i + 1).Value
            rjfdl_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(45, 2 * i + 1).Value
            xdctz_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(46, 2 * i + 1).Value
            xdczjgl_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(47, 2 * i + 1).Value
            nttz_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(48, 2 * i + 1).Value
            glgrl_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(49, 2 * i + 1).Value
            gftz_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(50, 2 * i + 1).Value
            gfzjgl_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(51, 2 * i + 1).Value
            rmfdl_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(52, 2 * i + 1).Value
            fdtz_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(53, 2 * i + 1).Value
            fdzjgl_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(54, 2 * i + 1).Value
            ljfdl_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(55, 2 * i + 1).Value
        Next
        '6-10次
        For i = 6 To 10
            tznf_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 2 * i - 9).Value
            jttz_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 2 * i - 9).Value
            dttz_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(74, 2 * i - 9).Value
            tzzbj_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(75, 2 * i - 9).Value
            dkje_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(76, 2 * i - 9).Value
            dklx_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(77, 2 * i - 9).Value
            ldzj_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(78, 2 * i - 9).Value
            pdldzj_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(79, 2 * i - 9).Value
            ldzjdk_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(80, 2 * i - 9).Value
            kdkzzs_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(81, 2 * i - 9).Value
            rjtz_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(91, 2 * i - 9).Value
            rjfdl_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(92, 2 * i - 9).Value
            xdctz_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(93, 2 * i - 9).Value
            xdczjgl_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(94, 2 * i - 9).Value
            nttz_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(95, 2 * i - 9).Value
            glgrl_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(96, 2 * i - 9).Value
            gftz_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(97, 2 * i - 9).Value
            gfzjgl_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(98, 2 * i - 9).Value
            rmfdl_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(99, 2 * i - 9).Value
            fdtz_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(100, 2 * i - 9).Value
            fdzjgl_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(101, 2 * i - 9).Value
            ljfdl_list(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(102, 2 * i - 9).Value
        Next

        '返回结果
        Dim ans(21)
        ans(0) = tznf_list
        ans(1) = jttz_list
        ans(2) = dttz_list
        ans(3) = tzzbj_list
        ans(4) = dkje_list
        ans(5) = dklx_list
        ans(6) = ldzj_list
        ans(7) = pdldzj_list
        ans(8) = ldzjdk_list
        ans(9) = kdkzzs_list
        ans(10) = rjtz_list
        ans(11) = rjfdl_list
        ans(12) = xdctz_list
        ans(13) = xdczjgl_list
        ans(14) = nttz_list
        ans(15) = glgrl_list
        ans(16) = gftz_list
        ans(17) = gfzjgl_list
        ans(18) = fdtz_list
        ans(19) = fdzjgl_list
        ans(20) = ljfdl_list
        ans(21) = rmfdl_list
        Return ans
    End Function
    Function 读取计算常量设置(ExcelApp As Object)
        '读取第3列(C列)，内容分别是
        '资本金基准收益率(%)
        '全投资基准收益率(%)
        '投资方基准收益率(%)
        Dim ans_C(3)
        For i = 1 To 3
            ans_C(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 4, 3).Value
        Next
        '读取第5列(E列)，内容分别是
        '资本金比例(%)
        '自有流动资金比例(%)
        '固定资产残值率(%)
        '无形资产所占比例(%)
        '常规设备修理费率（%）
        '保险费率(%)
        '增值税退税比例(%)
        '提取盈余公积(%)
        '城市维护建设税率(%)
        '教育附加税率(%)
        Dim ans_E(10)
        For i = 1 To 10
            ans_E(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 4, 5).Value
        Next
        '读取第7列(G列)，内容分别是
        '本项目计算年限(年)
        '固定资产折旧年限(年)
        '长期贷款还款年限(年)
        '无形资产摊销年限(年)
        '建设期增值税抵扣年限(年)
        '长期贷款还款宽限年限(年)
        '所得税免除年份数(年)
        '所得税减征年份数(年)
        '所得税率(%)
        '所得税减少征收比例(%)
        Dim ans_G(10)
        For i = 1 To 10
            ans_G(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 4, 7).Value
        Next
        '读取第9列(I列)，内容分别是
        '长期贷款年计息次数
        '流动资金贷款年计息次数
        '短期贷款年计息次数
        '长期贷款利率(%)
        '流动资金贷款利率(%)
        '短期贷款利率(%)
        '平均资本成长率(%)
        '经济增加值折现率(%)
        '其他制造费率(%)
        '其他管理费率(%)
        Dim ans_I(10)
        For i = 1 To 10
            ans_I(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 4, 9).Value
        Next
        Dim ans(3)
        ans(0) = ans_C
        ans(1) = ans_E
        ans(2) = ans_G
        ans(3) = ans_I
        Return ans
    End Function
    Function 读取计算方法设置(ExcelApp As Object)
        '读取第11列(K列)，内容分别是
        '长期借款计算方法选择
        '折旧&摊销计算方法选择
        '修理费计算方法选择
        '建设期贷款利息计算方法
        '固定成本是否按月份折算
        '宽限期内贷款是否付息
        '材料和其它费计算方式
        '投产后抵扣建设期增值税
        '补贴收入是否缴纳所得税
        '电力/ 市政计算方法选择
        Dim ans_K(10) As String
        For i = 1 To 10
            ans_K(i) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 4, 11).Value
        Next
        '返回结果
        Return ans_K
    End Function
    Function 读取建设期时间计划数据(ExcelApp As Object)
        '开始年份列表
        Dim year_start_list(10) As Integer
        '结束年份列表
        Dim year_end_list(10) As Integer
        '开始月份列表
        Dim month_start_list(10) As Integer
        '结束月份列表
        Dim month_end_list(10) As Integer
        '计数
        Dim js_1 As Integer = 1
        For i = 3 To 131 Step 32 '行
            For j = 2 To 10 Step 8 '列
                year_start_list(js_1) = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(i, j).Value
                year_end_list(js_1) = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(i + 1, j).Value
                month_start_list(js_1) = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(i, j + 2).Value
                month_end_list(js_1) = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(i + 1, j + 2).Value
                js_1 += 1
            Next
        Next
        '手动输入的投产月份，列表
        Dim year_shuru_list(10) As Integer
        Dim month_shuru_list(10) As Integer
        '计数
        Dim js_2 As Integer = 1
        For i = 172 To 180 Step 2 '行
            For j = 2 To 10 Step 8 '列
                year_shuru_list(js_2) = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(i, j).Value
                month_shuru_list(js_2) = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(i, j + 2).Value
                js_2 += 1
            Next
        Next
        '读取手动输入的建设期贷款利息金额
        Dim dklx_shuru_list(10) As Double
        For i = 1 To 5
            dklx_shuru_list(i) = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(162 + i, 12).Value
        Next
        For i = 6 To 10
            dklx_shuru_list(i) = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(157 + i, 16).Value
        Next
        '读取投资各方出资比例
        Dim tzgfczbl_list(5) As Double
        For i = 1 To 5
            tzgfczbl_list(i) = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(162 + i, 2).Value
        Next
        '返回结果
        Dim ans(7)
        ans(0) = year_start_list
        ans(1) = year_end_list
        ans(2) = month_start_list
        ans(3) = month_end_list
        ans(4) = year_shuru_list
        ans(5) = month_shuru_list
        ans(6) = dklx_shuru_list
        ans(7) = tzgfczbl_list
        Return ans
    End Function
    Function 读取逐年负荷率(ExcelApp As Object)
        Dim ans_fhl(31) As Double
        '1-15年
        For i = 1 To 15
            ans_fhl(i) = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i + 2).Value
        Next
        '16-31年
        For i = 16 To 31
            ans_fhl(i) = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 14).Value
        Next
        '返回结果
        Return ans_fhl
    End Function
End Module
