Module 保险费计算
    Sub 保险费计算(ExcelApp As Object, hscz As Boolean, zbj_model As Integer, bxf_model As Integer, kcje_bxf_model As Integer)
        '每次计算完折旧摊销和长期贷款后，都要计算一次本SUB
        '保险费金额只与固定资产原值OR固定资产净值有关
        'hscy：计算期末，是否回收资产残值
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        'bxf_model：保险费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_bxf_model：保险费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        '————————————————————————————————————————————————————————————————————————————————————————
        '项目计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '分项逐年保险费金额计算，列表长度31
        Dim ans_bxf = 分项逐年保险费金额计算(ExcelApp, hscz, zbj_model, bxf_model, kcje_bxf_model)
        Dim ans_bxf_cg = ans_bxf(0)
        Dim ans_bxf_gf = ans_bxf(1)
        Dim ans_bxf_fd = ans_bxf(2)
        Dim ans_bxf_xdc = ans_bxf(3)
        Dim ans_bxf_rj = ans_bxf(4)
        Dim ans_bxf_nt = ans_bxf(5)
        '保险费加和
        Dim ans_bxf_all(31) As Double
        For i = 1 To 31
            ans_bxf_all(i) = ans_bxf_cg(i) + ans_bxf_gf(i) + ans_bxf_fd(i) + ans_bxf_xdc(i) + ans_bxf_rj(i) + ans_bxf_nt(i)
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算结果写入Excel
        '前15年
        For i = 1 To 15
            If i <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(211, 4 + i).Value = ans_bxf_all(i)
            Else
                ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(211, 4 + i).Value = 0
            End If
        Next
        '16—31年
        For i = 16 To 31
            If i <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(215, i - 12).Value = ans_bxf_all(i)
            Else
                ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(215, i - 12).Value = 0
            End If
        Next
    End Sub

    Function 分项逐年保险费金额计算(ExcelApp As Object, hscz As Boolean, zbj_model As Integer, bxf_model As Integer, kcje_bxf_model As Integer)
        'On Error Resume Next
        '光伏、风电、蓄电池的保险费要单独计算
        '只计算出分项保险费金额数值，不写入EXCEL
        '————————————————————————————————————————————————————————————————————————————————————————
        'hscy：计算期末，是否回收资产残值
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        'bxf_model：保险费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_bxf_model：保险费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        '————————————————————————————————————————————————————————————————————————————————————————
        '读取输入数据
        Dim GSBSJ = 读取估算表数据(ExcelApp)
        '10次投资年份序号列表，列表，长度10
        Dim tznf_list = GSBSJ(0)
        '分项折旧摊销计算
        Dim ans_zjtx = 分项逐年折旧摊销金额计算(ExcelApp, hscz, zbj_model)
        '资产原值，列表长度10
        Dim gdzcyz_cg = ans_zjtx(2)
        Dim gdzcyz_rj = ans_zjtx(3)
        Dim gdzcyz_xdc = ans_zjtx(4)
        Dim gdzcyz_nt = ans_zjtx(5)
        Dim gdzcyz_gf = ans_zjtx(6)
        Dim gdzcyz_fd = ans_zjtx(7)
        '逐年剩余固定资产净值，列表长度31
        Dim sygdzcjz_cg = ans_zjtx(14)(2)
        Dim sygdzcjz_rj = ans_zjtx(15)(2)
        Dim sygdzcjz_xdc = ans_zjtx(16)(2)
        Dim sygdzcjz_nt = ans_zjtx(17)(2)
        Dim sygdzcjz_gf = ans_zjtx(18)(2)
        Dim sygdzcjz_fd = ans_zjtx(19)(2)
        '将10次投资的资产原值转为长度31的列表
        Dim gdzcyz_cg_list = 基础计算功能_10_to_31(tznf_list, gdzcyz_cg)
        Dim gdzcyz_rj_list = 基础计算功能_10_to_31(tznf_list, gdzcyz_rj)
        Dim gdzcyz_xdc_list = 基础计算功能_10_to_31(tznf_list, gdzcyz_xdc)
        Dim gdzcyz_nt_list = 基础计算功能_10_to_31(tznf_list, gdzcyz_nt)
        Dim gdzcyz_gf_list = 基础计算功能_10_to_31(tznf_list, gdzcyz_gf)
        Dim gdzcyz_fd_list = 基础计算功能_10_to_31(tznf_list, gdzcyz_fd)
        '————————————————————————————————————————————————————————————————————————————————————————
        '读取保险费计算基数设置
        Dim bxf_mode As String = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(209, 22).Value
        '读取保险费率
        Dim bxfl_cg_list(31) As Double
        Dim bxfl_gf_list(31) As Double
        Dim bxfl_fd_list(31) As Double
        Dim bxfl_xdc_list(31) As Double
        Dim bxfl_rj_list(31) As Double
        Dim bxfl_nt_list(31) As Double
        If bxf_mode = 0 Then
            Dim bxfl_mr = 默认逐年保险费率(ExcelApp)
            bxfl_cg_list = bxfl_mr(0)
            bxfl_gf_list = bxfl_mr(1)
            bxfl_fd_list = bxfl_mr(2)
            bxfl_xdc_list = bxfl_mr(3)
            bxfl_rj_list = bxfl_mr(4)
            bxfl_nt_list = bxfl_mr(5)
        Else
            For i = 3 To 33
                '燃机修保险率（%）
                bxfl_rj_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(135, i).Value
                '蓄电池保险率（%）
                bxfl_xdc_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(136, i).Value
                '光伏设备保险率（%）
                bxfl_gf_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(137, i).Value
                '暖通设备保险率（%）
                bxfl_nt_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(138, i).Value
                '风电设备保险率（%）
                bxfl_fd_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(139, i).Value
                '常规设备保险率（%）
                bxfl_cg_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(140, i).Value
            Next
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '读取默认值
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
        If kcje_bxf_model = 0 Then
            Dim kcje_mr = 设备保险费计算基数扣除默认设置(ExcelApp)
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
            yynx_rj = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(109, 3).Value
            yynx_xdc = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(109, 4).Value
            yynx_nt = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(109, 5).Value
            yynx_gf = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(109, 6).Value
            yynx_fd = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(109, 7).Value
            yynx_cg = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(109, 8).Value
            kcbl_rj = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(110, 3).Value
            kcbl_xdc = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(110, 4).Value
            kcbl_nt = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(110, 5).Value
            kcbl_gf = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(110, 6).Value
            kcbl_fd = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(110, 7).Value
            kcbl_cg = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(110, 8).Value
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算逐年保险费
        Dim ans_bxf_cg(31) As Double
        Dim ans_bxf_gf(31) As Double
        Dim ans_bxf_fd(31) As Double
        Dim ans_bxf_xdc(31) As Double
        Dim ans_bxf_rj(31) As Double
        Dim ans_bxf_nt(31) As Double
        If bxf_mode = "净值" Then
            '以固定资产净值为基础进行计算
            For i = 1 To 31
                ans_bxf_cg(i) = sygdzcjz_cg(i) * bxfl_cg_list(i)
                ans_bxf_gf(i) = sygdzcjz_gf(i) * bxfl_gf_list(i)
                ans_bxf_fd(i) = sygdzcjz_fd(i) * bxfl_fd_list(i)
                ans_bxf_xdc(i) = sygdzcjz_xdc(i) * bxfl_xdc_list(i)
                ans_bxf_rj(i) = sygdzcjz_rj(i) * bxfl_rj_list(i)
                ans_bxf_nt(i) = sygdzcjz_nt(i) * bxfl_nt_list(i)
            Next
        Else
            '以固定资产原值为基础进行计算(默认计算方式)
            Dim gdzcyz_lj_cg As Double = 0
            Dim gdzcyz_lj_rj As Double = 0
            Dim gdzcyz_lj_xdc As Double = 0
            Dim gdzcyz_lj_nt As Double = 0
            Dim gdzcyz_lj_gf As Double = 0
            Dim gdzcyz_lj_fd As Double = 0
            For i = 1 To 31
                gdzcyz_lj_cg += gdzcyz_cg_list(i - 1)
                gdzcyz_lj_rj += gdzcyz_rj_list(i - 1)
                gdzcyz_lj_xdc += gdzcyz_xdc_list(i - 1)
                gdzcyz_lj_nt += gdzcyz_nt_list(i - 1)
                gdzcyz_lj_gf += gdzcyz_gf_list(i - 1)
                gdzcyz_lj_fd += gdzcyz_fd_list(i - 1)
                ans_bxf_cg(i) = gdzcyz_lj_cg * bxfl_cg_list(i)
                ans_bxf_gf(i) = gdzcyz_lj_gf * bxfl_gf_list(i)
                ans_bxf_fd(i) = gdzcyz_lj_fd * bxfl_fd_list(i)
                ans_bxf_xdc(i) = gdzcyz_lj_xdc * bxfl_xdc_list(i)
                ans_bxf_rj(i) = gdzcyz_lj_rj * bxfl_rj_list(i)
                ans_bxf_nt(i) = gdzcyz_lj_nt * bxfl_nt_list(i)
            Next
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '返回结果
        Dim ans(5)
        ans(0) = ans_bxf_cg
        ans(1) = ans_bxf_gf
        ans(2) = ans_bxf_fd
        ans(3) = ans_bxf_xdc
        ans(4) = ans_bxf_rj
        ans(5) = ans_bxf_nt
        Return ans
    End Function
    Function 默认逐年保险费率(ExcelApp As Object)
        '读取Excel中输入的默认修理费
        Dim bxfl_cg As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 5).Value
        Dim bxfl_gf As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(211, 22).Value
        Dim bxfl_fd As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(212, 22).Value
        Dim bxfl_xdc As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(213, 22).Value
        Dim bxfl_rj As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(214, 22).Value
        Dim bxfl_nt As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(215, 22).Value
        '逐年修理费率默认值，每年都一样
        Dim bxfl_cg_list(31)
        Dim bxfl_rj_list(31)
        Dim bxfl_xdc_list(31)
        Dim bxfl_gf_list(31)
        Dim bxfl_nt_list(31)
        Dim bxfl_fd_list(31)
        For i = 1 To 31
            bxfl_cg_list(i) = bxfl_cg
            bxfl_rj_list(i) = bxfl_rj
            bxfl_xdc_list(i) = bxfl_xdc
            bxfl_gf_list(i) = bxfl_gf
            bxfl_fd_list(i) = bxfl_fd
            bxfl_nt_list(i) = bxfl_nt
        Next
        '返回结果
        Dim ans(5)
        ans(0) = bxfl_cg_list
        ans(1) = bxfl_gf_list
        ans(2) = bxfl_fd_list
        ans(3) = bxfl_xdc_list
        ans(4) = bxfl_rj_list
        ans(5) = bxfl_nt_list
        Return ans
    End Function
    Function 设备保险费计算基数扣除默认设置(ExcelApp As Object)
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
        Dim yynx_gf As Integer = jsnx - 1
        Dim yynx_fd As Integer = jsnx - 1
        Dim yynx_cg As Integer = jsnx - 1
        Dim kcbl_rj As Double = 0
        Dim kcbl_xdc As Double = 1
        Dim kcbl_nt As Double = 0
        Dim kcbl_gf As Double = 0
        Dim kcbl_fd As Double = 0
        Dim kcbl_cg As Double = 0

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
