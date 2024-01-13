Module 保险费计算
    Sub 保险费计算(ExcelApp As Object, hscz As Boolean, zbj_model As Integer)
        '每次计算完折旧摊销和长期贷款后，都要计算一次本SUB
        '保险费金额只与固定资产原值OR固定资产净值有关
        'hscy：计算期末，是否回收资产残值
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        '————————————————————————————————————————————————————————————————————————————————————————
        '项目计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '分项逐年保险费金额计算，列表长度31
        Dim ans_bxf = 分项逐年保险费金额计算(ExcelApp, hscz, zbj_model)
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

    Function 分项逐年保险费金额计算(ExcelApp As Object, hscz As Boolean, zbj_model As Integer)
        'On Error Resume Next
        '光伏、风电、蓄电池的保险费要单独计算
        '只计算出分项保险费金额数值，不写入EXCEL
        '————————————————————————————————————————————————————————————————————————————————————————
        'hscy：计算期末，是否回收资产残值
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
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
        '计算逐年累计固定资产原值
        Dim gdzcyz_lj_cg(31) As Double
        Dim gdzcyz_lj_rj(31) As Double
        Dim gdzcyz_lj_xdc(31) As Double
        Dim gdzcyz_lj_nt(31) As Double
        Dim gdzcyz_lj_gf(31) As Double
        Dim gdzcyz_lj_fd(31) As Double
        For i = 1 To 31
            gdzcyz_lj_cg(i) = gdzcyz_lj_cg(i - 1) + gdzcyz_cg_list(i)
            gdzcyz_lj_rj(i) = gdzcyz_lj_rj(i - 1) + gdzcyz_rj_list(i)
            gdzcyz_lj_xdc(i) = gdzcyz_lj_xdc(i - 1) + gdzcyz_xdc_list(i)
            gdzcyz_lj_nt(i) = gdzcyz_lj_nt(i - 1) + gdzcyz_nt_list(i)
            gdzcyz_lj_gf(i) = gdzcyz_lj_gf(i - 1) + gdzcyz_gf_list(i)
            gdzcyz_lj_fd(i) = gdzcyz_lj_fd(i - 1) + gdzcyz_fd_list(i)
        Next
        'Dim gdzcyz_lj_qt(31) As Double
        'For i = 1 To 31
        '    gdzcyz_lj_qt(i) = gdzcyz_lj_cg(i) + gdzcyz_lj_rj(i) + gdzcyz_lj_nt(i)
        'Next
        ''其它项逐年剩余固定资产净值
        'Dim sygdzcjz_qt(31) As Double
        'For i = 1 To 31
        '    sygdzcjz_qt(i) = sygdzcjz_cg(i) + sygdzcjz_rj(i) + sygdzcjz_nt(i)
        'Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '读取保险费计算基数设置
        Dim bxf_mode As String = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(209, 22).Value
        '读取保险费率
        Dim bxfl_cg As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 5).Value
        Dim bxfl_gf As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(211, 22).Value
        Dim bxfl_fd As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(212, 22).Value
        Dim bxfl_xdc As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(213, 22).Value
        Dim bxfl_rj As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(214, 22).Value
        Dim bxfl_nt As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(215, 22).Value
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
                ans_bxf_cg(i) = sygdzcjz_cg(i) * bxfl_cg
                ans_bxf_gf(i) = sygdzcjz_gf(i) * bxfl_gf
                ans_bxf_fd(i) = sygdzcjz_fd(i) * bxfl_fd
                ans_bxf_xdc(i) = sygdzcjz_xdc(i) * bxfl_xdc
                ans_bxf_rj(i) = sygdzcjz_rj(i) * bxfl_rj
                ans_bxf_nt(i) = sygdzcjz_nt(i) * bxfl_nt
            Next
        Else
            '以固定资产原值为基础进行计算(默认计算方式)
            For i = 1 To 31
                ans_bxf_cg(i) = gdzcyz_lj_cg(i) * bxfl_cg
                ans_bxf_gf(i) = gdzcyz_lj_gf(i) * bxfl_gf
                ans_bxf_fd(i) = gdzcyz_lj_fd(i) * bxfl_fd
                ans_bxf_xdc(i) = gdzcyz_lj_xdc(i) * bxfl_xdc
                ans_bxf_rj(i) = gdzcyz_lj_rj(i) * bxfl_rj
                ans_bxf_nt(i) = gdzcyz_lj_nt(i) * bxfl_nt
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

End Module
