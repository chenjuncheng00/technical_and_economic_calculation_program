Module 折旧摊销计算
    Sub 折旧摊销计算(ExcelApp As Object, hscz As Boolean, zbj_model As Integer)
        'hscy：计算期末，是否回收资产残值
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        '————————————————————————————————————————————————————————————————————————————————————————
        '项目计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '————————————————————————————————————————————————————————————————————————————————————————
        '分项折旧摊销计算
        Dim ans_zjtx = 分项逐年折旧摊销金额计算(ExcelApp, hscz, zbj_model, jsnx)
        Dim gdzcyz = ans_zjtx(0)
        Dim wxzcyz = ans_zjtx(1)
        Dim ans_gdzczj_cg = ans_zjtx(14)
        Dim ans_gdzczj_rj = ans_zjtx(15)
        Dim ans_gdzczj_xdc = ans_zjtx(16)
        Dim ans_gdzczj_nt = ans_zjtx(17)
        Dim ans_gdzczj_gf = ans_zjtx(18)
        Dim ans_gdzczj_fd = ans_zjtx(19)
        Dim ans_wxzctx_cg = ans_zjtx(20)
        Dim ans_wxzctx_rj = ans_zjtx(21)
        Dim ans_wxzctx_xdc = ans_zjtx(22)
        Dim ans_wxzctx_nt = ans_zjtx(23)
        Dim ans_wxzctx_gf = ans_zjtx(24)
        Dim ans_wxzctx_fd = ans_zjtx(25)
        '————————————————————————————————————————————————————————————————————————————————————————   
        '结果写入Excel
        '写入固定资产、无形资产原值
        Dim gdzcyz_all As Double
        Dim wxzcyz_all As Double
        For i = 1 To 10
            gdzcyz_all += gdzcyz(i)
            wxzcyz_all += wxzcyz(i)
        Next
        ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(5, 4).Value = gdzcyz_all
        ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(11, 4).Value = wxzcyz_all
        '各种方法结果的累加
        Dim gdzcyz_list(31) As Double '固定资产折旧
        Dim ans_gdzczj_zjtxf(31) As Double '逐年折旧摊销费
        Dim ans_gdzczj_zjtxflj(31) As Double '逐年折旧摊销费累计值
        Dim ans_gdzczj_syjz(31) As Double '逐年剩余固定资产净值
        '无形资产摊销
        Dim wxzcyz_list(31) As Double '无形资产摊销
        Dim ans_wxzctx_zjtxf(31) As Double '逐年折旧摊销费
        Dim ans_wxzctx_zjtxflj(31) As Double '逐年折旧摊销费累计值
        Dim ans_wxzctx_syjz(31) As Double '逐年剩余无形资产净值
        For i = 1 To 31
            ans_gdzczj_zjtxf(i) = ans_gdzczj_cg(0)(i) + ans_gdzczj_rj(0)(i) + ans_gdzczj_xdc(0)(i) + ans_gdzczj_nt(0)(i) + ans_gdzczj_gf(0)(i) + ans_gdzczj_fd(0)(i)
            ans_gdzczj_zjtxflj(i) = ans_gdzczj_cg(1)(i) + ans_gdzczj_rj(1)(i) + ans_gdzczj_xdc(1)(i) + ans_gdzczj_nt(1)(i) + ans_gdzczj_gf(1)(i) + ans_gdzczj_fd(1)(i)
            ans_gdzczj_syjz(i) = ans_gdzczj_cg(2)(i) + ans_gdzczj_rj(2)(i) + ans_gdzczj_xdc(2)(i) + ans_gdzczj_nt(2)(i) + ans_gdzczj_gf(2)(i) + ans_gdzczj_fd(2)(i)
            ans_wxzctx_zjtxf(i) = ans_wxzctx_cg(0)(i) + ans_wxzctx_rj(0)(i) + ans_wxzctx_xdc(0)(i) + ans_wxzctx_nt(0)(i) + ans_wxzctx_gf(0)(i) + ans_wxzctx_fd(0)(i)
            ans_wxzctx_zjtxflj(i) = ans_wxzctx_cg(1)(i) + ans_wxzctx_rj(1)(i) + ans_wxzctx_xdc(1)(i) + ans_wxzctx_nt(1)(i) + ans_wxzctx_gf(1)(i) + ans_wxzctx_fd(1)(i)
            ans_wxzctx_syjz(i) = ans_wxzctx_cg(2)(i) + ans_wxzctx_rj(2)(i) + ans_wxzctx_xdc(2)(i) + ans_wxzctx_nt(2)(i) + ans_wxzctx_gf(2)(i) + ans_wxzctx_fd(2)(i)
            gdzcyz_list(i) = ans_gdzczj_zjtxflj(i) + ans_gdzczj_syjz(i)
            wxzcyz_list(i) = ans_wxzctx_zjtxflj(i) + ans_wxzctx_syjz(i)
        Next
        '前15年
        For i = 1 To 15
            If i <= jsnx Then
                '固定资产折旧
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(8, 4 + i).Value = ans_gdzczj_zjtxf(i)
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(9, 4 + i).Value = ans_gdzczj_zjtxflj(i)
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(10, 4 + i).Value = ans_gdzczj_syjz(i)
                '无形资产摊销
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(14, 4 + i).Value = ans_wxzctx_zjtxf(i)
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(15, 4 + i).Value = ans_wxzctx_zjtxflj(i)
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(16, 4 + i).Value = ans_wxzctx_syjz(i)
            Else
                '固定资产折旧
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(8, 4 + i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(9, 4 + i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(10, 4 + i).Value = 0
                '无形资产摊销
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(14, 4 + i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(15, 4 + i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(16, 4 + i).Value = 0
            End If
        Next
        '16—31年
        For i = 16 To 31
            If i <= jsnx Then
                '固定资产折旧
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(25, i - 12).Value = ans_gdzczj_zjtxf(i)
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(26, i - 12).Value = ans_gdzczj_zjtxflj(i)
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(27, i - 12).Value = ans_gdzczj_syjz(i)
                '无形资产摊销
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(31, i - 12).Value = ans_wxzctx_zjtxf(i)
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(32, i - 12).Value = ans_wxzctx_zjtxflj(i)
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(33, i - 12).Value = ans_wxzctx_syjz(i)
            Else
                '固定资产折旧
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(25, i - 12).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(26, i - 12).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(27, i - 12).Value = 0
                '无形资产摊销
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(31, i - 12).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(32, i - 12).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(33, i - 12).Value = 0
            End If
        Next
        '写入期末回收资产余值
        '前15年
        For i = 1 To 15
            If i = jsnx Then
                '固定资产
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(26, 22).Value = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(10, 4 + i).Value
                '无形资产
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(31, 22).Value = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(16, 4 + i).Value
            End If
        Next
        '16—31年
        For i = 16 To 31
            If i = jsnx Then
                '固定资产
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(26, 22).Value = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(27, i - 12).Value
                '无形资产
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(31, 22).Value = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(33, i - 12).Value
            End If
        Next
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '逐年固定资产原值
        '结果写入Excel
        '前15年
        For i = 1 To 15
            If i <= jsnx Then
                '固定资产
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(5, 4 + i).Value = gdzcyz_list(i)
                '无形资产
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(11, 4 + i).Value = wxzcyz_list(i)
            Else
                '固定资产
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(5, 4 + i).Value = 0
                '无形资产
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(11, 4 + i).Value = 0
            End If
        Next
        '16—31年
        For i = 16 To 31
            If i <= jsnx Then
                '固定资产
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(22, i - 12).Value = gdzcyz_list(i)
                '无形资产
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(28, i - 12).Value = wxzcyz_list(i)
            Else
                '固定资产
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(22, i - 12).Value = 0
                '无形资产
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(28, i - 12).Value = 0
            End If
        Next
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '计算一次Excel
        ExcelApp.Calculate()
    End Sub
    Function 分项逐年折旧摊销金额计算(ExcelApp As Object, hscz As Boolean, zbj_model As Integer, jsnx As Integer)
        '只计算出分项建设期资金运用的金额数值，不写入EXCEL
        'hscy：计算期末，是否回收资产残值
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        'jsnx：项目计算年限
        '————————————————————————————————————————————————————————————————————————————————————————
        '每次投资资产原值计算
        '读取输入数据
        Dim GSBSJ = 读取估算表数据(ExcelApp)
        '10次投资年份序号列表，列表，长度10
        Dim tznf_list = GSBSJ(0)
        '每次投资的可抵扣增值税金额
        Dim kdkzzs_list = GSBSJ(9)
        'jttz_list：逐年静态投资，列表
        Dim jttz = GSBSJ(1)
        '燃机总投资(万元)
        Dim rjtz = GSBSJ(10)
        '蓄电池总投资(万元)
        Dim xdctz = GSBSJ(12)
        '暖通总投资(万元)
        Dim nttz = GSBSJ(14)
        '光伏总投资(万元)
        Dim gftz = GSBSJ(16)
        '风电总投资(万元)
        Dim fdtz = GSBSJ(18)
        '获取10次投资,每一次投资的投产月份数
        Dim ans_month = 逐年投产月份数_10次投资(ExcelApp)
        Dim month_10_list = ans_month(1)
        '无形资产所占比例
        Dim wxzcbl As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 5).Value
        '————————————————————————————————————————————————————————————————————————————————————————
        '10次投资的资产原值
        Dim gdzcyz(10) As Double
        Dim wxzcyz(10) As Double
        '读取开始年份和资产原值，前5年
        Dim dttz_list = GSBSJ(2)
        For i = 1 To 5
            gdzcyz(i) = (dttz_list(i) - kdkzzs_list(i)) * (1 - wxzcbl)
            wxzcyz(i) = (dttz_list(i) - kdkzzs_list(i)) * wxzcbl
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '分项建设期资金运用计算，列表长度31
        Dim ans_zjyy = 分项建设期资金运用计算(ExcelApp, zbj_model, GSBSJ)
        Dim jsqzjyy_cg = ans_zjyy(0)
        Dim jsqzjyy_rj = ans_zjyy(1)
        Dim jsqzjyy_xdc = ans_zjyy(2)
        Dim jsqzjyy_nt = ans_zjyy(3)
        Dim jsqzjyy_gf = ans_zjyy(4)
        Dim jsqzjyy_fd = ans_zjyy(5)
        '获取分项投资的动态投资金额，列表长度31
        Dim dttz_cg = jsqzjyy_cg(0)
        Dim dttz_rj = jsqzjyy_rj(0)
        Dim dttz_xdc = jsqzjyy_xdc(0)
        Dim dttz_nt = jsqzjyy_nt(0)
        Dim dttz_gf = jsqzjyy_gf(0)
        Dim dttz_fd = jsqzjyy_fd(0)
        '列表长度转为10
        Dim dttz_10_cg = 基础计算功能_31_to_10(tznf_list, dttz_cg)
        Dim dttz_10_rj = 基础计算功能_31_to_10(tznf_list, dttz_rj)
        Dim dttz_10_xdc = 基础计算功能_31_to_10(tznf_list, dttz_xdc)
        Dim dttz_10_nt = 基础计算功能_31_to_10(tznf_list, dttz_nt)
        Dim dttz_10_gf = 基础计算功能_31_to_10(tznf_list, dttz_gf)
        Dim dttz_10_fd = 基础计算功能_31_to_10(tznf_list, dttz_fd)
        '————————————————————————————————————————————————————————————————————————————————————————
        '根据静态投资比例，分摊“建设期可抵扣增值税金额”，长度10的列表
        Dim kdkzzs_cg(10) As Double
        Dim kdkzzs_rj(10) As Double
        Dim kdkzzs_xdc(10) As Double
        Dim kdkzzs_nt(10) As Double
        Dim kdkzzs_gf(10) As Double
        Dim kdkzzs_fd(10) As Double
        For i = 1 To 10
            kdkzzs_cg(i) = kdkzzs_list(i) * (jttz(i) - rjtz(i) - xdctz(i) - nttz(i) - gftz(i) - fdtz(i)) / jttz(i)
            kdkzzs_rj(i) = kdkzzs_list(i) * rjtz(i) / jttz(i)
            kdkzzs_xdc(i) = kdkzzs_list(i) * xdctz(i) / jttz(i)
            kdkzzs_nt(i) = kdkzzs_list(i) * nttz(i) / jttz(i)
            kdkzzs_gf(i) = kdkzzs_list(i) * gftz(i) / jttz(i)
            kdkzzs_fd(i) = kdkzzs_list(i) * fdtz(i) / jttz(i)
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算分项固定资产原值，长度10的列表
        Dim gdzcyz_cg(10) As Double
        Dim gdzcyz_rj(10) As Double
        Dim gdzcyz_xdc(10) As Double
        Dim gdzcyz_nt(10) As Double
        Dim gdzcyz_gf(10) As Double
        Dim gdzcyz_fd(10) As Double
        '计算分项无形资产原值，长度10的列表
        Dim wxzcyz_cg(10) As Double
        Dim wxzcyz_rj(10) As Double
        Dim wxzcyz_xdc(10) As Double
        Dim wxzcyz_nt(10) As Double
        Dim wxzcyz_gf(10) As Double
        Dim wxzcyz_fd(10) As Double
        For i = 1 To 10
            gdzcyz_cg(i) = (dttz_10_cg(i) - kdkzzs_cg(i)) * (1 - wxzcbl)
            gdzcyz_rj(i) = (dttz_10_rj(i) - kdkzzs_rj(i)) * (1 - wxzcbl)
            gdzcyz_xdc(i) = (dttz_10_xdc(i) - kdkzzs_xdc(i)) * (1 - wxzcbl)
            gdzcyz_nt(i) = (dttz_10_nt(i) - kdkzzs_nt(i)) * (1 - wxzcbl)
            gdzcyz_gf(i) = (dttz_10_gf(i) - kdkzzs_gf(i)) * (1 - wxzcbl)
            gdzcyz_fd(i) = (dttz_10_fd(i) - kdkzzs_fd(i)) * (1 - wxzcbl)
        Next
        For i = 1 To 10
            wxzcyz_cg(i) = (dttz_10_cg(i) - kdkzzs_cg(i)) * wxzcbl
            wxzcyz_rj(i) = (dttz_10_rj(i) - kdkzzs_rj(i)) * wxzcbl
            wxzcyz_xdc(i) = (dttz_10_xdc(i) - kdkzzs_xdc(i)) * wxzcbl
            wxzcyz_nt(i) = (dttz_10_nt(i) - kdkzzs_nt(i)) * wxzcbl
            wxzcyz_gf(i) = (dttz_10_gf(i) - kdkzzs_gf(i)) * wxzcbl
            wxzcyz_fd(i) = (dttz_10_fd(i) - kdkzzs_fd(i)) * wxzcbl
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '逐年折旧摊销计算
        '折旧摊销开始年份
        Dim ksnf(10) As Integer
        For i = 1 To 5
            ksnf(i) = tznf_list(i)
        Next
        '10次投资的固定资产折旧年限，无形资产摊销年限
        '常规设备
        Dim gdzczjnx_cg(10) As Integer
        Dim wxzctxnx_cg(10) As Integer
        '燃机
        Dim gdzczjnx_rj(10) As Integer
        Dim wxzctxnx_rj(10) As Integer
        '蓄电池
        Dim gdzczjnx_xdc(10) As Integer
        Dim wxzctxnx_xdc(10) As Integer
        '暖通
        Dim gdzczjnx_nt(10) As Integer
        Dim wxzctxnx_nt(10) As Integer
        '光伏
        Dim gdzczjnx_gf(10) As Integer
        Dim wxzctxnx_gf(10) As Integer
        '风电
        Dim gdzczjnx_fd(10) As Integer
        Dim wxzctxnx_fd(10) As Integer
        '10次投资的固定资产残值率，无形资产残值率
        '常规设备
        Dim gdzcczl_cg(10) As Double
        Dim wxzcczl_cg(10) As Double
        '燃机
        Dim gdzcczl_rj(10) As Double
        Dim wxzcczl_rj(10) As Double
        '蓄电池
        Dim gdzcczl_xdc(10) As Double
        Dim wxzcczl_xdc(10) As Double
        '暖通
        Dim gdzcczl_nt(10) As Double
        Dim wxzcczl_nt(10) As Double
        '光伏
        Dim gdzcczl_gf(10) As Double
        Dim wxzcczl_gf(10) As Double
        '风电
        Dim gdzcczl_fd(10) As Double
        Dim wxzcczl_fd(10) As Double
        '读取计算参数
        '常规设备
        For i = 4 To 13
            gdzczjnx_cg(i - 3) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 22).Value
            wxzctxnx_cg(i - 3) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 26).Value
            gdzcczl_cg(i - 3) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 24).Value
            wxzcczl_cg(i - 3) = 0
        Next
        '燃机
        For i = 1 To 10
            gdzczjnx_rj(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i + 3, 29).Value
            gdzcczl_rj(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i + 14, 29).Value
            wxzctxnx_rj(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i + 25, 29).Value
            wxzcczl_rj(i) = 0
        Next
        '蓄电池
        For i = 1 To 10
            gdzczjnx_xdc(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i + 3, 32).Value
            gdzcczl_xdc(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i + 14, 32).Value
            wxzctxnx_xdc(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i + 25, 32).Value
            wxzcczl_xdc(i) = 0
        Next
        '暖通
        For i = 1 To 10
            gdzczjnx_nt(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i + 3, 35).Value
            gdzcczl_nt(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i + 14, 35).Value
            wxzctxnx_nt(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i + 25, 35).Value
            wxzcczl_nt(i) = 0
        Next
        '光伏
        For i = 1 To 10
            gdzczjnx_gf(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i + 3, 38).Value
            gdzcczl_gf(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i + 14, 38).Value
            wxzctxnx_gf(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i + 25, 38).Value
            wxzcczl_gf(i) = 0
        Next
        '风电
        For i = 1 To 10
            gdzczjnx_fd(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i + 3, 41).Value
            gdzcczl_fd(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i + 14, 41).Value
            wxzctxnx_fd(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i + 25, 41).Value
            wxzcczl_fd(i) = 0
        Next
        '————————————————————————————————————————————————————————————————————————————————————————   
        '折旧摊销计算
        Dim ans_gdzczj_cg
        Dim ans_wxzctx_cg
        Dim ans_gdzczj_rj
        Dim ans_wxzctx_rj
        Dim ans_gdzczj_xdc
        Dim ans_wxzctx_xdc
        Dim ans_gdzczj_nt
        Dim ans_wxzctx_nt
        Dim ans_gdzczj_gf
        Dim ans_wxzctx_gf
        Dim ans_gdzczj_fd
        Dim ans_wxzctx_fd
        '方法一
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法一" Then
            '常规设备
            ans_gdzczj_cg = 折旧摊销计算_直线法_10次投资合并计算(gdzcyz_cg, gdzczjnx_cg(1), ksnf, jsnx, gdzcczl_cg(1), month_10_list, hscz)
            ans_wxzctx_cg = 折旧摊销计算_直线法_10次投资合并计算(wxzcyz_cg, wxzctxnx_cg(1), ksnf, jsnx, wxzcczl_cg(1), month_10_list, hscz)
            '燃机
            ans_gdzczj_rj = 折旧摊销计算_直线法_10次投资合并计算(gdzcyz_rj, gdzczjnx_rj(1), ksnf, jsnx, gdzcczl_rj(1), month_10_list, hscz)
            ans_wxzctx_rj = 折旧摊销计算_直线法_10次投资合并计算(wxzcyz_rj, wxzctxnx_rj(1), ksnf, jsnx, wxzcczl_rj(1), month_10_list, hscz)
            '蓄电池
            ans_gdzczj_xdc = 折旧摊销计算_直线法_10次投资合并计算(gdzcyz_xdc, gdzczjnx_xdc(1), ksnf, jsnx, gdzcczl_xdc(1), month_10_list, hscz)
            ans_wxzctx_xdc = 折旧摊销计算_直线法_10次投资合并计算(wxzcyz_xdc, wxzctxnx_xdc(1), ksnf, jsnx, wxzcczl_xdc(1), month_10_list, hscz)
            '暖通
            ans_gdzczj_nt = 折旧摊销计算_直线法_10次投资合并计算(gdzcyz_nt, gdzczjnx_nt(1), ksnf, jsnx, gdzcczl_nt(1), month_10_list, hscz)
            ans_wxzctx_nt = 折旧摊销计算_直线法_10次投资合并计算(wxzcyz_nt, wxzctxnx_nt(1), ksnf, jsnx, wxzcczl_nt(1), month_10_list, hscz)
            '光伏
            ans_gdzczj_gf = 折旧摊销计算_直线法_10次投资合并计算(gdzcyz_gf, gdzczjnx_gf(1), ksnf, jsnx, gdzcczl_gf(1), month_10_list, hscz)
            ans_wxzctx_gf = 折旧摊销计算_直线法_10次投资合并计算(wxzcyz_gf, wxzctxnx_gf(1), ksnf, jsnx, wxzcczl_gf(1), month_10_list, hscz)
            '风电
            ans_gdzczj_fd = 折旧摊销计算_直线法_10次投资合并计算(gdzcyz_fd, gdzczjnx_fd(1), ksnf, jsnx, gdzcczl_fd(1), month_10_list, hscz)
            ans_wxzctx_fd = 折旧摊销计算_直线法_10次投资合并计算(wxzcyz_fd, wxzctxnx_fd(1), ksnf, jsnx, wxzcczl_fd(1), month_10_list, hscz)
        End If
        '方法二
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法二" Then
            '常规设备
            ans_gdzczj_cg = 折旧摊销计算_年数总和法_10次投资合并计算(gdzcyz_cg, gdzczjnx_cg(1), ksnf, jsnx, gdzcczl_cg(1), month_10_list, hscz)
            ans_wxzctx_cg = 折旧摊销计算_年数总和法_10次投资合并计算(wxzcyz_cg, wxzctxnx_cg(1), ksnf, jsnx, wxzcczl_cg(1), month_10_list, hscz)
            '燃机
            ans_gdzczj_rj = 折旧摊销计算_年数总和法_10次投资合并计算(gdzcyz_rj, gdzczjnx_rj(1), ksnf, jsnx, gdzcczl_rj(1), month_10_list, hscz)
            ans_wxzctx_rj = 折旧摊销计算_年数总和法_10次投资合并计算(wxzcyz_rj, wxzctxnx_rj(1), ksnf, jsnx, wxzcczl_rj(1), month_10_list, hscz)
            '蓄电池
            ans_gdzczj_xdc = 折旧摊销计算_年数总和法_10次投资合并计算(gdzcyz_xdc, gdzczjnx_xdc(1), ksnf, jsnx, gdzcczl_xdc(1), month_10_list, hscz)
            ans_wxzctx_xdc = 折旧摊销计算_年数总和法_10次投资合并计算(wxzcyz_xdc, wxzctxnx_xdc(1), ksnf, jsnx, wxzcczl_xdc(1), month_10_list, hscz)
            '暖通
            ans_gdzczj_nt = 折旧摊销计算_年数总和法_10次投资合并计算(gdzcyz_nt, gdzczjnx_nt(1), ksnf, jsnx, gdzcczl_nt(1), month_10_list, hscz)
            ans_wxzctx_nt = 折旧摊销计算_年数总和法_10次投资合并计算(wxzcyz_nt, wxzctxnx_nt(1), ksnf, jsnx, wxzcczl_nt(1), month_10_list, hscz)
            '光伏
            ans_gdzczj_gf = 折旧摊销计算_年数总和法_10次投资合并计算(gdzcyz_gf, gdzczjnx_gf(1), ksnf, jsnx, gdzcczl_gf(1), month_10_list, hscz)
            ans_wxzctx_gf = 折旧摊销计算_年数总和法_10次投资合并计算(wxzcyz_gf, wxzctxnx_gf(1), ksnf, jsnx, wxzcczl_gf(1), month_10_list, hscz)
            '风电
            ans_gdzczj_fd = 折旧摊销计算_年数总和法_10次投资合并计算(gdzcyz_fd, gdzczjnx_fd(1), ksnf, jsnx, gdzcczl_fd(1), month_10_list, hscz)
            ans_wxzctx_fd = 折旧摊销计算_年数总和法_10次投资合并计算(wxzcyz_fd, wxzctxnx_fd(1), ksnf, jsnx, wxzcczl_fd(1), month_10_list, hscz)
        End If
        '方法三
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法三" Then
            '常规设备
            ans_gdzczj_cg = 折旧摊销计算_直线法_10次投资分开计算(gdzcyz_cg, gdzczjnx_cg, ksnf, jsnx, gdzcczl_cg, month_10_list, hscz)
            ans_wxzctx_cg = 折旧摊销计算_直线法_10次投资分开计算(wxzcyz_cg, wxzctxnx_cg, ksnf, jsnx, wxzcczl_cg, month_10_list, hscz)
            '燃机
            ans_gdzczj_rj = 折旧摊销计算_直线法_10次投资分开计算(gdzcyz_rj, gdzczjnx_rj, ksnf, jsnx, gdzcczl_rj, month_10_list, hscz)
            ans_wxzctx_rj = 折旧摊销计算_直线法_10次投资分开计算(wxzcyz_rj, wxzctxnx_rj, ksnf, jsnx, wxzcczl_rj, month_10_list, hscz)
            '蓄电池
            ans_gdzczj_xdc = 折旧摊销计算_直线法_10次投资分开计算(gdzcyz_xdc, gdzczjnx_xdc, ksnf, jsnx, gdzcczl_xdc, month_10_list, hscz)
            ans_wxzctx_xdc = 折旧摊销计算_直线法_10次投资分开计算(wxzcyz_xdc, wxzctxnx_xdc, ksnf, jsnx, wxzcczl_xdc, month_10_list, hscz)
            '暖通
            ans_gdzczj_nt = 折旧摊销计算_直线法_10次投资分开计算(gdzcyz_nt, gdzczjnx_nt, ksnf, jsnx, gdzcczl_nt, month_10_list, hscz)
            ans_wxzctx_nt = 折旧摊销计算_直线法_10次投资分开计算(wxzcyz_nt, wxzctxnx_nt, ksnf, jsnx, wxzcczl_nt, month_10_list, hscz)
            '光伏
            ans_gdzczj_gf = 折旧摊销计算_直线法_10次投资分开计算(gdzcyz_gf, gdzczjnx_gf, ksnf, jsnx, gdzcczl_gf, month_10_list, hscz)
            ans_wxzctx_gf = 折旧摊销计算_直线法_10次投资分开计算(wxzcyz_gf, wxzctxnx_gf, ksnf, jsnx, wxzcczl_gf, month_10_list, hscz)
            '风电
            ans_gdzczj_fd = 折旧摊销计算_直线法_10次投资分开计算(gdzcyz_fd, gdzczjnx_fd, ksnf, jsnx, gdzcczl_fd, month_10_list, hscz)
            ans_wxzctx_fd = 折旧摊销计算_直线法_10次投资分开计算(wxzcyz_fd, wxzctxnx_fd, ksnf, jsnx, wxzcczl_fd, month_10_list, hscz)
        End If
        '方法四
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法四" Then
            '常规设备
            ans_gdzczj_cg = 折旧摊销计算_年数总和法_10次投资分开计算(gdzcyz_cg, gdzczjnx_cg, ksnf, jsnx, gdzcczl_cg, month_10_list, hscz)
            ans_wxzctx_cg = 折旧摊销计算_年数总和法_10次投资分开计算(wxzcyz_cg, wxzctxnx_cg, ksnf, jsnx, wxzcczl_cg, month_10_list, hscz)
            '燃机
            ans_gdzczj_rj = 折旧摊销计算_年数总和法_10次投资分开计算(gdzcyz_rj, gdzczjnx_rj, ksnf, jsnx, gdzcczl_rj, month_10_list, hscz)
            ans_wxzctx_rj = 折旧摊销计算_年数总和法_10次投资分开计算(wxzcyz_rj, wxzctxnx_rj, ksnf, jsnx, wxzcczl_rj, month_10_list, hscz)
            '蓄电池
            ans_gdzczj_xdc = 折旧摊销计算_年数总和法_10次投资分开计算(gdzcyz_xdc, gdzczjnx_xdc, ksnf, jsnx, gdzcczl_xdc, month_10_list, hscz)
            ans_wxzctx_xdc = 折旧摊销计算_年数总和法_10次投资分开计算(wxzcyz_xdc, wxzctxnx_xdc, ksnf, jsnx, wxzcczl_xdc, month_10_list, hscz)
            '暖通
            ans_gdzczj_nt = 折旧摊销计算_年数总和法_10次投资分开计算(gdzcyz_nt, gdzczjnx_nt, ksnf, jsnx, gdzcczl_nt, month_10_list, hscz)
            ans_wxzctx_nt = 折旧摊销计算_年数总和法_10次投资分开计算(wxzcyz_nt, wxzctxnx_nt, ksnf, jsnx, wxzcczl_nt, month_10_list, hscz)
            '光伏
            ans_gdzczj_gf = 折旧摊销计算_年数总和法_10次投资分开计算(gdzcyz_gf, gdzczjnx_gf, ksnf, jsnx, gdzcczl_gf, month_10_list, hscz)
            ans_wxzctx_gf = 折旧摊销计算_年数总和法_10次投资分开计算(wxzcyz_gf, wxzctxnx_gf, ksnf, jsnx, wxzcczl_gf, month_10_list, hscz)
            '风电
            ans_gdzczj_fd = 折旧摊销计算_年数总和法_10次投资分开计算(gdzcyz_fd, gdzczjnx_fd, ksnf, jsnx, gdzcczl_fd, month_10_list, hscz)
            ans_wxzctx_fd = 折旧摊销计算_年数总和法_10次投资分开计算(wxzcyz_fd, wxzctxnx_fd, ksnf, jsnx, wxzcczl_fd, month_10_list, hscz)
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '返回结果
        Dim ans(25)
        ans(0) = gdzcyz
        ans(1) = wxzcyz
        ans(2) = gdzcyz_cg
        ans(3) = gdzcyz_rj
        ans(4) = gdzcyz_xdc
        ans(5) = gdzcyz_nt
        ans(6) = gdzcyz_gf
        ans(7) = gdzcyz_fd
        ans(8) = wxzcyz_cg
        ans(9) = wxzcyz_rj
        ans(10) = wxzcyz_xdc
        ans(11) = wxzcyz_nt
        ans(12) = wxzcyz_gf
        ans(13) = wxzcyz_fd
        ans(14) = ans_gdzczj_cg
        ans(15) = ans_gdzczj_rj
        ans(16) = ans_gdzczj_xdc
        ans(17) = ans_gdzczj_nt
        ans(18) = ans_gdzczj_gf
        ans(19) = ans_gdzczj_fd
        ans(20) = ans_wxzctx_cg
        ans(21) = ans_wxzctx_rj
        ans(22) = ans_wxzctx_xdc
        ans(23) = ans_wxzctx_nt
        ans(24) = ans_wxzctx_gf
        ans(25) = ans_wxzctx_fd
        Return ans
    End Function

    Function 折旧摊销计算_直线法_10次投资合并计算(zcyz_list As Array, zjtxnx_0 As Integer, ksnf_list As Array,
                                                  jsnx As Integer, czl As Double, month_10_list As Array, hscz As Boolean)
        'zcyz_list：资产原值，列表，长度10
        'zjtxnx_0：折旧摊销年限数初始值
        'ksnf_list：折旧摊销计算的开始年份，列表，长度10
        'jsnx：项目总的计算年限
        'czl：残值率
        'month_10_list：10次投资，计算每次投资的逐年投产月份数，列表，长度10
        'hscy：计算期末，是否回收资产残值

        '10次投资的固定资产折旧或者无形资产摊销合并在一起进行计算

        '第1次投资
        Dim zcyz_1 As Double = zcyz_list(1)
        Dim ksnf_1 As Integer = ksnf_list(1)
        Dim zjtxnx_0_1 As Integer
        If zjtxnx_0 + 1 - ksnf_1 > 0 Then
            zjtxnx_0_1 = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_1)
        Else
            zjtxnx_0_1 = Math.Min(zjtxnx_0, jsnx - ksnf_1)
        End If
        '第2次投资
        Dim zcyz_2 As Double = zcyz_list(2)
        Dim ksnf_2 As Integer = ksnf_list(2)
        Dim zjtxnx_0_2 As Integer
        If zjtxnx_0 + 1 - ksnf_2 > 0 Then
            zjtxnx_0_2 = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_2)
        Else
            zjtxnx_0_2 = Math.Min(zjtxnx_0, jsnx - ksnf_2)
        End If
        '第3次投资
        Dim zcyz_3 As Double = zcyz_list(3)
        Dim ksnf_3 As Integer = ksnf_list(3)
        Dim zjtxnx_0_3 As Integer
        If zjtxnx_0 + 1 - ksnf_3 > 0 Then
            zjtxnx_0_3 = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_3)
        Else
            zjtxnx_0_3 = Math.Min(zjtxnx_0, jsnx - ksnf_3)
        End If
        '第4次投资
        Dim zcyz_4 As Double = zcyz_list(4)
        Dim ksnf_4 As Integer = ksnf_list(4)
        Dim zjtxnx_0_4 As Integer
        If zjtxnx_0 + 1 - ksnf_4 > 0 Then
            zjtxnx_0_4 = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_4)
        Else
            zjtxnx_0_4 = Math.Min(zjtxnx_0, jsnx - ksnf_4)
        End If
        '第5次投资
        Dim zcyz_5 As Double = zcyz_list(5)
        Dim ksnf_5 As Integer = ksnf_list(5)
        Dim zjtxnx_0_5 As Integer
        If zjtxnx_0 + 1 - ksnf_5 > 0 Then
            zjtxnx_0_5 = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_5)
        Else
            zjtxnx_0_5 = Math.Min(zjtxnx_0, jsnx - ksnf_5)
        End If
        '第6次投资
        Dim zcyz_6 As Double = zcyz_list(6)
        Dim ksnf_6 As Integer = ksnf_list(6)
        Dim zjtxnx_0_6 As Integer
        If zjtxnx_0 + 1 - ksnf_6 > 0 Then
            zjtxnx_0_6 = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_6)
        Else
            zjtxnx_0_6 = Math.Min(zjtxnx_0, jsnx - ksnf_6)
        End If
        '第7次投资
        Dim zcyz_7 As Double = zcyz_list(7)
        Dim ksnf_7 As Integer = ksnf_list(7)
        Dim zjtxnx_0_7 As Integer
        If zjtxnx_0 + 1 - ksnf_7 > 0 Then
            zjtxnx_0_7 = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_7)
        Else
            zjtxnx_0_7 = Math.Min(zjtxnx_0, jsnx - ksnf_7)
        End If
        '第8次投资
        Dim zcyz_8 As Double = zcyz_list(8)
        Dim ksnf_8 As Integer = ksnf_list(8)
        Dim zjtxnx_0_8 As Integer
        If zjtxnx_0 + 1 - ksnf_8 > 0 Then
            zjtxnx_0_8 = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_8)
        Else
            zjtxnx_0_8 = Math.Min(zjtxnx_0, jsnx - ksnf_8)
        End If
        '第9次投资
        Dim zcyz_9 As Double = zcyz_list(9)
        Dim ksnf_9 As Integer = ksnf_list(9)
        Dim zjtxnx_0_9 As Integer
        If zjtxnx_0 + 1 - ksnf_9 > 0 Then
            zjtxnx_0_9 = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_9)
        Else
            zjtxnx_0_9 = Math.Min(zjtxnx_0, jsnx - ksnf_9)
        End If
        '第10次投资
        Dim zcyz_10 As Double = zcyz_list(10)
        Dim ksnf_10 As Integer = ksnf_list(10)
        Dim zjtxnx_0_10 As Integer
        If zjtxnx_0 + 1 - ksnf_10 > 0 Then
            zjtxnx_0_10 = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_10)
        Else
            zjtxnx_0_10 = Math.Min(zjtxnx_0, jsnx - ksnf_10)
        End If

        '列表，储存计算结果
        Dim ZJTXF(31) As Double '逐年折旧摊销费金额
        Dim ZJTXFLJ(31) As Double '逐年折旧摊销费金额累计
        Dim SYJZ(31) As Double '逐年剩余资产净值

        '分次进行计算
        Dim ans_1 = 折旧摊销计算_直线法(zcyz_1, zjtxnx_0_1, ksnf_1, jsnx, czl, month_10_list(1))
        Dim ans_2 = 折旧摊销计算_直线法(zcyz_2, zjtxnx_0_2, ksnf_2, jsnx, czl, month_10_list(2))
        Dim ans_3 = 折旧摊销计算_直线法(zcyz_3, zjtxnx_0_3, ksnf_3, jsnx, czl, month_10_list(3))
        Dim ans_4 = 折旧摊销计算_直线法(zcyz_4, zjtxnx_0_4, ksnf_4, jsnx, czl, month_10_list(4))
        Dim ans_5 = 折旧摊销计算_直线法(zcyz_5, zjtxnx_0_5, ksnf_5, jsnx, czl, month_10_list(5))
        Dim ans_6 = 折旧摊销计算_直线法(zcyz_6, zjtxnx_0_6, ksnf_6, jsnx, czl, month_10_list(6))
        Dim ans_7 = 折旧摊销计算_直线法(zcyz_7, zjtxnx_0_7, ksnf_7, jsnx, czl, month_10_list(7))
        Dim ans_8 = 折旧摊销计算_直线法(zcyz_8, zjtxnx_0_8, ksnf_8, jsnx, czl, month_10_list(8))
        Dim ans_9 = 折旧摊销计算_直线法(zcyz_9, zjtxnx_0_9, ksnf_9, jsnx, czl, month_10_list(9))
        Dim ans_10 = 折旧摊销计算_直线法(zcyz_10, zjtxnx_0_10, ksnf_10, jsnx, czl, month_10_list(10))
        '累加
        For i = 1 To 31 '第1年到第31年
            '逐年折旧摊销费金额
            ZJTXF(i) = ans_1(0)(i) + ans_2(0)(i) + ans_3(0)(i) + ans_4(0)(i) + ans_5(0)(i) + ans_6(0)(i) +
                       ans_7(0)(i) + ans_8(0)(i) + ans_9(0)(i) + ans_10(0)(i)
            '逐年折旧摊销费金额累计
            ZJTXFLJ(i) = ans_1(1)(i) + ans_2(1)(i) + ans_3(1)(i) + ans_4(1)(i) + ans_5(1)(i) + ans_6(1)(i) +
                         ans_7(1)(i) + ans_8(1)(i) + ans_9(1)(i) + ans_10(1)(i)
            '逐年剩余资产净值
            SYJZ(i) = ans_1(2)(i) + ans_2(2)(i) + ans_3(2)(i) + ans_4(2)(i) + ans_5(2)(i) + ans_6(2)(i) +
                      ans_7(2)(i) + ans_8(2)(i) + ans_9(2)(i) + ans_10(2)(i)
        Next

        '判断计算期末如果不回收残值，则进行调整
        If hscz = False Then
            '计算期最后1年的，折旧摊销费金额
            ZJTXF(jsnx) = SYJZ(jsnx)
            '计算期最后1年的，剩余资产净值
            SYJZ(jsnx) = 0
            '计算期最后1年的，折旧摊销费金额累计
            ZJTXFLJ(jsnx) += ZJTXF(jsnx)
        End If

        '返回结果
        Dim ans(2)
        ans(0) = ZJTXF
        ans(1) = ZJTXFLJ
        ans(2) = SYJZ
        Return ans
    End Function
    Function 折旧摊销计算_年数总和法_10次投资合并计算(zcyz_list As Array, zjtxnx_0 As Integer, ksnf_list As Array,
                                                      jsnx As Integer, czl As Double, month_10_list As Array, hscz As Boolean)

        'zcyz_list：资产原值，列表，长度10
        'zjtxnx_0：折旧摊销年限数初始值
        'ksnf_list：折旧摊销计算的开始年份初始值，列表，长度10
        'jsnx：项目总的计算年限
        'czl：残值率
        'month_10_list：10次投资，计算每次投资的逐年投产月份数，列表，长度10
        'hscy：计算期末，是否回收资产残值

        '10次投资的固定资产折旧或者无形资产摊销合并在一起进行计算

        '第1次投资
        Dim zcyz_1 As Double = zcyz_list(1)
        Dim ksnf_1 As Integer = ksnf_list(1)
        Dim zjtxnx_0_1 As Integer
        If zjtxnx_0 + 1 - ksnf_1 > 0 Then
            zjtxnx_0_1 = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_1)
        Else
            zjtxnx_0_1 = Math.Min(zjtxnx_0, jsnx - ksnf_1)
        End If
        '第2次投资
        Dim zcyz_2 As Double = zcyz_list(2)
        Dim ksnf_2 As Integer = ksnf_list(2)
        Dim zjtxnx_0_2 As Integer
        If zjtxnx_0 + 1 - ksnf_2 > 0 Then
            zjtxnx_0_2 = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_2)
        Else
            zjtxnx_0_2 = Math.Min(zjtxnx_0, jsnx - ksnf_2)
        End If
        '第3次投资
        Dim zcyz_3 As Double = zcyz_list(3)
        Dim ksnf_3 As Integer = ksnf_list(3)
        Dim zjtxnx_0_3 As Integer
        If zjtxnx_0 + 1 - ksnf_3 > 0 Then
            zjtxnx_0_3 = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_3)
        Else
            zjtxnx_0_3 = Math.Min(zjtxnx_0, jsnx - ksnf_3)
        End If
        '第4次投资
        Dim zcyz_4 As Double = zcyz_list(4)
        Dim ksnf_4 As Integer = ksnf_list(4)
        Dim zjtxnx_0_4 As Integer
        If zjtxnx_0 + 1 - ksnf_4 > 0 Then
            zjtxnx_0_4 = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_4)
        Else
            zjtxnx_0_4 = Math.Min(zjtxnx_0, jsnx - ksnf_4)
        End If
        '第5次投资
        Dim zcyz_5 As Double = zcyz_list(5)
        Dim ksnf_5 As Integer = ksnf_list(5)
        Dim zjtxnx_0_5 As Integer
        If zjtxnx_0 + 1 - ksnf_5 > 0 Then
            zjtxnx_0_5 = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_5)
        Else
            zjtxnx_0_5 = Math.Min(zjtxnx_0, jsnx - ksnf_5)
        End If
        '第6次投资
        Dim zcyz_6 As Double = zcyz_list(6)
        Dim ksnf_6 As Integer = ksnf_list(6)
        Dim zjtxnx_0_6 As Integer
        If zjtxnx_0 + 1 - ksnf_6 > 0 Then
            zjtxnx_0_6 = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_6)
        Else
            zjtxnx_0_6 = Math.Min(zjtxnx_0, jsnx - ksnf_6)
        End If
        '第7次投资
        Dim zcyz_7 As Double = zcyz_list(7)
        Dim ksnf_7 As Integer = ksnf_list(7)
        Dim zjtxnx_0_7 As Integer
        If zjtxnx_0 + 1 - ksnf_7 > 0 Then
            zjtxnx_0_7 = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_7)
        Else
            zjtxnx_0_7 = Math.Min(zjtxnx_0, jsnx - ksnf_7)
        End If
        '第8次投资
        Dim zcyz_8 As Double = zcyz_list(8)
        Dim ksnf_8 As Integer = ksnf_list(8)
        Dim zjtxnx_0_8 As Integer
        If zjtxnx_0 + 1 - ksnf_8 > 0 Then
            zjtxnx_0_8 = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_8)
        Else
            zjtxnx_0_8 = Math.Min(zjtxnx_0, jsnx - ksnf_8)
        End If
        '第9次投资
        Dim zcyz_9 As Double = zcyz_list(9)
        Dim ksnf_9 As Integer = ksnf_list(9)
        Dim zjtxnx_0_9 As Integer
        If zjtxnx_0 + 1 - ksnf_9 > 0 Then
            zjtxnx_0_9 = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_9)
        Else
            zjtxnx_0_9 = Math.Min(zjtxnx_0, jsnx - ksnf_9)
        End If
        '第10次投资
        Dim zcyz_10 As Double = zcyz_list(10)
        Dim ksnf_10 As Integer = ksnf_list(10)
        Dim zjtxnx_0_10 As Integer
        If zjtxnx_0 + 1 - ksnf_10 > 0 Then
            zjtxnx_0_10 = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_10)
        Else
            zjtxnx_0_10 = Math.Min(zjtxnx_0, jsnx - ksnf_10)
        End If

        '列表，储存计算过程量
        Dim ZJTXF(31) As Double '逐年折旧摊销费金额
        Dim ZJTXFLJ(31) As Double '逐年折旧摊销费金额累计
        Dim SYJZ(31) As Double '逐年剩余资产净值
        Dim ZJL(31) As Double '逐年折旧率

        '分次进行计算
        Dim ans_1 = 折旧摊销计算_年数总和法(zcyz_1, zjtxnx_0_1, ksnf_1, jsnx, czl, month_10_list(1))
        Dim ans_2 = 折旧摊销计算_年数总和法(zcyz_2, zjtxnx_0_2, ksnf_2, jsnx, czl, month_10_list(2))
        Dim ans_3 = 折旧摊销计算_年数总和法(zcyz_3, zjtxnx_0_3, ksnf_3, jsnx, czl, month_10_list(3))
        Dim ans_4 = 折旧摊销计算_年数总和法(zcyz_4, zjtxnx_0_4, ksnf_4, jsnx, czl, month_10_list(4))
        Dim ans_5 = 折旧摊销计算_年数总和法(zcyz_5, zjtxnx_0_5, ksnf_5, jsnx, czl, month_10_list(5))
        Dim ans_6 = 折旧摊销计算_年数总和法(zcyz_6, zjtxnx_0_6, ksnf_6, jsnx, czl, month_10_list(6))
        Dim ans_7 = 折旧摊销计算_年数总和法(zcyz_7, zjtxnx_0_7, ksnf_7, jsnx, czl, month_10_list(7))
        Dim ans_8 = 折旧摊销计算_年数总和法(zcyz_8, zjtxnx_0_8, ksnf_8, jsnx, czl, month_10_list(8))
        Dim ans_9 = 折旧摊销计算_年数总和法(zcyz_9, zjtxnx_0_9, ksnf_9, jsnx, czl, month_10_list(9))
        Dim ans_10 = 折旧摊销计算_年数总和法(zcyz_10, zjtxnx_0_10, ksnf_10, jsnx, czl, month_10_list(10))

        '累加
        For i = 1 To 31 '第1年到第31年
            '逐年折旧摊销费金额
            ZJTXF(i) = ans_1(0)(i) + ans_2(0)(i) + ans_3(0)(i) + ans_4(0)(i) + ans_5(0)(i) + ans_6(0)(i) +
                       ans_7(0)(i) + ans_8(0)(i) + ans_9(0)(i) + ans_10(0)(i)
            '逐年折旧摊销费金额累计
            ZJTXFLJ(i) = ans_1(1)(i) + ans_2(1)(i) + ans_3(1)(i) + ans_4(1)(i) + ans_5(1)(i) + ans_6(1)(i) +
                         ans_7(1)(i) + ans_8(1)(i) + ans_9(1)(i) + ans_10(1)(i)
            '逐年剩余资产净值
            SYJZ(i) = ans_1(2)(i) + ans_2(2)(i) + ans_3(2)(i) + ans_4(2)(i) + ans_5(2)(i) + ans_6(2)(i) +
                      ans_7(2)(i) + ans_8(2)(i) + ans_9(2)(i) + ans_10(2)(i)
            '逐年折旧率
            ZJL(i) = ans_1(3)(i) + ans_2(3)(i) + ans_3(3)(i) + ans_4(3)(i) + ans_5(3)(i) + ans_6(3)(i) +
                     ans_7(3)(i) + ans_8(3)(i) + ans_9(3)(i) + ans_10(3)(i)
        Next

        '判断计算期末如果不回收残值，则进行调整
        If hscz = False Then
            '计算期最后1年的，折旧摊销费金额
            ZJTXF(jsnx) = SYJZ(jsnx)
            '计算期最后1年的，剩余资产净值
            SYJZ(jsnx) = 0
            '计算期最后1年的，折旧摊销费金额累计
            ZJTXFLJ(jsnx) += ZJTXF(jsnx)
            '计算期最后1年的，折旧率
            ZJL(jsnx) = 1
        End If

        '返回结果
        Dim ans(3)
        ans(0) = ZJTXF
        ans(1) = ZJTXFLJ
        ans(2) = SYJZ
        ans(3) = ZJL
        Return ans
    End Function
    Function 折旧摊销计算_直线法_10次投资分开计算(zcyz_list As Array, zjtxnx_0_list As Array, ksnf_list As Array,
                                                  jsnx As Integer, czl_list As Array, month_10_list As Array, hscz As Boolean)
        'zcyz_list：资产原值，列表，长度10
        'zjtxnx_0_list：折旧摊销年限数初始值，列表，长度10
        'ksnf_list：折旧摊销计算的开始年份，列表，长度10
        'jsnx：项目总的计算年限
        'czl_list：残值率，列表，长度10
        'month_10_list：10次投资，计算每次投资的逐年投产月份数，列表，长度10
        'hscy：计算期末，是否回收资产残值

        '10次投资分别计算固定资产折旧或者无形资产摊销，再累加在一起

        '第1次投资
        Dim zcyz_1 As Double = zcyz_list(1)
        Dim zjtxnx_0_1 As Integer = zjtxnx_0_list(1)
        Dim ksnf_1 As Integer = ksnf_list(1)
        Dim czl_1 As Double = czl_list(1)
        '第2次投资
        Dim zcyz_2 As Double = zcyz_list(2)
        Dim zjtxnx_0_2 As Integer = zjtxnx_0_list(2)
        Dim ksnf_2 As Integer = ksnf_list(2)
        Dim czl_2 As Double = czl_list(2)
        '第3次投资
        Dim zcyz_3 As Double = zcyz_list(3)
        Dim zjtxnx_0_3 As Integer = zjtxnx_0_list(3)
        Dim ksnf_3 As Integer = ksnf_list(3)
        Dim czl_3 As Double = czl_list(3)
        '第4次投资
        Dim zcyz_4 As Double = zcyz_list(4)
        Dim zjtxnx_0_4 As Integer = zjtxnx_0_list(4)
        Dim ksnf_4 As Integer = ksnf_list(4)
        Dim czl_4 As Double = czl_list(4)
        '第5次投资
        Dim zcyz_5 As Double = zcyz_list(5)
        Dim zjtxnx_0_5 As Integer = zjtxnx_0_list(5)
        Dim ksnf_5 As Integer = ksnf_list(5)
        Dim czl_5 As Double = czl_list(5)
        '第6次投资
        Dim zcyz_6 As Double = zcyz_list(6)
        Dim zjtxnx_0_6 As Integer = zjtxnx_0_list(6)
        Dim ksnf_6 As Integer = ksnf_list(6)
        Dim czl_6 As Double = czl_list(6)
        '第7次投资
        Dim zcyz_7 As Double = zcyz_list(7)
        Dim zjtxnx_0_7 As Integer = zjtxnx_0_list(7)
        Dim ksnf_7 As Integer = ksnf_list(7)
        Dim czl_7 As Double = czl_list(7)
        '第8次投资
        Dim zcyz_8 As Double = zcyz_list(8)
        Dim zjtxnx_0_8 As Integer = zjtxnx_0_list(8)
        Dim ksnf_8 As Integer = ksnf_list(8)
        Dim czl_8 As Double = czl_list(8)
        '第9次投资
        Dim zcyz_9 As Double = zcyz_list(9)
        Dim zjtxnx_0_9 As Integer = zjtxnx_0_list(9)
        Dim ksnf_9 As Integer = ksnf_list(9)
        Dim czl_9 As Double = czl_list(9)
        '第10次投资
        Dim zcyz_10 As Double = zcyz_list(10)
        Dim zjtxnx_0_10 As Integer = zjtxnx_0_list(10)
        Dim ksnf_10 As Integer = ksnf_list(10)
        Dim czl_10 As Double = czl_list(10)

        '列表，储存计算结果
        Dim ZJTXF(31) As Double '逐年折旧摊销费金额
        Dim ZJTXFLJ(31) As Double '逐年折旧摊销费金额累计
        Dim SYJZ(31) As Double '逐年剩余资产净值

        '分次进行计算
        Dim ans_1 = 折旧摊销计算_直线法(zcyz_1, zjtxnx_0_1, ksnf_1, jsnx, czl_1, month_10_list(1))
        Dim ans_2 = 折旧摊销计算_直线法(zcyz_2, zjtxnx_0_2, ksnf_2, jsnx, czl_2, month_10_list(2))
        Dim ans_3 = 折旧摊销计算_直线法(zcyz_3, zjtxnx_0_3, ksnf_3, jsnx, czl_3, month_10_list(3))
        Dim ans_4 = 折旧摊销计算_直线法(zcyz_4, zjtxnx_0_4, ksnf_4, jsnx, czl_4, month_10_list(4))
        Dim ans_5 = 折旧摊销计算_直线法(zcyz_5, zjtxnx_0_5, ksnf_5, jsnx, czl_5, month_10_list(5))
        Dim ans_6 = 折旧摊销计算_直线法(zcyz_6, zjtxnx_0_6, ksnf_6, jsnx, czl_6, month_10_list(6))
        Dim ans_7 = 折旧摊销计算_直线法(zcyz_7, zjtxnx_0_7, ksnf_7, jsnx, czl_7, month_10_list(7))
        Dim ans_8 = 折旧摊销计算_直线法(zcyz_8, zjtxnx_0_8, ksnf_8, jsnx, czl_8, month_10_list(8))
        Dim ans_9 = 折旧摊销计算_直线法(zcyz_9, zjtxnx_0_9, ksnf_9, jsnx, czl_9, month_10_list(9))
        Dim ans_10 = 折旧摊销计算_直线法(zcyz_10, zjtxnx_0_10, ksnf_10, jsnx, czl_10, month_10_list(10))
        '累加
        For i = 1 To 31 '第1年到第31年
            '逐年折旧摊销费金额
            ZJTXF(i) = ans_1(0)(i) + ans_2(0)(i) + ans_3(0)(i) + ans_4(0)(i) + ans_5(0)(i) + ans_6(0)(i) +
                       ans_7(0)(i) + ans_8(0)(i) + ans_9(0)(i) + ans_10(0)(i)
            '逐年折旧摊销费金额累计
            ZJTXFLJ(i) = ans_1(1)(i) + ans_2(1)(i) + ans_3(1)(i) + ans_4(1)(i) + ans_5(1)(i) + ans_6(1)(i) +
                         ans_7(1)(i) + ans_8(1)(i) + ans_9(1)(i) + ans_10(1)(i)
            '逐年剩余资产净值
            SYJZ(i) = ans_1(2)(i) + ans_2(2)(i) + ans_3(2)(i) + ans_4(2)(i) + ans_5(2)(i) + ans_6(2)(i) +
                      ans_7(2)(i) + ans_8(2)(i) + ans_9(2)(i) + ans_10(2)(i)
        Next

        '判断计算期末如果不回收残值，则进行调整
        If hscz = False Then
            '计算期最后1年的，折旧摊销费金额
            ZJTXF(jsnx) = SYJZ(jsnx)
            '计算期最后1年的，剩余资产净值
            SYJZ(jsnx) = 0
            '计算期最后1年的，折旧摊销费金额累计
            ZJTXFLJ(jsnx) += ZJTXF(jsnx)
        End If

        '返回结果
        Dim ans(2)
        ans(0) = ZJTXF
        ans(1) = ZJTXFLJ
        ans(2) = SYJZ
        Return ans
    End Function
    Function 折旧摊销计算_年数总和法_10次投资分开计算(zcyz_list As Array, zjtxnx_0_list As Array, ksnf_list As Array,
                                                      jsnx As Integer, czl_list As Array, month_10_list As Array, hscz As Boolean)
        'zcyz_list：资产原值，列表，长度10
        'zjtxnx_0_list：折旧摊销年限数初始值，列表，长度10
        'ksnf_list：折旧摊销计算的开始年份初始值，列表，长度10
        'jsnx：项目总的计算年限
        'czl_list：残值率，列表，长度10
        'month_10_list：10次投资，计算每次投资的逐年投产月份数，列表，长度10
        'hscy：计算期末，是否回收资产残值

        '10次投资分别计算固定资产折旧或者无形资产摊销，再累加在一起

        '第1次投资
        Dim zcyz_1 As Double = zcyz_list(1)
        Dim zjtxnx_0_1 As Integer = zjtxnx_0_list(1)
        Dim ksnf_1 As Integer = ksnf_list(1)
        Dim czl_1 As Double = czl_list(1)
        '第2次投资
        Dim zcyz_2 As Double = zcyz_list(2)
        Dim zjtxnx_0_2 As Integer = zjtxnx_0_list(2)
        Dim ksnf_2 As Integer = ksnf_list(2)
        Dim czl_2 As Double = czl_list(2)
        '第3次投资
        Dim zcyz_3 As Double = zcyz_list(3)
        Dim zjtxnx_0_3 As Integer = zjtxnx_0_list(3)
        Dim ksnf_3 As Integer = ksnf_list(3)
        Dim czl_3 As Double = czl_list(3)
        '第4次投资
        Dim zcyz_4 As Double = zcyz_list(4)
        Dim zjtxnx_0_4 As Integer = zjtxnx_0_list(4)
        Dim ksnf_4 As Integer = ksnf_list(4)
        Dim czl_4 As Double = czl_list(4)
        '第5次投资
        Dim zcyz_5 As Double = zcyz_list(5)
        Dim zjtxnx_0_5 As Integer = zjtxnx_0_list(5)
        Dim ksnf_5 As Integer = ksnf_list(5)
        Dim czl_5 As Double = czl_list(5)
        '第6次投资
        Dim zcyz_6 As Double = zcyz_list(6)
        Dim zjtxnx_0_6 As Integer = zjtxnx_0_list(6)
        Dim ksnf_6 As Integer = ksnf_list(6)
        Dim czl_6 As Double = czl_list(6)
        '第7次投资
        Dim zcyz_7 As Double = zcyz_list(7)
        Dim zjtxnx_0_7 As Integer = zjtxnx_0_list(7)
        Dim ksnf_7 As Integer = ksnf_list(7)
        Dim czl_7 As Double = czl_list(7)
        '第8次投资
        Dim zcyz_8 As Double = zcyz_list(8)
        Dim zjtxnx_0_8 As Integer = zjtxnx_0_list(8)
        Dim ksnf_8 As Integer = ksnf_list(8)
        Dim czl_8 As Double = czl_list(8)
        '第9次投资
        Dim zcyz_9 As Double = zcyz_list(9)
        Dim zjtxnx_0_9 As Integer = zjtxnx_0_list(9)
        Dim ksnf_9 As Integer = ksnf_list(9)
        Dim czl_9 As Double = czl_list(9)
        '第10次投资
        Dim zcyz_10 As Double = zcyz_list(10)
        Dim zjtxnx_0_10 As Integer = zjtxnx_0_list(10)
        Dim ksnf_10 As Integer = ksnf_list(10)
        Dim czl_10 As Double = czl_list(10)

        '列表，储存计算过程量
        Dim ZJTXF(31) As Double '逐年折旧摊销费金额
        Dim ZJTXFLJ(31) As Double '逐年折旧摊销费金额累计
        Dim SYJZ(31) As Double '逐年剩余资产净值
        Dim ZJL(31) As Double '逐年折旧率

        '分次进行计算
        Dim ans_1 = 折旧摊销计算_年数总和法(zcyz_1, zjtxnx_0_1, ksnf_1, jsnx, czl_1, month_10_list(1))
        Dim ans_2 = 折旧摊销计算_年数总和法(zcyz_2, zjtxnx_0_2, ksnf_2, jsnx, czl_2, month_10_list(2))
        Dim ans_3 = 折旧摊销计算_年数总和法(zcyz_3, zjtxnx_0_3, ksnf_3, jsnx, czl_3, month_10_list(3))
        Dim ans_4 = 折旧摊销计算_年数总和法(zcyz_4, zjtxnx_0_4, ksnf_4, jsnx, czl_4, month_10_list(4))
        Dim ans_5 = 折旧摊销计算_年数总和法(zcyz_5, zjtxnx_0_5, ksnf_5, jsnx, czl_5, month_10_list(5))
        Dim ans_6 = 折旧摊销计算_年数总和法(zcyz_6, zjtxnx_0_6, ksnf_6, jsnx, czl_6, month_10_list(6))
        Dim ans_7 = 折旧摊销计算_年数总和法(zcyz_7, zjtxnx_0_7, ksnf_7, jsnx, czl_7, month_10_list(7))
        Dim ans_8 = 折旧摊销计算_年数总和法(zcyz_8, zjtxnx_0_8, ksnf_8, jsnx, czl_8, month_10_list(8))
        Dim ans_9 = 折旧摊销计算_年数总和法(zcyz_9, zjtxnx_0_9, ksnf_9, jsnx, czl_9, month_10_list(9))
        Dim ans_10 = 折旧摊销计算_年数总和法(zcyz_10, zjtxnx_0_10, ksnf_10, jsnx, czl_10, month_10_list(10))

        '累加
        For i = 1 To 31 '第1年到第31年
            '逐年折旧摊销费金额
            ZJTXF(i) = ans_1(0)(i) + ans_2(0)(i) + ans_3(0)(i) + ans_4(0)(i) + ans_5(0)(i) + ans_6(0)(i) +
                       ans_7(0)(i) + ans_8(0)(i) + ans_9(0)(i) + ans_10(0)(i)
            '逐年折旧摊销费金额累计
            ZJTXFLJ(i) = ans_1(1)(i) + ans_2(1)(i) + ans_3(1)(i) + ans_4(1)(i) + ans_5(1)(i) + ans_6(1)(i) +
                         ans_7(1)(i) + ans_8(1)(i) + ans_9(1)(i) + ans_10(1)(i)
            '逐年剩余资产净值
            SYJZ(i) = ans_1(2)(i) + ans_2(2)(i) + ans_3(2)(i) + ans_4(2)(i) + ans_5(2)(i) + ans_6(2)(i) +
                      ans_7(2)(i) + ans_8(2)(i) + ans_9(2)(i) + ans_10(2)(i)
            '逐年折旧率
            ZJL(i) = ans_1(3)(i) + ans_2(3)(i) + ans_3(3)(i) + ans_4(3)(i) + ans_5(3)(i) + ans_6(3)(i) +
                     ans_7(3)(i) + ans_8(3)(i) + ans_9(3)(i) + ans_10(3)(i)
        Next

        '判断计算期末如果不回收残值，则进行调整
        If hscz = False Then
            '计算期最后1年的，折旧摊销费金额
            ZJTXF(jsnx) = SYJZ(jsnx)
            '计算期最后1年的，剩余资产净值
            SYJZ(jsnx) = 0
            '计算期最后1年的，折旧摊销费金额累计
            ZJTXFLJ(jsnx) += ZJTXF(jsnx)
            '计算期最后1年的，折旧率
            ZJL(jsnx) = 1
        End If

        '返回结果
        Dim ans(3)
        ans(0) = ZJTXF
        ans(1) = ZJTXFLJ
        ans(2) = SYJZ
        ans(3) = ZJL
        Return ans
    End Function
    Function 折旧摊销计算_直线法(zcyz As Double, zjtxnx_0 As Integer, ksnf As Integer, jsnx As Integer, czl As Double, tcyfs_list As Array)
        'zcyz：资产原值
        'zjtxnx_0：折旧摊销年限数初始值
        'ksnf：折旧摊销计算的开始年份
        'jsnx：项目总的计算年限
        'czl：残值率
        'tcyfs_list：逐年投产的月份数，列表，长度31

        '列表，储存计算过程量
        Dim ZJTXF(31) As Double '逐年折旧摊销费金额
        Dim ZJTXFLJ(31) As Double '逐年折旧摊销费金额累计
        Dim SYJZ(31) As Double '逐年剩余资产净值

        '折旧摊销从资产原值产生年份的的当年开始计算，按照投产月份数进行折算
        If zcyz <> 0 And zjtxnx_0 <> 0 And ksnf <> 0 And jsnx <> 0 Then
            '对折旧摊销年限进行修正
            Dim zjtxnx As Integer
            If ksnf + zjtxnx_0 <= jsnx Then
                zjtxnx = zjtxnx_0
            Else
                zjtxnx = jsnx - ksnf
            End If

            '每年折旧摊销的金额
            Dim ZJTXJE As Double = zcyz * (1 - czl) / zjtxnx
            Dim YJZJTX As Double = 0 '已经折旧摊销的资产累计，初始值是0
            For i = ksnf To 31 'i表示年份序号
                If YJZJTX < zcyz * (1 - czl) Then
                    '还剩余的可以折旧摊销的资产
                    Dim SYKZJTX As Double = zcyz * (1 - czl) - YJZJTX
                    '当年折旧摊销费金额，每年是定值
                    ZJTXF(i) = Math.Min(ZJTXJE * (tcyfs_list(i) / 12), SYKZJTX)
                    '已经折旧摊销的资产累计
                    YJZJTX += ZJTXF(i)
                    '当年折旧摊销费金额累计
                    ZJTXFLJ(i) = YJZJTX
                    '剩余资产净值
                    SYJZ(i) = zcyz - YJZJTX
                Else
                    '当年折旧摊销费金额，每年是定值
                    ZJTXF(i) = 0
                    '当年折旧摊销费金额累计
                    ZJTXFLJ(i) = YJZJTX
                    '剩余资产净值
                    SYJZ(i) = zcyz - YJZJTX
                End If
            Next
        End If

        '返回结果
        Dim ans(2)
        ans(0) = ZJTXF
        ans(1) = ZJTXFLJ
        ans(2) = SYJZ
        Return ans
    End Function
    Function 折旧摊销计算_年数总和法(zcyz As Double, zjtxnx_0 As Integer, ksnf As Integer, jsnx As Integer, czl As Double, tcyfs_list As Array)
        'zcyz：资产原值
        'zjtxnx_0：折旧摊销年限数初始值
        'ksnf：折旧摊销计算的开始年份初始值
        'jsnx：项目总的计算年限
        'czl：残值率
        'tcyfs_list：逐年投产的月份数，列表，长度31

        '列表，储存计算过程量
        Dim ZJTXF(31) As Double '逐年折旧摊销费金额
        Dim ZJTXFLJ(31) As Double '逐年折旧摊销费金额累计
        Dim SYJZ(31) As Double '逐年剩余资产净值
        Dim ZJL(31) As Double '逐年折旧率

        '折旧摊销从资产原值产生年份的的当年开始计算，按照投产月份数进行折算
        If zcyz <> 0 And zjtxnx_0 <> 0 And ksnf <> 0 And jsnx <> 0 Then
            '对折旧摊销年限进行修正
            Dim zjtxnx As Integer
            If ksnf + zjtxnx_0 <= jsnx Then
                zjtxnx = zjtxnx_0
            Else
                zjtxnx = jsnx - ksnf
            End If

            '计算年份和
            Dim NFH As Integer = (1 + zjtxnx) * zjtxnx / 2
            Dim YJZJTX As Double = 0 '已经折旧摊销的资产累计，初始值是0
            For i = ksnf To 31 'i表示年份序号
                If YJZJTX < zcyz * (1 - czl) Then
                    '还剩余的可以折旧摊销的资产
                    Dim SYKZJTX As Double = zcyz * (1 - czl) - YJZJTX
                    '计算逐年折旧率
                    ZJL(i) = Math.Max((tcyfs_list(i) / 12) * (zjtxnx - (i - ksnf + 1) + 2) * (1 - czl) / NFH, 0)
                    '当年折旧摊销费金额
                    ZJTXF(i) = Math.Min(zcyz * ZJL(i), SYKZJTX)
                    '已经折旧摊销的资产累计
                    YJZJTX += ZJTXF(i)
                    '当年折旧摊销费金额累计
                    ZJTXFLJ(i) = YJZJTX
                    '剩余资产净值
                    SYJZ(i) = zcyz - YJZJTX
                Else
                    '当年折旧摊销费金额，每年是定值
                    ZJTXF(i) = 0
                    '当年折旧摊销费金额累计
                    ZJTXFLJ(i) = YJZJTX
                    '逐年折旧率
                    ZJL(i) = 0
                    '剩余资产净值
                    SYJZ(i) = zcyz - YJZJTX
                End If
            Next
        End If

        '返回结果
        Dim ans(3)
        ans(0) = ZJTXF
        ans(1) = ZJTXFLJ
        ans(2) = SYJZ
        ans(3) = ZJL
        Return ans
    End Function
End Module
