Module 折旧摊销计算
    Sub 折旧摊销计算(ExcelApp As Object, hscz As Boolean)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        'hscy：计算期末，是否回收资产残值

        Dim GSBSJ = 读取估算表数据(ExcelApp)
        '项目计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '无形资产所占比例
        Dim WXZCBL As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 5).Value
        '折旧摊销开始年份
        Dim KSNF(10) As Integer
        '10次投资的资产原值
        Dim GDZCYZ(10) As Double
        Dim WXZCYZ(10) As Double
        '10次投资的固定资产折旧年限，无形资产摊销年限
        Dim GDZCZJNX(10) As Integer
        Dim WXZCTXNX(10) As Integer
        '10次投资的固定资产残值率，无形资产残值率
        Dim GDZCCZL(10) As Double
        Dim WXZCCZL(10) As Double
        '每年投产月份数
        Dim tcyfs_list(31) As Integer
        '读取数据
        For i = 4 To 13
            GDZCZJNX(i - 3) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 22).Value
            WXZCTXNX(i - 3) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 26).Value
            GDZCCZL(i - 3) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 24).Value
            WXZCCZL(i - 3) = 0
        Next
        '读取开始年份和资产原值，前5年
        Dim tznf_list = GSBSJ(0)
        Dim dttz_list = GSBSJ(2)
        Dim kdkzzs_list = GSBSJ(9)
        For i = 1 To 5
            KSNF(i) = tznf_list(i)
            GDZCYZ(i) = (dttz_list(i) - kdkzzs_list(i)) * (1 - WXZCBL)
            WXZCYZ(i) = (dttz_list(i) - kdkzzs_list(i)) * WXZCBL
        Next
        '读取开始年份和资产原值，后5年
        For i = 6 To 10
            KSNF(i) = tznf_list(i)
            GDZCYZ(i) = (dttz_list(i) - kdkzzs_list(i)) * (1 - WXZCBL)
            WXZCYZ(i) = (dttz_list(i) - kdkzzs_list(i)) * WXZCBL
        Next
        '读取每年投产月份数
        For i = 1 To 31
            tcyfs_list(i) = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 2).Value
        Next
        '折旧摊销计算
        Dim ans_GDZCZJ
        Dim ans_WXZCTX
        '方法一
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法一" Then
            '固定资产折旧
            ans_GDZCZJ = 折旧摊销计算_直线法_10次投资合并计算(GDZCYZ, GDZCZJNX(1), KSNF, jsnx, GDZCCZL(1), tcyfs_list, hscz)
            '无形资产摊销
            ans_WXZCTX = 折旧摊销计算_直线法_10次投资合并计算(WXZCYZ, WXZCTXNX(1), KSNF, jsnx, WXZCCZL(1), tcyfs_list, hscz)
        End If
        '方法二
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法二" Then
            '固定资产折旧
            ans_GDZCZJ = 折旧摊销计算_年数总和法_10次投资合并计算(GDZCYZ, GDZCZJNX(1), KSNF, jsnx, GDZCCZL(1), tcyfs_list, hscz)
            '无形资产摊销
            ans_WXZCTX = 折旧摊销计算_年数总和法_10次投资合并计算(WXZCYZ, WXZCTXNX(1), KSNF, jsnx, WXZCCZL(1), tcyfs_list, hscz)
        End If
        '方法三
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法三" Then
            '固定资产折旧
            ans_GDZCZJ = 折旧摊销计算_直线法_10次投资分开计算(GDZCYZ, GDZCZJNX, KSNF, jsnx, GDZCCZL, tcyfs_list, hscz)
            '无形资产摊销
            ans_WXZCTX = 折旧摊销计算_直线法_10次投资分开计算(WXZCYZ, WXZCTXNX, KSNF, jsnx, WXZCCZL, tcyfs_list, hscz)
        End If
        '方法四
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法四" Then
            '固定资产折旧
            ans_GDZCZJ = 折旧摊销计算_年数总和法_10次投资分开计算(GDZCYZ, GDZCZJNX, KSNF, jsnx, GDZCCZL, tcyfs_list, hscz)
            '无形资产摊销
            ans_WXZCTX = 折旧摊销计算_年数总和法_10次投资分开计算(WXZCYZ, WXZCTXNX, KSNF, jsnx, WXZCCZL, tcyfs_list, hscz)
        End If
        '结果写入Excel
        '写入固定资产、无形资产原值
        Dim gdzcyz_all As Double
        Dim wxzcyz_all As Double
        For i = 1 To 10
            gdzcyz_all += GDZCYZ(i)
            wxzcyz_all += WXZCYZ(i)
        Next
        ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(5, 4).Value = gdzcyz_all
        ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(11, 4).Value = wxzcyz_all
        '前15年
        For i = 1 To 15
            If i <= jsnx Then
                '固定资产折旧
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(8, 4 + i).Value = ans_GDZCZJ(0)(i)
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(9, 4 + i).Value = ans_GDZCZJ(1)(i)
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(10, 4 + i).Value = ans_GDZCZJ(2)(i)
                '无形资产摊销
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(14, 4 + i).Value = ans_WXZCTX(0)(i)
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(15, 4 + i).Value = ans_WXZCTX(1)(i)
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(16, 4 + i).Value = ans_WXZCTX(2)(i)
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
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(25, i - 12).Value = ans_GDZCZJ(0)(i)
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(26, i - 12).Value = ans_GDZCZJ(1)(i)
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(27, i - 12).Value = ans_GDZCZJ(2)(i)
                '无形资产摊销
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(31, i - 12).Value = ans_WXZCTX(0)(i)
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(32, i - 12).Value = ans_WXZCTX(1)(i)
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(33, i - 12).Value = ans_WXZCTX(2)(i)
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
        '将10次投资的资产原值转为长度31的列表
        Dim gdzcyz_list = 基础计算功能_10_to_31(tznf_list, GDZCYZ)
        Dim wxzcyz_list = 基础计算功能_10_to_31(tznf_list, WXZCYZ)
        '计算逐年累计固定资产、无形资产原值
        Dim gdzcyz_lj(31) As Double
        Dim wxzcyz_lj(31) As Double
        For i = 1 To 31
            gdzcyz_lj(i) = gdzcyz_lj(i - 1) + gdzcyz_list(i)
            wxzcyz_lj(i) = wxzcyz_lj(i - 1) + wxzcyz_list(i)
        Next
        '结果写入Excel
        '前15年
        For i = 1 To 15
            If i <= jsnx Then
                '固定资产
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(5, 4 + i).Value = gdzcyz_lj(i)
                '无形资产
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(11, 4 + i).Value = wxzcyz_lj(i)
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
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(22, i - 12).Value = gdzcyz_lj(i)
                '无形资产
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(28, i - 12).Value = wxzcyz_lj(i)
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
    Function 折旧摊销计算_直线法_10次投资合并计算(zcyz_list As Array, zjtxnx_0 As Integer, ksnf_list As Array,
                                                  jsnx As Integer, czl As Double, tcyfs_list As Array, hscz As Boolean)
        'zcyz_list：资产原值，列表，长度10
        'zjtxnx_0：折旧摊销年限数初始值
        'ksnf_list：折旧摊销计算的开始年份，列表，长度10
        'jsnx：项目总的计算年限
        'czl：残值率
        'tcyfs_list：逐年投产的月份数，列表，长度31
        'hscy：计算期末，是否回收资产残值

        '10次投资的固定资产折旧或者无形资产摊销合并在一起进行计算

        '第1次投资
        Dim zcyz_1 As Double = zcyz_list(1)
        Dim ksnf_1 As Integer = ksnf_list(1)
        Dim zjtxnx_0_1 As Integer = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_1)
        '第2次投资
        Dim zcyz_2 As Double = zcyz_list(2)
        Dim ksnf_2 As Integer = ksnf_list(2)
        Dim zjtxnx_0_2 As Integer = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_2)
        '第3次投资
        Dim zcyz_3 As Double = zcyz_list(3)
        Dim ksnf_3 As Integer = ksnf_list(3)
        Dim zjtxnx_0_3 As Integer = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_3)
        '第4次投资
        Dim zcyz_4 As Double = zcyz_list(4)
        Dim ksnf_4 As Integer = ksnf_list(4)
        Dim zjtxnx_0_4 As Integer = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_4)
        '第5次投资
        Dim zcyz_5 As Double = zcyz_list(5)
        Dim ksnf_5 As Integer = ksnf_list(5)
        Dim zjtxnx_0_5 As Integer = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_5)
        '第6次投资
        Dim zcyz_6 As Double = zcyz_list(6)
        Dim ksnf_6 As Integer = ksnf_list(6)
        Dim zjtxnx_0_6 As Integer = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_6)
        '第7次投资
        Dim zcyz_7 As Double = zcyz_list(7)
        Dim ksnf_7 As Integer = ksnf_list(7)
        Dim zjtxnx_0_7 As Integer = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_7)
        '第8次投资
        Dim zcyz_8 As Double = zcyz_list(8)
        Dim ksnf_8 As Integer = ksnf_list(8)
        Dim zjtxnx_0_8 As Integer = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_8)
        '第9次投资
        Dim zcyz_9 As Double = zcyz_list(9)
        Dim ksnf_9 As Integer = ksnf_list(9)
        Dim zjtxnx_0_9 As Integer = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_9)
        '第10次投资
        Dim zcyz_10 As Double = zcyz_list(10)
        Dim ksnf_10 As Integer = ksnf_list(10)
        Dim zjtxnx_0_10 As Integer = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_10)

        '列表，储存计算结果
        Dim ZJTXF(31) As Double '逐年折旧摊销费金额
        Dim ZJTXFLJ(31) As Double '逐年折旧摊销费金额累计
        Dim SYJZ(31) As Double '逐年剩余资产净值

        '分次进行计算
        Dim ans_1 = 折旧摊销计算_直线法(zcyz_1, zjtxnx_0_1, ksnf_1, jsnx, czl, tcyfs_list)
        Dim ans_2 = 折旧摊销计算_直线法(zcyz_2, zjtxnx_0_2, ksnf_2, jsnx, czl, tcyfs_list)
        Dim ans_3 = 折旧摊销计算_直线法(zcyz_3, zjtxnx_0_3, ksnf_3, jsnx, czl, tcyfs_list)
        Dim ans_4 = 折旧摊销计算_直线法(zcyz_4, zjtxnx_0_4, ksnf_4, jsnx, czl, tcyfs_list)
        Dim ans_5 = 折旧摊销计算_直线法(zcyz_5, zjtxnx_0_5, ksnf_5, jsnx, czl, tcyfs_list)
        Dim ans_6 = 折旧摊销计算_直线法(zcyz_6, zjtxnx_0_6, ksnf_6, jsnx, czl, tcyfs_list)
        Dim ans_7 = 折旧摊销计算_直线法(zcyz_7, zjtxnx_0_7, ksnf_7, jsnx, czl, tcyfs_list)
        Dim ans_8 = 折旧摊销计算_直线法(zcyz_8, zjtxnx_0_8, ksnf_8, jsnx, czl, tcyfs_list)
        Dim ans_9 = 折旧摊销计算_直线法(zcyz_9, zjtxnx_0_9, ksnf_9, jsnx, czl, tcyfs_list)
        Dim ans_10 = 折旧摊销计算_直线法(zcyz_10, zjtxnx_0_10, ksnf_10, jsnx, czl, tcyfs_list)
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


    Function 折旧摊销计算_年数总和法_10次投资合并计算(zcyz_list As Array, zjtxnx_0 As Integer, ksnf_0_list As Array,
                                                      jsnx As Integer, czl As Double, tcyfs_list As Array, hscz As Boolean)

        'zcyz_list：资产原值，列表，长度10
        'zjtxnx_0：折旧摊销年限数初始值
        'ksnf_0_list：折旧摊销计算的开始年份初始值，列表，长度10
        'jsnx：项目总的计算年限
        'czl：残值率
        'tcyfs_list：逐年投产的月份数，列表，长度31
        'hscy：计算期末，是否回收资产残值

        '10次投资的固定资产折旧或者无形资产摊销合并在一起进行计算

        '第1次投资
        Dim zcyz_1 As Double = zcyz_list(1)
        Dim ksnf_0_1 As Integer = ksnf_0_list(1)
        Dim zjtxnx_0_1 As Integer = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_0_1)
        '第2次投资
        Dim zcyz_2 As Double = zcyz_list(2)
        Dim ksnf_0_2 As Integer = ksnf_0_list(2)
        Dim zjtxnx_0_2 As Integer = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_0_2)
        '第3次投资
        Dim zcyz_3 As Double = zcyz_list(3)
        Dim ksnf_0_3 As Integer = ksnf_0_list(3)
        Dim zjtxnx_0_3 As Integer = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_0_3)
        '第4次投资
        Dim zcyz_4 As Double = zcyz_list(4)
        Dim ksnf_0_4 As Integer = ksnf_0_list(4)
        Dim zjtxnx_0_4 As Integer = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_0_4)
        '第5次投资
        Dim zcyz_5 As Double = zcyz_list(5)
        Dim ksnf_0_5 As Integer = ksnf_0_list(5)
        Dim zjtxnx_0_5 As Integer = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_0_5)
        '第6次投资
        Dim zcyz_6 As Double = zcyz_list(6)
        Dim ksnf_0_6 As Integer = ksnf_0_list(6)
        Dim zjtxnx_0_6 As Integer = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_0_6)
        '第7次投资
        Dim zcyz_7 As Double = zcyz_list(7)
        Dim ksnf_0_7 As Integer = ksnf_0_list(7)
        Dim zjtxnx_0_7 As Integer = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_0_7)
        '第8次投资
        Dim zcyz_8 As Double = zcyz_list(8)
        Dim ksnf_0_8 As Integer = ksnf_0_list(8)
        Dim zjtxnx_0_8 As Integer = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_0_8)
        '第9次投资
        Dim zcyz_9 As Double = zcyz_list(9)
        Dim ksnf_0_9 As Integer = ksnf_0_list(9)
        Dim zjtxnx_0_9 As Integer = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_0_9)
        '第10次投资
        Dim zcyz_10 As Double = zcyz_list(10)
        Dim ksnf_0_10 As Integer = ksnf_0_list(10)
        Dim zjtxnx_0_10 As Integer = Math.Min(zjtxnx_0, zjtxnx_0 + 1 - ksnf_0_10)
        '列表，储存计算过程量
        Dim ZJTXF(31) As Double '逐年折旧摊销费金额
        Dim ZJTXFLJ(31) As Double '逐年折旧摊销费金额累计
        Dim SYJZ(31) As Double '逐年剩余资产净值
        Dim ZJL(31) As Double '逐年折旧率

        '分次进行计算
        Dim ans_1 = 折旧摊销计算_年数总和法(zcyz_1, zjtxnx_0_1, ksnf_0_1, jsnx, czl, tcyfs_list)
        Dim ans_2 = 折旧摊销计算_年数总和法(zcyz_2, zjtxnx_0_2, ksnf_0_2, jsnx, czl, tcyfs_list)
        Dim ans_3 = 折旧摊销计算_年数总和法(zcyz_3, zjtxnx_0_3, ksnf_0_3, jsnx, czl, tcyfs_list)
        Dim ans_4 = 折旧摊销计算_年数总和法(zcyz_4, zjtxnx_0_4, ksnf_0_4, jsnx, czl, tcyfs_list)
        Dim ans_5 = 折旧摊销计算_年数总和法(zcyz_5, zjtxnx_0_5, ksnf_0_5, jsnx, czl, tcyfs_list)
        Dim ans_6 = 折旧摊销计算_年数总和法(zcyz_6, zjtxnx_0_6, ksnf_0_6, jsnx, czl, tcyfs_list)
        Dim ans_7 = 折旧摊销计算_年数总和法(zcyz_7, zjtxnx_0_7, ksnf_0_7, jsnx, czl, tcyfs_list)
        Dim ans_8 = 折旧摊销计算_年数总和法(zcyz_8, zjtxnx_0_8, ksnf_0_8, jsnx, czl, tcyfs_list)
        Dim ans_9 = 折旧摊销计算_年数总和法(zcyz_9, zjtxnx_0_9, ksnf_0_9, jsnx, czl, tcyfs_list)
        Dim ans_10 = 折旧摊销计算_年数总和法(zcyz_10, zjtxnx_0_10, ksnf_0_10, jsnx, czl, tcyfs_list)

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
                                                  jsnx As Integer, czl_list As Array, tcyfs_list As Array, hscz As Boolean)
        'zcyz_list：资产原值，列表，长度10
        'zjtxnx_0_list：折旧摊销年限数初始值，列表，长度10
        'ksnf_list：折旧摊销计算的开始年份，列表，长度10
        'jsnx：项目总的计算年限
        'czl_list：残值率，列表，长度10
        'tcyfs_list：逐年投产的月份数，列表，长度31
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
        Dim ans_1 = 折旧摊销计算_直线法(zcyz_1, zjtxnx_0_1, ksnf_1, jsnx, czl_1, tcyfs_list)
        Dim ans_2 = 折旧摊销计算_直线法(zcyz_2, zjtxnx_0_2, ksnf_2, jsnx, czl_2, tcyfs_list)
        Dim ans_3 = 折旧摊销计算_直线法(zcyz_3, zjtxnx_0_3, ksnf_3, jsnx, czl_3, tcyfs_list)
        Dim ans_4 = 折旧摊销计算_直线法(zcyz_4, zjtxnx_0_4, ksnf_4, jsnx, czl_4, tcyfs_list)
        Dim ans_5 = 折旧摊销计算_直线法(zcyz_5, zjtxnx_0_5, ksnf_5, jsnx, czl_5, tcyfs_list)
        Dim ans_6 = 折旧摊销计算_直线法(zcyz_6, zjtxnx_0_6, ksnf_6, jsnx, czl_6, tcyfs_list)
        Dim ans_7 = 折旧摊销计算_直线法(zcyz_7, zjtxnx_0_7, ksnf_7, jsnx, czl_7, tcyfs_list)
        Dim ans_8 = 折旧摊销计算_直线法(zcyz_8, zjtxnx_0_8, ksnf_8, jsnx, czl_8, tcyfs_list)
        Dim ans_9 = 折旧摊销计算_直线法(zcyz_9, zjtxnx_0_9, ksnf_9, jsnx, czl_9, tcyfs_list)
        Dim ans_10 = 折旧摊销计算_直线法(zcyz_10, zjtxnx_0_10, ksnf_10, jsnx, czl_10, tcyfs_list)
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

    Function 折旧摊销计算_年数总和法_10次投资分开计算(zcyz_list As Array, zjtxnx_0_list As Array, ksnf_0_list As Array,
                                                      jsnx As Integer, czl_list As Array, tcyfs_list As Array, hscz As Boolean)
        'zcyz_list：资产原值，列表，长度10
        'zjtxnx_0_list：折旧摊销年限数初始值，列表，长度10
        'ksnf_0_list：折旧摊销计算的开始年份初始值，列表，长度10
        'jsnx：项目总的计算年限
        'czl_list：残值率，列表，长度10
        'tcyfs_list：逐年投产的月份数，列表，长度31
        'hscy：计算期末，是否回收资产残值

        '10次投资分别计算固定资产折旧或者无形资产摊销，再累加在一起

        '第1次投资
        Dim zcyz_1 As Double = zcyz_list(1)
        Dim zjtxnx_0_1 As Integer = zjtxnx_0_list(1)
        Dim ksnf_0_1 As Integer = ksnf_0_list(1)
        Dim czl_1 As Double = czl_list(1)
        '第2次投资
        Dim zcyz_2 As Double = zcyz_list(2)
        Dim zjtxnx_0_2 As Integer = zjtxnx_0_list(2)
        Dim ksnf_0_2 As Integer = ksnf_0_list(2)
        Dim czl_2 As Double = czl_list(2)
        '第3次投资
        Dim zcyz_3 As Double = zcyz_list(3)
        Dim zjtxnx_0_3 As Integer = zjtxnx_0_list(3)
        Dim ksnf_0_3 As Integer = ksnf_0_list(3)
        Dim czl_3 As Double = czl_list(3)
        '第4次投资
        Dim zcyz_4 As Double = zcyz_list(4)
        Dim zjtxnx_0_4 As Integer = zjtxnx_0_list(4)
        Dim ksnf_0_4 As Integer = ksnf_0_list(4)
        Dim czl_4 As Double = czl_list(4)
        '第5次投资
        Dim zcyz_5 As Double = zcyz_list(5)
        Dim zjtxnx_0_5 As Integer = zjtxnx_0_list(5)
        Dim ksnf_0_5 As Integer = ksnf_0_list(5)
        Dim czl_5 As Double = czl_list(5)
        '第6次投资
        Dim zcyz_6 As Double = zcyz_list(6)
        Dim zjtxnx_0_6 As Integer = zjtxnx_0_list(6)
        Dim ksnf_0_6 As Integer = ksnf_0_list(6)
        Dim czl_6 As Double = czl_list(6)
        '第7次投资
        Dim zcyz_7 As Double = zcyz_list(7)
        Dim zjtxnx_0_7 As Integer = zjtxnx_0_list(7)
        Dim ksnf_0_7 As Integer = ksnf_0_list(7)
        Dim czl_7 As Double = czl_list(7)
        '第8次投资
        Dim zcyz_8 As Double = zcyz_list(8)
        Dim zjtxnx_0_8 As Integer = zjtxnx_0_list(8)
        Dim ksnf_0_8 As Integer = ksnf_0_list(8)
        Dim czl_8 As Double = czl_list(8)
        '第9次投资
        Dim zcyz_9 As Double = zcyz_list(9)
        Dim zjtxnx_0_9 As Integer = zjtxnx_0_list(9)
        Dim ksnf_0_9 As Integer = ksnf_0_list(9)
        Dim czl_9 As Double = czl_list(9)
        '第10次投资
        Dim zcyz_10 As Double = zcyz_list(10)
        Dim zjtxnx_0_10 As Integer = zjtxnx_0_list(10)
        Dim ksnf_0_10 As Integer = ksnf_0_list(10)
        Dim czl_10 As Double = czl_list(10)

        '列表，储存计算过程量
        Dim ZJTXF(31) As Double '逐年折旧摊销费金额
        Dim ZJTXFLJ(31) As Double '逐年折旧摊销费金额累计
        Dim SYJZ(31) As Double '逐年剩余资产净值
        Dim ZJL(31) As Double '逐年折旧率

        '分次进行计算
        Dim ans_1 = 折旧摊销计算_年数总和法(zcyz_1, zjtxnx_0_1, ksnf_0_1, jsnx, czl_1, tcyfs_list)
        Dim ans_2 = 折旧摊销计算_年数总和法(zcyz_2, zjtxnx_0_2, ksnf_0_2, jsnx, czl_2, tcyfs_list)
        Dim ans_3 = 折旧摊销计算_年数总和法(zcyz_3, zjtxnx_0_3, ksnf_0_3, jsnx, czl_3, tcyfs_list)
        Dim ans_4 = 折旧摊销计算_年数总和法(zcyz_4, zjtxnx_0_4, ksnf_0_4, jsnx, czl_4, tcyfs_list)
        Dim ans_5 = 折旧摊销计算_年数总和法(zcyz_5, zjtxnx_0_5, ksnf_0_5, jsnx, czl_5, tcyfs_list)
        Dim ans_6 = 折旧摊销计算_年数总和法(zcyz_6, zjtxnx_0_6, ksnf_0_6, jsnx, czl_6, tcyfs_list)
        Dim ans_7 = 折旧摊销计算_年数总和法(zcyz_7, zjtxnx_0_7, ksnf_0_7, jsnx, czl_7, tcyfs_list)
        Dim ans_8 = 折旧摊销计算_年数总和法(zcyz_8, zjtxnx_0_8, ksnf_0_8, jsnx, czl_8, tcyfs_list)
        Dim ans_9 = 折旧摊销计算_年数总和法(zcyz_9, zjtxnx_0_9, ksnf_0_9, jsnx, czl_9, tcyfs_list)
        Dim ans_10 = 折旧摊销计算_年数总和法(zcyz_10, zjtxnx_0_10, ksnf_0_10, jsnx, czl_10, tcyfs_list)

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
    Function 折旧摊销计算_直线法(zcyz As Double, zjtxnx_0 As Integer, ksnf As Integer, jsnx As Integer,
                                 czl As Double, tcyfs_list As Array)
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

        '折旧摊销一定从资产原值产生年份的后一年开始计算，从后一年开始，按照投产月份数进行折算

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
            For i = ksnf + 1 To 31 'i表示年份序号
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

    Function 折旧摊销计算_年数总和法(zcyz As Double, zjtxnx_0 As Integer, ksnf_0 As Integer, jsnx As Integer,
                                     czl As Double, tcyfs_list As Array)
        'zcyz：资产原值
        'zjtxnx_0：折旧摊销年限数初始值
        'ksnf_0：折旧摊销计算的开始年份初始值
        'jsnx：项目总的计算年限
        'czl：残值率
        'tcyfs_list：逐年投产的月份数，列表，长度31

        '列表，储存计算过程量
        Dim ZJTXF(31) As Double '逐年折旧摊销费金额
        Dim ZJTXFLJ(31) As Double '逐年折旧摊销费金额累计
        Dim SYJZ(31) As Double '逐年剩余资产净值
        Dim ZJL(31) As Double '逐年折旧率

        '折旧摊销一定从资产原值产生年份的后一年开始计算，从后一年开始，按照投产月份数进行折算

        If zcyz <> 0 And zjtxnx_0 <> 0 And ksnf_0 <> 0 And jsnx <> 0 Then
            '对开始年限进行修正，如果处于建设期内，及时有投产月份数，但是不足12个月，年数总和法不计算折旧率
            Dim ksnf As Integer = ksnf_0
            For i = ksnf + 1 To 31
                If tcyfs_list(i) = 12 Then
                    ksnf = i - 1
                    Exit For
                End If
            Next

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
            For i = ksnf + 1 To 31 'i表示年份序号
                If YJZJTX < zcyz * (1 - czl) Then
                    '还剩余的可以折旧摊销的资产
                    Dim SYKZJTX As Double = zcyz * (1 - czl) - YJZJTX
                    '计算逐年折旧率
                    ZJL(i) = (zjtxnx - (i - ksnf + 1) + 2) * (1 - czl) / NFH
                    '当年折旧摊销费金额，每年是定值
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
