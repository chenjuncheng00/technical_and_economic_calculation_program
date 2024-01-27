Module 材料费其它费计算
    Sub 材料费其它费计算(ExcelApp As Object, clfl_qtfl_model As Integer, kcje_clf_qtf_model As Integer)
        On Error Resume Next
        'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        '————————————————————————————————————————————————————————————————————————————————————————        
        '项目计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '计算逐年分项材料费其它费金额
        Dim ans_clfqtf = 分项逐年材料费其它费计算(ExcelApp, clfl_qtfl_model, kcje_clf_qtf_model)
        Dim ans_znclf = ans_clfqtf(0)
        Dim ans_znqtf = ans_clfqtf(1)
        '————————————————————————————————————————————————————————————————————————————————————————
        '结果写入Excel
        '前15年
        For i = 1 To 15
            If i <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(225, 4 + i).Value = ans_znclf(i)
                ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(227, 4 + i).Value = ans_znqtf(i)
            Else
                ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(225, 4 + i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(227, 4 + i).Value = 0
            End If
        Next
        '16—31年
        For i = 16 To 31
            If i <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(233, i - 12).Value = ans_znclf(i)
                ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(235, i - 12).Value = ans_znqtf(i)
            Else
                ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(233, i - 12).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(235, i - 12).Value = 0
            End If
        Next
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '计算一次Excel
        ExcelApp.Calculate()
    End Sub

    Function 分项逐年材料费其它费计算(ExcelApp As Object, clfl_qtfl_model As Integer, kcje_clf_qtf_model As Integer)
        'On Error Resume Next
        '只计算出分项逐年材料费和其它费的金额数值，不写入EXCEL
        'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        '————————————————————————————————————————————————————————————————————————————————————————        
        '读取材料费其它费计算默认值
        Dim clqtfl = 默认逐年材料费率其它费率(ExcelApp)
        '逐年材料费默认值
        Dim clfl_mr_rj_list = clqtfl(0)
        Dim clfl_mr_rm_list = clqtfl(1)
        Dim clfl_mr_ljfd_list = clqtfl(2)
        Dim clfl_mr_glgr_list = clqtfl(3)
        Dim clfl_mr_gf_list = clqtfl(4)
        Dim clfl_mr_fd_list = clqtfl(5)
        Dim clfl_mr_xdc_list = clqtfl(12)
        '逐年其它费默认值
        Dim qtfl_mr_rj_list = clqtfl(6)
        Dim qtfl_mr_rm_list = clqtfl(7)
        Dim qtfl_mr_ljfd_list = clqtfl(8)
        Dim qtfl_mr_glgr_list = clqtfl(9)
        Dim qtfl_mr_gf_list = clqtfl(10)
        Dim qtfl_mr_fd_list = clqtfl(11)
        Dim qtfl_mr_xdc_list = clqtfl(13)
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算采用的费率
        Dim clfl_rj_list(31) As Double
        Dim clfl_rm_list(31) As Double
        Dim clfl_ljfd_list(31) As Double
        Dim clfl_glgr_list(31) As Double
        Dim clfl_gf_list(31) As Double
        Dim clfl_fd_list(31) As Double
        Dim clfl_xdc_list(31) As Double
        Dim qtfl_rj_list(31) As Double
        Dim qtfl_rm_list(31) As Double
        Dim qtfl_ljfd_list(31) As Double
        Dim qtfl_glgr_list(31) As Double
        Dim qtfl_gf_list(31) As Double
        Dim qtfl_fd_list(31) As Double
        Dim qtfl_xdc_list(31) As Double
        If clfl_qtfl_model = 0 Then
            clfl_rj_list = clfl_mr_rj_list
            clfl_rm_list = clfl_mr_rm_list
            clfl_ljfd_list = clfl_mr_ljfd_list
            clfl_glgr_list = clfl_mr_glgr_list
            clfl_gf_list = clfl_mr_gf_list
            clfl_fd_list = clfl_mr_fd_list
            clfl_xdc_list = clfl_mr_xdc_list
            qtfl_rj_list = qtfl_mr_rj_list
            qtfl_rm_list = qtfl_mr_rm_list
            qtfl_ljfd_list = qtfl_mr_ljfd_list
            qtfl_glgr_list = qtfl_mr_glgr_list
            qtfl_gf_list = qtfl_mr_gf_list
            qtfl_fd_list = qtfl_mr_fd_list
            qtfl_xdc_list = qtfl_mr_xdc_list
        Else
            For i = 3 To 33 '列
                clfl_rj_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(182, i).Value
                clfl_rm_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(183, i).Value
                clfl_ljfd_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(184, i).Value
                clfl_glgr_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(185, i).Value
                clfl_gf_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(186, i).Value
                clfl_fd_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(187, i).Value
                qtfl_rj_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(188, i).Value
                qtfl_rm_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(189, i).Value
                qtfl_ljfd_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(190, i).Value
                qtfl_glgr_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(191, i).Value
                qtfl_gf_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(192, i).Value
                qtfl_fd_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(193, i).Value
                clfl_xdc_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(194, i).Value
                qtfl_xdc_list(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(195, i).Value
            Next
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '设备材料费其它费计算基数扣除计算模式
        Dim yynx_rj As Integer
        Dim yynx_xdc As Integer
        Dim yynx_nt As Integer
        Dim yynx_gf As Integer
        Dim yynx_rm As Integer
        Dim yynx_fd As Integer
        Dim yynx_ljfd As Integer
        Dim kcbl_rj As Double
        Dim kcbl_xdc As Double
        Dim kcbl_nt As Double
        Dim kcbl_gf As Double
        Dim kcbl_rm As Double
        Dim kcbl_fd As Double
        Dim kcbl_ljfd As Double
        Dim kcje_mr = 材料费其它费计算基数扣除默认设置(ExcelApp)
        If kcje_clf_qtf_model = 0 Then
            yynx_rj = kcje_mr(0)
            yynx_xdc = kcje_mr(1)
            yynx_nt = kcje_mr(2)
            yynx_gf = kcje_mr(3)
            yynx_rm = kcje_mr(4)
            yynx_fd = kcje_mr(5)
            yynx_ljfd = kcje_mr(6)
            kcbl_rj = kcje_mr(7)
            kcbl_xdc = kcje_mr(8)
            kcbl_nt = kcje_mr(9)
            kcbl_gf = kcje_mr(10)
            kcbl_rm = kcje_mr(11)
            kcbl_fd = kcje_mr(12)
            kcbl_ljfd = kcje_mr(13)
        Else
            yynx_rj = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(106, 3).Value
            yynx_xdc = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(106, 4).Value
            yynx_nt = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(106, 5).Value
            yynx_gf = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(106, 6).Value
            yynx_rm = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(106, 7).Value
            yynx_fd = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(106, 8).Value
            yynx_ljfd = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(106, 9).Value
            kcbl_rj = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(107, 3).Value
            kcbl_xdc = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(107, 4).Value
            kcbl_nt = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(107, 5).Value
            kcbl_gf = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(107, 6).Value
            kcbl_rm = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(107, 7).Value
            kcbl_fd = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(107, 8).Value
            kcbl_ljfd = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(107, 9).Value
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '项目计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '材料费其它费计算模式
        Dim GSBSJ = 读取估算表数据(ExcelApp)
        '投资年份，10次投资的情况
        Dim tznf_list = GSBSJ(0)
        '计算10次投资，每次建设年份的投产月份数
        Dim tcyf_list = 逐年投产月份数_10次投资(ExcelApp)(2)
        '燃机总发电量(万kWh)，10次投资的情况
        Dim rjfdl = GSBSJ(11)
        '蓄电池总装机功率(kW)，10次投资的情况
        Dim xdczjgl = GSBSJ(13)
        '供冷供热总量(万kWh)，10次投资的情况
        Dim glgrl = GSBSJ(15)
        '光伏总装机功率(kW)，10次投资的情况
        Dim gfzjgl = GSBSJ(17)
        '燃煤机组总发电量(万kWh)，10次投资的情况
        Dim rmfdl = GSBSJ(21)
        '风电总装机功率(kW)，10次投资的情况
        Dim fdzjgl = GSBSJ(19)
        '垃圾发电总发电量(万kWh)，10次投资的情况
        Dim ljfdl = GSBSJ(20)
        '读取逐年负荷率
        Dim fhl_list = 读取逐年负荷率(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        '光伏、风电、蓄电池的材料费和其它费只有一种计算模式，不区分“负荷率” OR “投资量”
        Dim ans_clf_gf = 计算逐年总金额_10次投资(tznf_list, tcyf_list, gfzjgl, yynx_gf, kcbl_gf, True, clfl_gf_list, 10000)
        Dim ans_qtf_gf = 计算逐年总金额_10次投资(tznf_list, tcyf_list, gfzjgl, yynx_gf, kcbl_gf, True, qtfl_gf_list, 10000)
        Dim ans_clf_fd = 计算逐年总金额_10次投资(tznf_list, tcyf_list, fdzjgl, yynx_fd, kcbl_fd, True, clfl_fd_list, 10000)
        Dim ans_qtf_fd = 计算逐年总金额_10次投资(tznf_list, tcyf_list, fdzjgl, yynx_fd, kcbl_fd, True, qtfl_fd_list, 10000)
        Dim ans_clf_xdc = 计算逐年总金额_10次投资(tznf_list, tcyf_list, xdczjgl, yynx_xdc, kcbl_xdc, True, clfl_xdc_list, 10000)
        Dim ans_qtf_xdc = 计算逐年总金额_10次投资(tznf_list, tcyf_list, xdczjgl, yynx_xdc, kcbl_xdc, True, qtfl_xdc_list, 10000)
        '燃机、垃圾发电、燃煤、供冷供热的材料费和其它费区分“负荷率” OR “投资量”，有两种不同的计算模式
        Dim ans_clf_rj(31) As Double
        Dim ans_clf_lj(31) As Double
        Dim ans_clf_rm(31) As Double
        Dim ans_clf_glgr(31) As Double
        Dim ans_qtf_rj(31) As Double
        Dim ans_qtf_lj(31) As Double
        Dim ans_qtf_rm(31) As Double
        Dim ans_qtf_glgr(31) As Double
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(11, 11).Value = "负荷率" Then
            '将逐年负荷率和材料费率其它费率相乘，得到综合费率
            Dim fh_clfl_rj_list(31) As Double
            Dim fh_qtfl_rj_list(31) As Double
            Dim fh_clfl_ljfd_list(31) As Double
            Dim fh_qtfl_ljfd_list(31) As Double
            Dim fh_clfl_rm_list(31) As Double
            Dim fh_qtfl_rm_list(31) As Double
            Dim fh_clfl_glgr_list(31) As Double
            Dim fh_qtfl_glgr_list(31) As Double
            For i = 1 To 31
                fh_clfl_rj_list(i) = clfl_rj_list(i) * fhl_list(i)
                fh_qtfl_rj_list(i) = qtfl_rj_list(i) * fhl_list(i)
                fh_clfl_ljfd_list(i) = clfl_ljfd_list(i) * fhl_list(i)
                fh_qtfl_ljfd_list(i) = qtfl_ljfd_list(i) * fhl_list(i)
                fh_clfl_rm_list(i) = clfl_rm_list(i) * fhl_list(i)
                fh_qtfl_rm_list(i) = qtfl_rm_list(i) * fhl_list(i)
                fh_clfl_glgr_list(i) = clfl_glgr_list(i) * fhl_list(i)
                fh_qtfl_glgr_list(i) = qtfl_glgr_list(i) * fhl_list(i)
            Next
            '计算
            ans_clf_rj = 计算逐年总金额_10次投资(tznf_list, tcyf_list, rjfdl, yynx_rj, kcbl_rj, False, fh_clfl_rj_list, 1000)
            ans_qtf_rj = 计算逐年总金额_10次投资(tznf_list, tcyf_list, rjfdl, yynx_rj, kcbl_rj, False, fh_qtfl_rj_list, 1000)
            ans_clf_lj = 计算逐年总金额_10次投资(tznf_list, tcyf_list, ljfdl, yynx_ljfd, kcbl_ljfd, False, fh_clfl_ljfd_list, 1000)
            ans_qtf_lj = 计算逐年总金额_10次投资(tznf_list, tcyf_list, ljfdl, yynx_ljfd, kcbl_ljfd, False, fh_qtfl_ljfd_list, 1000)
            ans_clf_rm = 计算逐年总金额_10次投资(tznf_list, tcyf_list, rmfdl, yynx_rm, kcbl_rm, False, fh_clfl_rm_list, 1000)
            ans_qtf_rm = 计算逐年总金额_10次投资(tznf_list, tcyf_list, rmfdl, yynx_rm, kcbl_rm, False, fh_qtfl_rm_list, 1000)
            ans_clf_glgr = 计算逐年总金额_10次投资(tznf_list, tcyf_list, glgrl, yynx_nt, kcbl_nt, False, fh_clfl_glgr_list, 1000)
            ans_qtf_glgr = 计算逐年总金额_10次投资(tznf_list, tcyf_list, glgrl, yynx_nt, kcbl_nt, False, fh_qtfl_glgr_list, 1000)
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(11, 11).Value = "投资量" Then
            ans_clf_rj = 计算逐年总金额_10次投资(tznf_list, tcyf_list, rjfdl, yynx_rj, kcbl_rj, True, clfl_rj_list, 1000)
            ans_qtf_rj = 计算逐年总金额_10次投资(tznf_list, tcyf_list, rjfdl, yynx_rj, kcbl_rj, True, qtfl_rj_list, 1000)
            ans_clf_lj = 计算逐年总金额_10次投资(tznf_list, tcyf_list, ljfdl, yynx_ljfd, kcbl_ljfd, True, clfl_ljfd_list, 1000)
            ans_qtf_lj = 计算逐年总金额_10次投资(tznf_list, tcyf_list, ljfdl, yynx_ljfd, kcbl_ljfd, True, qtfl_ljfd_list, 1000)
            ans_clf_rm = 计算逐年总金额_10次投资(tznf_list, tcyf_list, rmfdl, yynx_rm, kcbl_rm, True, clfl_rm_list, 1000)
            ans_qtf_rm = 计算逐年总金额_10次投资(tznf_list, tcyf_list, rmfdl, yynx_rm, kcbl_rm, True, qtfl_rm_list, 1000)
            ans_clf_glgr = 计算逐年总金额_10次投资(tznf_list, tcyf_list, glgrl, yynx_nt, kcbl_nt, True, clfl_glgr_list, 1000)
            ans_qtf_glgr = 计算逐年总金额_10次投资(tznf_list, tcyf_list, glgrl, yynx_nt, kcbl_nt, True, qtfl_glgr_list, 1000)
        End If
        '数据加和
        Dim ans_znclf(31) As Double
        Dim ans_znqtf(31) As Double
        For i = 1 To 31
            ans_znclf(i) = ans_clf_rj(i) + ans_clf_lj(i) + ans_clf_rm(i) + ans_clf_glgr(i) + ans_clf_gf(i) + ans_clf_fd(i) + ans_clf_xdc(i)
            ans_znqtf(i) = ans_qtf_rj(i) + ans_qtf_lj(i) + ans_qtf_rm(i) + ans_qtf_glgr(i) + ans_qtf_gf(i) + ans_qtf_fd(i) + ans_qtf_xdc(i)
        Next
        '返回结果
        Dim ans(15)
        ans(0) = ans_znclf
        ans(1) = ans_znqtf
        ans(2) = ans_clf_rj
        ans(3) = ans_clf_lj
        ans(4) = ans_clf_rm
        ans(5) = ans_clf_glgr
        ans(6) = ans_clf_gf
        ans(7) = ans_clf_fd
        ans(8) = ans_clf_xdc
        ans(9) = ans_qtf_rj
        ans(10) = ans_qtf_lj
        ans(11) = ans_qtf_rm
        ans(12) = ans_qtf_glgr
        ans(13) = ans_qtf_gf
        ans(14) = ans_qtf_fd
        ans(15) = ans_qtf_xdc
        Return ans
    End Function
    Function 默认逐年材料费率其它费率(ExcelApp As Object)
        '读取Excel中输入的默认材料费率
        '燃机材料费， 元 / MWh
        '燃煤发电材料费， 元 / MWh
        '垃圾发电材料费， 元 / MWh
        '供冷供热材料费， 元 / MWh
        '光伏材料费， 元 / kW
        '风电材料费， 元 / kW
        '蓄电池材料费， 元 / kW，预留功能
        Dim clf_set_rj As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(221, 22).Value
        Dim clf_set_rm As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(223, 22).Value
        Dim clf_set_ljfd As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(225, 22).Value
        Dim clf_set_glgr As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(227, 22).Value
        Dim clf_set_gf As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(229, 22).Value
        Dim clf_set_fd As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(231, 22).Value
        Dim clf_set_xdc As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(233, 22).Value
        '读取Excel中输入的默认其他费率
        '燃机其它费用， 元 / MWh
        '燃煤发电其它费用， 元 / MWh
        '垃圾发电其它费用， 元 / MWh
        '供冷供热其它费， 元 / MWh
        '光伏其它费用， 元 / kW
        '风电其它费用， 元 / kW
        '蓄电池其它费， 元 / kW，预留功能
        Dim qtf_set_rj As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(222, 22).Value
        Dim qtf_set_rm As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(224, 22).Value
        Dim qtf_set_ljfd As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(226, 22).Value
        Dim qtf_set_glgr As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(228, 22).Value
        Dim qtf_set_gf As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(230, 22).Value
        Dim qtf_set_fd As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(232, 22).Value
        Dim qtf_set_xdc As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(234, 22).Value
        '逐年材料费默认值
        Dim clfl_rj_list(31)
        Dim clfl_rm_list(31)
        Dim clfl_ljfd_list(31)
        Dim clfl_glgr_list(31)
        Dim clfl_gf_list(31)
        Dim clfl_fd_list(31)
        Dim clfl_xdc_list(31)
        '逐年其它费默认值
        Dim qtfl_rj_list(31)
        Dim qtfl_rm_list(31)
        Dim qtfl_ljfd_list(31)
        Dim qtfl_glgr_list(31)
        Dim qtfl_gf_list(31)
        Dim qtfl_fd_list(31)
        Dim qtfl_xdc_list(31)
        '逐年费率
        For i = 1 To 31
            '逐年材料费默认值
            clfl_rj_list(i) = clf_set_rj
            clfl_rm_list(i) = clf_set_rm
            clfl_ljfd_list(i) = clf_set_ljfd
            clfl_glgr_list(i) = clf_set_glgr
            clfl_gf_list(i) = clf_set_gf
            clfl_fd_list(i) = clf_set_fd
            clfl_xdc_list(i) = clf_set_xdc
            '逐年其它费默认值
            qtfl_rj_list(i) = qtf_set_rj
            qtfl_rm_list(i) = qtf_set_rm
            qtfl_ljfd_list(i) = qtf_set_ljfd
            qtfl_glgr_list(i) = qtf_set_glgr
            qtfl_gf_list(i) = qtf_set_gf
            qtfl_fd_list(i) = qtf_set_fd
            qtfl_xdc_list(i) = qtf_set_xdc
        Next
        '返回结果
        Dim ans(13)
        '逐年材料费默认值
        ans(0) = clfl_rj_list
        ans(1) = clfl_rm_list
        ans(2) = clfl_ljfd_list
        ans(3) = clfl_glgr_list
        ans(4) = clfl_gf_list
        ans(5) = clfl_fd_list
        ans(12) = clfl_xdc_list
        '逐年其它费默认值
        ans(6) = qtfl_rj_list
        ans(7) = qtfl_rm_list
        ans(8) = qtfl_ljfd_list
        ans(9) = qtfl_glgr_list
        ans(10) = qtfl_gf_list
        ans(11) = qtfl_fd_list
        ans(13) = qtfl_xdc_list
        Return ans
    End Function
    Function 材料费其它费计算基数扣除默认设置(ExcelApp As Object)
        'yynx_rj：燃机每次投资运营年限数量（年）
        'yynx_xdc：蓄电池每次投资运营年限数量（年）
        'yynx_nt：暖通每次投资运营年限数量（年）
        'yynx_gf：光伏每次投资运营年限数量（年）
        'yynx_rm：燃煤每次投资运营年限数量（年）
        'yynx_fd：风电每次投资运营年限数量（年）
        'yynx_ljfd：垃圾发电每次投资运营年限数量（年）
        'kcbl_rj：燃机每次投资运营年限结束后，上次投资计算材料费其它费的扣除比例（%）
        'kcbl_xdc：蓄电池每次投资运营年限结束后，上次投资计算材料费其它费的扣除比例（%）
        'kcbl_nt：暖通每次投资运营年限结束后，上次投资计算材料费其它费的扣除比例（%）
        'kcbl_gf：光伏每次投资运营年限结束后，上次投资计算材料费其它费的扣除比例（%）
        'kcbl_rm：燃煤每次投资运营年限结束后，上次投资计算材料费其它费的扣除比例（%）
        'kcbl_fd：风电每次投资运营年限结束后，上次投资计算材料费其它费的扣除比例（%）
        'kcbl_ljfd：垃圾发电每次投资运营年限结束后，上次投资计算材料费其它费的扣除比例（%）

        '读取计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        Dim yynx_rj As Integer = jsnx - 1
        Dim yynx_xdc As Integer = 10
        Dim yynx_nt As Integer = jsnx - 1
        Dim yynx_gf As Integer = 25
        Dim yynx_rm As Integer = jsnx - 1
        Dim yynx_fd As Integer = 20
        Dim yynx_ljfd As Integer = jsnx - 1
        Dim kcbl_rj As Double = 1
        Dim kcbl_xdc As Double = 1
        Dim kcbl_nt As Double = 1
        Dim kcbl_gf As Double = 1
        Dim kcbl_rm As Double = 1
        Dim kcbl_fd As Double = 1
        Dim kcbl_ljfd As Double = 1

        '返回结果
        Dim ans(13)
        ans(0) = yynx_rj
        ans(1) = yynx_xdc
        ans(2) = yynx_nt
        ans(3) = yynx_gf
        ans(4) = yynx_rm
        ans(5) = yynx_fd
        ans(6) = yynx_ljfd
        ans(7) = kcbl_rj
        ans(8) = kcbl_xdc
        ans(9) = kcbl_nt
        ans(10) = kcbl_gf
        ans(11) = kcbl_rm
        ans(12) = kcbl_fd
        ans(13) = kcbl_ljfd
        Return ans
    End Function
End Module
