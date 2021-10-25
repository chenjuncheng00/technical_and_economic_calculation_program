Module 材料费其它费计算
    Sub 材料费其它费计算(ExcelApp As Object, clfl_qtfl_model As Integer, kcje_clf_qtf_model As Integer)
        On Error Resume Next
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
        '逐年其它费默认值
        Dim qtfl_mr_rj_list = clqtfl(6)
        Dim qtfl_mr_rm_list = clqtfl(7)
        Dim qtfl_mr_ljfd_list = clqtfl(8)
        Dim qtfl_mr_glgr_list = clqtfl(9)
        Dim qtfl_mr_gf_list = clqtfl(10)
        Dim qtfl_mr_fd_list = clqtfl(11)
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算采用的费率
        Dim clfl_rj_list(31) As Double
        Dim clfl_rm_list(31) As Double
        Dim clfl_ljfd_list(31) As Double
        Dim clfl_glgr_list(31) As Double
        Dim clfl_gf_list(31) As Double
        Dim clfl_fd_list(31) As Double
        Dim qtfl_rj_list(31) As Double
        Dim qtfl_rm_list(31) As Double
        Dim qtfl_ljfd_list(31) As Double
        Dim qtfl_glgr_list(31) As Double
        Dim qtfl_gf_list(31) As Double
        Dim qtfl_fd_list(31) As Double
        If clfl_qtfl_model = 0 Then
            clfl_rj_list = clfl_mr_rj_list
            clfl_rm_list = clfl_mr_rm_list
            clfl_ljfd_list = clfl_mr_ljfd_list
            clfl_glgr_list = clfl_mr_glgr_list
            clfl_gf_list = clfl_mr_gf_list
            clfl_fd_list = clfl_mr_fd_list
            qtfl_rj_list = qtfl_mr_rj_list
            qtfl_rm_list = qtfl_mr_rm_list
            qtfl_ljfd_list = qtfl_mr_ljfd_list
            qtfl_glgr_list = qtfl_mr_glgr_list
            qtfl_gf_list = qtfl_mr_gf_list
            qtfl_fd_list = qtfl_mr_fd_list
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
        '燃机总发电量(万kWh)，10次投资的情况
        Dim rjfdl = GSBSJ(11)
        Dim rjfdl_list = 基础计算功能_10_to_31(tznf_list, rjfdl)
        '蓄电池总发电量(万kWh)，10次投资的情况
        Dim xdcfdl = GSBSJ(13)
        Dim xdcfdl_list = 基础计算功能_10_to_31(tznf_list, xdcfdl)
        '供冷供热总量(万kWh)，10次投资的情况
        Dim glgrl = GSBSJ(15)
        Dim glgrl_list = 基础计算功能_10_to_31(tznf_list, glgrl)
        '光伏总装机功率(kW)，10次投资的情况
        Dim gfzjgl = GSBSJ(17)
        Dim gfzjgl_list = 基础计算功能_10_to_31(tznf_list, gfzjgl)
        '燃煤机组总发电量(万kWh)，10次投资的情况
        Dim rmfdl = GSBSJ(21)
        Dim rmfdl_list = 基础计算功能_10_to_31(tznf_list, rmfdl)
        '风电总装机功率(kW)，10次投资的情况
        Dim fdzjgl = GSBSJ(19)
        Dim fdzjgl_list = 基础计算功能_10_to_31(tznf_list, fdzjgl)
        '垃圾发电总发电量(万kWh)，10次投资的情况
        Dim ljfdl = GSBSJ(20)
        Dim ljfdl_list = 基础计算功能_10_to_31(tznf_list, ljfdl)
        '读取逐年负荷率
        Dim fhl_list = 读取逐年负荷率(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        '根据扣除设置，修改逐年计算基数
        '燃机
        Dim ans_clfl_rj = 计算费率修正计算基础功能(tznf_list, rjfdl, rjfdl_list, clfl_rj_list, yynx_rj, kcbl_rj)
        Dim ans_qtfl_rj = 计算费率修正计算基础功能(tznf_list, rjfdl, rjfdl_list, qtfl_rj_list, yynx_rj, kcbl_rj)
        clfl_rj_list = ans_clfl_rj(0)
        qtfl_rj_list = ans_qtfl_rj(0)
        '暖通
        Dim ans_clfl_nt = 计算费率修正计算基础功能(tznf_list, glgrl, glgrl_list, clfl_glgr_list, yynx_nt, kcbl_nt)
        Dim ans_qtfl_nt = 计算费率修正计算基础功能(tznf_list, glgrl, glgrl_list, qtfl_glgr_list, yynx_nt, kcbl_nt)
        clfl_glgr_list = ans_clfl_nt(0)
        qtfl_glgr_list = ans_qtfl_nt(0)
        '光伏
        Dim ans_clfl_gf = 计算费率修正计算基础功能(tznf_list, gfzjgl, gfzjgl_list, clfl_gf_list, yynx_gf, kcbl_gf)
        Dim ans_qtfl_gf = 计算费率修正计算基础功能(tznf_list, gfzjgl, gfzjgl_list, qtfl_gf_list, yynx_gf, kcbl_gf)
        clfl_gf_list = ans_clfl_gf(0)
        qtfl_gf_list = ans_qtfl_gf(0)
        '燃煤
        Dim ans_clfl_rm = 计算费率修正计算基础功能(tznf_list, rmfdl, rmfdl_list, clfl_rm_list, yynx_rm, kcbl_rm)
        Dim ans_qtfl_rm = 计算费率修正计算基础功能(tznf_list, rmfdl, rmfdl_list, qtfl_rm_list, yynx_rm, kcbl_rm)
        clfl_rm_list = ans_clfl_rm(0)
        qtfl_rm_list = ans_qtfl_rm(0)
        '风电
        Dim ans_clfl_fd = 计算费率修正计算基础功能(tznf_list, fdzjgl, fdzjgl_list, clfl_fd_list, yynx_fd, kcbl_fd)
        Dim ans_qtfl_fd = 计算费率修正计算基础功能(tznf_list, fdzjgl, fdzjgl_list, qtfl_fd_list, yynx_fd, kcbl_fd)
        clfl_fd_list = ans_clfl_fd(0)
        qtfl_fd_list = ans_qtfl_fd(0)
        '垃圾发电
        Dim ans_clfl_ljfd = 计算费率修正计算基础功能(tznf_list, ljfdl, ljfdl_list, clfl_ljfd_list, yynx_ljfd, kcbl_ljfd)
        Dim ans_qtfl_ljfd = 计算费率修正计算基础功能(tznf_list, ljfdl, ljfdl_list, qtfl_ljfd_list, yynx_ljfd, kcbl_ljfd)
        clfl_ljfd_list = ans_clfl_ljfd(0)
        qtfl_ljfd_list = ans_qtfl_ljfd(0)
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算
        Dim ans_clfqtf
        Dim ans_znclf
        Dim ans_znqtf
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(11, 11).Value = "负荷率" Then
            ans_clfqtf = 材料费其它费计算_负荷率(rjfdl_list, glgrl_list, gfzjgl_list, rmfdl_list, fdzjgl_list, ljfdl_list, clfl_rj_list, clfl_rm_list, clfl_ljfd_list,
                                                 clfl_glgr_list, clfl_gf_list, clfl_fd_list, qtfl_rj_list, qtfl_rm_list, qtfl_ljfd_list, qtfl_glgr_list, qtfl_gf_list,
                                                 qtfl_fd_list, fhl_list)
            ans_znclf = ans_clfqtf(0)
            ans_znqtf = ans_clfqtf(1)
        ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(11, 11).Value = "投产量" Then
            ans_clfqtf = 材料费其它费计算_投资量(rjfdl_list, glgrl_list, gfzjgl_list, rmfdl_list, fdzjgl_list, ljfdl_list, clfl_rj_list, clfl_rm_list, clfl_ljfd_list,
                                                 clfl_glgr_list, clfl_gf_list, clfl_fd_list, qtfl_rj_list, qtfl_rm_list, qtfl_ljfd_list, qtfl_glgr_list, qtfl_gf_list,
                                                 qtfl_fd_list)
            ans_znclf = ans_clfqtf(0)
            ans_znqtf = ans_clfqtf(1)
        End If
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
    Function 材料费其它费计算_负荷率(rjfdl_list As Array, glgrl_list As Array, gfzjgl_list As Array, rmfdl_list As Array,
                                     fdzjgl_list As Array, ljfdl_list As Array, clfl_rj_list As Array, clfl_rm_list As Array,
                                     clfl_ljfd_list As Array, clfl_glgr_list As Array, clfl_gf_list As Array,
                                     clfl_fd_list As Array, qtfl_rj_list As Array, qtfl_rm_list As Array,
                                     qtfl_ljfd_list As Array, qtfl_glgr_list As Array, qtfl_gf_list As Array,
                                     qtfl_fd_list As Array, fhl_list As Array)
        'rjfdl_list, glgrl_list, gfzjgl_list, rmfdl_list, fdzjgl_list, ljfdl_list：与下面的量一一对应
        '燃机总发电量(万kWh)，10次投资的情况，列表，长度31
        '供冷供热总量(万kWh)，10次投资的情况，列表，长度31
        '光伏总装机功率(kW)，10次投资的情况，列表，长度31
        '燃煤机组总发电量(万kWh)，10次投资的情况，列表，长度31
        '风电总装机功率(kW)，10次投资的情况，列表，长度31
        '垃圾发电总发电量(万kWh)，10次投资的情况，列表，长度31
        'clfl_rj_list, clfl_rm_list, clfl_ljfd_list, clfl_glgr_list, clfl_gf_list, clfl_fd_list：与下面的材料费率一一对应
        '燃机材料费， 元 / MWh，列表，长度31
        '燃煤发电材料费， 元 / MWh，列表，长度31
        '垃圾发电材料费， 元 / MWh，列表，长度31
        '供冷供热材料费， 元 / MWh，列表，长度31
        '光伏材料费， 元 / kW，列表，长度31
        '风电材料费， 元 / kW，列表，长度31
        'qtfl_rj_list, qtfl_rm_list, qtfl_ljfd_list, qtfl_glgr_list, qtfl_gf_list, qtfl_fd_list：与下面的其它费率一一对应
        '燃机其它费用， 元 / MWh，列表，长度31
        '燃煤发电其它费用， 元 / MWh，列表，长度31
        '垃圾发电其它费用， 元 / MWh，列表，长度31
        '供冷供热其它费， 元 / MWh，列表，长度31
        '光伏其它费用， 元 / kW，列表，长度31
        '风电其它费用， 元 / kW，列表，长度31

        'fhl_list：除了光伏、风电外的设备乘以这个负荷率，列表，长度31

        '各项内容的累计值
        Dim rjfdl_lj As Double = 0
        Dim glgrl_lj As Double = 0
        Dim rmfdl_lj As Double = 0
        Dim ljfdl_lj As Double = 0
        '这几项计算31年总的累计值，然后计算中再乘以负荷率即可
        For i = 1 To 31
            rjfdl_lj += rjfdl_list(i)
            glgrl_lj += glgrl_list(i)
            rmfdl_lj += rmfdl_list(i)
            ljfdl_lj += ljfdl_list(i)
        Next
        '风电和光伏不乘以负荷率，按照装机kW的累计情况计算
        Dim gfzjgl_lj As Double = 0
        Dim fdzjgl_lj As Double = 0
        '计算逐年材料费其它费金额
        Dim ans_znclf(31) As Double
        Dim ans_znqtf(31) As Double
        For i = 1 To 31
            '计算前一年累积量，用前一年的累积量进行计算
            gfzjgl_lj += gfzjgl_list(i - 1)
            fdzjgl_lj += fdzjgl_list(i - 1)
            '计算材料费、其它费
            ans_znclf(i) = ((rjfdl_lj * clfl_rj_list(i) + ljfdl_lj * clfl_ljfd_list(i) + rmfdl_lj * clfl_rm_list(i) +
                             glgrl_lj * clfl_glgr_list(i)) * fhl_list(i) + gfzjgl_lj * clfl_gf_list(i) +
                             fdzjgl_lj * clfl_fd_list(i)) / 10000
            ans_znqtf(i) = ((rjfdl_lj * qtfl_rj_list(i) + ljfdl_lj * qtfl_ljfd_list(i) + rmfdl_lj * qtfl_rm_list(i) +
                             glgrl_lj * qtfl_glgr_list(i)) * fhl_list(i) + gfzjgl_lj * qtfl_gf_list(i) +
                             fdzjgl_lj * qtfl_fd_list(i)) / 10000
        Next
        '返回结果
        Dim ans(1)
        ans(0) = ans_znclf
        ans(1) = ans_znqtf
        Return ans
    End Function
    Function 材料费其它费计算_投资量(rjfdl_list As Array, glgrl_list As Array, gfzjgl_list As Array, rmfdl_list As Array,
                                     fdzjgl_list As Array, ljfdl_list As Array, clfl_rj_list As Array, clfl_rm_list As Array,
                                     clfl_ljfd_list As Array, clfl_glgr_list As Array, clfl_gf_list As Array,
                                     clfl_fd_list As Array, qtfl_rj_list As Array, qtfl_rm_list As Array,
                                     qtfl_ljfd_list As Array, qtfl_glgr_list As Array, qtfl_gf_list As Array,
                                     qtfl_fd_list As Array)
        'rjfdl_list, glgrl_list, gfzjgl_list, rmfdl_list, fdzjgl_list, ljfdl_list：与下面的量一一对应
        '燃机总发电量(万kWh)，10次投资的情况，列表，长度31
        '供冷供热总量(万kWh)，10次投资的情况，列表，长度31
        '光伏总装机功率(kW)，10次投资的情况，列表，长度31
        '燃煤机组总发电量(万kWh)，10次投资的情况，列表，长度31
        '风电总装机功率(kW)，10次投资的情况，列表，长度31
        '垃圾发电总发电量(万kWh)，10次投资的情况，列表，长度31
        'clfl_rj_list, clfl_rm_list, clfl_ljfd_list, clfl_glgr_list, clfl_gf_list, clfl_fd_list：与下面的材料费率一一对应
        '燃机材料费， 元 / MWh，列表，长度31
        '燃煤发电材料费， 元 / MWh，列表，长度31
        '垃圾发电材料费， 元 / MWh，列表，长度31
        '供冷供热材料费， 元 / MWh，列表，长度31
        '光伏材料费， 元 / kW，列表，长度31
        '风电材料费， 元 / kW，列表，长度31
        'qtfl_rj_list, qtfl_rm_list, qtfl_ljfd_list, qtfl_glgr_list, qtfl_gf_list, qtfl_fd_list：与下面的其它费率一一对应
        '燃机其它费用， 元 / MWh，列表，长度31
        '燃煤发电其它费用， 元 / MWh，列表，长度31
        '垃圾发电其它费用， 元 / MWh，列表，长度31
        '供冷供热其它费， 元 / MWh，列表，长度31
        '光伏其它费用， 元 / kW，列表，长度31
        '风电其它费用， 元 / kW，列表，长度31

        '各项内容的累计值
        Dim rjfdl_lj As Double = 0
        Dim glgrl_lj As Double = 0
        Dim gfzjgl_lj As Double = 0
        Dim rmfdl_lj As Double = 0
        Dim fdzjgl_lj As Double = 0
        Dim ljfdl_lj As Double = 0
        '计算逐年材料费其它费金额
        Dim ans_znclf(31) As Double
        Dim ans_znqtf(31) As Double
        For i = 1 To 31
            '计算前一年累积量，用前一年的累积量进行计算
            rjfdl_lj += rjfdl_list(i - 1)
            glgrl_lj += glgrl_list(i - 1)
            gfzjgl_lj += gfzjgl_list(i - 1)
            rmfdl_lj += rmfdl_list(i - 1)
            fdzjgl_lj += fdzjgl_list(i - 1)
            ljfdl_lj += ljfdl_list(i - 1)
            '计算材料费、其它费
            ans_znclf(i) = (rjfdl_lj * clfl_rj_list(i) + ljfdl_lj * clfl_ljfd_list(i) + rmfdl_lj * clfl_rm_list(i) +
                            glgrl_lj * clfl_glgr_list(i) + gfzjgl_lj * clfl_gf_list(i) + fdzjgl_lj * clfl_fd_list(i)) / 10000
            ans_znqtf(i) = (rjfdl_lj * qtfl_rj_list(i) + ljfdl_lj * qtfl_ljfd_list(i) + rmfdl_lj * qtfl_rm_list(i) +
                            glgrl_lj * qtfl_glgr_list(i) + gfzjgl_lj * qtfl_gf_list(i) + fdzjgl_lj * qtfl_fd_list(i)) / 10000
        Next
        '返回结果
        Dim ans(1)
        ans(0) = ans_znclf
        ans(1) = ans_znqtf
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
        Dim clf_set_rj As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(222, 22).Value
        Dim clf_set_rm As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(224, 22).Value
        Dim clf_set_ljfd As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(226, 22).Value
        Dim clf_set_glgr As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(228, 22).Value
        Dim clf_set_gf As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(230, 22).Value
        Dim clf_set_fd As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(232, 22).Value
        '读取Excel中输入的默认其他费率
        '燃机其它费用， 元 / MWh
        '燃煤发电其它费用， 元 / MWh
        '垃圾发电其它费用， 元 / MWh
        '供冷供热其它费， 元 / MWh
        '光伏其它费用， 元 / kW
        '风电其它费用， 元 / kW
        Dim qtf_set_rj As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(223, 22).Value
        Dim qtf_set_rm As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(225, 22).Value
        Dim qtf_set_ljfd As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(227, 22).Value
        Dim qtf_set_glgr As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(229, 22).Value
        Dim qtf_set_gf As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(231, 22).Value
        Dim qtf_set_fd As Double = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(233, 22).Value
        '逐年材料费默认值
        Dim clfl_rj_list(31)
        Dim clfl_rm_list(31)
        Dim clfl_ljfd_list(31)
        Dim clfl_glgr_list(31)
        Dim clfl_gf_list(31)
        Dim clfl_fd_list(31)
        '逐年其它费默认值
        Dim qtfl_rj_list(31)
        Dim qtfl_rm_list(31)
        Dim qtfl_ljfd_list(31)
        Dim qtfl_glgr_list(31)
        Dim qtfl_gf_list(31)
        Dim qtfl_fd_list(31)
        '逐年费率
        For i = 1 To 31
            '逐年材料费默认值
            clfl_rj_list(i) = clf_set_rj
            clfl_rm_list(i) = clf_set_rm
            clfl_ljfd_list(i) = clf_set_ljfd
            clfl_glgr_list(i) = clf_set_glgr
            clfl_gf_list(i) = clf_set_gf
            clfl_fd_list(i) = clf_set_fd
            '逐年其它费默认值
            qtfl_rj_list(i) = qtf_set_rj
            qtfl_rm_list(i) = qtf_set_rm
            qtfl_ljfd_list(i) = qtf_set_ljfd
            qtfl_glgr_list(i) = qtf_set_glgr
            qtfl_gf_list(i) = qtf_set_gf
            qtfl_fd_list(i) = qtf_set_fd
        Next
        '返回结果
        Dim ans(11)
        '逐年材料费默认值
        ans(0) = clfl_rj_list
        ans(1) = clfl_rm_list
        ans(2) = clfl_ljfd_list
        ans(3) = clfl_glgr_list
        ans(4) = clfl_gf_list
        ans(5) = clfl_fd_list
        '逐年其它费默认值
        ans(6) = qtfl_rj_list
        ans(7) = qtfl_rm_list
        ans(8) = qtfl_ljfd_list
        ans(9) = qtfl_glgr_list
        ans(10) = qtfl_gf_list
        ans(11) = qtfl_fd_list
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
        Dim yynx_gf As Integer = jsnx - 1
        Dim yynx_rm As Integer = jsnx - 1
        Dim yynx_fd As Integer = jsnx - 1
        Dim yynx_ljfd As Integer = jsnx - 1
        Dim kcbl_rj As Double = 0
        Dim kcbl_xdc As Double = 1
        Dim kcbl_nt As Double = 0
        Dim kcbl_gf As Double = 0
        Dim kcbl_rm As Double = 0
        Dim kcbl_fd As Double = 0
        Dim kcbl_ljfd As Double = 0

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
