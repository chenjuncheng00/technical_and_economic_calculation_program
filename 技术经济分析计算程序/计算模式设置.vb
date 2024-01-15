Module 计算模式设置
    Function 计算模式_从EXCEL读取(ExcelApp As Object)
        'zbj_model：资本金计算模式，0：以动态投资为计算基础（默认值），1：以静态投资为计算基础，数据来自用户设置
        Dim zbj_model As Integer
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 7).Value = "动态" Then
            zbj_model = 0
        Else
            zbj_model = 1
        End If
        'hscy：计算期末，是否回收资产残值（默认回收）
        Dim hscz As Boolean
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(53, 18).Value = "期末不回收残值" Then
            hscz = False
        Else
            hscz = True
        End If
        'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim xlfl_cg_model As Integer = 1
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim xlfl_qt_model As Integer = 1
        'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim clfl_qtfl_model As Integer = 1
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim sdsl_model As Integer = 1
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_xlf_model As Integer = 1
        'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_clf_qtf_model As Integer = 1
        'ldzj_model：流动资金的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim ldzj_model As Integer = 1
        'kcje_ldzj_model：流动资金计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_ldzj_model As Integer = 1
        'bxf_model：保险费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim bxf_model As Integer = 1
        'kcje_bxf_model：保险费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_bxf_model As Integer = 1
        '————————————————————————————————————————————————————————————————————————————————————————
        '返回结果
        Dim ans(11)
        ans(0) = zbj_model
        ans(1) = hscz
        ans(2) = xlfl_cg_model
        ans(3) = xlfl_qt_model
        ans(4) = clfl_qtfl_model
        ans(5) = sdsl_model
        ans(6) = kcje_xlf_model
        ans(7) = kcje_clf_qtf_model
        ans(8) = ldzj_model
        ans(9) = kcje_ldzj_model
        ans(10) = bxf_model
        ans(11) = kcje_bxf_model
        Return ans
    End Function
    Function 计算模式_默认值()
        'zbj_model：资本金计算模式，0：以动态投资为计算基础（默认值），1：以静态投资为计算基础，数据来自用户设置
        Dim zbj_model As Integer = 0
        'hscy：计算期末，是否回收资产残值（默认回收）
        Dim hscz As Boolean = True
        'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim xlfl_cg_model As Integer = 0
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim xlfl_qt_model As Integer = 0
        'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim clfl_qtfl_model As Integer = 0
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim sdsl_model As Integer = 0
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_xlf_model As Integer = 0
        'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_clf_qtf_model As Integer = 0
        'ldzj_model：流动资金的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim ldzj_model As Integer = 0
        'kcje_ldzj_model：流动资金计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_ldzj_model As Integer = 0
        'bxf_model：保险费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim bxf_model As Integer = 0
        'kcje_bxf_model：保险费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_bxf_model As Integer = 0
        '————————————————————————————————————————————————————————————————————————————————————————
        '返回结果
        Dim ans(11)
        ans(0) = zbj_model
        ans(1) = hscz
        ans(2) = xlfl_cg_model
        ans(3) = xlfl_qt_model
        ans(4) = clfl_qtfl_model
        ans(5) = sdsl_model
        ans(6) = kcje_xlf_model
        ans(7) = kcje_clf_qtf_model
        ans(8) = ldzj_model
        ans(9) = kcje_ldzj_model
        ans(10) = bxf_model
        ans(11) = kcje_bxf_model
        Return ans
    End Function
End Module
