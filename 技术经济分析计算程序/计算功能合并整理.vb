Module 计算功能合并整理
    Sub 确定估算表参数设置(ExcelApp As Object, zbj_model As Integer, hscz As Boolean, xlfl_cg_model As Integer, xlfl_qt_model As Integer, clfl_qtfl_model As Integer,
                           sdsl_model As Integer, kcje_xlf_model As Integer, kcje_clf_qtf_model As Integer, ldzj_model As Integer, kcje_ldzj_model As Integer,
                           bxf_model As Integer, kcje_bxf_model As Integer)
        '<估算表>中：项目计算年限、长期贷款相关年限、折旧摊销相关年限、可抵扣增值税年限变化后相关计算
        '————————————————————————————————————————————————————————————————————————————————————————
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        'hscy：计算期末，是否回收资产残值
        'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'ldzj_model：流动资金的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_ldzj_model：流动资金计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'bxf_model：保险费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_bxf_model：保险费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        '————————————————————————————————————————————————————————————————————————————————————————
        Call 年限系数相关计算.年限系数相关计算(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        Call 确定投资数据输入(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
    End Sub
    Sub 确定投资数据输入(ExcelApp As Object, zbj_model As Integer, hscz As Boolean, xlfl_cg_model As Integer, xlfl_qt_model As Integer, clfl_qtfl_model As Integer,
                         sdsl_model As Integer, kcje_xlf_model As Integer, kcje_clf_qtf_model As Integer, ldzj_model As Integer, kcje_ldzj_model As Integer,
                         bxf_model As Integer, kcje_bxf_model As Integer)
        '包括：改变建设期投资金额数值后相关计算+材料费其它费计算
        '————————————————————————————————————————————————————————————————————————————————————————
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        'hscy：计算期末，是否回收资产残值
        'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'ldzj_model：流动资金的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_ldzj_model：流动资金计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'bxf_model：保险费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_bxf_model：保险费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        '————————————————————————————————————————————————————————————————————————————————————————
        '建设期资金运用计算
        Call 建设期资金运用计算.建设期资金运用计算(ExcelApp, zbj_model)
        '长期贷款计算
        Call 长期贷款计算.长期贷款计算(ExcelApp, zbj_model)
        '折旧摊销计算
        Call 折旧摊销计算.折旧摊销计算(ExcelApp, hscz, zbj_model)
        '保险费计算
        Call 保险费计算.保险费计算(ExcelApp, hscz, zbj_model, bxf_model, kcje_bxf_model)
        '计算还款利润
        Call 计算还款利润(ExcelApp)
        '修理费计算
        Call 修理费计算.修理费计算(ExcelApp, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model)
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim ans_fhl_base = 默认逐年达产率()
        '风电默认逐年负荷率
        Dim fhl_fd = ans_fhl_base(0)
        '光伏默认逐年负荷率
        Dim fhl_gf = ans_fhl_base(1)
        '蓄电池默认逐年负荷率
        Dim fhl_xdc = ans_fhl_base(2)
        '————————————————————————————————————————————————————————————————————————————————————————       
        '光伏逐年衰减系数
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(46, 18).Value = "逐年投产月份比例" Then
            Dim tcyfzs_gf As Boolean = True
            Call 光伏逐年综合达产率计算(ExcelApp, fhl_gf, tcyfzs_gf)
        ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(46, 18).Value = "常规设置" Then
            Dim tcyfzs_gf As Boolean = False
            Call 光伏逐年综合达产率计算(ExcelApp, fhl_gf, tcyfzs_gf)
        End If
        '—————————————————————————————————————————————————————————————————
        '蓄电池逐年衰减系数
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(47, 18).Value = "逐年投产月份比例" Then
            Dim tcyfzs_xdc As Boolean = True
            Call 蓄电池逐年综合达产率计算(ExcelApp, fhl_xdc, "供电和购电", tcyfzs_xdc)
        ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(47, 18).Value = "常规设置" Then
            Dim tcyfzs_xdc As Boolean = False
            Call 蓄电池逐年综合达产率计算(ExcelApp, fhl_xdc, "供电和购电", tcyfzs_xdc)
        End If
        '—————————————————————————————————————————————————————————————————
        '风力发电逐年系数	
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(48, 18).Value = "逐年投产月份比例" Then
            Dim tcyfzs_fd As Boolean = True
            Call 风电逐年综合达产率计算(ExcelApp, fhl_fd, tcyfzs_fd)
        ElseIf ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(48, 18).Value = "常规设置" Then
            Dim tcyfzs_fd As Boolean = False
            Call 风电逐年综合达产率计算(ExcelApp, fhl_fd, tcyfzs_fd)
        End If
        '—————————————————————————————————————————————————————————————————
        '计算材料费和其它费
        Call 材料费其它费计算.材料费其它费计算(ExcelApp, clfl_qtfl_model, kcje_clf_qtf_model)
        '————————————————————————————————————————————————————————————————————————————————————————
        '资金筹措表计算
        Call 投资计划与资金筹措表计算(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        '增值税相关计算
        Call 增值税相关计算(ExcelApp)
        '所得税相关计算
        Call 所得税相关计算(ExcelApp, sdsl_model)
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算流动资金
        Call 流动资金相关计算(ExcelApp, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        '资产负债表计算
        Call 资产负债表计算.资产负债表计算(ExcelApp)
        '重新计算投资收益率和投资回收期
        Call 投资收益率和投资回收期计算(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
    Sub 建设投资相关计算(ExcelApp As Object, zbj_model As Integer, hscz As Boolean, xlfl_cg_model As Integer, xlfl_qt_model As Integer, sdsl_model As Integer, kcje_xlf_model As Integer,
                         clfl_qtfl_model As Integer, kcje_clf_qtf_model As Integer, ldzj_model As Integer, kcje_ldzj_model As Integer, bxf_model As Integer, kcje_bxf_model As Integer)
        '仅改变建设期投资金额数值后相关计算
        '————————————————————————————————————————————————————————————————————————————————————————
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        'hscy：计算期末，是否回收资产残值
        'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'ldzj_model：流动资金的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_ldzj_model：流动资金计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'bxf_model：保险费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_bxf_model：保险费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'bxf_model：保险费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_bxf_model：保险费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        '————————————————————————————————————————————————————————————————————————————————————————   
        '建设期资金运用计算
        Call 建设期资金运用计算.建设期资金运用计算(ExcelApp, zbj_model)
        '长期贷款计算
        Call 长期贷款计算.长期贷款计算(ExcelApp, zbj_model)
        '折旧摊销计算
        Call 折旧摊销计算.折旧摊销计算(ExcelApp, hscz, zbj_model)
        '保险费计算
        Call 保险费计算.保险费计算(ExcelApp, hscz, zbj_model, bxf_model, kcje_bxf_model)
        '修理费计算
        Call 修理费计算.修理费计算(ExcelApp, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model)
        '—————————————————————————————————————————————————————————————————
        '计算材料费和其它费
        Call 材料费其它费计算.材料费其它费计算(ExcelApp, clfl_qtfl_model, kcje_clf_qtf_model)
        '—————————————————————————————————————————————————————————————————
        '计算还款利润
        Call 计算还款利润(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        '资金筹措表计算
        Call 投资计划与资金筹措表计算(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        '增值税相关计算
        Call 增值税相关计算(ExcelApp)
        '所得税相关计算
        Call 所得税相关计算(ExcelApp, sdsl_model)
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算流动资金，包括：材料费其它费+修理费+保险费
        Call 流动资金相关计算(ExcelApp, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        '资产负债表计算
        Call 资产负债表计算.资产负债表计算(ExcelApp)
        '重新计算投资收益率和投资回收期
        Call 投资收益率和投资回收期计算(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
    Sub 折旧摊销相关计算(ExcelApp As Object, hscz As Boolean, sdsl_model As Integer, zbj_model As Integer, xlfl_cg_model As Integer, xlfl_qt_model As Integer, kcje_xlf_model As Integer,
                         clfl_qtfl_model As Integer, kcje_clf_qtf_model As Integer, ldzj_model As Integer, kcje_ldzj_model As Integer, bxf_model As Integer, kcje_bxf_model As Integer)
        'hscy：计算期末，是否回收资产残值
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        'xlfl_cg_model： 常规设备修理费率的计算方式， 0： 使用默认值， 1： 从Excel中读取已有的值
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'clfl_qtfl_model： 材料费率、其它费率的计算方式，0：使用默认值， 1： 从Excel中读取已有的值
        'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'ldzj_model：流动资金的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_ldzj_model：流动资金计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'bxf_model：保险费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_bxf_model：保险费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'bxf_model：保险费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_bxf_model：保险费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        '————————————————————————————————————————————————————————————————————————————————————————
        '折旧摊销计算
        Call 折旧摊销计算.折旧摊销计算(ExcelApp, hscz, zbj_model)
        '保险费计算
        Call 保险费计算.保险费计算(ExcelApp, hscz, zbj_model, bxf_model, kcje_bxf_model)
        '修理费计算
        Call 修理费计算.修理费计算(ExcelApp, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model)
        '计算还款利润
        Call 计算还款利润(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        '增值税相关计算
        Call 增值税相关计算(ExcelApp)
        '所得税相关计算
        Call 所得税相关计算(ExcelApp, sdsl_model)
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算流动资金
        Call 流动资金相关计算(ExcelApp, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        '资产负债表计算
        Call 资产负债表计算.资产负债表计算(ExcelApp)
        '重新计算投资收益率和投资回收期
        Call 投资收益率和投资回收期计算(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
    Sub 长期贷款相关计算(ExcelApp As Object, sdsl_model As Integer, zbj_model As Integer, xlfl_cg_model As Integer, xlfl_qt_model As Integer, kcje_xlf_model As Integer, hscz As Boolean,
                         clfl_qtfl_model As Integer, kcje_clf_qtf_model As Integer, ldzj_model As Integer, kcje_ldzj_model As Integer, bxf_model As Integer, kcje_bxf_model As Integer)
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        'xlfl_cg_model： 常规设备修理费率的计算方式， 0： 使用默认值， 1： 从Excel中读取已有的值
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'hscy：计算期末，是否回收资产残值
        'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'ldzj_model：流动资金的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_ldzj_model：流动资金计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'bxf_model：保险费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_bxf_model：保险费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'bxf_model：保险费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_bxf_model：保险费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        '————————————————————————————————————————————————————————————————————————————————————————
        '长期贷款计算
        Call 长期贷款计算.长期贷款计算(ExcelApp, zbj_model)
        '保险费计算
        Call 保险费计算.保险费计算(ExcelApp, hscz, zbj_model, bxf_model, kcje_bxf_model)
        '修理费计算
        Call 修理费计算.修理费计算(ExcelApp, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model)
        '计算还款利润
        Call 计算还款利润(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        '增值税相关计算
        Call 增值税相关计算(ExcelApp)
        '所得税相关计算
        Call 所得税相关计算(ExcelApp, sdsl_model)
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算流动资金
        Call 流动资金相关计算(ExcelApp, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        '重新计算投资收益率和投资回收期
        Call 投资收益率和投资回收期计算(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
    Sub 修理费相关计算(ExcelApp As Object, xlfl_cg_model As Integer, xlfl_qt_model As Integer, sdsl_model As Integer, kcje_xlf_model As Integer, hscz As Boolean, zbj_model As Integer,
                       clfl_qtfl_model As Integer, kcje_clf_qtf_model As Integer, ldzj_model As Integer, kcje_ldzj_model As Integer, bxf_model As Integer, kcje_bxf_model As Integer)
        'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'hscy：计算期末，是否回收资产残值
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'ldzj_model：流动资金的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_ldzj_model：流动资金计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'bxf_model：保险费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_bxf_model：保险费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        '————————————————————————————————————————————————————————————————————————————————————————
        '修理费计算
        Call 修理费计算.修理费计算(ExcelApp, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model)
        '————————————————————————————————————————————————————————————————————————————————————————
        '增值税相关计算
        Call 增值税相关计算(ExcelApp)
        '所得税相关计算
        Call 所得税相关计算(ExcelApp, sdsl_model)
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算流动资金，包括：材料费其它费+修理费+保险费
        Call 流动资金相关计算(ExcelApp, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        '重新计算投资收益率和投资回收期
        Call 投资收益率和投资回收期计算(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
    Sub 收入成本相关计算(ExcelApp As Object, sdsl_model As Integer, xlfl_cg_model As Integer, xlfl_qt_model As Integer, kcje_xlf_model As Integer, hscz As Boolean, zbj_model As Integer,
                         clfl_qtfl_model As Integer, kcje_clf_qtf_model As Integer, ldzj_model As Integer, kcje_ldzj_model As Integer, bxf_model As Integer, kcje_bxf_model As Integer)
        '仅<收入&成本输入表>中的内容改变后相关计算
        '————————————————————————————————————————————————————————————————————————————————————————
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'hscy：计算期末，是否回收资产残值
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'ldzj_model：流动资金的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_ldzj_model：流动资金计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'bxf_model：保险费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_bxf_model：保险费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次负荷率
        Call 逐年达产率计算_main(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        '材料费其它费计算，可能和负荷率有关，所以需要计算
        Call 材料费其它费计算.材料费其它费计算(ExcelApp, clfl_qtfl_model, kcje_clf_qtf_model)
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '增值税相关计算
        Call 增值税相关计算(ExcelApp)
        '所得税相关计算
        Call 所得税相关计算(ExcelApp, sdsl_model)
        '————————————————————————————————————————————————————————————————————————————————————————
        Call 隐藏收入税收表中收入为0的行(ExcelApp)
        Call 隐藏总成本表中成本为0的行(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算流动资金
        Call 流动资金相关计算(ExcelApp, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        '重新计算投资收益率和投资回收期
        Call 投资收益率和投资回收期计算(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
    Sub 税收相关计算(ExcelApp As Object, zbj_model As Integer, hscz As Boolean, xlfl_cg_model As Integer, xlfl_qt_model As Integer, clfl_qtfl_model As Integer,
                     sdsl_model As Integer, kcje_xlf_model As Integer, kcje_clf_qtf_model As Integer, ldzj_model As Integer, kcje_ldzj_model As Integer,
                     bxf_model As Integer, kcje_bxf_model As Integer)
        '影响增值税或者所得税的相关内容计算，包括投资相关、收入成本相关的全部内容
        '————————————————————————————————————————————————————————————————————————————————————————
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        'hscy：计算期末，是否回收资产残值
        'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'ldzj_model：流动资金的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_ldzj_model：流动资金计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'bxf_model：保险费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_bxf_model：保险费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        '————————————————————————————————————————————————————————————————————————————————————————
        '建设期资金运用计算
        Call 建设期资金运用计算.建设期资金运用计算(ExcelApp, zbj_model)
        '长期贷款计算
        Call 长期贷款计算.长期贷款计算(ExcelApp, zbj_model)
        '折旧摊销计算
        Call 折旧摊销计算.折旧摊销计算(ExcelApp, hscz, zbj_model)
        '保险费计算
        Call 保险费计算.保险费计算(ExcelApp, hscz, zbj_model, bxf_model, kcje_bxf_model)
        '修理费计算
        Call 修理费计算.修理费计算(ExcelApp, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model)
        '计算还款利润
        Call 计算还款利润(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算材料费和其它费
        Call 材料费其它费计算.材料费其它费计算(ExcelApp, clfl_qtfl_model, kcje_clf_qtf_model)
        '————————————————————————————————————————————————————————————————————————————————————————
        '资金筹措表计算
        Call 投资计划与资金筹措表计算(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次负荷率
        Call 逐年达产率计算_main(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '增值税相关计算
        Call 增值税相关计算(ExcelApp)
        '所得税相关计算
        Call 所得税相关计算(ExcelApp, sdsl_model)
        '————————————————————————————————————————————————————————————————————————————————————————
        Call 隐藏收入税收表中收入为0的行(ExcelApp)
        Call 隐藏总成本表中成本为0的行(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算流动资金，包括：材料费其它费+修理费+保险费
        Call 流动资金相关计算(ExcelApp, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        '重新计算投资收益率和投资回收期
        Call 投资收益率和投资回收期计算(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
    Sub 保险费相关计算(ExcelApp As Object, zbj_model As Integer, hscz As Boolean, xlfl_cg_model As Integer, xlfl_qt_model As Integer, clfl_qtfl_model As Integer,
                       sdsl_model As Integer, kcje_xlf_model As Integer, kcje_clf_qtf_model As Integer, ldzj_model As Integer, kcje_ldzj_model As Integer,
                       bxf_model As Integer, kcje_bxf_model As Integer)
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        'hscy：计算期末，是否回收资产残值
        'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'ldzj_model：流动资金的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_ldzj_model：流动资金计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'bxf_model：保险费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_bxf_model：保险费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        '————————————————————————————————————————————————————————————————————————————————————————
        '保险费计算
        Call 保险费计算.保险费计算(ExcelApp, hscz, zbj_model, bxf_model, kcje_bxf_model)
        '————————————————————————————————————————————————————————————————————————————————————————
        '增值税相关计算
        Call 增值税相关计算(ExcelApp)
        '所得税相关计算
        Call 所得税相关计算(ExcelApp, sdsl_model)
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算流动资金，包括：材料费其它费+修理费+保险费
        Call 流动资金相关计算(ExcelApp, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        '重新计算投资收益率和投资回收期
        Call 投资收益率和投资回收期计算(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
End Module
