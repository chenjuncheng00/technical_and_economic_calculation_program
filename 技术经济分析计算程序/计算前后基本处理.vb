Module 计算前后基本处理
    Sub 计算前基本处理(ExcelApp As Object, jbcl_mode As Integer)
        '启动迭代计算
        ExcelApp.Application.Iteration = True
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        '屏蔽屏幕更新
        ExcelApp.Application.ScreenUpdating = False
        '屏蔽事件
        ExcelApp.Application.EnableEvents = False
        '不一定每次都需要关闭Excel自动计算
        If jbcl_mode = 1 Then
            '手动计算，关闭excel的自动计算
            ExcelApp.Application.Calculation = XlCalculation.xlCalculationManual
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
        '解锁表格
        Call 解锁表格(ExcelApp)
    End Sub
    Sub 计算后基本处理(ExcelApp As Object, js_mode As Integer, xlfl_cg_model As Integer, xlfl_qt_model As Integer, kcje_xlf_model As Integer, hscz As Boolean, zbj_model As Integer,
                       clfl_qtfl_model As Integer, kcje_clf_qtf_model As Integer, ldzj_model As Integer, kcje_ldzj_model As Integer, bxf_model As Integer, kcje_bxf_model As Integer)
        'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'hscy：计算期末，是否回收资产残值
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        'ldzj_model：流动资金的计算方式，0：使用默认值，1：从Excel中读取已有的值，预留功能
        'kcje_ldzj_model：流动资金计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值，预留功能
        'bxf_model：保险费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        'kcje_bxf_model：保险费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        '————————————————————————————————————————————————————————————————————————————————————————  
        '计算一次工作簿
        ExcelApp.Calculate()
        '打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
        '打开事件
        ExcelApp.Application.EnableEvents = True
        '打开excel的自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '不是每次都需要进行税收、流动资金、收益率、回收期计算
        If js_mode = 1 Then
            '增值税相关计算
            Call 增值税相关计算(ExcelApp)
            '流动资金计算
            Call 流动资金相关计算(ExcelApp, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            '投资回收期计算
            Call 投资收益率和投资回收期计算(ExcelApp)
        End If
        '锁定表格
        Call 锁定表格(ExcelApp)
    End Sub
End Module
