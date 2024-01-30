Public Class 设置建设期资金运用方式
    Sub 建设期默认资金运用模式(ZBJBL As Double, DKLL As Double)
        '每次投资的资本金比例和建设期贷款利率均相同，资本金比例为占动态投资比例
        '————————————————————————————————————————————————————————————————————————————————————————
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '建设期贷款计息次数
        Dim jsqdkjxcs = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 9).Value
        '————————————————————————————————————————————————————————————————————————————————————————
        '将资本金比例写入Excel
        For i = 1 To 10
            '常规设备
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 4, 22).Value = ZBJBL
            '其它设备
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 15, 22).Value = ZBJBL
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 26, 22).Value = ZBJBL
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 4, 25).Value = ZBJBL
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 15, 25).Value = ZBJBL
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 26, 25).Value = ZBJBL
        Next
        '计算复利贷款利率
        Dim dkll_fl As Double = (1 + DKLL / jsqdkjxcs) ^ jsqdkjxcs - 1
        '将实际的建设期贷款利率写入Excel
        '常规设备
        For i = 1 To 10
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i + 11, 39).Value = dkll_fl
        Next
        '其它设备
        For i = 5 To 14
            '长期贷款利率
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 42).Value = dkll_fl
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 45).Value = dkll_fl
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 48).Value = dkll_fl
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 51).Value = dkll_fl
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 54).Value = dkll_fl
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
        '——————————————————————————————————————————————————————————————————————————————————————
        '在表格中写入当前采用的资金运用模式
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 7).Value = "动态"
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 7).Value = "相同"
        '——————————————————————————————————————————————————————————————————————————————————————
        '建设期时间计划计算
        Call 建设期时间计划计算.建设期时间计划计算(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算模式
        Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
        Dim zbj_model = ans_mode(0)
        Dim hscz = ans_mode(1)
        Dim xlfl_cg_model = ans_mode(2)
        Dim xlfl_qt_model = ans_mode(3)
        Dim clfl_qtfl_model = ans_mode(4)
        Dim sdsl_model = ans_mode(5)
        Dim kcje_xlf_model = ans_mode(6)
        Dim kcje_clf_qtf_model = ans_mode(7)
        Dim ldzj_model = ans_mode(8)
        Dim kcje_ldzj_model = ans_mode(9)
        Dim bxf_model = ans_mode(10)
        Dim kcje_bxf_model = ans_mode(11)
        '————————————————————————————————————————————————————————————————————————————————————————
        '确定估算表参数设置(包括了投资金额变化后计算、收入成本变化后计算)
        Call 确定估算表参数设置(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
    End Sub

    Private Sub 资本金比例为占动态投资比例_统一设置_Click(sender As Object, e As EventArgs) Handles 资本金比例为占动态投资比例_统一设置.Click
        Dim XZ = MsgBox("是否需要将资本金比例计算模式切换为占动态总投资的比例，同时每次投资的资本金比例和建设期贷款利率均相同？", vbOKCancel)
        If XZ = vbOK Then
            Dim ZBJBL = CType(Me.zbjbl_a.Text, Double) / 100
            Dim DKLL = CType(Me.dkll_a.Text, Double) / 100
            Call 建设期默认资金运用模式(ZBJBL, DKLL)
            Me.RichTextBox1.Text = "计算完成：资本金比例计算模式切换为占动态总投资的比例，同时每次投资的资本金比例和建设期贷款利率均相同！"
        End If
    End Sub

    Private Sub 资本金比例为占静态投资比例_统一设置_Click(sender As Object, e As EventArgs) Handles 资本金比例为占静态投资比例_统一设置.Click
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp        
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '读取资本金(动态)比例，建设期贷款利率
        Dim ZBJBL = CType(Me.zbjbl_a.Text, Double) / 100
        Dim DKLL = CType(Me.dkll_a.Text, Double) / 100
        '建设期贷款计息次数
        Dim jsqdkjxcs = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 9).Value
        '————————————————————————————————————————————————————————————————————————————————————————
        'Excel表格中的资本金比例均为占动态投资的比例，设置的如果是占静态投资比例，则实际上动态投资比例比设定的值要低，则往下减，直到满足要求
        Dim XZ = MsgBox("是否需要将资本金比例计算模式切换为占静态总投资的比例，同时每次投资的资本金比例和建设期贷款利率均相同？", vbOKCancel)
        If XZ = vbOK Then
            '计算复利贷款利率
            Dim dkll_fl As Double = (1 + DKLL / jsqdkjxcs) ^ jsqdkjxcs - 1
            '将实际的建设期贷款利率写入Excel
            '常规设备
            For i = 1 To 10
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i + 11, 39).Value = dkll_fl
            Next
            '其它设备
            For i = 5 To 14
                '长期贷款利率
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 42).Value = dkll_fl
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 45).Value = dkll_fl
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 48).Value = dkll_fl
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 51).Value = dkll_fl
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 54).Value = dkll_fl
            Next
            '将资本金比例写入Excel
            For i = 1 To 10
                '常规设备
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 4, 22).Value = ZBJBL
                '其它设备
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 15, 22).Value = ZBJBL
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 26, 22).Value = ZBJBL
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 4, 25).Value = ZBJBL
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 15, 25).Value = ZBJBL
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 26, 25).Value = ZBJBL
            Next
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算一次工作簿
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————
            '在表格中写入当前采用的资金运用模式
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 7).Value = "静态"
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 7).Value = "相同"
            '————————————————————————————————————————————————————————————————————————————————————————
            '建设期时间计划计算
            Call 建设期时间计划计算.建设期时间计划计算(ExcelApp)
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            '确定估算表参数设置(包括了投资金额变化后计算、收入成本变化后计算)
            Call 确定估算表参数设置(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            '————————————————————————————————————————————————————————————————————————————————————————  
            Me.RichTextBox1.Text = "计算完成：资本金比例计算模式切换为占静态总投资的比例，同时每次投资的资本金比例和建设期贷款利率均相同！"
        End If
    End Sub

    Private Sub 资本金比例为占动态投资比例_逐次计算_Click(sender As Object, e As EventArgs) Handles 资本金比例为占动态投资比例_逐次计算.Click
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————       
        Dim XZ = MsgBox("是否需要将资本金比例计算模式切换为占动态总投资的比例，同时每次投资的资本金比例和建设期贷款利率均不相同？", vbOKCancel)
        If XZ = vbOK Then
            '计算一次工作簿
            ExcelApp.Calculate()
            '——————————————————————————————————————————————————————————————————————————————————————
            '在表格中写入当前采用的资金运用模式
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 7).Value = "动态"
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 7).Value = "不相同"
            '——————————————————————————————————————————————————————————————————————————————————————
            '建设期时间计划计算
            Call 建设期时间计划计算.建设期时间计划计算(ExcelApp)
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            '确定估算表参数设置(包括了投资金额变化后计算、收入成本变化后计算)
            Call 确定估算表参数设置(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            '————————————————————————————————————————————————————————————————————————————————————————  
            Me.RichTextBox1.Text = "计算完成：资本金比例计算模式切换为占动态总投资的比例，同时每次投资的资本金比例和建设期贷款利率均不相同！"
        End If
    End Sub

    Private Sub 资本金比例为占静态投资比例_逐次计算_Click(sender As Object, e As EventArgs) Handles 资本金比例为占静态投资比例_逐次计算.Click
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        Dim XZ = MsgBox("是否需要将资本金比例计算模式切换为占静态总投资的比例，同时每次投资的资本金比例和建设期贷款利率均不相同？", vbOKCancel)
        If XZ = vbOK Then
            '计算一次工作簿
            ExcelApp.Calculate()
            '——————————————————————————————————————————————————————————————————————————————————————
            '在表格中写入当前采用的资金运用模式
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 7).Value = "静态"
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 7).Value = "不相同"
            '——————————————————————————————————————————————————————————————————————————————————————
            '建设期时间计划计算
            Call 建设期时间计划计算.建设期时间计划计算(ExcelApp)
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            '确定估算表参数设置(包括了投资金额变化后计算、收入成本变化后计算)
            Call 确定估算表参数设置(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            '————————————————————————————————————————————————————————————————————————————————————————  
            Me.RichTextBox1.Text = "计算完成：资本金比例计算模式切换为占静态总投资的比例，同时每次投资的资本金比例和建设期贷款利率均不相同！"
        End If
    End Sub

    Private Sub 重置回默认计算模式_Click(sender As Object, e As EventArgs) Handles 重置回默认计算模式.Click
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp        
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '读取资本金(动态)比例，建设期贷款利率
        Dim ZBJBL = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 5).Value
        Dim DKLL = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 9).Value
        Dim XZ = MsgBox("确定要重置回默认计算模式？默认计算模式为每次投资的资本金比例和建设期贷款利率均相同,资本金比例为占动态投资的比例。", vbOKCancel)
        If XZ = vbOK Then
            Call 建设期默认资金运用模式(ZBJBL, DKLL)
            Me.RichTextBox1.Text = "计算完成：每次投资的资本金比例和建设期贷款利率均相同,资本金比例为占动态投资的比例！"
        End If
    End Sub

    Private Sub 清空数据_Click(sender As Object, e As EventArgs) Handles 清空数据.Click
        Dim XZ = MsgBox("确定要清空本窗体输入的的全部内容？", vbOKCancel)
        If XZ = vbOK Then
            Me.zbjbl1.Clear()
            Me.zbjbl2.Clear()
            Me.zbjbl3.Clear()
            Me.zbjbl4.Clear()
            Me.zbjbl5.Clear()
            Me.zbjbl6.Clear()
            Me.zbjbl7.Clear()
            Me.zbjbl8.Clear()
            Me.zbjbl9.Clear()
            Me.zbjbl10.Clear()
            Me.dkll1.Clear()
            Me.dkll2.Clear()
            Me.dkll3.Clear()
            Me.dkll4.Clear()
            Me.dkll5.Clear()
            Me.dkll6.Clear()
            Me.dkll7.Clear()
            Me.dkll8.Clear()
            Me.dkll9.Clear()
            Me.dkll10.Clear()
            '选择
            Me.changgui.Checked = False
            Me.ranji.Checked = False
            Me.xudianchi.Checked = False
            Me.nuantong.Checked = False
            Me.guangfu.Checked = False
            Me.fengdian.Checked = False
            '显示
            Me.RichTextBox1.Clear()
        End If
    End Sub

    Private Sub 设置建设期资金运用方式_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '清空已有的全部数据
        Me.zbjbl1.Clear()
        Me.zbjbl2.Clear()
        Me.zbjbl3.Clear()
        Me.zbjbl4.Clear()
        Me.zbjbl5.Clear()
        Me.zbjbl6.Clear()
        Me.zbjbl7.Clear()
        Me.zbjbl8.Clear()
        Me.zbjbl9.Clear()
        Me.zbjbl10.Clear()
        Me.dkll1.Clear()
        Me.dkll2.Clear()
        Me.dkll3.Clear()
        Me.dkll4.Clear()
        Me.dkll5.Clear()
        Me.dkll6.Clear()
        Me.dkll7.Clear()
        Me.dkll8.Clear()
        Me.dkll9.Clear()
        Me.dkll10.Clear()
        Me.RichTextBox1.Clear()
        '建设期贷款计息次数
        Dim jsqdkjxcs = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 9).Value
        '根据已经输入的投资情况，载入默认值
        '载入默认值
        '每次的资本金(动态)比例和建设期贷款利率均一致的情况
        Me.zbjbl_a.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 5).Value * 100, String)
        Me.dkll_a.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 9).Value * 100, String)
        '读取上次修改后的数据
        '第1次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 3).Value > 0 Then
            Me.zbjbl1.Text = CType(ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(5, 22).Value * 100, String)
            Me.dkll1.Text = CType(Math.Round(((ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(12, 39).Value + 1) ^ (1 / jsqdkjxcs) - 1) * jsqdkjxcs * 100, 2), String)
        End If
        '第2次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value > 0 Then
            Me.zbjbl2.Text = CType(ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(6, 22).Value * 100, String)
            Me.dkll2.Text = CType(Math.Round(((ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(13, 39).Value + 1) ^ (1 / jsqdkjxcs) - 1) * jsqdkjxcs * 100, 2), String)
        End If
        '第3次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value > 0 Then
            Me.zbjbl3.Text = CType(ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(7, 22).Value * 100, String)
            Me.dkll3.Text = CType(Math.Round(((ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(14, 39).Value + 1) ^ (1 / jsqdkjxcs) - 1) * jsqdkjxcs * 100, 2), String)
        End If
        '第4次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value > 0 Then
            Me.zbjbl4.Text = CType(ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(8, 22).Value * 100, String)
            Me.dkll4.Text = CType(Math.Round(((ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(15, 39).Value + 1) ^ (1 / jsqdkjxcs) - 1) * jsqdkjxcs * 100, 2), String)
        End If
        '第5次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value > 0 Then
            Me.zbjbl5.Text = CType(ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(9, 22).Value * 100, String)
            Me.dkll5.Text = CType(Math.Round(((ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(16, 39).Value + 1) ^ (1 / jsqdkjxcs) - 1) * jsqdkjxcs * 100, 2), String)
        End If
        '第6次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value > 0 Then
            Me.zbjbl6.Text = CType(ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(10, 22).Value * 100, String)
            Me.dkll6.Text = CType(Math.Round(((ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(17, 39).Value + 1) ^ (1 / jsqdkjxcs) - 1) * jsqdkjxcs * 100, 2), String)
        End If
        '第7次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value > 0 Then
            Me.zbjbl7.Text = CType(ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(11, 22).Value * 100, String)
            Me.dkll7.Text = CType(Math.Round(((ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(18, 39).Value + 1) ^ (1 / jsqdkjxcs) - 1) * jsqdkjxcs * 100, 2), String)
        End If
        '第8次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value > 0 Then
            Me.zbjbl8.Text = CType(ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(12, 22).Value * 100, String)
            Me.dkll8.Text = CType(Math.Round(((ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(19, 39).Value + 1) ^ (1 / jsqdkjxcs) - 1) * jsqdkjxcs * 100, 2), String)
        End If
        '第9次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value > 0 Then
            Me.zbjbl9.Text = CType(ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(13, 22).Value * 100, String)
            Me.dkll9.Text = CType(Math.Round(((ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(20, 39).Value + 1) ^ (1 / jsqdkjxcs) - 1) * jsqdkjxcs * 100, 2), String)
        End If
        '第10次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value > 0 Then
            Me.zbjbl10.Text = CType(ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(14, 22).Value * 100, String)
            Me.dkll10.Text = CType(Math.Round(((ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(21, 39).Value + 1) ^ (1 / jsqdkjxcs) - 1) * jsqdkjxcs * 100, 2), String)
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '根据输入的投资情况，确定每个checkbox和投资比例输入是否可以选择和输入，并载入默认值
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(44, 1).Value > 0 Then
            Me.ranji.Enabled = True
            Me.ranji.Checked = False
        Else
            Me.ranji.Enabled = False
            Me.ranji.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(46, 1).Value > 0 Then
            Me.xudianchi.Enabled = True
            Me.xudianchi.Checked = False
        Else
            Me.xudianchi.Enabled = False
            Me.xudianchi.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(48, 1).Value > 0 Then
            Me.nuantong.Enabled = True
            Me.nuantong.Checked = False
        Else
            Me.nuantong.Enabled = False
            Me.nuantong.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(50, 1).Value > 0 Then
            Me.guangfu.Enabled = True
            Me.guangfu.Checked = False
        Else
            Me.guangfu.Enabled = False
            Me.guangfu.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(53, 1).Value > 0 Then
            Me.fengdian.Enabled = True
            Me.fengdian.Checked = False
        Else
            Me.fengdian.Enabled = False
            Me.fengdian.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
            Me.changgui.Enabled = True
            Me.changgui.Checked = False
        Else
            Me.changgui.Enabled = False
            Me.changgui.Checked = False
        End If
    End Sub

    Private Sub 写入参数_Click(sender As Object, e As EventArgs) Handles 写入参数.Click
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————       
        '定义局部变量
        Dim zbjbl1, zbjbl2, zbjbl3, zbjbl4, zbjbl5, zbjbl6, zbjbl7, zbjbl8, zbjbl9, zbjbl10 As Double '资本金比例
        Dim dkll1, dkll2, dkll3, dkll4, dkll5, dkll6, dkll7, dkll8, dkll9, dkll10 As Double '建设期贷款利率
        '————————————————————————————————————————————————————————————————————————————————————————  
        '读取输入的资本金(动态)比例
        If Me.zbjbl1.Text = "" Then
            zbjbl1 = CType(Me.zbjbl_a.Text, Double) / 100
        Else
            zbjbl1 = CType(Me.zbjbl1.Text, Double) / 100
        End If
        If Me.zbjbl2.Text = "" Then
            zbjbl2 = CType(Me.zbjbl_a.Text, Double) / 100
        Else
            zbjbl2 = CType(Me.zbjbl2.Text, Double) / 100
        End If
        If Me.zbjbl3.Text = "" Then
            zbjbl3 = CType(Me.zbjbl_a.Text, Double) / 100
        Else
            zbjbl3 = CType(Me.zbjbl3.Text, Double) / 100
        End If
        If Me.zbjbl4.Text = "" Then
            zbjbl4 = CType(Me.zbjbl_a.Text, Double) / 100
        Else
            zbjbl4 = CType(Me.zbjbl4.Text, Double) / 100
        End If
        If Me.zbjbl5.Text = "" Then
            zbjbl5 = CType(Me.zbjbl_a.Text, Double) / 100
        Else
            zbjbl5 = CType(Me.zbjbl5.Text, Double) / 100
        End If
        If Me.zbjbl6.Text = "" Then
            zbjbl6 = CType(Me.zbjbl_a.Text, Double) / 100
        Else
            zbjbl6 = CType(Me.zbjbl6.Text, Double) / 100
        End If
        If Me.zbjbl7.Text = "" Then
            zbjbl7 = CType(Me.zbjbl_a.Text, Double) / 100
        Else
            zbjbl7 = CType(Me.zbjbl7.Text, Double) / 100
        End If
        If Me.zbjbl8.Text = "" Then
            zbjbl8 = CType(Me.zbjbl_a.Text, Double) / 100
        Else
            zbjbl8 = CType(Me.zbjbl8.Text, Double) / 100
        End If
        If Me.zbjbl9.Text = "" Then
            zbjbl9 = CType(Me.zbjbl_a.Text, Double) / 100
        Else
            zbjbl9 = CType(Me.zbjbl9.Text, Double) / 100
        End If
        If Me.zbjbl10.Text = "" Then
            zbjbl10 = CType(Me.zbjbl_a.Text, Double) / 100
        Else
            zbjbl10 = CType(Me.zbjbl10.Text, Double) / 100
        End If
        '读取输入的建设期贷款利率
        If Me.dkll1.Text = "" Then
            dkll1 = CType(Me.dkll_a.Text, Double) / 100
        Else
            dkll1 = CType(Me.dkll1.Text, Double) / 100
        End If
        If Me.dkll2.Text = "" Then
            dkll2 = CType(Me.dkll_a.Text, Double) / 100
        Else
            dkll2 = CType(Me.dkll2.Text, Double) / 100
        End If
        If Me.dkll3.Text = "" Then
            dkll3 = CType(Me.dkll_a.Text, Double) / 100
        Else
            dkll3 = CType(Me.dkll3.Text, Double) / 100
        End If
        If Me.dkll4.Text = "" Then
            dkll4 = CType(Me.dkll_a.Text, Double) / 100
        Else
            dkll4 = CType(Me.dkll4.Text, Double) / 100
        End If
        If Me.dkll5.Text = "" Then
            dkll5 = CType(Me.dkll_a.Text, Double) / 100
        Else
            dkll5 = CType(Me.dkll5.Text, Double) / 100
        End If
        If Me.dkll6.Text = "" Then
            dkll6 = CType(Me.dkll_a.Text, Double) / 100
        Else
            dkll6 = CType(Me.dkll6.Text, Double) / 100
        End If
        If Me.dkll7.Text = "" Then
            dkll7 = CType(Me.dkll_a.Text, Double) / 100
        Else
            dkll7 = CType(Me.dkll7.Text, Double) / 100
        End If
        If Me.dkll8.Text = "" Then
            dkll8 = CType(Me.dkll_a.Text, Double) / 100
        Else
            dkll8 = CType(Me.dkll8.Text, Double) / 100
        End If
        If Me.dkll9.Text = "" Then
            dkll9 = CType(Me.dkll_a.Text, Double) / 100
        Else
            dkll9 = CType(Me.dkll9.Text, Double) / 100
        End If
        If Me.dkll10.Text = "" Then
            dkll10 = CType(Me.dkll_a.Text, Double) / 100
        Else
            dkll10 = CType(Me.dkll10.Text, Double) / 100
        End If
        '建设期贷款计息次数
        Dim jsqdkjxcs = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 9).Value
        '————————————————————————————————————————————————————————————————————————————————————————  
        '检测有投资的年份，是否都输入的资本金(动态)比例和建设期贷款利率，如果不正确报错，正确则写入数据（贷款利息全部采用复利计算）
        '第1次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 3).Value > 0 And (zbjbl1 <= 0 Or dkll1 <= 0) Then
            MsgBox("第1次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
            Exit Sub
        End If
        '第2次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 5).Value > 0 And (zbjbl2 <= 0 Or dkll2 <= 0) Then
            MsgBox("第2次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
            Exit Sub
        End If
        '第3次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 7).Value > 0 And (zbjbl3 <= 0 Or dkll3 <= 0) Then
            MsgBox("第3次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
            Exit Sub
        End If
        '第4次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 9).Value > 0 And (zbjbl4 <= 0 Or dkll4 <= 0) Then
            MsgBox("第4次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
            Exit Sub
        End If
        '第5次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 11).Value > 0 And (zbjbl5 <= 0 Or dkll5 <= 0) Then
            MsgBox("第5次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
            Exit Sub
        End If
        '第6次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 3).Value > 0 And (zbjbl6 <= 0 Or dkll6 <= 0) Then
            MsgBox("第6次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
            Exit Sub
        End If
        '第7次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 5).Value > 0 And (zbjbl7 <= 0 Or dkll7 <= 0) Then
            MsgBox("第7次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
            Exit Sub
        End If
        '第8次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 7).Value > 0 And (zbjbl8 <= 0 Or dkll8 <= 0) Then
            MsgBox("第8次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
            Exit Sub
        End If
        '第9次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 9).Value > 0 And (zbjbl9 <= 0 Or dkll9 <= 0) Then
            MsgBox("第9次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
            Exit Sub
        End If
        '第10次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 11).Value > 0 And (zbjbl10 <= 0 Or dkll10 <= 0) Then
            MsgBox("第10次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————  
        '常规设备
        If Me.changgui.Checked = True Then
            '第1次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(5, 22).Value = zbjbl1
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(12, 39).Value = (1 + dkll1 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第2次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(6, 22).Value = zbjbl2
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(13, 39).Value = (1 + dkll2 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第3次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(7, 22).Value = zbjbl3
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(14, 39).Value = (1 + dkll3 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第4次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(8, 22).Value = zbjbl4
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(15, 39).Value = (1 + dkll4 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第5次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(9, 22).Value = zbjbl5
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(16, 39).Value = (1 + dkll5 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第6次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(10, 22).Value = zbjbl6
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(17, 39).Value = (1 + dkll6 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第7次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(11, 22).Value = zbjbl7
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(18, 39).Value = (1 + dkll7 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第8次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(12, 22).Value = zbjbl8
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(19, 39).Value = (1 + dkll8 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第9次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(13, 22).Value = zbjbl9
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(20, 39).Value = (1 + dkll9 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第10次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(14, 22).Value = zbjbl10
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(21, 39).Value = (1 + dkll10 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '显示
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<常规设备>建设期资金运用方式计算参数设置写入完成！"
        End If
        '燃机
        If Me.ranji.Checked = True Then
            '第1次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(16, 22).Value = zbjbl1
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(5, 42).Value = (1 + dkll1 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第2次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(17, 22).Value = zbjbl2
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(6, 42).Value = (1 + dkll2 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第3次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(18, 22).Value = zbjbl3
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(7, 42).Value = (1 + dkll3 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第4次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(19, 22).Value = zbjbl4
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(8, 42).Value = (1 + dkll4 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第5次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(20, 22).Value = zbjbl5
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(9, 42).Value = (1 + dkll5 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第6次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(21, 22).Value = zbjbl6
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(10, 42).Value = (1 + dkll6 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第7次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(22, 22).Value = zbjbl7
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(11, 42).Value = (1 + dkll7 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第8次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(23, 22).Value = zbjbl8
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(12, 42).Value = (1 + dkll8 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第9次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(24, 22).Value = zbjbl9
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(13, 42).Value = (1 + dkll9 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第10次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(25, 22).Value = zbjbl10
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(14, 42).Value = (1 + dkll10 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '显示
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<燃机设备>建设期资金运用方式计算参数设置写入完成！"
        End If
        '蓄电池
        If Me.xudianchi.Checked = True Then
            '第1次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(27, 22).Value = zbjbl1
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(5, 45).Value = (1 + dkll1 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第2次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(28, 22).Value = zbjbl2
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(6, 45).Value = (1 + dkll2 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第3次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(29, 22).Value = zbjbl3
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(7, 45).Value = (1 + dkll3 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第4次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(30, 22).Value = zbjbl4
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(8, 45).Value = (1 + dkll4 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第5次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(31, 22).Value = zbjbl5
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(9, 45).Value = (1 + dkll5 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第6次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(32, 22).Value = zbjbl6
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(10, 45).Value = (1 + dkll6 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第7次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(33, 22).Value = zbjbl7
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(11, 45).Value = (1 + dkll7 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第8次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(34, 22).Value = zbjbl8
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(12, 45).Value = (1 + dkll8 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第9次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(35, 22).Value = zbjbl9
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(13, 45).Value = (1 + dkll9 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第10次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(36, 22).Value = zbjbl10
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(14, 45).Value = (1 + dkll10 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '显示
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<蓄电池设备>建设期资金运用方式计算参数设置写入完成！"
        End If
        '暖通
        If Me.nuantong.Checked = True Then
            '第1次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(5, 25).Value = zbjbl1
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(5, 48).Value = (1 + dkll1 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第2次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(6, 25).Value = zbjbl2
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(6, 48).Value = (1 + dkll2 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第3次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(7, 25).Value = zbjbl3
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(7, 48).Value = (1 + dkll3 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第4次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(8, 25).Value = zbjbl4
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(8, 48).Value = (1 + dkll4 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第5次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(9, 25).Value = zbjbl5
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(9, 48).Value = (1 + dkll5 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第6次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(10, 25).Value = zbjbl6
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(10, 48).Value = (1 + dkll6 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第7次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(11, 25).Value = zbjbl7
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(11, 48).Value = (1 + dkll7 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第8次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(12, 25).Value = zbjbl8
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(12, 48).Value = (1 + dkll8 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第9次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(13, 25).Value = zbjbl9
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(13, 48).Value = (1 + dkll9 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第10次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(14, 25).Value = zbjbl10
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(14, 48).Value = (1 + dkll10 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '显示
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<暖通设备>建设期资金运用方式计算参数设置写入完成！"
        End If
        '光伏
        If Me.guangfu.Checked = True Then
            '第1次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(16, 25).Value = zbjbl1
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(5, 51).Value = (1 + dkll1 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第2次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(17, 25).Value = zbjbl2
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(6, 51).Value = (1 + dkll2 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第3次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(18, 25).Value = zbjbl3
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(7, 51).Value = (1 + dkll3 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第4次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(19, 25).Value = zbjbl4
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(8, 51).Value = (1 + dkll4 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第5次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(20, 25).Value = zbjbl5
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(9, 51).Value = (1 + dkll5 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第6次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(21, 25).Value = zbjbl6
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(10, 51).Value = (1 + dkll6 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第7次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(22, 25).Value = zbjbl7
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(11, 51).Value = (1 + dkll7 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第8次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(23, 25).Value = zbjbl8
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(12, 51).Value = (1 + dkll8 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第9次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(24, 25).Value = zbjbl9
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(13, 51).Value = (1 + dkll9 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第10次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(25, 25).Value = zbjbl10
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(14, 51).Value = (1 + dkll10 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '显示
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<光伏设备>建设期资金运用方式计算参数设置写入完成！"
        End If
        '风电
        If Me.fengdian.Checked = True Then
            '第1次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(27, 25).Value = zbjbl1
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(5, 54).Value = (1 + dkll1 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第2次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(28, 25).Value = zbjbl2
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(6, 54).Value = (1 + dkll2 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第3次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(29, 25).Value = zbjbl3
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(7, 54).Value = (1 + dkll3 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第4次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(30, 25).Value = zbjbl4
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(8, 54).Value = (1 + dkll4 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第5次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(31, 25).Value = zbjbl5
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(9, 54).Value = (1 + dkll5 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第6次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(32, 25).Value = zbjbl6
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(10, 54).Value = (1 + dkll6 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第7次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(33, 25).Value = zbjbl7
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(11, 54).Value = (1 + dkll7 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第8次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(34, 25).Value = zbjbl8
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(12, 54).Value = (1 + dkll8 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第9次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(35, 25).Value = zbjbl9
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(13, 54).Value = (1 + dkll9 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '第10次投资
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(36, 25).Value = zbjbl10
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(14, 54).Value = (1 + dkll10 / jsqdkjxcs) ^ jsqdkjxcs - 1
            '显示
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<风电设备>建设期资金运用方式计算参数设置写入完成！"
        End If
        '————————————————————————————————————————————————————————————————————————————————————————  
        '等于0的参数改为默认值
        '资本金比例
        For i = 1 To 10
            '常规设备
            If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 4, 22).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 4, 22).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 5).Value
            End If
            '其它设备
            If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 15, 22).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 15, 22).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 5).Value
            End If
            If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 26, 22).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 26, 22).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 5).Value
            End If
            If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 4, 25).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 4, 25).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 5).Value
            End If
            If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 15, 25).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 15, 25).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 5).Value
            End If
            If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 26, 25).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 26, 25).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 5).Value
            End If
        Next
        '长期贷款年利率
        '计算复利贷款利率
        Dim dkll = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 9).Value
        Dim dkll_fl As Double = (1 + dkll / jsqdkjxcs) ^ jsqdkjxcs - 1
        '常规设备
        For i = 12 To 21
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 39).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 39).Value = dkll_fl
            End If
        Next
        '其它设备
        For i = 5 To 14
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 42).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 42).Value = dkll_fl
            End If
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 45).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 45).Value = dkll_fl
            End If
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 48).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 48).Value = dkll_fl
            End If
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 51).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 51).Value = dkll_fl
            End If
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 54).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 54).Value = dkll_fl
            End If
        Next
    End Sub
End Class