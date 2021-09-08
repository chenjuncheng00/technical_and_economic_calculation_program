Public Class 建设期资金运用方式设置
    Sub 建设期默认资金运用模式(ZBJBL As Double, DKLL As Double)
        '每次投资的资本金比例和建设期贷款利率均相同，资本金比例为占动态投资比例
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '建设期贷款计息次数
        Dim jsqdkjxcs = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 9).Value
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算复利贷款利率
        Dim dkll_fl As Double = (1 + DKLL / jsqdkjxcs) ^ jsqdkjxcs - 1
        '将实际的建设期贷款利率写入Excel
        For i = 1 To 10
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i + 11, 39).Value = dkll_fl
        Next
        '将资本金比例写入Excel
        For i = 1 To 10
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 4, 22).Value = ZBJBL
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
        '建设期时间计划计算
        Call 建设期时间计划计算.建设期时间计划计算(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        Dim zbj_model As Integer = 0
        'hscy：计算期末，是否回收资产残值，默认回收
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
        '————————————————————————————————————————————————————————————————————————————————————————
        '确定估算表参数设置(包括了投资金额变化后计算、收入成本变化后计算)
        Call 确定估算表参数设置(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model)
        '——————————————————————————————————————————————————————————————————————————————————————
        '在表格中写入当前采用的资金运用模式
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 7).Value = "动态"
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 7).Value = "相同"
    End Sub

    Private Sub 资本金比例为占动态投资比例_统一设置_Click(sender As Object, e As EventArgs) Handles 资本金比例为占动态投资比例_统一设置.Click
        Dim XZ = MsgBox("是否需要将资本金比例计算模式切换为占动态总投资的比例，同时每次投资的资本金比例和建设期贷款利率均相同？", vbOKCancel)
        If XZ = vbOK Then
            Dim ZBJBL = CType(Me.zbjbl_a.Text, Double) / 100
            Dim DKLL = CType(Me.dkll_a.Text, Double) / 100
            Call 建设期默认资金运用模式(ZBJBL, DKLL)
            MsgBox("已经将资本金比例计算模式切换为占动态总投资的比例，同时每次投资的资本金比例和建设期贷款利率均相同！")
            Me.Close()
        End If
    End Sub

    Private Sub 资本金比例为占静态投资比例_统一设置_Click(sender As Object, e As EventArgs) Handles 资本金比例为占静态投资比例_统一设置.Click
        On Error Resume Next
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
            '如果没有任何一次投资计划年份或者投资金额，则不计算，防止死循环
            Dim TZNF_zero As Integer = 0
            Dim TZJE_zero As Integer = 0
            For i = 3 To 11 Step 2
                For j = 28 To 71 Step 43
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(j, i).Value = 0 Then
                        TZNF_zero = TZNF_zero + 1
                    End If
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(j + 2, i).Value = 0 Then
                        TZJE_zero = TZJE_zero + 1
                    End If
                Next
            Next
            If TZNF_zero = 10 Or TZJE_zero = 10 Then
                MsgBox("没输入任何投资计划年份或者投资金额，请检查！")
                Exit Sub
            End If
            '如果某一次投资没有输入静态投资金额或者投资计划年份，不允许计算，防止死循环
            For i = 3 To 11 Step 2
                For j = 28 To 71 Step 43
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(j, i).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(j + 2, i).Value > 0 Then
                        '不执行任何操作
                    ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(j, i).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(j + 2, i).Value = 0 Then
                        '不执行任何操作
                    Else
                        MsgBox("有投资计划的年份没有投资金额，或者有投资金额的年份没有投资计划，请检查并重新输入！")
                        Exit Sub
                    End If

                Next
            Next
            '———————————————————————————————————————————————————————————————————————————————————————— 
            '计算复利贷款利率
            Dim dkll_fl As Double = (1 + DKLL / jsqdkjxcs) ^ jsqdkjxcs - 1
            '将实际的建设期贷款利率写入Excel
            For i = 1 To 10
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i + 11, 39).Value = dkll_fl
            Next
            '将资本金比例写入Excel
            For i = 1 To 10
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(i + 4, 22).Value = ZBJBL
            Next
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算一次工作簿
            ExcelApp.Calculate()
            '建设期时间计划计算
            Call 建设期时间计划计算.建设期时间计划计算(ExcelApp)
            '————————————————————————————————————————————————————————————————————————————————————————
            'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
            Dim zbj_model As Integer = 1
            'hscy：计算期末，是否回收资产残值，默认回收
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
            '————————————————————————————————————————————————————————————————————————————————————————
            '确定估算表参数设置(包括了投资金额变化后计算、收入成本变化后计算)
            Call 确定估算表参数设置(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model)
            '————————————————————————————————————————————————————————————————————————————————————————  
            '在表格中写入当前采用的资金运用模式
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 7).Value = "静态"
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 7).Value = "相同"
            MsgBox("已经将资本金比例计算模式切换为占静态总投资的比例，同时每次投资的资本金比例和建设期贷款利率均相同！")
        End If
        Me.Close()
    End Sub

    Private Sub 资本金比例为占动态投资比例_逐次计算_Click(sender As Object, e As EventArgs) Handles 资本金比例为占动态投资比例_逐次计算.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————       
        '定义局部变量
        Dim zbjbl1, zbjbl2, zbjbl3, zbjbl4, zbjbl5, zbjbl6, zbjbl7, zbjbl8, zbjbl9, zbjbl10 As Double '资本金比例
        Dim dkll1, dkll2, dkll3, dkll4, dkll5, dkll6, dkll7, dkll8, dkll9, dkll10 As Double '建设期贷款利率
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————  
        '长期贷款计算方法需要采用方法三或者方法四
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法一" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法二" Then
            MsgBox("长期贷款计算方式，请选择计算方法三或者方法四，计算终止！")
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————  
        '读取输入的资本金(动态)比例
        zbjbl1 = CType(Me.zbjbl1.Text, Double) / 100
        zbjbl2 = CType(Me.zbjbl2.Text, Double) / 100
        zbjbl3 = CType(Me.zbjbl3.Text, Double) / 100
        zbjbl4 = CType(Me.zbjbl4.Text, Double) / 100
        zbjbl5 = CType(Me.zbjbl5.Text, Double) / 100
        zbjbl6 = CType(Me.zbjbl6.Text, Double) / 100
        zbjbl7 = CType(Me.zbjbl7.Text, Double) / 100
        zbjbl8 = CType(Me.zbjbl8.Text, Double) / 100
        zbjbl9 = CType(Me.zbjbl9.Text, Double) / 100
        zbjbl10 = CType(Me.zbjbl10.Text, Double) / 100
        '读取输入的建设期贷款利率
        dkll1 = CType(Me.dkll1.Text, Double) / 100
        dkll2 = CType(Me.dkll2.Text, Double) / 100
        dkll3 = CType(Me.dkll3.Text, Double) / 100
        dkll4 = CType(Me.dkll4.Text, Double) / 100
        dkll5 = CType(Me.dkll5.Text, Double) / 100
        dkll6 = CType(Me.dkll6.Text, Double) / 100
        dkll7 = CType(Me.dkll7.Text, Double) / 100
        dkll8 = CType(Me.dkll8.Text, Double) / 100
        dkll9 = CType(Me.dkll9.Text, Double) / 100
        dkll10 = CType(Me.dkll10.Text, Double) / 100
        '建设期贷款计息次数
        Dim jsqdkjxcs = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 9).Value
        '———————————————————————————————————————————————————————————————————————————————————————— 
        Dim XZ = MsgBox("是否需要将资本金比例计算模式切换为占动态总投资的比例，同时每次投资的资本金比例和建设期贷款利率均不相同？", vbOKCancel)
        If XZ = vbOK Then
            '检测有投资的年份，是否都输入的资本金(动态)比例和建设期贷款利率，如果不正确报错，正确则写入数据（贷款利息全部采用复利计算）
            '第1次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 3).Value > 0 And (zbjbl1 <= 0 Or dkll1 <= 0) Then
                MsgBox("第1次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(5, 22).Value = zbjbl1
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(12, 39).Value = (1 + dkll1 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第2次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 5).Value > 0 And (zbjbl2 <= 0 Or dkll2 <= 0) Then
                MsgBox("第2次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(6, 22).Value = zbjbl2
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(13, 39).Value = (1 + dkll2 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第3次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 7).Value > 0 And (zbjbl3 <= 0 Or dkll3 <= 0) Then
                MsgBox("第3次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(7, 22).Value = zbjbl3
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(14, 39).Value = (1 + dkll3 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第4次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 9).Value > 0 And (zbjbl4 <= 0 Or dkll4 <= 0) Then
                MsgBox("第4次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(8, 22).Value = zbjbl4
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(15, 39).Value = (1 + dkll4 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第5次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 11).Value > 0 And (zbjbl5 <= 0 Or dkll5 <= 0) Then
                MsgBox("第5次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(9, 22).Value = zbjbl5
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(16, 39).Value = (1 + dkll5 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第6次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 3).Value > 0 And (zbjbl6 <= 0 Or dkll6 <= 0) Then
                MsgBox("第6次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(10, 22).Value = zbjbl6
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(17, 39).Value = (1 + dkll6 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第7次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 5).Value > 0 And (zbjbl7 <= 0 Or dkll7 <= 0) Then
                MsgBox("第7次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(11, 22).Value = zbjbl7
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(18, 39).Value = (1 + dkll7 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第8次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 7).Value > 0 And (zbjbl8 <= 0 Or dkll8 <= 0) Then
                MsgBox("第8次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(12, 22).Value = zbjbl8
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(19, 39).Value = (1 + dkll8 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第9次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 9).Value > 0 And (zbjbl9 <= 0 Or dkll9 <= 0) Then
                MsgBox("第9次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(13, 22).Value = zbjbl9
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(20, 39).Value = (1 + dkll9 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第10次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 11).Value > 0 And (zbjbl10 <= 0 Or dkll10 <= 0) Then
                MsgBox("第10次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(14, 22).Value = zbjbl10
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(21, 39).Value = (1 + dkll10 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '————————————————————————————————————————————————————————————————————————————————————————  
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算一次工作簿
            ExcelApp.Calculate()
            '建设期时间计划计算
            Call 建设期时间计划计算.建设期时间计划计算(ExcelApp)
            '————————————————————————————————————————————————————————————————————————————————————————
            'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
            Dim zbj_model As Integer = 0
            'hscy：计算期末，是否回收资产残值，默认回收
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
            '————————————————————————————————————————————————————————————————————————————————————————
            '确定估算表参数设置(包括了投资金额变化后计算、收入成本变化后计算)
            Call 确定估算表参数设置(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model)
            '————————————————————————————————————————————————————————————————————————————————————————  
            '在表格中写入当前采用的资金运用模式
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 7).Value = "动态"
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 7).Value = "不相同"
            MsgBox("已经将资本金比例计算模式切换为占动态总投资的比例，同时每次投资的资本金比例和建设期贷款利率均不相同！")
            Me.Close()
        End If
    End Sub

    Private Sub 资本金比例为占静态投资比例_逐次计算_Click(sender As Object, e As EventArgs) Handles 资本金比例为占静态投资比例_逐次计算.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '定义局部变量
        Dim zbjbl1, zbjbl2, zbjbl3, zbjbl4, zbjbl5, zbjbl6, zbjbl7, zbjbl8, zbjbl9, zbjbl10 As Double '资本金比例
        Dim dkll1, dkll2, dkll3, dkll4, dkll5, dkll6, dkll7, dkll8, dkll9, dkll10 As Double '建设期贷款利率
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————  
        '长期贷款计算方法需要采用方法三或者方法四
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法一" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法二" Then
            MsgBox("长期借款计算方式，请选择计算方法三或者方法四，计算终止！")
            Exit Sub
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '读取输入的资本金(动态)比例
        zbjbl1 = CType(Me.zbjbl1.Text, Double) / 100
        zbjbl2 = CType(Me.zbjbl2.Text, Double) / 100
        zbjbl3 = CType(Me.zbjbl3.Text, Double) / 100
        zbjbl4 = CType(Me.zbjbl4.Text, Double) / 100
        zbjbl5 = CType(Me.zbjbl5.Text, Double) / 100
        zbjbl6 = CType(Me.zbjbl6.Text, Double) / 100
        zbjbl7 = CType(Me.zbjbl7.Text, Double) / 100
        zbjbl8 = CType(Me.zbjbl8.Text, Double) / 100
        zbjbl9 = CType(Me.zbjbl9.Text, Double) / 100
        zbjbl10 = CType(Me.zbjbl10.Text, Double) / 100
        '读取输入的建设期贷款利率
        dkll1 = CType(Me.dkll1.Text, Double) / 100
        dkll2 = CType(Me.dkll2.Text, Double) / 100
        dkll3 = CType(Me.dkll3.Text, Double) / 100
        dkll4 = CType(Me.dkll4.Text, Double) / 100
        dkll5 = CType(Me.dkll5.Text, Double) / 100
        dkll6 = CType(Me.dkll6.Text, Double) / 100
        dkll7 = CType(Me.dkll7.Text, Double) / 100
        dkll8 = CType(Me.dkll8.Text, Double) / 100
        dkll9 = CType(Me.dkll9.Text, Double) / 100
        dkll10 = CType(Me.dkll10.Text, Double) / 100
        '建设期贷款计息次数
        Dim jsqdkjxcs = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 9).Value
        '———————————————————————————————————————————————————————————————————————————————————————— 
        Dim XZ = MsgBox("是否需要将资本金比例计算模式切换为占静态总投资的比例，同时每次投资的资本金比例和建设期贷款利率均不相同？", vbOKCancel)
        If XZ = vbOK Then
            '如果没有任何一次投资计划年份或者投资金额，则不计算，防止死循环
            Dim TZNF_zero As Integer = 0
            Dim TZJE_zero As Integer = 0
            For i = 3 To 11 Step 2
                For j = 28 To 71 Step 43
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(j, i).Value = 0 Then
                        TZNF_zero = TZNF_zero + 1
                    End If
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(j + 2, i).Value = 0 Then
                        TZJE_zero = TZJE_zero + 1
                    End If
                Next
            Next
            If TZNF_zero = 10 Or TZJE_zero = 10 Then
                MsgBox("没输入任何投资计划年份或者投资金额，请检查！")
                Exit Sub
            End If
            '如果某一次投资没有输入静态投资金额或者投资计划年份，不允许计算，防止死循环
            For i = 3 To 11 Step 2
                For j = 28 To 71 Step 43
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(j, i).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(j + 2, i).Value > 0 Then
                        '不执行任何操作
                    ElseIf ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(j, i).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(j + 2, i).Value = 0 Then
                        '不执行任何操作
                    Else
                        MsgBox("有投资计划的年份没有投资金额，或者有投资金额的年份没有投资计划，请检查并重新输入！")
                        Exit Sub
                    End If

                Next
            Next
            '———————————————————————————————————————————————————————————————————————————————————————— 
            '第1次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 3).Value > 0 And (zbjbl1 <= 0 Or dkll1 <= 0) Then
                MsgBox("第1次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(5, 22).Value = zbjbl1
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(12, 39).Value = (1 + dkll1 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第2次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 5).Value > 0 And (zbjbl2 <= 0 Or dkll2 <= 0) Then
                MsgBox("第2次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(6, 22).Value = zbjbl2
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(13, 39).Value = (1 + dkll2 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第3次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 7).Value > 0 And (zbjbl3 <= 0 Or dkll3 <= 0) Then
                MsgBox("第3次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(7, 22).Value = zbjbl3
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(14, 39).Value = (1 + dkll3 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第4次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 9).Value > 0 And (zbjbl4 <= 0 Or dkll4 <= 0) Then
                MsgBox("第4次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(8, 22).Value = zbjbl4
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(15, 39).Value = (1 + dkll4 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第5次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 11).Value > 0 And (zbjbl5 <= 0 Or dkll5 <= 0) Then
                MsgBox("第5次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(9, 22).Value = zbjbl5
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(16, 39).Value = (1 + dkll5 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第6次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 3).Value > 0 And (zbjbl6 <= 0 Or dkll6 <= 0) Then
                MsgBox("第6次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(10, 22).Value = zbjbl6
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(17, 39).Value = (1 + dkll6 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第7次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 5).Value > 0 And (zbjbl7 <= 0 Or dkll7 <= 0) Then
                MsgBox("第7次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(11, 22).Value = zbjbl7
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(18, 39).Value = (1 + dkll7 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第8次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 7).Value > 0 And (zbjbl8 <= 0 Or dkll8 <= 0) Then
                MsgBox("第8次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(12, 22).Value = zbjbl8
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(19, 39).Value = (1 + dkll8 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第9次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 9).Value > 0 And (zbjbl9 <= 0 Or dkll9 <= 0) Then
                MsgBox("第9次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(13, 22).Value = zbjbl9
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(20, 39).Value = (1 + dkll9 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第10次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 11).Value > 0 And (zbjbl10 <= 0 Or dkll10 <= 0) Then
                MsgBox("第10次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(14, 22).Value = zbjbl10
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(21, 39).Value = (1 + dkll10 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '————————————————————————————————————————————————————————————————————————————————————————  
            '计算一次工作簿
            ExcelApp.Calculate()
            '建设期时间计划计算
            Call 建设期时间计划计算.建设期时间计划计算(ExcelApp)
            '————————————————————————————————————————————————————————————————————————————————————————
            'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
            Dim zbj_model As Integer = 1
            'hscy：计算期末，是否回收资产残值，默认回收
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
            '————————————————————————————————————————————————————————————————————————————————————————
            '确定估算表参数设置(包括了投资金额变化后计算、收入成本变化后计算)
            Call 确定估算表参数设置(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model)
            '————————————————————————————————————————————————————————————————————————————————————————  
            '在表格中写入当前采用的资金运用模式
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 7).Value = "静态"
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 7).Value = "不相同"
            MsgBox("已经将资本金比例计算模式切换为占静态总投资的比例，同时每次投资的资本金比例和建设期贷款利率均不相同！")
            Me.Close()
        End If
    End Sub

    Private Sub 重置回默认计算模式_Click(sender As Object, e As EventArgs) Handles 重置回默认计算模式.Click
        On Error Resume Next
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
            MsgBox("设置已完成！")
            Me.Close()
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
            MsgBox("清空窗体已完成！")
        End If
    End Sub

    Private Sub 每次投资设置不同的资本金比例和建设期贷款利率_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error Resume Next
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
    End Sub
End Class