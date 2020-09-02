Public Class 建设期资金运用方式设置
    Sub 建设期默认资金运用模式(ZBJBL As Double, DKLL As Double)
        '每次投资的资本金比例和建设期贷款利率均相同，资本金比例为占动态投资比例
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '建设期贷款计息次数
        Dim jsqdkjxcs = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 9).Value
        '————————————————————————————————————————————————————————————————————————————————————————
        '解锁表格
        ExcelApp.ThisWorkbook.Worksheets("估算表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").unProtect(Password:="wscjc")
        '——————————————————————————————————————————————————————————————————————————————————————
        '将计算方式设置为模式1
        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(170, 9).Value = 1
        '将已有的数据清零
        For i = 3 To 21 Step 2
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(176, i).Value = 0
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(177, i).Value = 0
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, i).Value = 0
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, i).Value = 0
        Next
        For i = 3 To 21 Step 2
            '方法一的结果
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(176, i).Value = ZBJBL
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(177, i).Value = (1 + DKLL / jsqdkjxcs) ^ jsqdkjxcs - 1
            '方法二的结果
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, i).Value = 0
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, i).Value = 0
        Next
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
            '解锁表格
            ExcelApp.ThisWorkbook.Worksheets("估算表").unProtect(Password:="wscjc")
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").unProtect(Password:="wscjc")
            '————————————————————————————————————————————————————————————————————————————————————————
            '将已有的数据清零
            For i = 3 To 21 Step 2
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(176, i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(177, i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, i).Value = 0
            Next
            '部分区域重新计算
            ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
            '———————————————————————————————————————————————————————————————————————————————————————— 
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
            '将计算方式设置为模式1
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(170, 9).Value = 1
            '部分区域重新计算
            ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
            '将动态投资比例写入
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(176, 3).Value = ZBJBL
            Dim ZBJBL_Tem As Double '资本金比例计算的过程量
            Dim ZBJBL_JT As Double '静态资本金比例系数
            '写入建设期贷款利率
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(177, 3).Value = (1 + DKLL / jsqdkjxcs) ^ jsqdkjxcs - 1
            '部分区域重新计算
            ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
            For i = 1 To 200 '每次变化千分之2
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(176, 3).Value = ZBJBL * (200 - i) / 200
                '部分区域重新计算
                ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 3).Value <= ZBJBL Then
                    ZBJBL_Tem = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(176, 3).Value
                    Exit For
                End If
            Next
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 3).Value < ZBJBL Then
                ZBJBL_Tem = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(176, 3).Value
                '微调资本金比例
                For j = 1 To 10000
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(176, 3).Value = ZBJBL_Tem + j * (0.01 / 100)
                    '部分区域重新计算
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 3).Value >= ZBJBL Then
                        ZBJBL_JT = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(176, 3).Value  '记录下此时的动态投资比例系数
                        Exit For
                    End If
                Next
            End If
            '将计算结果写入表格
            For i = 3 To 21 Step 2
                '方法一的结果
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(176, i).Value = ZBJBL_JT
                '方法二的结果
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, i).Value = 0
            Next
            '————————————————————————————————————————————————————————————————————————————————————————  
            '在表格中写入当前采用的资金运用模式
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 7).Value = "静态"
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 7).Value = "相同"
            MsgBox("已经将资本金比例计算模式切换为占静态总投资的比例，同时每次投资的资本金比例和建设期贷款利率均相同！")
            Me.Close()
        End If
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
            '解锁表格
            ExcelApp.ThisWorkbook.Worksheets("估算表").unProtect(Password:="wscjc")
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").unProtect(Password:="wscjc")
            '将已有的数据清零
            For i = 3 To 21 Step 2
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(176, i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(177, i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, i).Value = 0
            Next
            '将计算方式设置为模式2
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(170, 9).Value = 2
            '————————————————————————————————————————————————————————————————————————————————————————  
            '检测有投资的年份，是否都输入的资本金(动态)比例和建设期贷款利率，如果不正确报错，正确则写入数据（贷款利息全部采用复利计算）
            '第1次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 3).Value > 0 And (zbjbl1 <= 0 Or dkll1 <= 0) Then
                MsgBox("第1次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 3).Value = zbjbl1
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, 3).Value = (1 + dkll1 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第2次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 5).Value > 0 And (zbjbl2 <= 0 Or dkll2 <= 0) Then
                MsgBox("第2次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 5).Value = zbjbl2
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, 5).Value = (1 + dkll2 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第3次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 7).Value > 0 And (zbjbl3 <= 0 Or dkll3 <= 0) Then
                MsgBox("第3次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 7).Value = zbjbl3
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, 7).Value = (1 + dkll3 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第4次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 9).Value > 0 And (zbjbl4 <= 0 Or dkll4 <= 0) Then
                MsgBox("第4次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 9).Value = zbjbl4
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, 9).Value = (1 + dkll4 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第5次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 11).Value > 0 And (zbjbl5 <= 0 Or dkll5 <= 0) Then
                MsgBox("第5次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 11).Value = zbjbl5
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, 11).Value = (1 + dkll5 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第6次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 3).Value > 0 And (zbjbl6 <= 0 Or dkll6 <= 0) Then
                MsgBox("第6次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 13).Value = zbjbl6
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, 13).Value = (1 + dkll6 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第7次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 5).Value > 0 And (zbjbl7 <= 0 Or dkll7 <= 0) Then
                MsgBox("第7次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 15).Value = zbjbl7
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, 15).Value = (1 + dkll7 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第8次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 7).Value > 0 And (zbjbl8 <= 0 Or dkll8 <= 0) Then
                MsgBox("第8次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 17).Value = zbjbl8
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, 17).Value = (1 + dkll8 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第9次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 9).Value > 0 And (zbjbl9 <= 0 Or dkll9 <= 0) Then
                MsgBox("第9次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 19).Value = zbjbl9
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, 19).Value = (1 + dkll9 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '第10次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 11).Value > 0 And (zbjbl10 <= 0 Or dkll10 <= 0) Then
                MsgBox("第10次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 21).Value = zbjbl10
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, 21).Value = (1 + dkll10 / jsqdkjxcs) ^ jsqdkjxcs - 1
            End If
            '————————————————————————————————————————————————————————————————————————————————————————  
            '将设置次数计数器+1
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(170, 11).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(170, 11).Value + 1
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
            '解锁表格
            ExcelApp.ThisWorkbook.Worksheets("估算表").unProtect(Password:="wscjc")
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").unProtect(Password:="wscjc")
            '将计算方式设置为模式2
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(170, 9).Value = 2
            '将已有的数据清零
            For i = 3 To 21 Step 2
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(176, i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(177, i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, i).Value = 0
            Next
            '部分区域重新计算
            ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
            '———————————————————————————————————————————————————————————————————————————————————————— 
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
                Dim zbjbl1_Tem As Double '资本金比例计算的过程量
                Dim zbjbl1_jt As Double '计算出的静态投资
                '写入建设期贷款利率
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, 3).Value = (1 + dkll1 / jsqdkjxcs) ^ jsqdkjxcs - 1
                '将动态投资比例写入
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 3).Value = zbjbl1
                '部分区域重新计算
                ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                For i = 1 To 200 '每次变化千分之5
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 3).Value = zbjbl1 * (200 - i) / 200
                    '部分区域重新计算
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 3).Value <= zbjbl1 Then
                        zbjbl1_Tem = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 3).Value
                        Exit For
                    End If
                Next
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 3).Value < zbjbl1 Then
                    zbjbl1_Tem = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 3).Value
                    '微调资本金比例
                    For j = 1 To 10000
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 3).Value = zbjbl1_Tem + j * (0.01 / 100)
                        '部分区域重新计算
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 3).Value >= zbjbl1 Then
                            zbjbl1_jt = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 3).Value  '记录下此时的动态投资比例系数
                            Exit For
                        End If
                    Next
                End If
                '将计算结果写入表格
                '方法1的结果
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(176, 3).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(177, 3).Value = 0
                '方法2的结果
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 3).Value = zbjbl1_jt
            End If
            '第2次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 5).Value > 0 And (zbjbl2 <= 0 Or dkll2 <= 0) Then
                MsgBox("第2次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                Dim zbjbl2_Tem As Double '资本金比例计算的过程量
                Dim zbjbl2_jt As Double '计算出的静态投资
                '写入建设期贷款利率
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, 5).Value = (1 + dkll2 / jsqdkjxcs) ^ jsqdkjxcs - 1
                '将动态投资比例写入
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 5).Value = zbjbl2
                '部分区域重新计算
                ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                For i = 1 To 200 '每次变化千分之5
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 5).Value = zbjbl2 * (200 - i) / 200
                    '部分区域重新计算
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 5).Value <= zbjbl2 Then
                        zbjbl2_Tem = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 3).Value
                        Exit For
                    End If
                Next
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 5).Value < zbjbl2 Then
                    zbjbl2_Tem = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 3).Value
                    '微调资本金比例
                    For j = 1 To 10000
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 5).Value = zbjbl2_Tem + j * (0.01 / 100)
                        '部分区域重新计算
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 5).Value >= zbjbl2 Then
                            zbjbl2_jt = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 5).Value  '记录下此时的动态投资比例系数
                            Exit For
                        End If
                    Next
                End If
                '将计算结果写入表格
                '方法1的结果
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(176, 5).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(177, 5).Value = 0
                '方法2的结果
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 5).Value = zbjbl2_jt
            End If
            '第3次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 7).Value > 0 And (zbjbl3 <= 0 Or dkll3 <= 0) Then
                MsgBox("第3次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                Dim zbjbl3_Tem As Double '资本金比例计算的过程量
                Dim zbjbl3_jt As Double '计算出的静态投资
                '写入建设期贷款利率
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, 7).Value = (1 + dkll3 / jsqdkjxcs) ^ jsqdkjxcs - 1
                '将动态投资比例写入
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 7).Value = zbjbl3
                '部分区域重新计算
                ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                For i = 1 To 200 '每次变化千分之5
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 7).Value = zbjbl3 * (200 - i) / 200
                    '部分区域重新计算
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 7).Value <= zbjbl3 Then
                        zbjbl3_Tem = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 7).Value
                        Exit For
                    End If
                Next
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 7).Value < zbjbl3 Then
                    zbjbl3_Tem = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 7).Value
                    '微调资本金比例
                    For j = 1 To 10000
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 7).Value = zbjbl3_Tem + j * (0.01 / 100)
                        '部分区域重新计算
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 7).Value >= zbjbl3 Then
                            zbjbl3_jt = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 7).Value  '记录下此时的动态投资比例系数
                            Exit For
                        End If
                    Next
                End If
                '将计算结果写入表格
                '方法1的结果
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(176, 7).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(177, 7).Value = 0
                '方法2的结果
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 7).Value = zbjbl3_jt
            End If
            '第4次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 9).Value > 0 And (zbjbl4 <= 0 Or dkll4 <= 0) Then
                MsgBox("第4次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                Dim zbjbl4_Tem As Double '资本金比例计算的过程量
                Dim zbjbl4_jt As Double '计算出的静态投资
                '写入建设期贷款利率
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, 9).Value = (1 + dkll4 / jsqdkjxcs) ^ jsqdkjxcs - 1
                '将动态投资比例写入
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 9).Value = zbjbl4
                '部分区域重新计算
                ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                For i = 1 To 200 '每次变化千分之5
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 9).Value = zbjbl4 * (200 - i) / 200
                    '部分区域重新计算
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 9).Value <= zbjbl4 Then
                        zbjbl4_Tem = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 9).Value
                        Exit For
                    End If
                Next
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 9).Value < zbjbl4 Then
                    zbjbl4_Tem = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 9).Value
                    '微调资本金比例
                    For j = 1 To 10000
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 9).Value = zbjbl4_Tem + j * (0.01 / 100)
                        '部分区域重新计算
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 9).Value >= zbjbl4 Then
                            zbjbl4_jt = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 9).Value  '记录下此时的动态投资比例系数
                            Exit For
                        End If
                    Next
                End If
                '将计算结果写入表格
                '方法1的结果
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(176, 9).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(177, 9).Value = 0
                '方法2的结果
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 9).Value = zbjbl4_jt
            End If
            '第5次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 11).Value > 0 And (zbjbl5 <= 0 Or dkll5 <= 0) Then
                MsgBox("第5次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                Dim zbjbl5_Tem As Double '资本金比例计算的过程量
                Dim zbjbl5_jt As Double '计算出的静态投资
                '写入建设期贷款利率
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, 11).Value = (1 + dkll5 / jsqdkjxcs) ^ jsqdkjxcs - 1
                '将动态投资比例写入
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 11).Value = zbjbl5
                '部分区域重新计算
                ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                For i = 1 To 200 '每次变化千分之5
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 11).Value = zbjbl5 * (200 - i) / 200
                    '部分区域重新计算
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 11).Value <= zbjbl5 Then
                        zbjbl5_Tem = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 11).Value
                        Exit For
                    End If
                Next
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 11).Value < zbjbl5 Then
                    zbjbl5_Tem = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 11).Value
                    '微调资本金比例
                    For j = 1 To 10000
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 11).Value = zbjbl5_Tem + j * (0.01 / 100)
                        '部分区域重新计算
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 11).Value >= zbjbl5 Then
                            zbjbl5_jt = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 11).Value  '记录下此时的动态投资比例系数
                            Exit For
                        End If
                    Next
                End If
                '将计算结果写入表格
                '方法1的结果
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(176, 11).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(177, 11).Value = 0
                '方法2的结果
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 11).Value = zbjbl5_jt
            End If
            '第6次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 3).Value > 0 And (zbjbl6 <= 0 Or dkll6 <= 0) Then
                MsgBox("第6次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                Dim zbjbl6_Tem As Double '资本金比例计算的过程量
                Dim zbjbl6_jt As Double '计算出的静态投资
                '写入建设期贷款利率
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, 13).Value = (1 + dkll6 / jsqdkjxcs) ^ jsqdkjxcs - 1
                '将动态投资比例写入
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 13).Value = zbjbl6
                '部分区域重新计算
                ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                For i = 1 To 200 '每次变化千分之5
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 13).Value = zbjbl6 * (200 - i) / 200
                    '部分区域重新计算
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 13).Value <= zbjbl6 Then
                        zbjbl6_Tem = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 13).Value
                        Exit For
                    End If
                Next
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 13).Value < zbjbl6 Then
                    zbjbl6_Tem = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 13).Value
                    '微调资本金比例
                    For j = 1 To 10000
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 13).Value = zbjbl6_Tem + j * (0.01 / 100)
                        '部分区域重新计算
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 13).Value >= zbjbl6 Then
                            zbjbl6_jt = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 13).Value  '记录下此时的动态投资比例系数
                            Exit For
                        End If
                    Next
                End If
                '将计算结果写入表格
                '方法1的结果
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(176, 13).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(177, 13).Value = 0
                '方法2的结果
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 13).Value = zbjbl6_jt
            End If
            '第7次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 5).Value > 0 And (zbjbl7 <= 0 Or dkll7 <= 0) Then
                MsgBox("第7次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                Dim zbjbl7_Tem As Double '资本金比例计算的过程量
                Dim zbjbl7_jt As Double '计算出的静态投资
                '写入建设期贷款利率
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, 15).Value = (1 + dkll7 / jsqdkjxcs) ^ jsqdkjxcs - 1
                '将动态投资比例写入
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 15).Value = zbjbl7
                '部分区域重新计算
                ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                For i = 1 To 200 '每次变化千分之5
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 15).Value = zbjbl7 * (200 - i) / 200
                    '部分区域重新计算
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 15).Value <= zbjbl7 Then
                        zbjbl7_Tem = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 15).Value
                        Exit For
                    End If
                Next
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 15).Value < zbjbl7 Then
                    zbjbl7_Tem = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 15).Value
                    '微调资本金比例
                    For j = 1 To 10000
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 15).Value = zbjbl7_Tem + j * (0.01 / 100)
                        '部分区域重新计算
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 15).Value >= zbjbl7 Then
                            zbjbl7_jt = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 15).Value  '记录下此时的动态投资比例系数
                            Exit For
                        End If
                    Next
                End If
                '将计算结果写入表格
                '方法1的结果
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(176, 15).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(177, 15).Value = 0
                '方法2的结果
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 15).Value = zbjbl7_jt
            End If
            '第8次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 7).Value > 0 And (zbjbl8 <= 0 Or dkll8 <= 0) Then
                MsgBox("第8次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                Dim zbjbl8_Tem As Double '资本金比例计算的过程量
                Dim zbjbl8_jt As Double '计算出的静态投资
                '写入建设期贷款利率
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, 17).Value = (1 + dkll8 / jsqdkjxcs) ^ jsqdkjxcs - 1
                '将动态投资比例写入
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 17).Value = zbjbl8
                '部分区域重新计算
                ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                For i = 1 To 200 '每次变化千分之5
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 17).Value = zbjbl8 * (200 - i) / 200
                    '部分区域重新计算
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 17).Value <= zbjbl8 Then
                        zbjbl8_Tem = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 17).Value
                        Exit For
                    End If
                Next
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 17).Value < zbjbl8 Then
                    zbjbl8_Tem = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 17).Value
                    '微调资本金比例
                    For j = 1 To 10000
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 17).Value = zbjbl8_Tem + j * (0.01 / 100)
                        '部分区域重新计算
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 17).Value >= zbjbl8 Then
                            zbjbl8_jt = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 17).Value  '记录下此时的动态投资比例系数
                            Exit For
                        End If
                    Next
                End If
                '将计算结果写入表格
                '方法1的结果
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(176, 17).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(177, 17).Value = 0
                '方法2的结果
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 17).Value = zbjbl8_jt
            End If
            '第9次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 9).Value > 0 And (zbjbl9 <= 0 Or dkll9 <= 0) Then
                MsgBox("第9次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                Dim zbjbl9_Tem As Double '资本金比例计算的过程量
                Dim zbjbl9_jt As Double '计算出的静态投资
                '写入建设期贷款利率
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, 19).Value = (1 + dkll9 / jsqdkjxcs) ^ jsqdkjxcs - 1
                '将动态投资比例写入
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 19).Value = zbjbl9
                '部分区域重新计算
                ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                For i = 1 To 200 '每次变化千分之5
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 19).Value = zbjbl9 * (200 - i) / 200
                    '部分区域重新计算
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 19).Value <= zbjbl9 Then
                        zbjbl9_Tem = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 19).Valuet
                        Exit For
                    End If
                Next
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 19).Value < zbjbl9 Then
                    zbjbl9_Tem = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 19).Value
                    '微调资本金比例
                    For j = 1 To 10000
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 19).Value = zbjbl9_Tem + j * (0.01 / 100)
                        '部分区域重新计算
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 19).Value >= zbjbl9 Then
                            zbjbl9_jt = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 19).Value  '记录下此时的动态投资比例系数
                            Exit For
                        End If
                    Next
                End If
                '将计算结果写入表格
                '方法1的结果
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(176, 19).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(177, 19).Value = 0
                '方法2的结果
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 19).Value = zbjbl9_jt
            End If
            '第10次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 11).Value > 0 And (zbjbl10 <= 0 Or dkll10 <= 0) Then
                MsgBox("第10次投资的资本金比例或者建设期贷款利率不可以小于等于0，请重新输入！")
                Exit Sub
            Else
                Dim zbjbl10_Tem As Double '资本金比例计算的过程量
                Dim zbjbl10_jt As Double '计算出的静态投资
                '写入建设期贷款利率
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(180, 21).Value = (1 + dkll10 / jsqdkjxcs) ^ jsqdkjxcs - 1
                '将动态投资比例写入
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 21).Value = zbjbl10
                '部分区域重新计算
                ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                For i = 1 To 200 '每次变化千分之5
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 21).Value = zbjbl10 * (200 - i) / 200
                    '部分区域重新计算
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 21).Value <= zbjbl10 Then
                        zbjbl10_Tem = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 21).Value
                        Exit For
                    End If
                Next
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 21).Value < zbjbl10 Then
                    zbjbl10_Tem = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 21).Value
                    '微调资本金比例
                    For j = 1 To 10000
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 21).Value = zbjbl10_Tem + j * (0.01 / 100)
                        '部分区域重新计算
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Calculate
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(183, 21).Value >= zbjbl10 Then
                            zbjbl10_jt = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 21).Value  '记录下此时的动态投资比例系数
                            Exit For
                        End If
                    Next
                End If
                '将计算结果写入表格
                '方法1的结果
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(176, 21).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(177, 21).Value = 0
                '方法2的结果
                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(179, 21).Value = zbjbl10_jt
            End If
            '————————————————————————————————————————————————————————————————————————————————————————  
            '将设置次数计数器+1
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(170, 11).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(170, 11).Value + 1
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
        '根据已经输入的投资情况，载入默认值
        '读取逐次设置的次数
        Dim SZCS As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(170, 11).Value
        If SZCS >= 1 Then
            '读取贷款利息计息次数
            Dim jsqdkjxcs = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 9).Value
            '读取上次修改后的数据
            '第1次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 3).Value > 0 Then
                Me.zbjbl1.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(173, 3).Value * 100, String)
                Me.dkll1.Text = CType(Math.Round(((ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(174, 3).Value + 1) ^ (1 / jsqdkjxcs) - 1) * jsqdkjxcs * 100, 2), String)
            End If
            '第2次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value > 0 Then
                Me.zbjbl2.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(173, 5).Value * 100, String)
                Me.dkll2.Text = CType(Math.Round(((ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(174, 5).Value + 1) ^ (1 / jsqdkjxcs) - 1) * jsqdkjxcs * 100, 2), String)
            End If
            '第3次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value > 0 Then
                Me.zbjbl3.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(173, 7).Value * 100, String)
                Me.dkll3.Text = CType(Math.Round(((ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(174, 7).Value + 1) ^ (1 / jsqdkjxcs) - 1) * jsqdkjxcs * 100, 2), String)
            End If
            '第4次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value > 0 Then
                Me.zbjbl4.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(173, 9).Value * 100, String)
                Me.dkll4.Text = CType(Math.Round(((ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(174, 9).Value + 1) ^ (1 / jsqdkjxcs) - 1) * jsqdkjxcs * 100, 2), String)
            End If
            '第5次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value > 0 Then
                Me.zbjbl5.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(173, 11).Value * 100, String)
                Me.dkll5.Text = CType(Math.Round(((ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(174, 11).Value + 1) ^ (1 / jsqdkjxcs) - 1) * jsqdkjxcs * 100, 2), String)
            End If
            '第6次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value > 0 Then
                Me.zbjbl6.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(173, 13).Value * 100, String)
                Me.dkll6.Text = CType(Math.Round(((ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(174, 13).Value + 1) ^ (1 / jsqdkjxcs) - 1) * jsqdkjxcs * 100, 2), String)
            End If
            '第7次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value > 0 Then
                Me.zbjbl7.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(173, 15).Value * 100, String)
                Me.dkll7.Text = CType(Math.Round(((ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(174, 15).Value + 1) ^ (1 / jsqdkjxcs) - 1) * jsqdkjxcs * 100, 2), String)
            End If
            '第8次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value > 0 Then
                Me.zbjbl8.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(173, 17).Value * 100, String)
                Me.dkll8.Text = CType(Math.Round(((ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(174, 17).Value + 1) ^ (1 / jsqdkjxcs) - 1) * jsqdkjxcs * 100, 2), String)
            End If
            '第9次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value > 0 Then
                Me.zbjbl9.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(173, 19).Value * 100, String)
                Me.dkll9.Text = CType(Math.Round(((ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(174, 19).Value + 1) ^ (1 / jsqdkjxcs) - 1) * jsqdkjxcs * 100, 2), String)
            End If
            '第10次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value > 0 Then
                Me.zbjbl10.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(173, 21).Value * 100, String)
                Me.dkll10.Text = CType(Math.Round(((ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(174, 21).Value + 1) ^ (1 / jsqdkjxcs) - 1) * jsqdkjxcs * 100, 2), String)
            End If
        Else
            '载入默认值
            '每次的资本金(动态)比例和建设期贷款利率均一致的情况
            Me.zbjbl_a.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 5).Value * 100, String)
            Me.dkll_a.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 9).Value * 100, String)
            '第1次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 3).Value > 0 Then
                Me.zbjbl1.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 5).Value * 100, String)
                Me.dkll1.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 9).Value * 100, String)
            End If
            '第2次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value > 0 Then
                Me.zbjbl2.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 5).Value * 100, String)
                Me.dkll2.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 9).Value * 100, String)
            End If
            '第3次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value > 0 Then
                Me.zbjbl3.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 5).Value * 100, String)
                Me.dkll3.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 9).Value * 100, String)
            End If
            '第4次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value > 0 Then
                Me.zbjbl4.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 5).Value * 100, String)
                Me.dkll4.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 9).Value * 100, String)
            End If
            '第5次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value > 0 Then
                Me.zbjbl5.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 5).Value * 100, String)
                Me.dkll5.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 9).Value * 100, String)
            End If
            '第6次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value > 0 Then
                Me.zbjbl6.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 5).Value * 100, String)
                Me.dkll6.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 9).Value * 100, String)
            End If
            '第7次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value > 0 Then
                Me.zbjbl7.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 5).Value * 100, String)
                Me.dkll7.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 9).Value * 100, String)
            End If
            '第8次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value > 0 Then
                Me.zbjbl8.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 5).Value * 100, String)
                Me.dkll8.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 9).Value * 100, String)
            End If
            '第9次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value > 0 Then
                Me.zbjbl9.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 5).Value * 100, String)
                Me.dkll9.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 9).Value * 100, String)
            End If
            '第10次投资
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value > 0 Then
                Me.zbjbl10.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 5).Value * 100, String)
                Me.dkll10.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 9).Value * 100, String)
            End If
        End If
    End Sub
End Class