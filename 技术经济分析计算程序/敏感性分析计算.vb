Imports Microsoft.Office.Interop
Module 敏感性分析计算
    Sub 静态投资敏感性分析计算(ExcelApp As Object, MGXFXBHL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '实例化计算进度显示窗体
        Dim Form1 As New 计算进度显示
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(7, 29).Value = 1 Then
            '计算进度提醒
            Form1.Show()
            Form1.Label1.Text = "正在读取静态投资敏感性分析基本数据。"
            Form1.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            '————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————  
            'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
            Dim zbj_model As Integer
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 7).Value = "动态" Then
                zbj_model = 0
            Else
                zbj_model = 1
            End If
            'hscy：计算期末，是否回收资产残值
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
            'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
            Dim sdsl_model As Integer = 1
            'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
            Dim kcje_xlf_model As Integer = 1
            '————————————————————————————————————————————————————————————————————————————————————————
            '将总静态投资写入敏感性分析表格中
            ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(7, 8).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value
            Dim TZHH(50) As Integer '数组，储存投资所在的行号
            Dim TZHHJS As Integer = 0 '投资行号计数
            Dim TZJECC(200) '数组，储存各种投资金额
            For i = 30 To 110 '检索估算表中投资编号，投资编号都是1表示
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, 27).Value = 1 Then
                    TZHHJS = TZHHJS + 1
                    TZHH(TZHHJS) = i '行号储存进数组，从下标1开始
                End If
            Next
            Dim TZJESLJS As Integer = 0 '投资金额数量计数
            '将各种投资金额储存在数组中
            For i = 1 To TZHHJS '行
                For j = 3 To 11 Step 2 '列
                    TZJESLJS = TZJESLJS + 1
                    TZJECC(TZJESLJS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(i), j).Value
                Next
            Next
            Dim JS As Integer = 0
            '计算静态总投资敏感性分析
            For i = 1 To 5 '敏感性分析循环5次
                '计算进度提醒
                Form1.Show()
                Form1.Label1.Text = "正在计算静态投资敏感性分析，计算进度：(" & i & "/ 5)"
                Form1.TopMost = True
                System.Windows.Forms.Application.DoEvents()
                '遍历所有的投资，每个值都变化
                For j = 1 To TZHHJS '行
                    For m = 3 To 11 Step 2 '列
                        JS = (j - 1) * 5 + Int(m / 2)
                        If TZJECC(JS) > 0 Then
                            '改变投资额
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = TZJECC(JS) * (（1 - 2 * MGXFXBHL） + (i - 1) * MGXFXBHL)
                        End If
                    Next
                Next
                '投资金额变化后计算
                Call 建设投资相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, sdsl_model, kcje_xlf_model)
                '读取计算的结果，内部收益率和回收年限
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(6 + i, 11).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 19).Value
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(6 + i, 12).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(13, 19).Value
            Next
            '静态总投资重置回默认值
            Form1.Show()
            Form1.Label1.Text = "正在将静态投资敏感性基本数据重置回默认值。"
            Form1.TopMost = True
            System.Windows.Forms.Application.DoEvents()
            '静态投资重置回默认值
            JS = 0
            For j = 1 To TZHHJS '行
                For m = 3 To 11 Step 2 '列
                    JS = (j - 1) * 5 + Int(m / 2)
                    If TZJECC(JS) > 0 Then
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(TZHH(j), m).Value = TZJECC(JS)
                    End If
                Next
            Next
            '投资金额变化后计算
            Call 建设投资相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, sdsl_model, kcje_xlf_model)
        End If
        '计算一次工作簿
        ExcelApp.Calculate()
        Form1.Close()
    End Sub
    Sub 收入敏感性分析计算(ExcelApp As Object, MGXFXBHL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '实例化计算进度显示窗体
        Dim Form1 As New 计算进度显示
        Dim SRBH As Integer '收入编号
        Dim SRDJ '收入单价
        '————————————————————————————————————————————————————————————————————————————————————————  
        'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim clfl_qtfl_model As Integer = 1
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim sdsl_model As Integer = 1
        '————————————————————————————————————————————————————————————————————————————————————————  
        For i = 13 To 25 '收入&成本表中的有收入的项
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 7).Value > 0 Then '存在收入
                SRBH = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 27).Value
                SRDJ = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value
                For j = 6 To 150
                    If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(j, 27).Value = SRBH And SRBH > 0 And SRDJ > 0 Then '找到编号相同的收入
                        Dim BQ As String = ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(j, 18).Value '标签，用户计算进度显示
                        'Dim XZ1 = MsgBox("是否要进行" & BQ & "敏感性分析计算？", vbOKCancel)
                        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(j, 29).Value = 1 Then
                            ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(j, 8).Value = SRDJ '收入单价写入表格
                            For m = 1 To 5 '敏感性分析
                                '计算进度显示
                                Form1.Show()
                                Form1.Label1.Text = "正在计算" & BQ & "敏感性分析，计算进度：(" & m & "/ 5)"
                                Form1.TopMost = True
                                System.Windows.Forms.Application.DoEvents()
                                '敏感性分析
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJ * (（1 - 2 * MGXFXBHL） + (m - 1) * MGXFXBHL)
                                '收入变化后相关计算
                                Call 计算功能合并整理.收入成本相关计算(ExcelApp, sdsl_model)
                                '读取收益率和回收期
                                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(j - 1 + m, 11).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 19).Value
                                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(j - 1 + m, 12).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(13, 19).Value
                            Next
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 5).Value = SRDJ '收入单价初始值返回
                            '收入变化后相关计算
                            Call 计算功能合并整理.收入成本相关计算(ExcelApp, sdsl_model)
                        End If
                    End If
                Next
            End If
        Next
        '计算一次工作簿
        ExcelApp.Calculate()
        Form1.Close()
    End Sub
    Sub 成本敏感性分析计算(ExcelApp As Object, MGXFXBHL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '实例化计算进度显示窗体
        Dim Form1 As New 计算进度显示
        Dim CBBH As Integer '成本编号
        Dim CBDJ '成本单价
        '————————————————————————————————————————————————————————————————————————————————————————        
        'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim clfl_qtfl_model As Integer = 1
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim sdsl_model As Integer = 1
        '————————————————————————————————————————————————————————————————————————————————————————  
        For i = 14 To 27 '收入&成本表中的有成本的项
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 14).Value > 0 Then '存在成本
                CBBH = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 28).Value
                CBDJ = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value
                For j = 6 To 150
                    If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(j, 28).Value = CBBH And CBBH > 0 And CBDJ > 0 Then '找到编号相同的成本
                        Dim BQ = ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(j, 18).Value '标签，用户计算进度显示
                        'Dim XZ1 = MsgBox("是否要进行" & BQ & "敏感性分析计算？", vbOKCancel)
                        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(j, 29).Value = 1 Then
                            ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(j, 8).Value = CBDJ '成本单价写入表格
                            For m = 1 To 5 '敏感性分析
                                '计算进度显示
                                Form1.Show()
                                Form1.Label1.Text = "正在计算" & BQ & "敏感性分析，计算进度：(" & m & "/ 5)"
                                Form1.TopMost = True
                                System.Windows.Forms.Application.DoEvents()
                                '敏感性分析
                                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJ * (（1 - 2 * MGXFXBHL） + (m - 1) * MGXFXBHL)
                                '成本变化后相关计算
                                Call 计算功能合并整理.收入成本相关计算(ExcelApp, sdsl_model)
                                '读取收益率和回收期
                                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(j - 1 + m, 11).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 19).Value
                                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(j - 1 + m, 12).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(13, 19).Value
                            Next
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 12).Value = CBDJ '成本单价返回初始值
                            '成本变化后相关计算
                            Call 计算功能合并整理.收入成本相关计算(ExcelApp, sdsl_model)
                        End If
                    End If
                Next
            End If
        Next
        '计算一次工作簿
        ExcelApp.Calculate()
        Form1.Close()
    End Sub
    Sub 年运行小时数敏感性分析(ExcelApp As Object, MGXFXBHL As Double)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '实例化计算进度显示窗体
        Dim Form1 As New 计算进度显示
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(137, 29).Value = 1 Then
            'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
            Dim clfl_qtfl_model As Integer = 1
            'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
            Dim sdsl_model As Integer = 1
            '————————————————————————————————————————————————————————————————————————————————————————  
            '只有当输入了年运行小时数时，才计算
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(137, 8).Value > 0 Then
                '计算进度显示
                Form1.Show()
                Form1.Label1.Text = "正在读取年利用小时数敏感性分析基本数据。"
                Form1.TopMost = True
                System.Windows.Forms.Application.DoEvents()
                '基本数据读入
                Dim NLHH(50) As Integer '数组，储存各种年量所在的行号
                Dim NLHHJS As Integer = 0 '各种年量行号计数
                Dim NLCC(200) '数组，储存各种年量
                For i = 30 To 110 '检索估算表中各种年量编号，各种年量编号都是2表示
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, 28).Value = 2 Then
                        NLHHJS = NLHHJS + 1
                        NLHH(NLHHJS) = i '行号储存进数组，从下标1开始
                    End If
                Next
                Dim NLSLJS As Integer = 0 '各种年量数量计数
                '将各种年量储存在数组中
                For i = 1 To NLHHJS '行
                    For j = 3 To 11 Step 2 '列
                        NLSLJS = NLSLJS + 1
                        NLCC(NLSLJS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(NLHH(i), j).Value
                    Next
                Next
                Dim SRNL(20) '收入年量
                Dim CBNL（20） '成本年量
                For i = 13 To 25
                    SRNL(i - 12) = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 6).Value
                Next
                For i = 14 To 27
                    CBNL(i - 12) = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 13).Value
                Next
                Dim JS As Integer = 0
                '计算年运行小时数敏感性分析
                For i = 1 To 5 '敏感性分析循环5次
                    '计算进度显示
                    Form1.Show()
                    Form1.Label1.Text = "正在计算年运行小时数敏感性分析，计算进度：(" & i & "/ 5)"
                    Form1.TopMost = True
                    System.Windows.Forms.Application.DoEvents()
                    '遍历所有的各种年量，每个值都变化
                    For j = 1 To NLHHJS '行
                        For m = 3 To 11 Step 2 '列
                            JS = (j - 1) * 5 + Int(m / 2)
                            If NLCC(JS) > 0 Then
                                ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(NLHH(j), m).Value = NLCC(JS) * (（1 - 2 * MGXFXBHL） + (i - 1) * MGXFXBHL)
                            End If
                        Next
                    Next
                    '遍历各种收入和成本的全年输入，每个值都变化
                    '收入
                    For j = 13 To 25 '行号
                        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 6).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 30).Value = 1 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 6).Value = SRNL(j - 12) * (（1 - 2 * MGXFXBHL） + (i - 1) * MGXFXBHL)
                        End If
                    Next
                    '成本
                    For j = 14 To 27 '行号
                        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 13).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 31).Value = 1 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 13).Value = CBNL(j - 12) * (（1 - 2 * MGXFXBHL） + (i - 1) * MGXFXBHL)
                        End If
                    Next
                    '收入和成本变化后相关计算
                    Call 计算功能合并整理.收入成本相关计算(ExcelApp, sdsl_model)
                    '读取计算的结果，内部收益率和回收年限
                    ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(136 + i, 11).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12, 19).Value
                    ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(136 + i, 12).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(13, 19).Value
                Next
                '计算进度显示
                Form1.Show()
                Form1.Label1.Text = "正将年利用小时数敏感性分析基本数据重置回默认值。"
                Form1.TopMost = True
                System.Windows.Forms.Application.DoEvents()
                '各种重置回默认值
                JS = 0
                For j = 1 To NLHHJS '行
                    For m = 3 To 11 Step 2 '列
                        JS = (j - 1) * 5 + Int(m / 2)
                        If NLCC(JS) > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(NLHH(j), m).Value = NLCC(JS)
                        End If
                    Next
                Next
                '收入成本量重置回默认值
                For j = 13 To 25 '行号
                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 6).Value > 0 Then
                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 6).Value = SRNL(j - 12)
                    End If
                Next
                For j = 14 To 27 '行号
                    If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 13).Value > 0 Then
                        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(j, 13).Value = CBNL(j - 12)
                    End If
                Next
                '收入和成本变化后相关计算
                Call 计算功能合并整理.收入成本相关计算(ExcelApp, sdsl_model)
            End If
        End If
        '计算一次工作簿
        ExcelApp.Calculate()
        Form1.Close()
    End Sub
    Sub 绘制单因素敏感性分析图(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '绘制单因素敏感性分析图
        ExcelApp.ThisWorkbook.Worksheets("指标数据").Shapes.AddChart.Name = "单因素敏感性分析图" '创建图表并重命名
        ExcelApp.ThisWorkbook.Worksheets("指标数据").Shapes("单因素敏感性分析图").Select
        ExcelApp.ActiveChart.ChartType = Excel.XlChartType.xlLineMarkers '选择图表类型
        ExcelApp.ActiveSheet.Shapes("单因素敏感性分析图").IncrementLeft(-415)
        ExcelApp.ActiveSheet.Shapes("单因素敏感性分析图").IncrementTop(340)
        ExcelApp.ThisWorkbook.Worksheets("指标数据").ChartObjects("单因素敏感性分析图").Activate
        ExcelApp.ActiveChart.ChartTitle.Text = "单因素敏感性分析图表"
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Dim js As Integer = 0
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(7, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$7"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$7:$K$11"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(12, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$12"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$12:$K$16"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(17, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$17"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$17:$K$21"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(22, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$22"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$22:$K$26"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(27, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$27"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$27:$K$31"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(32, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$32"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$32:$K$36"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(37, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$37"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$37:$K$41"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(42, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$42"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$42:$K$46"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(47, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$47"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$47:$K$51"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(52, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$52"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$52:$K$56"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(57, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$57"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$57:$K$61"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(62, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$62"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$62:$K$66"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(67, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$67"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$67:$K$71"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(72, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$72"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$72:$K$76"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(77, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$77"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$77:$K$81"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(82, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$82"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$82:$K$86"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(87, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$87"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$87:$K$91"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(92, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$92"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$92:$K$96"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(97, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$97"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$97:$K$101"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(102, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$102"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$102:$K$106"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(107, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$107"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$107:$K$111"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(112, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$112"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$112:$K$116"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(117, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$117"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$117:$K$121"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(122, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$122"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$122:$K$126"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(127, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$127"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$127:$K$131"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(132, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$132"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$132:$K$136"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(137, 8).Value > 0 Then
            js = js + 1
            ExcelApp.ActiveChart.SeriesCollection.NewSeries
            ExcelApp.ActiveChart.SeriesCollection(js).Name = "=指标数据!$R$137"
            ExcelApp.ActiveChart.SeriesCollection(js).Values = "=指标数据!$K$137:$K$141"
            ExcelApp.ActiveChart.SeriesCollection(js).XValues = "=指标数据!$I$7:$I$11"
        End If
    End Sub
    Sub 设置敏感性分析图格式(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '求取计算出的内部收益率中的最小值和最大值
        Dim NBSYLmin As Double = ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(7, 11).Value '内部收益率最小值
        Dim NBSYLmax As Double = ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(7, 11).Value '内部收益率最大值
        For i = 8 To 141 '求最小值
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i, 11).Value < NBSYLmin And ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i, 11).Value <> Nothing Then
                NBSYLmin = ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i, 11).Value
            End If
        Next
        For i = 8 To 141 '求最大值
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i, 11).Value > NBSYLmax And ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i, 11).Value <> Nothing Then
                NBSYLmax = ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i, 11).Value
            End If
        Next
        ExcelApp.ThisWorkbook.Worksheets("指标数据").ChartObjects("单因素敏感性分析图").Activate
        'Y轴数字显示格式
        ExcelApp.ActiveChart.HasAxis(Excel.XlAxisType.xlValue) = True
        ExcelApp.ActiveChart.Axes(Excel.XlAxisType.xlValue).Select
        ExcelApp.Selection.TickLabels.NumberFormatLocal = "0%"
        ExcelApp.ActiveChart.Axes(Excel.XlAxisType.xlValue).MinimumScale = Math.Round(NBSYLmin - 0.005, 2) '设置折线图坐标下限
        ExcelApp.ActiveChart.Axes(Excel.XlAxisType.xlValue).MaximumScale = Math.Round(NBSYLmax + 0.005, 2) '设置折线图坐标上限
        '设置横坐标标题、纵坐标标题、图表总标题
        ExcelApp.ActiveSheet.ChartObjects("单因素敏感性分析图").Activate
        'ExcelApp.ActiveChart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionBottom '设置图例出现在底部
        ExcelApp.ActiveChart.SetElement(MsoChartElementType.msoElementChartTitleAboveChart)
        ExcelApp.Selection.Format.TextFrame2.TextRange.Characters.Text = "单因素敏感性分析折线图"
        With ExcelApp.Selection.Format.TextFrame2.TextRange.Characters(1, 11).ParagraphFormat
            .TextDirection = MsoTextDirection.msoTextDirectionLeftToRight
            .Alignment = MsoAlignCmd.msoAlignCenters
        End With
        With ExcelApp.Selection.Format.TextFrame2.TextRange.Characters(1, 11).Font
            .BaselineOffset = 0
            .Bold = MsoTriState.msoTrue
            .NameComplexScript = "+mn-cs"
            .NameFarEast = "+mn-ea"
            .Fill.Visible = MsoTriState.msoTrue
            .Fill.ForeColor.RGB = RGB(0, 0, 0)
            .Fill.Transparency = 0
            .Fill.Solid
            .Size = 18
            .Italic = MsoTriState.msoFalse
            .Kerning = 12
            .Name = "+mn-lt"
            .UnderlineStyle = MsoTextUnderlineType.msoNoUnderline
            .Strike = MsoTextStrike.msoNoStrike
        End With
        ExcelApp.ActiveChart.ChartTitle.Text = "单因素敏感性分析折线图"
        ExcelApp.Selection.Format.TextFrame2.TextRange.Characters.Text = "单因素敏感性分析折线图"
        With ExcelApp.Selection.Format.TextFrame2.TextRange.Characters(1, 11).ParagraphFormat
            .TextDirection = MsoTextDirection.msoTextDirectionLeftToRight
            .Alignment = MsoAlignCmd.msoAlignCenters
        End With
        With ExcelApp.Selection.Format.TextFrame2.TextRange.Characters(1, 11).Font
            .BaselineOffset = 0
            .Bold = MsoTriState.msoTrue
            .NameComplexScript = "+mn-cs"
            .NameFarEast = "+mn-ea"
            .Fill.Visible = MsoTriState.msoTrue
            .Fill.ForeColor.RGB = RGB(0, 0, 0)
            .Fill.Transparency = 0
            .Fill.Solid
            .Size = 18
            .Italic = MsoTriState.msoFalse
            .Kerning = 12
            .Name = "+mn-lt"
            .UnderlineStyle = MsoTextUnderlineType.msoNoUnderline
            .Strike = MsoTextStrike.msoNoStrike
        End With
        ExcelApp.ActiveChart.ChartArea.Select()
        ExcelApp.ActiveChart.SetElement(MsoChartElementType.msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
        ExcelApp.Selection.Format.TextFrame2.TextRange.Characters.Text = "单因素值变化率（%）"
        With ExcelApp.Selection.Format.TextFrame2.TextRange.Characters(1, 7).ParagraphFormat
            .TextDirection = MsoTextDirection.msoTextDirectionLeftToRight
            .Alignment = MsoAlignCmd.msoAlignCenters
        End With
        With ExcelApp.Selection.Format.TextFrame2.TextRange.Characters(1, 7).Font
            .BaselineOffset = 0
            .Bold = MsoTriState.msoTrue
            .NameComplexScript = "+mn-cs"
            .NameFarEast = "+mn-ea"
            .Fill.Visible = MsoTriState.msoTrue
            .Fill.ForeColor.RGB = RGB(0, 0, 0)
            .Fill.Transparency = 0
            .Fill.Solid
            .Size = 10
            .Italic = MsoTriState.msoFalse
            .Kerning = 12
            .Name = "+mn-lt"
            .UnderlineStyle = MsoTextUnderlineType.msoNoUnderline
            .Strike = MsoTextStrike.msoNoStrike
        End With
        ExcelApp.ActiveChart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = "单因素值变化率（%）"
        ExcelApp.ActiveChart.ChartArea.Select()
        ExcelApp.ActiveChart.SetElement(MsoChartElementType.msoElementPrimaryValueAxisTitleVertical)
        ExcelApp.Selection.Format.TextFrame2.TextRange.Characters.Text = "资本金"
        With ExcelApp.Selection.Format.TextFrame2.TextRange.Characters(1, 3).ParagraphFormat
            .TextDirection = MsoTextDirection.msoTextDirectionLeftToRight
            .Alignment = MsoAlignCmd.msoAlignCenters
        End With
        With ExcelApp.Selection.Format.TextFrame2.TextRange.Characters(1, 3).Font
            .BaselineOffset = 0
            .Bold = MsoTriState.msoTrue
            .NameComplexScript = "+mn-cs"
            .NameFarEast = "+mn-ea"
            .Fill.Visible = MsoTriState.msoTrue
            .Fill.ForeColor.RGB = RGB(0, 0, 0)
            .Fill.Transparency = 0
            .Fill.Solid
            .Size = 10
            .Italic = MsoTriState.msoFalse
            .Kerning = 12
            .Name = "+mn-lt"
            .UnderlineStyle = MsoTextUnderlineType.msoNoUnderline
            .Strike = MsoTextStrike.msoNoStrike
        End With
        ExcelApp.Selection.Format.TextFrame2.TextRange.Characters.Text = "资本金所得税后内部收益率（%）"
        With ExcelApp.Selection.Format.TextFrame2.TextRange.Characters(1, 12).ParagraphFormat
            .TextDirection = MsoTextDirection.msoTextDirectionLeftToRight
            .Alignment = MsoAlignCmd.msoAlignCenters
        End With
        With ExcelApp.Selection.Format.TextFrame2.TextRange.Characters(1, 12).Font
            .BaselineOffset = 0
            .Bold = MsoTriState.msoTrue
            .NameComplexScript = "+mn-cs"
            .NameFarEast = "+mn-ea"
            .Fill.Visible = MsoTriState.msoTrue
            .Fill.ForeColor.RGB = RGB(0, 0, 0)
            .Fill.Transparency = 0
            .Fill.Solid
            .Size = 10
            .Italic = MsoTriState.msoFalse
            .Kerning = 12
            .Name = "+mn-lt"
            .UnderlineStyle = MsoTextUnderlineType.msoNoUnderline
            .Strike = MsoTextStrike.msoNoStrike
        End With
        ExcelApp.ActiveChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).AxisTitle.Text = "资本金所得税后内部收益率（%）"
        ExcelApp.ActiveChart.ChartArea.Select()
        '移动图表至新的位置并放大
        ExcelApp.ActiveSheet.Shapes("单因素敏感性分析图").ScaleWidth(1.9, MsoTriState.msoFalse, MsoScaleFrom.msoScaleFromTopLeft)
        ExcelApp.ActiveSheet.Shapes("单因素敏感性分析图").ScaleHeight(1.8, MsoTriState.msoFalse, MsoScaleFrom.msoScaleFromTopLeft)
        ExcelApp.ActiveSheet.Shapes("单因素敏感性分析图").IncrementLeft(890)
        ExcelApp.ActiveSheet.Shapes("单因素敏感性分析图").IncrementTop(-80)
        '检测是否有多余的数据标签，有多余的就删除
        Dim js As Integer = 0
        For i = 7 To 137 Step 5 '统计应该有多少个标签
            If ExcelApp.ThisWorkbook.Worksheets("指标数据").cells(i, 8).Value > 0 Then
                js = js + 1
            End If
        Next
        ExcelApp.ThisWorkbook.Worksheets("指标数据").ChartObjects("单因素敏感性分析图").Activate
        For i = 1 To 30
            For j = 1 To 30
                ExcelApp.ActiveChart.SeriesCollection(js + j).Delete
            Next
        Next
    End Sub

End Module
