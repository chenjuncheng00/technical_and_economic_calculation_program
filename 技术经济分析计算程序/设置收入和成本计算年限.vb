Public Class 设置收入和成本计算年限
    '开始年份和结束年份的列表
    Public ksnf_list As New List(Of Integer)
    Public jsnf_list As New List(Of Integer)
    Private Sub 设置收入和成本计算年限_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '读取计算年限
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '清空窗体
        Me.开始年份列表.Items.Clear()
        Me.结束年份列表.Items.Clear()
        Me.开始年份tmp.Clear()
        Me.结束年份tmp.Clear()
        Me.充电桩收入.Text = Nothing
        Me.购电容量费成本.Text = Nothing
        Me.城市管廊成本.Text = Nothing
        Me.CheckBox1.Checked = False
        Me.CheckBox2.Checked = False
        Me.CheckBox3.Checked = False
        Me.人员工资递增比例.Clear()
        Me.RichTextBox1.Rtf = Nothing
        Me.RichTextBox1.Clear()
        '载入按钮文本
        Me.充电桩收入.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 2).Value, String)
        Me.购电容量费成本.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 9).Value, String)
        Me.城市管廊成本.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 9).Value, String)
        '是否勾选按照逐年投产月份数折算
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(9, 11).Value = "折算" Then
            Me.CheckBox2.Checked = True
        End If
        '载入默认值
        '各种年限系数的计算开始年份（补贴收入、销售收入和成本）
        Dim JSKSNF As Integer = 0
        For i = 1 To 15
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(5, 4 + i).Value > 0 Then
                JSKSNF = i
                Exit For
            End If
        Next
        Me.开始年份tmp.Text = CType(JSKSNF, String)
        Me.结束年份tmp.Text = CType(jsnx, String)
    End Sub

    Private Sub 充电桩收入_Click(sender As Object, e As EventArgs) Handles 充电桩收入.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '读取计算年限
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '输入的年份总数量
        Dim n_nf As Integer = ksnf_list.LongCount
        '添加报错功能
        For i = 0 To n_nf - 1
            If ksnf_list(i) <> 0 Or jsnf_list(i) <> 0 Then
                If ksnf_list(i) < 1 Or jsnf_list(i) < 1 Then
                    MsgBox("开始年份或者结束年份不可以存在小于1的情况，请重新输入！")
                    Exit Sub
                End If
            End If
        Next
        For i = 0 To n_nf - 1
            If ksnf_list(i) > jsnx Or jsnf_list(i) > jsnx Then
                MsgBox("开始年份或者结束年份存在大于计算年限的情况，程序会继续计算，但会自动忽略大于计算年限的年份的值！")
                Exit For
            End If
        Next
        If n_nf > 1 Then
            For i = 1 To n_nf - 1
                If jsnf_list(i - 1) > ksnf_list(i) Then
                    MsgBox("输入的后一个开始年份不可以小于上一个结束年份，请重新输入！")
                    Exit Sub
                End If
            Next
        End If
        '不可以同时勾选
        If Me.CheckBox2.Checked = True And Me.CheckBox3.Checked = True Then
            MsgBox("充电桩收入计算设置，不可以同时勾选两种计算模式，请重新选择！")
            Exit Sub
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '计算在输入的开始年份和结束年份之外的年份序号
        Dim qtnf_tmp_list As New List(Of Integer)
        For i = 1 To jsnx
            qtnf_tmp_list.Add(i)
        Next
        Dim tmp_list As New List(Of Integer)
        For i = 0 To n_nf - 1
            For j = ksnf_list(i) To jsnf_list(i)
                tmp_list.Add(j)
            Next
        Next
        Dim qtnf_list = qtnf_tmp_list.Except(tmp_list).ToList()
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '确定计算模式
        If Me.CheckBox2.Checked = True Then
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(52, 18).Value = "逐年投产月份比例"
        ElseIf Me.CheckBox3.Checked = True Then
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(52, 18).Value = "逐年达产率"
        Else
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(52, 18).Value = "保持每年100%"
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '获取10次投资,每一次投资的投产月份数
        Dim ans_month = 逐年投产月份数_10次投资(ExcelApp)
        Dim month_10_list = ans_month(1)
        '读取基础负荷率
        Dim fhl_base = 读取逐年负荷率(ExcelApp)
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '写入数值    
        Call 充电桩收入逐年达产率计算(ExcelApp, jsnx, month_10_list, fhl_base, qtnf_list)
        '计算一次Excel
        ExcelApp.Calculate()
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
        '收入和成本变化后相关计算
        Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '清空已有内容
        Me.RichTextBox1.Rtf = Nothing
        Me.RichTextBox1.Clear()
        '显示补贴收入计算结果
        Dim srbl1 As Double
        Dim nf As Integer
        Dim srbq1 As String = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 2).Value, String)
        For i = 3 To 33  '根据数组中的元素数量循环
            If i - 2 <= jsnx Then
                nf = i - 2 '年份序号
                srbl1 = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value) * 100, 2) '读取收入负荷率
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & srbl1 & "%(" & nf & ") " '输出到RichTextBox1
            End If
        Next
        Me.RichTextBox1.Text = srbq1 & "逐年负荷率：" & Me.RichTextBox1.Text
    End Sub
    Private Sub 购电容量费成本_Click(sender As Object, e As EventArgs) Handles 购电容量费成本.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '读取计算年限
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '输入的年份总数量
        Dim n_nf As Integer = ksnf_list.LongCount
        '添加报错功能
        For i = 0 To n_nf - 1
            If ksnf_list(i) <> 0 Or jsnf_list(i) <> 0 Then
                If ksnf_list(i) < 1 Or jsnf_list(i) < 1 Then
                    MsgBox("开始年份或者结束年份不可以存在小于1的情况，请重新输入！")
                    Exit Sub
                End If
            End If
        Next
        For i = 0 To n_nf - 1
            If ksnf_list(i) > jsnx Or jsnf_list(i) > jsnx Then
                MsgBox("开始年份或者结束年份存在大于计算年限的情况，程序会继续计算，但会自动忽略大于计算年限的年份的值！")
                Exit For
            End If
        Next
        If n_nf > 1 Then
            For i = 1 To n_nf - 1
                If jsnf_list(i - 1) > ksnf_list(i) Then
                    MsgBox("输入的后一个开始年份不可以小于上一个结束年份，请重新输入！")
                    Exit Sub
                End If
            Next
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '计算在输入的开始年份和结束年份之外的年份序号
        Dim qtnf_tmp_list As New List(Of Integer)
        For i = 1 To jsnx
            qtnf_tmp_list.Add(i)
        Next
        Dim tmp_list As New List(Of Integer)
        For i = 0 To n_nf - 1
            For j = ksnf_list(i) To jsnf_list(i)
                tmp_list.Add(j)
            Next
        Next
        Dim qtnf_list = qtnf_tmp_list.Except(tmp_list).ToList()
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '确定计算模式
        If Me.CheckBox2.Checked = True Then
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(49, 18).Value = "逐年投产月份比例"
        Else
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(49, 18).Value = "保持每年100%"
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '获取10次投资,每一次投资的投产月份数
        Dim ans_month = 逐年投产月份数_10次投资(ExcelApp)
        Dim month_10_list = ans_month(1)
        '读取基础负荷率
        Dim fhl_base = 读取逐年负荷率(ExcelApp)
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        Call 购电容量费成本逐年达产率计算(ExcelApp, jsnx, month_10_list, qtnf_list)
        '计算一次Excel
        ExcelApp.Calculate()
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
        '收入和成本变化后相关计算
        Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        '————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '清空已有内容
        Me.RichTextBox1.Rtf = Nothing
        Me.RichTextBox1.Clear()
        '显示补贴收入计算结果
        Dim cbbl1 As Double
        Dim nf As Integer
        Dim cbbq1 As String = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 9).Value, String)
        For i = 3 To 33  '根据数组中的元素数量循环
            If i - 2 <= jsnx Then
                nf = i - 2 '年份序号
                cbbl1 = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(145, i).Value) * 100, 2) '读取收入负荷率
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & cbbl1 & "%(" & nf & ") " '输出到RichTextBox1
            End If
        Next
        Me.RichTextBox1.Text = cbbq1 & "逐年负荷率：" & Me.RichTextBox1.Text
    End Sub

    Private Sub 城市管廊成本_Click(sender As Object, e As EventArgs) Handles 城市管廊成本.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '读取计算年限
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '输入的年份总数量
        Dim n_nf As Integer = ksnf_list.LongCount
        '添加报错功能
        For i = 0 To n_nf - 1
            If ksnf_list(i) <> 0 Or jsnf_list(i) <> 0 Then
                If ksnf_list(i) < 1 Or jsnf_list(i) < 1 Then
                    MsgBox("开始年份或者结束年份不可以存在小于1的情况，请重新输入！")
                    Exit Sub
                End If
            End If
        Next
        For i = 0 To n_nf - 1
            If ksnf_list(i) > jsnx Or jsnf_list(i) > jsnx Then
                MsgBox("开始年份或者结束年份存在大于计算年限的情况，程序会继续计算，但会自动忽略大于计算年限的年份的值！")
                Exit For
            End If
        Next
        If n_nf > 1 Then
            For i = 1 To n_nf - 1
                If jsnf_list(i - 1) > ksnf_list(i) Then
                    MsgBox("输入的后一个开始年份不可以小于上一个结束年份，请重新输入！")
                    Exit Sub
                End If
            Next
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '计算在输入的开始年份和结束年份之外的年份序号
        Dim qtnf_tmp_list As New List(Of Integer)
        For i = 1 To jsnx
            qtnf_tmp_list.Add(i)
        Next
        Dim tmp_list As New List(Of Integer)
        For i = 0 To n_nf - 1
            For j = ksnf_list(i) To jsnf_list(i)
                tmp_list.Add(j)
            Next
        Next
        Dim qtnf_list = qtnf_tmp_list.Except(tmp_list).ToList()
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '确定计算模式
        If Me.CheckBox2.Checked = True Then
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(50, 18).Value = "逐年投产月份比例"
        Else
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(50, 18).Value = "保持每年100%"
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '获取10次投资,每一次投资的投产月份数
        Dim ans_month = 逐年投产月份数_10次投资(ExcelApp)
        Dim month_10_list = ans_month(1)
        '读取基础负荷率
        Dim fhl_base = 读取逐年负荷率(ExcelApp)
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '写入数值
        Call 城市管廊成本逐年达产率计算(ExcelApp, jsnx, month_10_list, qtnf_list)
        '计算一次Excel
        ExcelApp.Calculate()
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
        '收入和成本变化后相关计算
        Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '清空已有内容
        Me.RichTextBox1.Rtf = Nothing
        Me.RichTextBox1.Clear()
        '显示补贴收入计算结果
        Dim cbbl2 As Double
        Dim nf As Integer
        Dim cbbq2 As String = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 9).Value, String)
        For i = 3 To 33  '根据数组中的元素数量循环
            If i - 2 <= jsnx Then
                nf = i - 2 '年份序号
                cbbl2 = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(146, i).Value) * 100, 2) '读取收入负荷率
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & cbbl2 & "%(" & nf & ") " '输出到RichTextBox1
            End If
        Next
        Me.RichTextBox1.Text = cbbq2 & "逐年负荷率：" & Me.RichTextBox1.Text
    End Sub

    Private Sub 人员工资_Click(sender As Object, e As EventArgs) Handles 人员工资.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '读取计算年限
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '输入的年份总数量
        Dim n_nf As Integer = ksnf_list.LongCount
        '添加报错功能
        For i = 0 To n_nf - 1
            If ksnf_list(i) <> 0 Or jsnf_list(i) <> 0 Then
                If ksnf_list(i) < 1 Or jsnf_list(i) < 1 Then
                    MsgBox("开始年份或者结束年份不可以存在小于1的情况，请重新输入！")
                    Exit Sub
                End If
            End If
        Next
        For i = 0 To n_nf - 1
            If ksnf_list(i) > jsnx Or jsnf_list(i) > jsnx Then
                MsgBox("开始年份或者结束年份存在大于计算年限的情况，程序会继续计算，但会自动忽略大于计算年限的年份的值！")
                Exit For
            End If
        Next
        If n_nf > 1 Then
            For i = 1 To n_nf - 1
                If jsnf_list(i - 1) > ksnf_list(i) Then
                    MsgBox("输入的后一个开始年份不可以小于上一个结束年份，请重新输入！")
                    Exit Sub
                End If
            Next
        End If
        '不可以同时勾选
        If Me.CheckBox1.Checked = True And Me.CheckBox2.Checked = True Then
            MsgBox("人员工资计算设置，不可以同时勾选两种计算模式，请重新选择！")
            Exit Sub
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '计算在输入的开始年份和结束年份之外的年份序号
        Dim qtnf_tmp_list As New List(Of Integer)
        For i = 1 To jsnx
            qtnf_tmp_list.Add(i)
        Next
        Dim tmp_list As New List(Of Integer)
        For i = 0 To n_nf - 1
            For j = ksnf_list(i) To jsnf_list(i)
                tmp_list.Add(j)
            Next
        Next
        Dim qtnf_list = qtnf_tmp_list.Except(tmp_list).ToList()
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '确定计算模式
        If Me.CheckBox1.Checked = True Then
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(51, 18).Value = "逐年递增"
        ElseIf Me.CheckBox2.Checked = True Then
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(51, 18).Value = "逐年投产月份比例"
        Else
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(51, 18).Value = "保持每年100%"
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '逐年递增
        Dim ZNDZBL As Double = CType(Me.人员工资递增比例.Text, Double) / 100
        '获取10次投资,每一次投资的投产月份数
        Dim ans_month = 逐年投产月份数_10次投资(ExcelApp)
        Dim month_10_list = ans_month(1)
        '读取基础负荷率
        Dim fhl_base = 读取逐年负荷率(ExcelApp)
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '写入数值
        Call 人员工资成本逐年达产率计算(ExcelApp, jsnx, month_10_list, ZNDZBL, qtnf_list)
        '计算一次Excel
        ExcelApp.Calculate()
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
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
        '收入和成本变化后相关计算
        Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '清空已有内容
        Me.RichTextBox1.Rtf = Nothing
        Me.RichTextBox1.Clear()
        '显示补贴收入计算结果
        Dim cbbl4 As Double
        Dim nf As Integer
        For i = 3 To 33  '根据数组中的元素数量循环
            If i - 2 <= jsnx Then
                nf = i - 2 '年份序号
                cbbl4 = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(147, i).Value) * 100, 2) '读取收入负荷率
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & cbbl4 & "%(" & nf & ") " '输出到RichTextBox1
            End If
        Next
        Me.RichTextBox1.Text = "人员工资逐年计算比例结果：" & Me.RichTextBox1.Text
    End Sub

    Private Sub 添加输入_Click(sender As Object, e As EventArgs) Handles 添加输入.Click
        On Error Resume Next
        '判断1
        If 开始年份tmp.Text = Nothing Or 结束年份tmp.Text = Nothing Then
            MsgBox("输入的开始年份和结束年份都必须不能为空！，请重新输入")
            Exit Sub
        End If
        '开始年份
        Dim ksnf_text As String = 开始年份tmp.Text
        Dim ksnf As Integer = CType(ksnf_text, Integer)
        '结束年份
        Dim jsnf_text As String = 结束年份tmp.Text
        Dim jsnf As Integer = CType(jsnf_text, Integer)
        '判断2
        If ksnf > jsnf Then
            MsgBox("输入的开始年份必须小于等于结束年份！，请重新输入")
            Exit Sub
        End If
        '添加数据
        开始年份列表.Items.Add(ksnf_text)
        ksnf_list.Add(ksnf)
        开始年份tmp.Clear()
        结束年份列表.Items.Add(jsnf_text)
        jsnf_list.Add(jsnf)
        结束年份tmp.Clear()
    End Sub

    Private Sub 清空输入_Click(sender As Object, e As EventArgs) Handles 清空输入.Click
        On Error Resume Next
        '清空窗体
        Me.开始年份列表.Items.Clear()
        Me.结束年份列表.Items.Clear()
        Me.开始年份tmp.Clear()
        Me.结束年份tmp.Clear()
        Me.人员工资递增比例.Clear()
        ksnf_list.Clear()
        jsnf_list.Clear()
    End Sub
End Class