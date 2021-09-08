Public Class 收入和成本计算年限设置
    Private Sub 设置部分销售收入和经营成本计算年限_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '读取计算年限
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '清空窗体
        Me.ksnf1.Clear()
        Me.ksnf2.Clear()
        Me.ksnf3.Clear()
        Me.jsnf1.Clear()
        Me.jsnf2.Clear()
        Me.jsnf3.Clear()
        Me.充电桩收入.Text = Nothing
        Me.风力发电收入.Text = Nothing
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
        Me.风力发电收入.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 2).Value, String)
        Me.购电容量费成本.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 9).Value, String)
        Me.城市管廊成本.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 9).Value, String)
        '是否勾选按照逐年投产月份数折算
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(9, 11).Value = "折算" Then
            Me.CheckBox2.Checked = True
        End If
        '载入默认值
        '各种年限系数的计算开始年份（补贴收入、销售收入和成本）
        Dim JSKSNF As Integer = 0
        For i = 3 To 33
            If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value > 0 Then
                JSKSNF = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(156, i).Value
                Exit For
            End If
        Next
        Me.ksnf1.Text = CType(JSKSNF, String)
        Me.jsnf1.Text = CType(jsnx, String)
    End Sub

    Private Sub 充电桩收入_Click(sender As Object, e As EventArgs) Handles 充电桩收入.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '定义局部变量
        Dim ksnf1, ksnf2, ksnf3 As Integer '开始年份
        Dim jsnf1, jsnf2, jsnf3 As Integer '结束年份
        '读取数据
        ksnf1 = CType(Me.ksnf1.Text, Integer)
        ksnf2 = CType(Me.ksnf2.Text, Integer)
        ksnf3 = CType(Me.ksnf3.Text, Integer)
        jsnf1 = CType(Me.jsnf1.Text, Integer)
        jsnf2 = CType(Me.jsnf2.Text, Integer)
        jsnf3 = CType(Me.jsnf3.Text, Integer)
        '读取计算年限
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '添加报错功能
        If ksnf1 > jsnf1 Or ksnf2 > jsnf2 Or ksnf3 > jsnf3 Then
            MsgBox("开始年份不可以大于结束年份，请重新输入！")
            Exit Sub
        End If
        If ksnf1 <> 0 Or jsnf1 <> 0 Then
            If ksnf1 < 1 Or jsnf1 < 1 Then
                MsgBox("开始年份1或者结束年份1不可以存在小于1的情况，请重新输入！")
                Exit Sub
            End If
        End If
        If ksnf2 <> 0 Or jsnf2 <> 0 Then
            If ksnf2 < 1 Or jsnf2 < 1 Then
                MsgBox("开始年份2或者结束年份2不可以存在小于1的情况，请重新输入！")
                Exit Sub
            End If
        End If
        If ksnf3 <> 0 Or jsnf3 <> 0 Then
            If ksnf3 < 1 Or jsnf3 < 1 Then
                MsgBox("开始年份3或者结束年份3不可以存在小于1的情况，请重新输入！")
                Exit Sub
            End If
        End If
        If ksnf1 > jsnx Or ksnf2 > jsnx Or ksnf3 > jsnx Or jsnf1 > jsnx Or jsnf2 > jsnx Or jsnf3 > jsnx Then
            MsgBox("开始年份或者结束年份存在大于计算年限的情况，程序会继续计算，但会自动忽略大于计算年限的年份的值！")
        End If
        '不可以同时勾选
        If Me.CheckBox2.Checked = True And Me.CheckBox3.Checked = True Then
            MsgBox("充电桩收入计算设置，不可以同时勾选两种计算模式，请重新选择！")
            Exit Sub
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '写入数值    
        '按照逐年投产月份数比例计算
        If Me.CheckBox2.Checked = True Then
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 0
                End If
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(52, 18).Value = "逐年投产月份比例"
        ElseIf Me.CheckBox3.Checked = True Then
            '乘以逐年达产率
            '前15年
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 0
                End If
            Next
            '后16年
            For i = 18 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 16).Value
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 16).Value
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 16).Value
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 0
                End If
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(52, 18).Value = "逐年达产率"
        Else
            '每年都是100%
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 1
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 1
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value = 0
                End If
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(52, 18).Value = "保持每年100%"
        End If
        '计算一次Excel
        ExcelApp.Calculate()
        '———————————————————————————————————————————————————————————————————————————————————————— 
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim sdsl_model As Integer = 1
        '收入和成本变化后相关计算
        Call 计算功能合并整理.收入成本相关计算(ExcelApp, sdsl_model)
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
                nf = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value '年份序号
                srbl1 = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(144, i).Value) * 100, 2) '读取收入负荷率
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & srbl1 & "%(" & nf & ") " '输出到RichTextBox1
            End If
        Next
        Me.RichTextBox1.Text = srbq1 & "逐年负荷率：" & Me.RichTextBox1.Text
    End Sub

    Private Sub 风力发电收入_Click(sender As Object, e As EventArgs) Handles 风力发电收入.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '定义局部变量
        Dim ksnf1, ksnf2, ksnf3 As Integer '开始年份
        Dim jsnf1, jsnf2, jsnf3 As Integer '结束年份
        '读取数据
        ksnf1 = CType(Me.ksnf1.Text, Integer)
        ksnf2 = CType(Me.ksnf2.Text, Integer)
        ksnf3 = CType(Me.ksnf3.Text, Integer)
        jsnf1 = CType(Me.jsnf1.Text, Integer)
        jsnf2 = CType(Me.jsnf2.Text, Integer)
        jsnf3 = CType(Me.jsnf3.Text, Integer)
        '读取计算年限
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '添加报错功能
        If ksnf1 > jsnf1 Or ksnf2 > jsnf2 Or ksnf3 > jsnf3 Then
            MsgBox("开始年份不可以大于结束年份，请重新输入！")
            Exit Sub
        End If
        If ksnf1 <> 0 Or jsnf1 <> 0 Then
            If ksnf1 < 1 Or jsnf1 < 1 Then
                MsgBox("开始年份1或者结束年份1不可以存在小于1的情况，请重新输入！")
                Exit Sub
            End If
        End If
        If ksnf2 <> 0 Or jsnf2 <> 0 Then
            If ksnf2 < 1 Or jsnf2 < 1 Then
                MsgBox("开始年份2或者结束年份2不可以存在小于1的情况，请重新输入！")
                Exit Sub
            End If
        End If
        If ksnf3 <> 0 Or jsnf3 <> 0 Then
            If ksnf3 < 1 Or jsnf3 < 1 Then
                MsgBox("开始年份3或者结束年份3不可以存在小于1的情况，请重新输入！")
                Exit Sub
            End If
        End If
        If ksnf1 > jsnx Or ksnf2 > jsnx Or ksnf3 > jsnx Or jsnf1 > jsnx Or jsnf2 > jsnx Or jsnf3 > jsnx Then
            MsgBox("开始年份或者结束年份存在大于计算年限的情况，程序会继续计算，但会自动忽略大于计算年限的年份的值！")
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '写入数值
        '按照逐年投产月份数比例计算
        If Me.CheckBox2.Checked = True Then
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(143, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(143, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(143, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(143, i).Value = 0
                End If
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(48, 18).Value = "逐年投产月份比例"
        Else
            '每年都是100%
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(143, i).Value = 1
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(143, i).Value = 1
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(143, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(143, i).Value = 0
                End If
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(48, 18).Value = "保持每年100%"
        End If
        '计算一次Excel
        ExcelApp.Calculate()
        '———————————————————————————————————————————————————————————————————————————————————————— 
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim sdsl_model As Integer = 1
        '收入和成本变化后相关计算
        Call 计算功能合并整理.收入成本相关计算(ExcelApp, sdsl_model)
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '清空已有内容
        Me.RichTextBox1.Rtf = Nothing
        Me.RichTextBox1.Clear()
        '显示补贴收入计算结果
        Dim srbl2 As Double
        Dim nf As Integer
        Dim srbq2 As String = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 2).Value, String)
        For i = 3 To 33  '根据数组中的元素数量循环
            If i - 2 <= jsnx Then
                nf = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value '年份序号
                srbl2 = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(143, i).Value) * 100, 2) '读取收入负荷率
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & srbl2 & "%(" & nf & ") " '输出到RichTextBox1
            End If
        Next
        Me.RichTextBox1.Text = srbq2 & "逐年负荷率：" & Me.RichTextBox1.Text
    End Sub

    Private Sub 购电容量费成本_Click(sender As Object, e As EventArgs) Handles 购电容量费成本.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '定义局部变量
        Dim ksnf1, ksnf2, ksnf3 As Integer '开始年份
        Dim jsnf1, jsnf2, jsnf3 As Integer '结束年份
        '读取数据
        ksnf1 = CType(Me.ksnf1.Text, Integer)
        ksnf2 = CType(Me.ksnf2.Text, Integer)
        ksnf3 = CType(Me.ksnf3.Text, Integer)
        jsnf1 = CType(Me.jsnf1.Text, Integer)
        jsnf2 = CType(Me.jsnf2.Text, Integer)
        jsnf3 = CType(Me.jsnf3.Text, Integer)
        '读取计算年限
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '添加报错功能
        If ksnf1 > jsnf1 Or ksnf2 > jsnf2 Or ksnf3 > jsnf3 Then
            MsgBox("开始年份不可以大于结束年份，请重新输入！")
            Exit Sub
        End If
        If ksnf1 <> 0 Or jsnf1 <> 0 Then
            If ksnf1 < 1 Or jsnf1 < 1 Then
                MsgBox("开始年份1或者结束年份1不可以存在小于1的情况，请重新输入！")
                Exit Sub
            End If
        End If
        If ksnf2 <> 0 Or jsnf2 <> 0 Then
            If ksnf2 < 1 Or jsnf2 < 1 Then
                MsgBox("开始年份2或者结束年份2不可以存在小于1的情况，请重新输入！")
                Exit Sub
            End If
        End If
        If ksnf3 <> 0 Or jsnf3 <> 0 Then
            If ksnf3 < 1 Or jsnf3 < 1 Then
                MsgBox("开始年份3或者结束年份3不可以存在小于1的情况，请重新输入！")
                Exit Sub
            End If
        End If
        If ksnf1 > jsnx Or ksnf2 > jsnx Or ksnf3 > jsnx Or jsnf1 > jsnx Or jsnf2 > jsnx Or jsnf3 > jsnx Then
            MsgBox("开始年份或者结束年份存在大于计算年限的情况，程序会继续计算，但会自动忽略大于计算年限的年份的值！")
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '写入数值
        '按照逐年投产月份数比例计算
        If Me.CheckBox2.Checked = True Then
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(145, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(145, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(145, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(145, i).Value = 0
                End If
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(49, 18).Value = "逐年投产月份比例"
        Else
            '每年都是100%
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(145, i).Value = 1
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(145, i).Value = 1
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(145, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(145, i).Value = 0
                End If
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(49, 18).Value = "保持每年100%"
        End If
        '计算一次Excel
        ExcelApp.Calculate()
        '———————————————————————————————————————————————————————————————————————————————————————— 
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim sdsl_model As Integer = 1
        '收入和成本变化后相关计算
        Call 计算功能合并整理.收入成本相关计算(ExcelApp, sdsl_model)
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '清空已有内容
        Me.RichTextBox1.Rtf = Nothing
        Me.RichTextBox1.Clear()
        '显示补贴收入计算结果
        Dim cbbl1 As Double
        Dim nf As Integer
        Dim cbbq1 As String = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 9).Value, String)
        For i = 3 To 33  '根据数组中的元素数量循环
            If i - 2 <= jsnx Then
                nf = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value '年份序号
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
        '定义局部变量
        Dim ksnf1, ksnf2, ksnf3 As Integer '开始年份
        Dim jsnf1, jsnf2, jsnf3 As Integer '结束年份
        '读取数据
        ksnf1 = CType(Me.ksnf1.Text, Integer)
        ksnf2 = CType(Me.ksnf2.Text, Integer)
        ksnf3 = CType(Me.ksnf3.Text, Integer)
        jsnf1 = CType(Me.jsnf1.Text, Integer)
        jsnf2 = CType(Me.jsnf2.Text, Integer)
        jsnf3 = CType(Me.jsnf3.Text, Integer)
        '读取计算年限
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '添加报错功能
        If ksnf1 > jsnf1 Or ksnf2 > jsnf2 Or ksnf3 > jsnf3 Then
            MsgBox("开始年份不可以大于结束年份，请重新输入！")
            Exit Sub
        End If
        If ksnf1 <> 0 Or jsnf1 <> 0 Then
            If ksnf1 < 1 Or jsnf1 < 1 Then
                MsgBox("开始年份1或者结束年份1不可以存在小于1的情况，请重新输入！")
                Exit Sub
            End If
        End If
        If ksnf2 <> 0 Or jsnf2 <> 0 Then
            If ksnf2 < 1 Or jsnf2 < 1 Then
                MsgBox("开始年份2或者结束年份2不可以存在小于1的情况，请重新输入！")
                Exit Sub
            End If
        End If
        If ksnf3 <> 0 Or jsnf3 <> 0 Then
            If ksnf3 < 1 Or jsnf3 < 1 Then
                MsgBox("开始年份3或者结束年份3不可以存在小于1的情况，请重新输入！")
                Exit Sub
            End If
        End If
        If ksnf1 > jsnx Or ksnf2 > jsnx Or ksnf3 > jsnx Or jsnf1 > jsnx Or jsnf2 > jsnx Or jsnf3 > jsnx Then
            MsgBox("开始年份或者结束年份存在大于计算年限的情况，程序会继续计算，但会自动忽略大于计算年限的年份的值！")
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '写入数值
        '按照逐年投产月份数比例计算
        If Me.CheckBox2.Checked = True Then
            '每年都是100%
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(146, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(146, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(146, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(146, i).Value = 0
                End If
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(50, 18).Value = "逐年投产月份比例"
        Else
            '每年都是100%
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(146, i).Value = 1
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(146, i).Value = 1
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(146, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(146, i).Value = 0
                End If
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(50, 18).Value = "保持每年100%"
        End If
        '计算一次Excel
        ExcelApp.Calculate()
        '———————————————————————————————————————————————————————————————————————————————————————— 
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim sdsl_model As Integer = 1
        '收入和成本变化后相关计算
        Call 计算功能合并整理.收入成本相关计算(ExcelApp, sdsl_model)
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
                nf = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value '年份序号
                cbbl2 = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(146, i).Value) * 100, 2) '读取收入负荷率
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & cbbl2 & "%(" & nf & ") " '输出到RichTextBox1
            End If
        Next
        Me.RichTextBox1.Text = cbbq2 & "逐年负荷率：" & Me.RichTextBox1.Text
    End Sub

    Private Sub 设备修理费_Click(sender As Object, e As EventArgs) Handles 设备修理费.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '定义局部变量
        Dim ksnf1, ksnf2, ksnf3 As Integer '开始年份
        Dim jsnf1, jsnf2, jsnf3 As Integer '结束年份
        '读取数据
        ksnf1 = CType(Me.ksnf1.Text, Integer)
        ksnf2 = CType(Me.ksnf2.Text, Integer)
        ksnf3 = CType(Me.ksnf3.Text, Integer)
        jsnf1 = CType(Me.jsnf1.Text, Integer)
        jsnf2 = CType(Me.jsnf2.Text, Integer)
        jsnf3 = CType(Me.jsnf3.Text, Integer)
        '读取计算年限
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '添加报错功能
        If ksnf1 > jsnf1 Or ksnf2 > jsnf2 Or ksnf3 > jsnf3 Then
            MsgBox("开始年份不可以大于结束年份，请重新输入！")
            Exit Sub
        End If
        If ksnf1 <> 0 Or jsnf1 <> 0 Then
            If ksnf1 < 1 Or jsnf1 < 1 Then
                MsgBox("开始年份1或者结束年份1不可以存在小于1的情况，请重新输入！")
                Exit Sub
            End If
        End If
        If ksnf2 <> 0 Or jsnf2 <> 0 Then
            If ksnf2 < 1 Or jsnf2 < 1 Then
                MsgBox("开始年份2或者结束年份2不可以存在小于1的情况，请重新输入！")
                Exit Sub
            End If
        End If
        If ksnf3 <> 0 Or jsnf3 <> 0 Then
            If ksnf3 < 1 Or jsnf3 < 1 Then
                MsgBox("开始年份3或者结束年份3不可以存在小于1的情况，请重新输入！")
                Exit Sub
            End If
        End If
        If ksnf1 > jsnx Or ksnf2 > jsnx Or ksnf3 > jsnx Or jsnf1 > jsnx Or jsnf2 > jsnx Or jsnf3 > jsnx Then
            MsgBox("开始年份或者结束年份存在大于计算年限的情况，程序会继续计算，但会自动忽略大于计算年限的年份的值！")
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '写入数值
        '每年都是100%
        For i = 3 To 33
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(148, i).Value = 1
            ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(148, i).Value = 1
            ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(148, i).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(148, i).Value = 0
            End If
        Next
        '计算一次Excel
        ExcelApp.Calculate()
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '修改逐年设备修理费
        '前15年
        For i = 1 To 15
            ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(137, 4 + i).Value = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(137, 4 + i).Value * ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(148, i + 2).Value
        Next
        '16—31年
        For i = 16 To 31
            ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(141, i - 12).Value = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(141, i - 12).Value * ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(148, i + 2).Value
        Next
        '计算一次Excel
        ExcelApp.Calculate()
        '———————————————————————————————————————————————————————————————————————————————————————— 
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim sdsl_model As Integer = 1
        '收入和成本变化后相关计算
        Call 计算功能合并整理.收入成本相关计算(ExcelApp, sdsl_model)
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '清空已有内容
        Me.RichTextBox1.Rtf = Nothing
        Me.RichTextBox1.Clear()
        '显示补贴收入计算结果
        Dim cbbl3 As Double
        Dim nf As Integer
        For i = 3 To 33  '根据数组中的元素数量循环
            If i - 2 <= jsnx Then
                nf = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value '年份序号
                cbbl3 = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(148, i).Value) * 100, 2) '读取收入负荷率
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & cbbl3 & "%(" & nf & ") " '输出到RichTextBox1
            End If
        Next
        Me.RichTextBox1.Text = "设备修理费计算年份(显示100%的年份表示有设备修理费，显示0%的表示没有设备修理费)：" & Me.RichTextBox1.Text
    End Sub

    Private Sub 人员工资_Click(sender As Object, e As EventArgs) Handles 人员工资.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '定义局部变量
        Dim ksnf1, ksnf2, ksnf3 As Integer '开始年份
        Dim jsnf1, jsnf2, jsnf3 As Integer '结束年份
        '读取数据
        ksnf1 = CType(Me.ksnf1.Text, Integer)
        ksnf2 = CType(Me.ksnf2.Text, Integer)
        ksnf3 = CType(Me.ksnf3.Text, Integer)
        jsnf1 = CType(Me.jsnf1.Text, Integer)
        jsnf2 = CType(Me.jsnf2.Text, Integer)
        jsnf3 = CType(Me.jsnf3.Text, Integer)
        '读取计算年限
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '添加报错功能
        If ksnf1 > jsnf1 Or ksnf2 > jsnf2 Or ksnf3 > jsnf3 Then
            MsgBox("开始年份不可以大于结束年份，请重新输入！")
            Exit Sub
        End If
        If ksnf1 <> 0 Or jsnf1 <> 0 Then
            If ksnf1 < 1 Or jsnf1 < 1 Then
                MsgBox("开始年份1或者结束年份1不可以存在小于1的情况，请重新输入！")
                Exit Sub
            End If
        End If
        If ksnf2 <> 0 Or jsnf2 <> 0 Then
            If ksnf2 < 1 Or jsnf2 < 1 Then
                MsgBox("开始年份2或者结束年份2不可以存在小于1的情况，请重新输入！")
                Exit Sub
            End If
        End If
        If ksnf3 <> 0 Or jsnf3 <> 0 Then
            If ksnf3 < 1 Or jsnf3 < 1 Then
                MsgBox("开始年份3或者结束年份3不可以存在小于1的情况，请重新输入！")
                Exit Sub
            End If
        End If
        If ksnf1 > jsnx Or ksnf2 > jsnx Or ksnf3 > jsnx Or jsnf1 > jsnx Or jsnf2 > jsnx Or jsnf3 > jsnx Then
            MsgBox("开始年份或者结束年份存在大于计算年限的情况，程序会继续计算，但会自动忽略大于计算年限的年份的值！")
        End If
        '不可以同时勾选
        If Me.CheckBox1.Checked = True And Me.CheckBox2.Checked = True Then
            MsgBox("人员工资计算设置，不可以同时勾选两种计算模式，请重新选择！")
            Exit Sub
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '写入数值
        If Me.CheckBox1.Checked = True Then
            '逐年递增
            Dim ZNDZBL As Double = CType(Me.人员工资递增比例.Text, Double) / 100
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(147, i).Value = 1 * (1 + ZNDZBL * (i - 4))
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(147, i).Value = 0
                End If
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(51, 18).Value = "逐年递增"
        ElseIf Me.CheckBox2.Checked = True Then
            '按照逐年投产月份数比例计算
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(147, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(147, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(147, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(147, i).Value = 0
                End If
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(51, 18).Value = "逐年投产月份比例"
        Else
            '人员工资每年都是100%，不逐年递增
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(147, i).Value = 1
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(147, i).Value = 1
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(147, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(147, i).Value = 0
                End If
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(51, 18).Value = "保持每年100%"
        End If
        '计算一次Excel
        ExcelApp.Calculate()
        '———————————————————————————————————————————————————————————————————————————————————————— 
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim sdsl_model As Integer = 1
        '收入和成本变化后相关计算
        Call 计算功能合并整理.收入成本相关计算(ExcelApp, sdsl_model)
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '清空已有内容
        Me.RichTextBox1.Rtf = Nothing
        Me.RichTextBox1.Clear()
        '显示补贴收入计算结果
        Dim cbbl4 As Double
        Dim nf As Integer
        For i = 3 To 33  '根据数组中的元素数量循环
            If i - 2 <= jsnx Then
                nf = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(142, i).Value '年份序号
                cbbl4 = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(147, i).Value) * 100, 2) '读取收入负荷率
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & cbbl4 & "%(" & nf & ") " '输出到RichTextBox1
            End If
        Next
        Me.RichTextBox1.Text = "人员工资逐年计算比例结果：：" & Me.RichTextBox1.Text
    End Sub
End Class