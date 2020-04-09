Public Class 设置补贴收入计算年限
    Private Sub 补贴收入1_Click(sender As Object, e As EventArgs) Handles 补贴收入1.Click
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
        ksnf1 = CType(Com技术经济分析计算程序.Form7.ksnf1.Text, Integer)
        ksnf2 = CType(Com技术经济分析计算程序.Form7.ksnf2.Text, Integer)
        ksnf3 = CType(Com技术经济分析计算程序.Form7.ksnf3.Text, Integer)
        jsnf1 = CType(Com技术经济分析计算程序.Form7.jsnf1.Text, Integer)
        jsnf2 = CType(Com技术经济分析计算程序.Form7.jsnf2.Text, Integer)
        jsnf3 = CType(Com技术经济分析计算程序.Form7.jsnf3.Text, Integer)
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
        '3个勾选框不可以同时被选择
        If Com技术经济分析计算程序.Form7.CheckBox1.Checked = True And Com技术经济分析计算程序.Form7.CheckBox2.Checked = True Then
            MsgBox("补贴收入逐年计算比例仅可以勾选一种计算模式，请重新选择！")
            Exit Sub
        End If
        If Com技术经济分析计算程序.Form7.CheckBox1.Checked = True And Com技术经济分析计算程序.Form7.CheckBox3.Checked = True Then
            MsgBox("补贴收入逐年计算比例仅可以勾选一种计算模式，请重新选择！")
            Exit Sub
        End If
        If Com技术经济分析计算程序.Form7.CheckBox2.Checked = True And Com技术经济分析计算程序.Form7.CheckBox3.Checked = True Then
            MsgBox("补贴收入逐年计算比例仅可以勾选一种计算模式，请重新选择！")
            Exit Sub
        End If
        '解锁表格
        ExcelApp.ThisWorkbook.Worksheets("收入税收表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").unProtect(Password:="wscjc")
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '写入数值
        '按照逐年达产率进行计算
        If Com技术经济分析计算程序.Form7.CheckBox1.Checked = True Then
            '前15年
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 0
                End If
            Next
            '后15年
            For i = 18 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 16).Value
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 16).Value
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 16).Value
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 0
                End If
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 18).Value = "逐年达产率"
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '按照逐年达产率增加值进行计算
        ElseIf Com技术经济分析计算程序.Form7.CheckBox2.Checked = True Then
            '第1年
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, 3).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, 3).Value
            '第2-15年
            For i = 4 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i - 1).Value)
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i - 1).Value)
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i - 1).Value)
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 0
                End If
            Next
            '第16年
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, 18).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, 2).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, 17).Value)
            '第17-30年
            For i = 19 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15 - 1).Value)
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15 - 1).Value)
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15 - 1).Value)
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 0
                End If
            Next
            '将小于0的结果设置为0
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value < 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 0
                End If
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 18).Value = "逐年达产率增加值"
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '按照逐年投产月份比例计算
        ElseIf Com技术经济分析计算程序.Form7.CheckBox3.Checked = True Then
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 0
                End If
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 18).Value = "逐年投产月份比例"
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        Else
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 1
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 1
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value = 0
                End If
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(42, 18).Value = "保持每年100%"
        End If
        '计算一次Excel
        ExcelApp.Calculate()
        '清空已有内容
        Com技术经济分析计算程序.Form7.RichTextBox1.Rtf = Nothing
        Com技术经济分析计算程序.Form7.RichTextBox1.Clear()
        '显示补贴收入计算结果
        Dim btsrbl1 As Double
        Dim nf As Integer
        Dim bq1 As String = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 2).Value, String)
        For i = 3 To 33  '根据数组中的元素数量循环
            If i - 2 <= jsnx Then
                nf = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value '年份序号
                btsrbl1 = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(163, i).Value) * 100, 2) '读取补贴收入负荷率
                Com技术经济分析计算程序.Form7.RichTextBox1.Text = Com技术经济分析计算程序.Form7.RichTextBox1.Text & btsrbl1 & "%(" & nf & ") " '输出到RichTextBox1
            End If
        Next
        Com技术经济分析计算程序.Form7.RichTextBox1.Text = bq1 & "逐年负荷率：" & Com技术经济分析计算程序.Form7.RichTextBox1.Text
    End Sub

    Private Sub 补贴收入2_Click(sender As Object, e As EventArgs) Handles 补贴收入2.Click
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
        ksnf1 = CType(Com技术经济分析计算程序.Form7.ksnf1.Text, Integer)
        ksnf2 = CType(Com技术经济分析计算程序.Form7.ksnf2.Text, Integer)
        ksnf3 = CType(Com技术经济分析计算程序.Form7.ksnf3.Text, Integer)
        jsnf1 = CType(Com技术经济分析计算程序.Form7.jsnf1.Text, Integer)
        jsnf2 = CType(Com技术经济分析计算程序.Form7.jsnf2.Text, Integer)
        jsnf3 = CType(Com技术经济分析计算程序.Form7.jsnf3.Text, Integer)
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
        '3个勾选框不可以同时被选择
        If Com技术经济分析计算程序.Form7.CheckBox1.Checked = True And Com技术经济分析计算程序.Form7.CheckBox2.Checked = True Then
            MsgBox("补贴收入逐年计算比例仅可以勾选一种计算模式，请重新选择！")
            Exit Sub
        End If
        If Com技术经济分析计算程序.Form7.CheckBox1.Checked = True And Com技术经济分析计算程序.Form7.CheckBox3.Checked = True Then
            MsgBox("补贴收入逐年计算比例仅可以勾选一种计算模式，请重新选择！")
            Exit Sub
        End If
        If Com技术经济分析计算程序.Form7.CheckBox2.Checked = True And Com技术经济分析计算程序.Form7.CheckBox3.Checked = True Then
            MsgBox("补贴收入逐年计算比例仅可以勾选一种计算模式，请重新选择！")
            Exit Sub
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '解锁表格
        ExcelApp.ThisWorkbook.Worksheets("收入税收表").unProtect(Password:="wscjc")
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").unProtect(Password:="wscjc")
        '写入数值
        '按照逐年达产率进行计算
        If Com技术经济分析计算程序.Form7.CheckBox1.Checked = True Then
            '前15年
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 0
                End If
            Next
            '后15年
            For i = 18 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 16).Value
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 16).Value
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 16).Value
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 0
                End If
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(43, 18).Value = "逐年达产率"
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '按照逐年达产率增加值进行计算
        ElseIf Com技术经济分析计算程序.Form7.CheckBox2.Checked = True Then
            '第1年
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, 3).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, 3).Value
            '第2-15年
            For i = 4 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i - 1).Value)
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i - 1).Value)
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i - 1).Value)
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 0
                End If
            Next
            '第16年
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, 18).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, 2).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, 17).Value)
            '第17-30年
            For i = 19 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15 - 1).Value)
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15 - 1).Value)
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15 - 1).Value)
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 0
                End If
            Next
            '将小于0的结果设置为0
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value < 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 0
                End If
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(43, 18).Value = "逐年达产率增加值"
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '按照逐年投产月份比例计算
        ElseIf Com技术经济分析计算程序.Form7.CheckBox3.Checked = True Then
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 0
                End If
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(43, 18).Value = "逐年投产月份比例"
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        Else
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 1
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 1
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value = 0
                End If
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(43, 18).Value = "保持每年100%"
        End If
        '计算一次Excel
        ExcelApp.Calculate()
        '清空已有内容
        Com技术经济分析计算程序.Form7.RichTextBox1.Rtf = Nothing
        Com技术经济分析计算程序.Form7.RichTextBox1.Clear()
        '显示补贴收入计算结果
        Dim btsrbl2 As Double
        Dim nf As Integer
        Dim bq2 As String = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 2).Value, String)
        For i = 3 To 33  '根据数组中的元素数量循环
            If i - 2 <= jsnx Then
                nf = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value '年份序号
                btsrbl2 = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(164, i).Value) * 100, 2) '读取补贴收入负荷率
                Com技术经济分析计算程序.Form7.RichTextBox1.Text = Com技术经济分析计算程序.Form7.RichTextBox1.Text & btsrbl2 & "%(" & nf & ") " '输出到RichTextBox1
            End If
        Next
        Com技术经济分析计算程序.Form7.RichTextBox1.Text = bq2 & "逐年负荷率：" & Com技术经济分析计算程序.Form7.RichTextBox1.Text
    End Sub
    Private Sub 光伏补贴收入_Click(sender As Object, e As EventArgs) Handles 光伏补贴收入.Click
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
        ksnf1 = CType(Com技术经济分析计算程序.Form7.ksnf1.Text, Integer)
        ksnf2 = CType(Com技术经济分析计算程序.Form7.ksnf2.Text, Integer)
        ksnf3 = CType(Com技术经济分析计算程序.Form7.ksnf3.Text, Integer)
        jsnf1 = CType(Com技术经济分析计算程序.Form7.jsnf1.Text, Integer)
        jsnf2 = CType(Com技术经济分析计算程序.Form7.jsnf2.Text, Integer)
        jsnf3 = CType(Com技术经济分析计算程序.Form7.jsnf3.Text, Integer)
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
        '两个勾选框都不可以被选择
        If Com技术经济分析计算程序.Form7.CheckBox1.Checked = True Or Com技术经济分析计算程序.Form7.CheckBox2.Checked = True Then
            MsgBox("光伏补贴收入逐年计算比例不可以按照逐年达产率计算或者按照逐年达产率增加值计算，请重新选择！")
            Exit Sub
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '解锁表格
        ExcelApp.ThisWorkbook.Worksheets("收入税收表").unProtect(Password:="wscjc")
        '写入数值
        For i = 3 To 33
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(165, i).Value = 1
            ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(165, i).Value = 1
            ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(165, i).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(165, i).Value = 0
            End If
        Next
        '计算一次Excel
        ExcelApp.Calculate()
        '清空已有内容
        Com技术经济分析计算程序.Form7.RichTextBox1.Rtf = Nothing
        Com技术经济分析计算程序.Form7.RichTextBox1.Clear()
        '显示补贴收入计算结果
        Dim gfbtsr As Double
        Dim nf As Integer
        For i = 3 To 33  '根据数组中的元素数量循环
            If i - 2 <= jsnx Then
                nf = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(162, i).Value '年份序号
                gfbtsr = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(165, i).Value) * 100, 2) '读取补贴收入负荷率
                Com技术经济分析计算程序.Form7.RichTextBox1.Text = Com技术经济分析计算程序.Form7.RichTextBox1.Text & gfbtsr & "%(" & nf & ") " '输出到RichTextBox1
            End If
        Next
        Com技术经济分析计算程序.Form7.RichTextBox1.Text = "光伏补贴计算年份(显示100%的年份表示有光伏补贴，显示0%的表示没有光伏补贴)：" & Com技术经济分析计算程序.Form7.RichTextBox1.Text
    End Sub

    Private Sub 清空窗体_Click(sender As Object, e As EventArgs) Handles 清空窗体.Click
        Dim XZ = MsgBox("确定要清空本窗体输入的的全部内容？", vbOKCancel)
        If XZ = vbOK Then
            Com技术经济分析计算程序.Form7.ksnf1.Clear()
            Com技术经济分析计算程序.Form7.ksnf2.Clear()
            Com技术经济分析计算程序.Form7.ksnf3.Clear()
            Com技术经济分析计算程序.Form7.jsnf1.Clear()
            Com技术经济分析计算程序.Form7.jsnf2.Clear()
            Com技术经济分析计算程序.Form7.jsnf3.Clear()
            Com技术经济分析计算程序.Form7.RichTextBox1.Rtf = Nothing
            Com技术经济分析计算程序.Form7.RichTextBox1.Clear()
            Com技术经济分析计算程序.Form7.CheckBox1.Checked = False
            Com技术经济分析计算程序.Form7.CheckBox2.Checked = False
            Com技术经济分析计算程序.Form7.CheckBox3.Checked = False
            MsgBox("清空窗体已完成！")
        End If
    End Sub

    Private Sub 设置补贴收入计算年限_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '读取计算年限
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '清空窗体
        Com技术经济分析计算程序.Form7.ksnf1.Clear()
        Com技术经济分析计算程序.Form7.ksnf2.Clear()
        Com技术经济分析计算程序.Form7.ksnf3.Clear()
        Com技术经济分析计算程序.Form7.jsnf1.Clear()
        Com技术经济分析计算程序.Form7.jsnf2.Clear()
        Com技术经济分析计算程序.Form7.jsnf3.Clear()
        Com技术经济分析计算程序.Form7.补贴收入1.Text = Nothing
        Com技术经济分析计算程序.Form7.补贴收入2.Text = Nothing
        Com技术经济分析计算程序.Form7.CheckBox1.Checked = False
        Com技术经济分析计算程序.Form7.CheckBox2.Checked = False
        Com技术经济分析计算程序.Form7.CheckBox3.Checked = False
        Com技术经济分析计算程序.Form7.RichTextBox1.Rtf = Nothing
        Com技术经济分析计算程序.Form7.RichTextBox1.Clear()
        '载入按钮文本
        Com技术经济分析计算程序.Form7.补贴收入1.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 2).Value, String)
        Com技术经济分析计算程序.Form7.补贴收入2.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 2).Value, String)
        '载入默认值
        '各种年限系数的计算开始年份（补贴收入、销售收入和成本）
        Dim JSKSNF As Integer = 0
        For i = 3 To 33
            If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value > 0 Then
                JSKSNF = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(156, i).Value
                Exit For
            End If
        Next
        Com技术经济分析计算程序.Form7.ksnf1.Text = CType(JSKSNF, String)
        Com技术经济分析计算程序.Form7.jsnf1.Text = CType(jsnx, String)
    End Sub
End Class