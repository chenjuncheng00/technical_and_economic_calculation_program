Public Class 设置接入费计算方式
    Private Sub 设置接入费计算方式_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        Me.CheckBox1.Checked = False
        Me.CheckBox2.Checked = False
        Me.RichTextBox1.Rtf = Nothing
        Me.RichTextBox1.Clear()
        '载入默认值
        '各种年限系数的计算开始年份（补贴收入、销售收入和成本）
        Dim JSKSNF As Integer = 0
        For i = 1 To 15
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(5, 4 + i).Value > 0 Then
                JSKSNF = i
                Exit For
            End If
        Next
        Me.ksnf1.Text = CType(JSKSNF, String)
        Me.jsnf1.Text = CType(jsnx, String)
    End Sub

    Private Sub 清空窗体_Click(sender As Object, e As EventArgs) Handles 清空窗体.Click
        Dim XZ = MsgBox("确定要清空本窗体输入的的全部内容？", vbOKCancel)
        If XZ = vbOK Then
            Me.ksnf1.Clear()
            Me.ksnf2.Clear()
            Me.ksnf3.Clear()
            Me.jsnf1.Clear()
            Me.jsnf2.Clear()
            Me.jsnf3.Clear()
            Me.CheckBox1.Checked = False
            Me.CheckBox2.Checked = False
            Me.RichTextBox1.Rtf = Nothing
            Me.RichTextBox1.Clear()
            MsgBox("清空窗体已完成！")
        End If
    End Sub

    Private Sub 确定参数_Click(sender As Object, e As EventArgs) Handles 确定参数.Click
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
        '不可以两个都勾选
        If Me.CheckBox1.Checked = True And Me.CheckBox2.Checked = True Then
            MsgBox("接入费逐年计算比例仅可以勾选一种计算模式，请重新选择！")
            Exit Sub
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '写入数值
        '按照逐年达产率进行计算
        If Me.CheckBox1.Checked = True Then
            '前15年
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
                End If
            Next
            '后16年
            For i = 18 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 16).Value
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 16).Value
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 16).Value
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
                End If
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(41, 18).Value = "逐年达产率"
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        ElseIf Me.CheckBox2.Checked = True Then
            '逐年投产月份比例
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
                End If
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(41, 18).Value = "逐年投产月份比例"
        Else
            '每年都是100%
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1
                ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
                End If
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(41, 18).Value = "保持每年100%"
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
        Dim jrfbl As Double
        Dim nf As Integer
        For i = 3 To 33  '根据数组中的元素数量循环
            If i - 2 <= jsnx Then
                nf = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value '年份序号
                jrfbl = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value) * 100, 2) '读取接入费收入负荷率
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & jrfbl & "%(" & nf & ") " '输出到RichTextBox1
            End If
        Next
        Me.RichTextBox1.Text = "接入费收入逐年负荷率：" & Me.RichTextBox1.Text
    End Sub

    Private Sub 默认方式_Click(sender As Object, e As EventArgs) Handles 默认方式.Click
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
        '第1年
        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, 3).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, 3).Value
        '第2-15年
        For i = 4 To 17
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i - 1).Value)
            ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i - 1).Value)
            ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i - 1).Value)
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
            End If
        Next
        '第16年
        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, 18).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, 2).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, 17).Value)
        '第17-31年
        For i = 19 To 33
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= ksnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnf1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15 - 1).Value)
            ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= ksnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnf2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15 - 1).Value)
            ElseIf ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= ksnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnf3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15 - 1).Value)
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
            End If
        Next
        '将小于0的结果设置为0
        For i = 3 To 33
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value < 0 Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
            End If
        Next
        ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(41, 18).Value = "逐年达产率增加值"
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
        Dim jrfbl As Double
        Dim nf As Integer
        For i = 3 To 33  '根据数组中的元素数量循环
            If i - 2 <= jsnx Then
                nf = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value '年份序号
                jrfbl = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value) * 100, 2) '读取接入费收入负荷率
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & jrfbl & "%(" & nf & ") " '输出到RichTextBox1
            End If
        Next
        Me.RichTextBox1.Text = "接入费收入逐年负荷率：" & Me.RichTextBox1.Text
    End Sub
End Class