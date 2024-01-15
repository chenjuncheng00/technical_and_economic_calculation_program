Public Class 设置接入费计算方式
    '开始年份和结束年份的列表
    Public ksnf_list As New List(Of Integer)
    Public jsnf_list As New List(Of Integer)
    Private Sub 设置接入费计算方式_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        Me.开始年份tmp.Text = CType(JSKSNF, String)
        Me.结束年份tmp.Text = CType(jsnx, String)
    End Sub

    Private Sub 清空窗体_Click(sender As Object, e As EventArgs) Handles 清空窗体.Click
        Dim XZ = MsgBox("确定要清空本窗体输入的的全部内容？", vbOKCancel)
        If XZ = vbOK Then
            Me.开始年份列表.Items.Clear()
            Me.结束年份列表.Items.Clear()
            Me.开始年份tmp.Clear()
            Me.结束年份tmp.Clear()
            ksnf_list.Clear()
            jsnf_list.Clear()
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
        '不可以两个都勾选
        If Me.CheckBox1.Checked = True And Me.CheckBox2.Checked = True Then
            MsgBox("接入费逐年计算比例仅可以勾选一种计算模式，请重新选择！")
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
        Dim qtnf_list = qtnf_tmp_list.Except(tmp_list)
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '写入数值
        '按照逐年达产率进行计算
        If Me.CheckBox1.Checked = True Then
            '正常计算的年份
            For j = 0 To n_nf - 1
                '前15年
                For i = 3 To 17
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= ksnf_list(j) And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnf_list(j) And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value
                    End If
                Next
                '后16年
                For i = 18 To 33
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= ksnf_list(j) And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnf_list(j) And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 16).Value
                    End If
                Next
            Next
            '其他年份
            For Each qtnf In qtnf_list
                '前15年
                For i = 3 To 17
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value = qtnf Or ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value > jsnx Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
                    End If
                Next
                '后16年
                For i = 18 To 33
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value = qtnf Or ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value > jsnx Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
                    End If
                Next
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(41, 18).Value = "逐年达产率"
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        ElseIf Me.CheckBox2.Checked = True Then
            '逐年投产月份比例
            '正常计算的年份
            For j = 0 To n_nf - 1
                For i = 3 To 33
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= ksnf_list(j) And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnf_list(j) And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value / 12
                    End If
                Next
            Next
            '其他年份
            For Each qtnf In qtnf_list
                For i = 3 To 33
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value = qtnf Or ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value > jsnx Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
                    End If
                Next
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(41, 18).Value = "逐年投产月份比例"
        Else
            '每年都是100%
            '正常计算的年份
            For j = 0 To n_nf - 1
                For i = 3 To 33
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= ksnf_list(j) And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnf_list(j) And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1
                    End If
                Next
            Next
            '其他年份
            For Each qtnf In qtnf_list
                For i = 3 To 33
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value = qtnf Or ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value > jsnx Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
                    End If
                Next
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(41, 18).Value = "保持每年100%"
        End If
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
        '不可以两个都勾选
        If Me.CheckBox1.Checked = True And Me.CheckBox2.Checked = True Then
            MsgBox("接入费逐年计算比例仅可以勾选一种计算模式，请重新选择！")
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
        Dim qtnf_list = qtnf_tmp_list.Except(tmp_list)
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
        '正常计算的年份
        '第1年
        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, 3).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, 3).Value
        '第2-15年
        For j = 0 To n_nf - 1
            For i = 4 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= ksnf_list(j) And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnf_list(j) And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i - 1).Value)
                End If
            Next
        Next
        '第16年
        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, 18).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, 2).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, 17).Value)
        '第17-31年
        For j = 0 To n_nf - 1
            For i = 19 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= ksnf_list(j) And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnf_list(j) And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15 - 1).Value)
                End If
            Next
        Next
        '其他年份
        For Each qtnf In qtnf_list
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value = qtnf Or ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value > jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
                End If
            Next
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
        ksnf_list.Clear()
        jsnf_list.Clear()
    End Sub
End Class