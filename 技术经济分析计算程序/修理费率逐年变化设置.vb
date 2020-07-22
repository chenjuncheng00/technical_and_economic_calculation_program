Imports System.Diagnostics.Eventing.Reader

Public Class 修理费率逐年变化设置
    Private Sub 清空窗体_Click(sender As Object, e As EventArgs) Handles 清空窗体.Click
        Dim XZ = MsgBox("是否清空窗体中的全部内容？", vbOKCancel)
        If XZ = vbOK Then
            Me.ksnf1.Clear()
            Me.ksnf2.Clear()
            Me.ksnf3.Clear()
            Me.ksnf4.Clear()
            Me.ksnf5.Clear()
            Me.jsnf1.Clear()
            Me.jsnf2.Clear()
            Me.jsnf3.Clear()
            Me.jsnf4.Clear()
            Me.jsnf5.Clear()
            Me.ksfl1.Clear()
            Me.ksfl2.Clear()
            Me.ksfl3.Clear()
            Me.ksfl4.Clear()
            Me.ksfl5.Clear()
            Me.jsfl1.Clear()
            Me.jsfl2.Clear()
            Me.jsfl3.Clear()
            Me.jsfl4.Clear()
            Me.jsfl5.Clear()
            Me.RichTextBox1.Rtf = Nothing
            Me.TZ1.Checked = False
            Me.TZ2.Checked = False
            Me.TZ3.Checked = False
            Me.TZ4.Checked = False
            Me.TZ5.Checked = False
            Me.TZ6.Checked = False
            Me.TZ7.Checked = False
            Me.TZ8.Checked = False
            Me.TZ9.Checked = False
            Me.TZ10.Checked = False
            Me.TZBL1.Clear()
            Me.TZBL2.Clear()
            Me.TZBL3.Clear()
            Me.TZBL4.Clear()
            Me.TZBL5.Clear()
            Me.TZBL6.Clear()
            Me.TZBL7.Clear()
            Me.TZBL8.Clear()
            Me.TZBL9.Clear()
            Me.TZBL10.Clear()
        End If
    End Sub

    Private Sub 确定_Click(sender As Object, e As EventArgs) Handles 确定.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '定义局部变量
        Dim KSNF1, KSNF2, KSNF3, KSNF4, KSNF5 As Integer '开始年份
        Dim JSNF1, JSNF2, JSNF3, JSNF4, JSNF5 As Integer '结束年份
        Dim KSNFFL1, KSNFFL2, KSNFFL3, KSNFFL4, KSNFFL5 As Double '开始年份费率
        Dim JSNFFL1, JSNFFL2, JSNFFL3, JSNFFL4, JSNFFL5 As Double '结束年份费率
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————  
        KSNF1 = CType(Me.ksnf1.Text, Integer)
        KSNF2 = CType(Me.ksnf2.Text, Integer)
        KSNF3 = CType(Me.ksnf3.Text, Integer)
        KSNF4 = CType(Me.ksnf4.Text, Integer)
        KSNF5 = CType(Me.ksnf5.Text, Integer)
        JSNF1 = CType(Me.jsnf1.Text, Integer)
        JSNF2 = CType(Me.jsnf2.Text, Integer)
        JSNF3 = CType(Me.jsnf3.Text, Integer)
        JSNF4 = CType(Me.jsnf4.Text, Integer)
        JSNF5 = CType(Me.jsnf5.Text, Integer)
        KSNFFL1 = CType(Me.ksfl1.Text, Double)
        KSNFFL2 = CType(Me.ksfl2.Text, Double)
        KSNFFL3 = CType(Me.ksfl3.Text, Double)
        KSNFFL4 = CType(Me.ksfl4.Text, Double)
        KSNFFL5 = CType(Me.ksfl5.Text, Double)
        JSNFFL1 = CType(Me.jsfl1.Text, Double)
        JSNFFL2 = CType(Me.jsfl2.Text, Double)
        JSNFFL3 = CType(Me.jsfl3.Text, Double)
        JSNFFL4 = CType(Me.jsfl4.Text, Double)
        JSNFFL5 = CType(Me.jsfl5.Text, Double)
        '————————————————————————————————————————————————————————————————————————————————————————  
        Dim ZDJSNF1 As Integer = Math.Max(JSNF1, JSNF2)
        Dim ZDJSNF2 As Integer = Math.Max(JSNF3, JSNF4)
        Dim ZDJSNF3 = Math.Max(ZDJSNF1, ZDJSNF2)
        Dim ZDJSNF = Math.Max(ZDJSNF3, JSNF5)
        '————————————————————————————————————————————————————————————————————————————————————————
        '各种年限系数的计算开始年份（补贴收入、销售收入和成本）
        Dim JSKSNF As Integer = 0
        For i = 3 To 33
            If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value > 0 Then
                JSKSNF = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(156, i).Value
                Exit For
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否确定输入的修理费计算系数？", vbOKCancel)
        If XZ = vbOK Then
            '检查输入的开始年份和结束年份
            If KSNF1 > JSNF1 Or KSNF2 > JSNF2 Or KSNF3 > JSNF3 Or KSNF4 > JSNF4 Or KSNF5 > JSNF5 Then
                MsgBox("输入的开始年份不可以大于结束年份，请重新输入！")
                Exit Sub
            End If
            If (JSNF1 > KSNF2 And KSNF2 <> 0) Or （JSNF2 > KSNF3 And KSNF3 <> 0) Or （JSNF3 > KSNF4 And KSNF4 <> 0) Or （JSNF4 > KSNF5 And KSNF5 <> 0) Then
                MsgBox("输入的后一个开始年份不可以小于上一个结束年份，请重新输入！")
                Exit Sub
            End If
            If （JSNF1 = KSNF2 And JSNF1 <> 0 And KSNF2 <> 0） Or （JSNF2 = KSNF3 And JSNF2 <> 0 And KSNF3 <> 0） Or （JSNF3 = KSNF4 And JSNF3 <> 0 And KSNF4 <> 0） Or （JSNF4 = KSNF5 And JSNF4 <> 0 And KSNF5 <> 0） Then
                MsgBox("输入的后一个开始年份不可以等于上一个结束年份，请重新输入！")
                Exit Sub
            End If
            If KSNF2 - JSNF1 > 1 Or KSNF3 - JSNF2 > 1 Or KSNF4 - JSNF3 > 1 Or KSNF5 - JSNF4 > 1 Then
                MsgBox("输入的后一个开始年份不可以大于上一个结束年份+1，请重新输入！")
                Exit Sub
            End If
            If KSNF1 <> JSKSNF And KSNF1 <> 0 Then
                MsgBox("输入的第一个开始年份必需为有收入的第一个年份，请重新输入！")
                Exit Sub
            End If
            If ZDJSNF < jsnx Then
                MsgBox("输入的结束年份均小于项目计算年限，请重新输入！")
                Exit Sub
            End If
            If ZDJSNF > jsnx Then
                MsgBox("输入的结束年份存在大于项目计算年限的情况，请重新输入！")
                Exit Sub
            End If
            '————————————————————————————————————————————————————————————————————————————————————————  
            '解锁表格
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").unProtect(Password:="wscjc")
            '清空窗体
            Me.RichTextBox1.Clear()
            For i = 3 To 33 '清空已有的修理费率，防止出错
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value = 0
            Next
            Dim ZNBHL1 As Double '输入的第1条修理费的逐年变化率
            ZNBHL1 = (JSNFFL1 - KSNFFL1) / (JSNF1 - KSNF1)
            Dim ZNBHL2 As Double '输入的第2条修理费的逐年变化率
            ZNBHL2 = (JSNFFL2 - KSNFFL2) / (JSNF2 - KSNF2)
            Dim ZNBHL3 As Double '输入的第3条修理费的逐年变化率
            ZNBHL3 = (JSNFFL3 - KSNFFL3) / (JSNF3 - KSNF3)
            Dim ZNBHL4 As Double '输入的第4条修理费的逐年变化率
            ZNBHL4 = (JSNFFL4 - KSNFFL4) / (JSNF4 - KSNF4)
            Dim ZNBHL5 As Double '输入的第5条修理费的逐年变化率
            ZNBHL5 = (JSNFFL5 - KSNFFL5) / (JSNF5 - KSNF5)
            '计算期第一年到到KSNF1之间的年份，修理费率设置为0
            Dim js0 As Integer = 0
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value >= 1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value < KSNF1 Then
                    js0 = js0 + 1
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value = 0
                End If
            Next
            '计算输入的第1条逐年修理费变化率
            If KSNF1 > 0 And JSNF1 > 0 Then
                Dim js1 As Integer = 0
                For i = 3 To 33
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value >= KSNF1 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value <= JSNF1 Then
                        js1 = js1 + 1
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value = (KSNFFL1 + (js1 - 1) * ZNBHL1) / 100
                    End If
                Next
            End If
            '计算第2条逐年修理费变化率
            If KSNF2 > 0 And JSNF2 > 0 Then
                Dim js2 As Integer = 0
                For i = 3 To 33
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value >= KSNF2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value <= JSNF2 Then
                        js2 = js2 + 1
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value = (KSNFFL2 + (js2 - 1) * ZNBHL2) / 100
                    End If
                Next
            End If
            '计算第3条逐年修理费变化率
            If KSNF3 > 0 And JSNF3 > 0 Then
                Dim js3 As Integer = 0
                For i = 3 To 33
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value >= KSNF3 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value <= JSNF3 Then
                        js3 = js3 + 1
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value = (KSNFFL3 + (js3 - 1) * ZNBHL3) / 100
                    End If
                Next
            End If
            '计算第4条逐年修理费变化率
            If KSNF4 > 0 And JSNF4 > 0 Then
                Dim js4 As Integer = 0
                For i = 3 To 33
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value >= KSNF4 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value <= JSNF4 Then
                        js4 = js4 + 1
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value = (KSNFFL4 + (js4 - 1) * ZNBHL4) / 100
                    End If
                Next
            End If
            '计算第5条逐年修理费变化率
            If KSNF5 > 0 And JSNF5 > 0 Then
                Dim js5 As Integer = 0
                For i = 3 To 33
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value >= KSNF5 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value <= JSNF5 Then
                        js5 = js5 + 1
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value = (KSNFFL5 + (js5 - 1) * ZNBHL5) / 100
                    End If
                Next
            End If
            '超过最大结束年份的修理费设置为0
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value > ZDJSNF Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value = 0
                End If
            Next
            '————————————————————————————————————————————————————————————————————————————————————————        
            '————————————————————————————————————————————————————————————————————————————————————————        
            '暂时废弃下面的程序，直接跳过
            '读取窗体中输入的每次投资参与计算的比例
            Dim tzbl_1 As Double = CType(Me.TZBL1.Text, Double) / 100
            Dim tzbl_2 As Double = CType(Me.TZBL2.Text, Double) / 100
            Dim tzbl_3 As Double = CType(Me.TZBL3.Text, Double) / 100
            Dim tzbl_4 As Double = CType(Me.TZBL4.Text, Double) / 100
            Dim tzbl_5 As Double = CType(Me.TZBL5.Text, Double) / 100
            Dim tzbl_6 As Double = CType(Me.TZBL6.Text, Double) / 100
            Dim tzbl_7 As Double = CType(Me.TZBL7.Text, Double) / 100
            Dim tzbl_8 As Double = CType(Me.TZBL8.Text, Double) / 100
            Dim tzbl_9 As Double = CType(Me.TZBL9.Text, Double) / 100
            Dim tzbl_10 As Double = CType(Me.TZBL10.Text, Double) / 100
            If TZ1.Checked = False Then
                tzbl_1 = 0
            End If
            If TZ2.Checked = False Then
                tzbl_2 = 0
            End If
            If TZ3.Checked = False Then
                tzbl_3 = 0
            End If
            If TZ4.Checked = False Then
                tzbl_4 = 0
            End If
            If TZ5.Checked = False Then
                tzbl_5 = 0
            End If
            If TZ6.Checked = False Then
                tzbl_6 = 0
            End If
            If TZ7.Checked = False Then
                tzbl_7 = 0
            End If
            If TZ8.Checked = False Then
                tzbl_8 = 0
            End If
            If TZ9.Checked = False Then
                tzbl_9 = 0
            End If
            If TZ10.Checked = False Then
                tzbl_10 = 0
            End If
            '读取每次投资的金额
            Dim TZJE_1 As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 3).Value
            Dim TZJE_2 As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 5).Value
            Dim TZJE_3 As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 7).Value
            Dim TZJE_4 As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 9).Value
            Dim TZJE_5 As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 11).Value
            Dim TZJE_6 As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 3).Value
            Dim TZJE_7 As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 5).Value
            Dim TZJE_8 As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 7).Value
            Dim TZJE_9 As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 9).Value
            Dim TZJE_10 As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 11).Value
            '计算总投资
            Dim TZJE_ALL As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value
            '每次投资发生的年份序号
            Dim TZNF_1 As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 3).Value
            Dim TZNF_2 As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value
            Dim TZNF_3 As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value
            Dim TZNF_4 As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value
            Dim TZNF_5 As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value
            Dim TZNF_6 As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value
            Dim TZNF_7 As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value
            Dim TZNF_8 As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value
            Dim TZNF_9 As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value
            Dim TZNF_10 As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value
            '根据勾选的情况和输入的情况，确定截止到每年的累计投资额占总投资额比例（用静态投资比例虽然会有点误差，但是差距不大，忽略这部分误差）
            '投资累计金额
            Dim TZLJJE_1 As Double = TZJE_1 * tzbl_1
            Dim TZLJJE_2 As Double = TZJE_1 * tzbl_1 + TZJE_2 * tzbl_2
            Dim TZLJJE_3 As Double = TZJE_1 * tzbl_1 + TZJE_2 * tzbl_2 + TZJE_3 * tzbl_3
            Dim TZLJJE_4 As Double = TZJE_1 * tzbl_1 + TZJE_2 * tzbl_2 + TZJE_3 * tzbl_3 + TZJE_4 * tzbl_4
            Dim TZLJJE_5 As Double = TZJE_1 * tzbl_1 + TZJE_2 * tzbl_2 + TZJE_3 * tzbl_3 + TZJE_4 * tzbl_4 + TZJE_5 * tzbl_5
            Dim TZLJJE_6 As Double = TZJE_1 * tzbl_1 + TZJE_2 * tzbl_2 + TZJE_3 * tzbl_3 + TZJE_4 * tzbl_4 + TZJE_5 * tzbl_5 + TZJE_6 * tzbl_6
            Dim TZLJJE_7 As Double = TZJE_1 * tzbl_1 + TZJE_2 * tzbl_2 + TZJE_3 * tzbl_3 + TZJE_4 * tzbl_4 + TZJE_5 * tzbl_5 + TZJE_6 * tzbl_6 + TZJE_7 * tzbl_7
            Dim TZLJJE_8 As Double = TZJE_1 * tzbl_1 + TZJE_2 * tzbl_2 + TZJE_3 * tzbl_3 + TZJE_4 * tzbl_4 + TZJE_5 * tzbl_5 + TZJE_6 * tzbl_6 + TZJE_7 * tzbl_7 + TZJE_8 * tzbl_8
            Dim TZLJJE_9 As Double = TZJE_1 * tzbl_1 + TZJE_2 * tzbl_2 + TZJE_3 * tzbl_3 + TZJE_4 * tzbl_4 + TZJE_5 * tzbl_5 + TZJE_6 * tzbl_6 + TZJE_7 * tzbl_7 + TZJE_8 * tzbl_8 + TZJE_9 * tzbl_9
            Dim TZLJJE_10 As Double = TZJE_1 * tzbl_1 + TZJE_2 * tzbl_2 + TZJE_3 * tzbl_3 + TZJE_4 * tzbl_4 + TZJE_5 * tzbl_5 + TZJE_6 * tzbl_6 + TZJE_7 * tzbl_7 + TZJE_8 * tzbl_8 + TZJE_9 * tzbl_9 + TZJE_10 * tzbl_10
            '将已经写入Excel的逐年修理费率读取到数组内
            Dim ZNXLFL_a(50) As Double
            For i = 3 To 33
                ZNXLFL_a(i - 2) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value
            Next
            '根据每次投资的累计比例，将逐年修理费率进行整体折算
            '计算到的年份序号计数
            Dim NFXH As Integer = 1
            '第一次投资
            If TZNF_1 > 0 Then
                For i = 3 To 33
                    '直接算到最后一年
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value > TZNF_1 Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value = ZNXLFL_a(i - 2) * TZLJJE_1 / TZJE_1
                    End If
                Next
            End If
            '第二次投资
            If TZNF_2 > 0 Then
                For i = 3 To 33
                    '直接算到最后一年
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value > TZNF_2 Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value = ZNXLFL_a(i - 2) * TZLJJE_2 / (TZJE_1 + TZJE_2)
                    End If
                Next
            End If
            '第三次投资
            If TZNF_3 > 0 Then
                For i = 3 To 33
                    '直接算到最后一年
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value > TZNF_3 Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value = ZNXLFL_a(i - 2) * TZLJJE_3 / (TZJE_1 + TZJE_2 + TZJE_3)
                    End If
                Next
            End If
            '第四次投资
            If TZNF_4 > 0 Then
                For i = 3 To 33
                    '直接算到最后一年
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value > TZNF_4 Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value = ZNXLFL_a(i - 2) * TZLJJE_4 / (TZJE_1 + TZJE_2 + TZJE_3 + TZJE_4)
                    End If
                Next
            End If
            '第五次投资
            If TZNF_5 > 0 Then
                For i = 3 To 33
                    '直接算到最后一年
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value > TZNF_5 Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value = ZNXLFL_a(i - 2) * TZLJJE_5 / (TZJE_1 + TZJE_2 + TZJE_3 + TZJE_4 + TZJE_5)
                    End If
                Next
            End If
            '第六次投资
            If TZNF_6 > 0 Then
                For i = 3 To 33
                    '直接算到最后一年
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value > TZNF_6 Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value = ZNXLFL_a(i - 2) * TZLJJE_6 / (TZJE_1 + TZJE_2 + TZJE_3 + TZJE_4 + TZJE_5 + TZJE_6)
                    End If
                Next
            End If
            '第七次投资
            If TZNF_7 > 0 Then
                For i = 3 To 33
                    '直接算到最后一年
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value > TZNF_7 Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value = ZNXLFL_a(i - 2) * TZLJJE_7 / (TZJE_1 + TZJE_2 + TZJE_3 + TZJE_4 + TZJE_5 + TZJE_6 + TZJE_7)
                    End If
                Next
            End If
            '第八次投资
            If TZNF_8 > 0 Then
                For i = 3 To 33
                    '直接算到最后一年
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value > TZNF_8 Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value = ZNXLFL_a(i - 2) * TZLJJE_8 / (TZJE_1 + TZJE_2 + TZJE_3 + TZJE_4 + TZJE_5 + TZJE_6 + TZJE_7 + TZJE_8)
                    End If
                Next
            End If
            '第九次投资
            If TZNF_9 > 0 Then
                For i = 3 To 33
                    '直接算到最后一年
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value > TZNF_9 Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value = ZNXLFL_a(i - 2) * TZLJJE_9 / (TZJE_1 + TZJE_2 + TZJE_3 + TZJE_4 + TZJE_5 + TZJE_6 + TZJE_7 + TZJE_8 + TZJE_9)
                    End If
                Next
            End If
            '第十次投资
            If TZNF_10 > 0 Then
                For i = 3 To 33
                    '直接算到最后一年
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value > TZNF_10 Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value = ZNXLFL_a(i - 2) * TZLJJE_10 / (TZJE_1 + TZJE_2 + TZJE_3 + TZJE_4 + TZJE_5 + TZJE_6 + TZJE_7 + TZJE_8 + TZJE_9 + TZJE_10)
                    End If
                Next
            End If
            '超过计算年限的年份设置为0
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value > jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value = 0
                End If
            Next
            '————————————————————————————————————————————————————————————————————————————————————————        
            '读取并显示计算出的逐年修理费率
            Dim SJ As Double
                Dim nf
                For i = 3 To 33  '根据数组中的元素数量循环
                    nf = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value '年份序号
                    SJ = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value) * 100, 2)
                    Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "%(" & nf & ") " '输出到RichTextBox1
                Next
                Me.RichTextBox1.Text = "逐年设备修理费率：" & Me.RichTextBox1.Text
                MsgBox("修理费系数设置完成！")
            End If
    End Sub

    Private Sub 修理费率逐年变化设置_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '清空表格
        Me.ksnf1.Clear()
        Me.ksnf2.Clear()
        Me.ksnf3.Clear()
        Me.ksnf4.Clear()
        Me.ksnf5.Clear()
        Me.jsnf1.Clear()
        Me.jsnf2.Clear()
        Me.jsnf3.Clear()
        Me.jsnf4.Clear()
        Me.jsnf5.Clear()
        Me.ksfl1.Clear()
        Me.ksfl2.Clear()
        Me.ksfl3.Clear()
        Me.ksfl4.Clear()
        Me.ksfl5.Clear()
        Me.jsfl1.Clear()
        Me.jsfl2.Clear()
        Me.jsfl3.Clear()
        Me.jsfl4.Clear()
        Me.jsfl5.Clear()
        Me.RichTextBox1.Rtf = Nothing
        '载入默认值
        '各种年限系数的计算开始年份（补贴收入、销售收入和成本）
        Dim JSKSNF As Integer = 0
        For i = 3 To 33
            If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value > 0 Then
                JSKSNF = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(156, i).Value
                Exit For
            End If
        Next
        Me.ksnf1.Text = JSKSNF
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        Me.jsnf1.Text = jsnx
        Me.ksfl1.Text = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(9, 5).Value * 100
        Me.jsfl1.Text = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(9, 5).Value * 100
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '根据输入的投资情况，确定每个checkbox和投资比例输入是否可以选择和输入，并载入默认值
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 3).Value > 0 Then
            Me.TZ1.Enabled = True
            Me.TZ1.Checked = True
            Me.TZBL1.Enabled = True
            Me.TZBL1.Text = 100
        Else
            Me.TZ1.Enabled = False
            Me.TZ1.Checked = False
            Me.TZBL1.Enabled = False
            Me.TZBL1.Text = Nothing
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 5).Value > 0 Then
            Me.TZ2.Enabled = True
            Me.TZ2.Checked = True
            Me.TZBL2.Enabled = True
            Me.TZBL2.Text = 100
        Else
            Me.TZ2.Enabled = False
            Me.TZ2.Checked = False
            Me.TZBL2.Enabled = False
            Me.TZBL2.Text = Nothing
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 7).Value > 0 Then
            Me.TZ3.Enabled = True
            Me.TZ3.Checked = True
            Me.TZBL3.Enabled = True
            Me.TZBL3.Text = 100
        Else
            Me.TZ3.Enabled = False
            Me.TZ3.Checked = False
            Me.TZBL3.Enabled = False
            Me.TZBL3.Text = Nothing
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 9).Value > 0 Then
            Me.TZ4.Enabled = True
            Me.TZ4.Checked = True
            Me.TZBL4.Enabled = True
            Me.TZBL4.Text = 100
        Else
            Me.TZ4.Enabled = False
            Me.TZ4.Checked = False
            Me.TZBL4.Enabled = False
            Me.TZBL4.Text = Nothing
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 11).Value > 0 Then
            Me.TZ5.Enabled = True
            Me.TZ5.Checked = True
            Me.TZBL5.Enabled = True
            Me.TZBL5.Text = 100
        Else
            Me.TZ5.Enabled = False
            Me.TZ5.Checked = False
            Me.TZBL5.Enabled = False
            Me.TZBL5.Text = Nothing
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 3).Value > 0 Then
            Me.TZ6.Enabled = True
            Me.TZ6.Checked = True
            Me.TZBL6.Enabled = True
            Me.TZBL6.Text = 100
        Else
            Me.TZ6.Enabled = False
            Me.TZ6.Checked = False
            Me.TZBL6.Enabled = False
            Me.TZBL6.Text = Nothing
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 5).Value > 0 Then
            Me.TZ7.Enabled = True
            Me.TZ7.Checked = True
            Me.TZBL7.Enabled = True
            Me.TZBL7.Text = 100
        Else
            Me.TZ7.Enabled = False
            Me.TZ7.Checked = False
            Me.TZBL7.Enabled = False
            Me.TZBL7.Text = Nothing
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 7).Value > 0 Then
            Me.TZ8.Enabled = True
            Me.TZ8.Checked = True
            Me.TZBL8.Enabled = True
            Me.TZBL8.Text = 100
        Else
            Me.TZ8.Enabled = False
            Me.TZ8.Checked = False
            Me.TZBL8.Enabled = False
            Me.TZBL8.Text = Nothing
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 9).Value > 0 Then
            Me.TZ9.Enabled = True
            Me.TZ9.Checked = True
            Me.TZBL9.Enabled = True
            Me.TZBL9.Text = 100
        Else
            Me.TZ9.Enabled = False
            Me.TZ9.Checked = False
            Me.TZBL9.Enabled = False
            Me.TZBL9.Text = Nothing
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(73, 11).Value > 0 Then
            Me.TZ10.Enabled = True
            Me.TZ10.Checked = True
            Me.TZBL10.Enabled = True
            Me.TZBL10.Text = 100
        Else
            Me.TZ10.Enabled = False
            Me.TZ10.Checked = False
            Me.TZBL10.Enabled = False
            Me.TZBL10.Text = Nothing
        End If
    End Sub
End Class