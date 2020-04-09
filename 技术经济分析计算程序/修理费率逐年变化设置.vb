Public Class 修理费率逐年变化设置
    Private Sub 清空窗体_Click(sender As Object, e As EventArgs) Handles 清空窗体.Click
        Dim XZ = MsgBox("是否清空窗体中的全部内容？", vbOKCancel)
        If XZ = vbOK Then
            Com技术经济分析计算程序.Form4.ksnf1.Clear()
            Com技术经济分析计算程序.Form4.ksnf2.Clear()
            Com技术经济分析计算程序.Form4.ksnf3.Clear()
            Com技术经济分析计算程序.Form4.ksnf4.Clear()
            Com技术经济分析计算程序.Form4.ksnf5.Clear()
            Com技术经济分析计算程序.Form4.jsnf1.Clear()
            Com技术经济分析计算程序.Form4.jsnf2.Clear()
            Com技术经济分析计算程序.Form4.jsnf3.Clear()
            Com技术经济分析计算程序.Form4.jsnf4.Clear()
            Com技术经济分析计算程序.Form4.jsnf5.Clear()
            Com技术经济分析计算程序.Form4.ksfl1.Clear()
            Com技术经济分析计算程序.Form4.ksfl2.Clear()
            Com技术经济分析计算程序.Form4.ksfl3.Clear()
            Com技术经济分析计算程序.Form4.ksfl4.Clear()
            Com技术经济分析计算程序.Form4.ksfl5.Clear()
            Com技术经济分析计算程序.Form4.jsfl1.Clear()
            Com技术经济分析计算程序.Form4.jsfl2.Clear()
            Com技术经济分析计算程序.Form4.jsfl3.Clear()
            Com技术经济分析计算程序.Form4.jsfl4.Clear()
            Com技术经济分析计算程序.Form4.jsfl5.Clear()
            Com技术经济分析计算程序.Form4.RichTextBox1.Rtf = Nothing
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
        KSNF1 = CType(Com技术经济分析计算程序.Form4.ksnf1.Text, Integer)
        KSNF2 = CType(Com技术经济分析计算程序.Form4.ksnf2.Text, Integer)
        KSNF3 = CType(Com技术经济分析计算程序.Form4.ksnf3.Text, Integer)
        KSNF4 = CType(Com技术经济分析计算程序.Form4.ksnf4.Text, Integer)
        KSNF5 = CType(Com技术经济分析计算程序.Form4.ksnf5.Text, Integer)
        JSNF1 = CType(Com技术经济分析计算程序.Form4.jsnf1.Text, Integer)
        JSNF2 = CType(Com技术经济分析计算程序.Form4.jsnf2.Text, Integer)
        JSNF3 = CType(Com技术经济分析计算程序.Form4.jsnf3.Text, Integer)
        JSNF4 = CType(Com技术经济分析计算程序.Form4.jsnf4.Text, Integer)
        JSNF5 = CType(Com技术经济分析计算程序.Form4.jsnf5.Text, Integer)
        KSNFFL1 = CType(Com技术经济分析计算程序.Form4.ksfl1.Text, Double)
        KSNFFL2 = CType(Com技术经济分析计算程序.Form4.ksfl2.Text, Double)
        KSNFFL3 = CType(Com技术经济分析计算程序.Form4.ksfl3.Text, Double)
        KSNFFL4 = CType(Com技术经济分析计算程序.Form4.ksfl4.Text, Double)
        KSNFFL5 = CType(Com技术经济分析计算程序.Form4.ksfl5.Text, Double)
        JSNFFL1 = CType(Com技术经济分析计算程序.Form4.jsfl1.Text, Double)
        JSNFFL2 = CType(Com技术经济分析计算程序.Form4.jsfl2.Text, Double)
        JSNFFL3 = CType(Com技术经济分析计算程序.Form4.jsfl3.Text, Double)
        JSNFFL4 = CType(Com技术经济分析计算程序.Form4.jsfl4.Text, Double)
        JSNFFL5 = CType(Com技术经济分析计算程序.Form4.jsfl5.Text, Double)
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
            Com技术经济分析计算程序.Form4.RichTextBox1.Clear()
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
            '读取并显示计算出的逐年修理费率
            Dim SJ As Double
            Dim nf
            For i = 3 To 33  '根据数组中的元素数量循环
                nf = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(158, i).Value '年份序号
                SJ = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(159, i).Value) * 100, 2)
                Com技术经济分析计算程序.Form4.RichTextBox1.Text = Com技术经济分析计算程序.Form4.RichTextBox1.Text & SJ & "%(" & nf & ") " '输出到RichTextBox1
            Next
            Com技术经济分析计算程序.Form4.RichTextBox1.Text = "逐年设备修理费率：" & Com技术经济分析计算程序.Form4.RichTextBox1.Text
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
        Com技术经济分析计算程序.Form4.ksnf1.Clear()
        Com技术经济分析计算程序.Form4.ksnf2.Clear()
        Com技术经济分析计算程序.Form4.ksnf3.Clear()
        Com技术经济分析计算程序.Form4.ksnf4.Clear()
        Com技术经济分析计算程序.Form4.ksnf5.Clear()
        Com技术经济分析计算程序.Form4.jsnf1.Clear()
        Com技术经济分析计算程序.Form4.jsnf2.Clear()
        Com技术经济分析计算程序.Form4.jsnf3.Clear()
        Com技术经济分析计算程序.Form4.jsnf4.Clear()
        Com技术经济分析计算程序.Form4.jsnf5.Clear()
        Com技术经济分析计算程序.Form4.ksfl1.Clear()
        Com技术经济分析计算程序.Form4.ksfl2.Clear()
        Com技术经济分析计算程序.Form4.ksfl3.Clear()
        Com技术经济分析计算程序.Form4.ksfl4.Clear()
        Com技术经济分析计算程序.Form4.ksfl5.Clear()
        Com技术经济分析计算程序.Form4.jsfl1.Clear()
        Com技术经济分析计算程序.Form4.jsfl2.Clear()
        Com技术经济分析计算程序.Form4.jsfl3.Clear()
        Com技术经济分析计算程序.Form4.jsfl4.Clear()
        Com技术经济分析计算程序.Form4.jsfl5.Clear()
        Com技术经济分析计算程序.Form4.RichTextBox1.Rtf = Nothing
        '载入默认值
        '各种年限系数的计算开始年份（补贴收入、销售收入和成本）
        Dim JSKSNF As Integer = 0
        For i = 3 To 33
            If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value > 0 Then
                JSKSNF = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(156, i).Value
                Exit For
            End If
        Next
        Com技术经济分析计算程序.Form4.ksnf1.Text = JSKSNF
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        Com技术经济分析计算程序.Form4.jsnf1.Text = jsnx
        Com技术经济分析计算程序.Form4.ksfl1.Text = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(17, 7).Value * 100
        Com技术经济分析计算程序.Form4.jsfl1.Text = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(17, 7).Value * 100

    End Sub
End Class