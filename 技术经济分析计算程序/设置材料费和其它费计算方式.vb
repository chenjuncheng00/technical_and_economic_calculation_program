Public Class 设置材料费和其它费计算方式
    Function 逐年材料费率和其它费率计算_base(ExcelApp As Object)
        On Error Resume Next
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
        For i = 1 To 15
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(5, 4 + i).Value > 0 Then
                JSKSNF = i
                Exit For
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '检查输入的开始年份和结束年份
        If KSNF1 > JSNF1 Or KSNF2 > JSNF2 Or KSNF3 > JSNF3 Or KSNF4 > JSNF4 Or KSNF5 > JSNF5 Then
            MsgBox("输入的开始年份不可以大于结束年份，请重新输入！")
            Exit Function
        End If
        If (JSNF1 > KSNF2 And KSNF2 <> 0) Or （JSNF2 > KSNF3 And KSNF3 <> 0) Or （JSNF3 > KSNF4 And KSNF4 <> 0) Or （JSNF4 > KSNF5 And KSNF5 <> 0) Then
            MsgBox("输入的后一个开始年份不可以小于上一个结束年份，请重新输入！")
            Exit Function
        End If
        If （JSNF1 = KSNF2 And JSNF1 <> 0 And KSNF2 <> 0） Or （JSNF2 = KSNF3 And JSNF2 <> 0 And KSNF3 <> 0） Or （JSNF3 = KSNF4 And JSNF3 <> 0 And KSNF4 <> 0） Or （JSNF4 = KSNF5 And JSNF4 <> 0 And KSNF5 <> 0） Then
            MsgBox("输入的后一个开始年份不可以等于上一个结束年份，请重新输入！")
            Exit Function
        End If
        If KSNF2 - JSNF1 > 1 Or KSNF3 - JSNF2 > 1 Or KSNF4 - JSNF3 > 1 Or KSNF5 - JSNF4 > 1 Then
            MsgBox("输入的后一个开始年份不可以大于上一个结束年份+1，请重新输入！")
            Exit Function
        End If
        If KSNF1 <> JSKSNF And KSNF1 <> 0 Then
            MsgBox("输入的第一个开始年份必需为有收入的第一个年份，请重新输入！")
            Exit Function
        End If
        If ZDJSNF < jsnx Then
            MsgBox("输入的结束年份均小于项目计算年限，请重新输入！")
            Exit Function
        End If
        If ZDJSNF > jsnx Then
            MsgBox("输入的结束年份存在大于项目计算年限的情况，请重新输入！")
            Exit Function
        End If
        '————————————————————————————————————————————————————————————————————————————————————————   
        Dim ZNBHL1 As Double '输入的第1条材料费和其它费的逐年变化率
        If JSNF1 - KSNF1 = 0 Then
            ZNBHL1 = 0
        Else
            ZNBHL1 = (JSNFFL1 - KSNFFL1) / (JSNF1 - KSNF1)
        End If
        Dim ZNBHL2 As Double '输入的第2条材料费和其它费的逐年变化率
        If JSNF2 - KSNF2 = 0 Then
            ZNBHL2 = 0
        Else
            ZNBHL2 = (JSNFFL2 - KSNFFL2) / (JSNF2 - KSNF2)
        End If
        Dim ZNBHL3 As Double '输入的第3条材料费和其它费的逐年变化率
        If JSNF3 - KSNF3 = 0 Then
            ZNBHL3 = 0
        Else
            ZNBHL3 = (JSNFFL3 - KSNFFL3) / (JSNF3 - KSNF3)
        End If
        Dim ZNBHL4 As Double '输入的第4条材料费和其它费的逐年变化率
        If JSNF4 - KSNF4 = 0 Then
            ZNBHL4 = 0
        Else
            ZNBHL4 = (JSNFFL4 - KSNFFL4) / (JSNF4 - KSNF4)
        End If
        Dim ZNBHL5 As Double '输入的第5条材料费和其它费的逐年变化率
        If JSNF5 - KSNF5 = 0 Then
            ZNBHL5 = 0
        Else
            ZNBHL5 = (JSNFFL5 - KSNFFL5) / (JSNF5 - KSNF5)
        End If
        '设备材料费和其它费率列表
        Dim clfl_qtfl_list(31) As Double
        '计算期第一年到到KSNF1之间的年份，材料费和其它费率设置为0
        Dim js0 As Integer = 0
        For i = 1 To 31
            If i >= 1 And i < KSNF1 Then
                js0 = js0 + 1
                clfl_qtfl_list(i) = 0
            End If
        Next
        '计算输入的第1条逐年材料费和其它费变化率
        If KSNF1 > 0 And JSNF1 > 0 Then
            Dim js1 As Integer = 0
            For i = 1 To 31
                If i >= KSNF1 And i <= JSNF1 Then
                    js1 = js1 + 1
                    clfl_qtfl_list(i) = (KSNFFL1 + (js1 - 1) * ZNBHL1)
                End If
            Next
        End If
        '计算第2条逐年材料费和其它费变化率
        If KSNF2 > 0 And JSNF2 > 0 Then
            Dim js2 As Integer = 0
            For i = 1 To 31
                If i >= KSNF2 And i <= JSNF2 Then
                    js2 = js2 + 1
                    clfl_qtfl_list(i) = (KSNFFL2 + (js2 - 1) * ZNBHL2)
                End If
            Next
        End If
        '计算第3条逐年材料费和其它费变化率
        If KSNF3 > 0 And JSNF3 > 0 Then
            Dim js3 As Integer = 0
            For i = 1 To 31
                If i >= KSNF3 And i <= JSNF3 Then
                    js3 = js3 + 1
                    clfl_qtfl_list(i) = (KSNFFL3 + (js3 - 1) * ZNBHL3)
                End If
            Next
        End If
        '计算第4条逐年材料费和其它费变化率
        If KSNF4 > 0 And JSNF4 > 0 Then
            Dim js4 As Integer = 0
            For i = 1 To 31
                If i >= KSNF4 And i <= JSNF4 Then
                    js4 = js4 + 1
                    clfl_qtfl_list(i) = (KSNFFL4 + (js4 - 1) * ZNBHL4)
                End If
            Next
        End If
        '计算第5条逐年材料费和其它费变化率
        If KSNF5 > 0 And JSNF5 > 0 Then
            Dim js5 As Integer = 0
            For i = 1 To 31
                If i >= KSNF5 And i <= JSNF5 Then
                    js5 = js5 + 1
                    clfl_qtfl_list(i) = (KSNFFL5 + (js5 - 1) * ZNBHL5)
                End If
            Next
        End If
        '超过最大结束年份的材料费和其它费设置为0
        For i = 1 To 31
            If i > ZDJSNF Then
                clfl_qtfl_list(i) = 0
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————     
        '返回计算结果
        Return clfl_qtfl_list
    End Function
    Private Sub 设置材料费和其它费计算方式_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        '各种年限系数的计算开始年份
        Dim JSKSNF As Integer = 0
        For i = 1 To 15
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(5, 4 + i).Value > 0 Then
                JSKSNF = i
                Exit For
            End If
        Next
        Me.ksnf1.Text = JSKSNF
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        Me.jsnf1.Text = jsnx
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '根据输入的投资情况，确定每个checkbox和投资比例输入是否可以选择和输入，并载入默认值
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(45, 1).Value > 0 Then
            Me.ranji.Enabled = True
            Me.ranji.Checked = False
            Me.YYNX_RJ.Enabled = True
            Me.YYNX_RJ.Text = jsnx - 1
            Me.KCBL_RJ.Enabled = True
            Me.KCBL_RJ.Text = 100
        Else
            Me.ranji.Enabled = False
            Me.ranji.Checked = False
            Me.YYNX_RJ.Enabled = False
            Me.YYNX_RJ.Text = Nothing
            Me.KCBL_RJ.Enabled = False
            Me.KCBL_RJ.Text = Nothing
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(52, 1).Value > 0 Then
            Me.ranmei.Enabled = True
            Me.ranmei.Checked = False
            Me.YYNX_RM.Enabled = True
            Me.YYNX_RM.Text = jsnx - 1
            Me.KCBL_RM.Enabled = True
            Me.KCBL_RM.Text = 100
        Else
            Me.ranmei.Enabled = False
            Me.ranmei.Checked = False
            Me.YYNX_RM.Enabled = False
            Me.YYNX_RM.Text = Nothing
            Me.KCBL_RM.Enabled = False
            Me.KCBL_RM.Text = Nothing
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(49, 1).Value > 0 Then
            Me.nuantong.Enabled = True
            Me.nuantong.Checked = False
            Me.YYNX_NT.Enabled = True
            Me.YYNX_NT.Text = jsnx - 1
            Me.KCBL_NT.Enabled = True
            Me.KCBL_NT.Text = 100
        Else
            Me.nuantong.Enabled = False
            Me.nuantong.Checked = False
            Me.YYNX_NT.Enabled = False
            Me.YYNX_NT.Text = Nothing
            Me.KCBL_NT.Enabled = False
            Me.KCBL_NT.Text = Nothing
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(51, 1).Value > 0 Then
            Me.guangfu.Enabled = True
            Me.guangfu.Checked = False
            Me.YYNX_GF.Enabled = True
            Me.YYNX_GF.Text = jsnx - 1
            Me.KCBL_GF.Enabled = True
            Me.KCBL_GF.Text = 100
        Else
            Me.guangfu.Enabled = False
            Me.guangfu.Checked = False
            Me.YYNX_GF.Enabled = False
            Me.YYNX_GF.Text = Nothing
            Me.KCBL_GF.Enabled = False
            Me.KCBL_GF.Text = Nothing
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(54, 1).Value > 0 Then
            Me.fengdian.Enabled = True
            Me.fengdian.Checked = False
            Me.YYNX_FD.Enabled = True
            Me.YYNX_FD.Text = jsnx - 1
            Me.KCBL_FD.Enabled = True
            Me.KCBL_FD.Text = 100
        Else
            Me.fengdian.Enabled = False
            Me.fengdian.Checked = False
            Me.YYNX_FD.Enabled = False
            Me.YYNX_FD.Text = Nothing
            Me.KCBL_FD.Enabled = False
            Me.KCBL_FD.Text = Nothing
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(55, 1).Value > 0 Then
            Me.lajifadian.Enabled = True
            Me.lajifadian.Checked = False
            Me.YYNX_LJFD.Enabled = True
            Me.YYNX_LJFD.Text = jsnx - 1
            Me.KCBL_LJFD.Enabled = True
            Me.KCBL_LJFD.Text = 100
        Else
            Me.lajifadian.Enabled = False
            Me.lajifadian.Checked = False
            Me.YYNX_LJFD.Enabled = False
            Me.YYNX_LJFD.Text = Nothing
            Me.KCBL_LJFD.Enabled = False
            Me.KCBL_LJFD.Text = Nothing
        End If
    End Sub

    Private Sub 设置剔除_Click(sender As Object, e As EventArgs) Handles 设置剔除.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '燃机
        If ranji.Checked = True Then
            '运营年限
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(106, 3).Value = CInt(YYNX_RJ.Text)
            '扣除比例
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(107, 3).Value = CDbl(KCBL_RJ.Text) / 100
            '显示
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<燃机设备>材料费和其它费计算参数设置写入完成！"
        End If
        ''蓄电池
        'If xudianchi.Checked = True Then
        '    '运营年限
        '    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(106, 4).Value = CInt(YYNX_XDC.Text)
        '    '扣除比例
        '    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(107, 4).Value = CDbl(KCBL_XDC.Text) / 100
        '    '显示
        '    Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<蓄电池设备>材料费和其它费计算参数设置写入完成！"
        'End If
        '暖通
        If nuantong.Checked = True Then
            '运营年限
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(106, 5).Value = CInt(YYNX_NT.Text)
            '扣除比例
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(107, 5).Value = CDbl(KCBL_NT.Text) / 100
            '显示
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<暖通设备>材料费和其它费计算参数设置写入完成！"
        End If
        '光伏
        If guangfu.Checked = True Then
            '运营年限
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(106, 6).Value = CInt(YYNX_GF.Text)
            '扣除比例
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(107, 6).Value = CDbl(KCBL_GF.Text) / 100
            '显示
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<光伏设备>材料费和其它费计算参数设置写入完成！"
        End If
        '燃煤
        If ranmei.Checked = True Then
            '运营年限
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(106, 7).Value = CInt(YYNX_RM.Text)
            '扣除比例
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(107, 7).Value = CDbl(KCBL_RM.Text) / 100
            '显示
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<燃煤设备>材料费和其它费计算参数设置写入完成！"
        End If
        '风电
        If fengdian.Checked = True Then
            '运营年限
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(106, 8).Value = CInt(YYNX_FD.Text)
            '扣除比例
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(107, 8).Value = CDbl(KCBL_FD.Text) / 100
            '显示
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<风电设备>材料费和其它费计算参数设置写入完成！"
        End If
        '垃圾发电设备
        If lajifadian.Checked = True Then
            '运营年限
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(106, 9).Value = CInt(YYNX_LJFD.Text)
            '扣除比例
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(107, 9).Value = CDbl(KCBL_LJFD.Text) / 100
            '显示
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<垃圾发电设备>材料费和其它费计算参数设置写入完成！"
        End If
        '————————————————————————————————————————————————————————————————————————————————————————        
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub

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
            Me.ranji.Checked = False
            Me.ranmei.Checked = False
            Me.nuantong.Checked = False
            Me.guangfu.Checked = False
            Me.fengdian.Checked = False
            Me.lajifadian.Checked = False
            Me.YYNX_RJ.Clear()
            Me.YYNX_RM.Clear()
            Me.YYNX_NT.Clear()
            Me.YYNX_GF.Clear()
            Me.YYNX_FD.Clear()
            Me.YYNX_LJFD.Clear()
            Me.KCBL_RJ.Clear()
            Me.KCBL_RM.Clear()
            Me.KCBL_NT.Clear()
            Me.KCBL_GF.Clear()
            Me.KCBL_FD.Clear()
            Me.KCBL_LJFD.Clear()
        End If
    End Sub

    Private Sub 光伏_Click(sender As Object, e As EventArgs) Handles 光伏.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '清空窗体
        Me.RichTextBox1.Clear()
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '计算逐年材料费和其它费率，光伏设备只可以以投资百分比计算，选择无效
        Dim clfl_qtfl_gf_list = 逐年材料费率和其它费率计算_base(ExcelApp)
        '———————————————————————————————————————————————————————————————————————————————————————— 
        If ComboBox1.Text = "材料费率" Then
            '清空已有的数据，防止出错
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(186, i).Value = 0
            Next
            '———————————————————————————————————————————————————————————————————————————————————————— 
            '逐年材料费和其它费写入Excel
            For i = 3 To 33
                '设备材料费和其它费率
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(186, i).Value = clfl_qtfl_gf_list(i - 2) / 100
            Next
            '————————————————————————————————————————————————————————————————————————————————————————        
            '计算一次工作簿
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————     
            '读取并显示计算出的逐年材料费和其它费率
            Dim SJ As Double
            Dim nf
            For i = 1 To 31  '根据数组中的元素数量循环
                nf = i '年份序号
                SJ = Math.Round(clfl_qtfl_gf_list(i), 2)
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "元/kW(" & nf & ") " '输出到RichTextBox1
            Next
            Me.RichTextBox1.Text = "光伏设备逐年材料费率：" & Me.RichTextBox1.Text
        ElseIf ComboBox1.Text = "其它费率" Then
            '清空已有的数据，防止出错
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(192, i).Value = 0
            Next
            '————————————————————————————————————————————————————————————————————————————————————————    
            '逐年材料费和其它费写入Excel
            For i = 3 To 33
                '设备材料费和其它费率
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(192, i).Value = clfl_qtfl_gf_list(i - 2)
            Next
            '————————————————————————————————————————————————————————————————————————————————————————        
            '计算一次工作簿
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————     
            '读取并显示计算出的逐年材料费和其它费率
            Dim SJ As Double
            Dim nf
            For i = 1 To 31  '根据数组中的元素数量循环
                nf = i '年份序号
                SJ = Math.Round(clfl_qtfl_gf_list(i), 2)
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "元/kW(" & nf & ") " '输出到RichTextBox1
            Next
            Me.RichTextBox1.Text = "光伏设备逐年其它费率：" & Me.RichTextBox1.Text
        End If
    End Sub

    Private Sub 风电_Click(sender As Object, e As EventArgs) Handles 风电.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '清空窗体
        Me.RichTextBox1.Clear()
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '计算逐年材料费和其它费率，风电设备只可以以投资百分比计算，选择无效
        Dim clfl_qtfl_fd_list = 逐年材料费率和其它费率计算_base(ExcelApp)
        '———————————————————————————————————————————————————————————————————————————————————————— 
        If ComboBox1.Text = "材料费率" Then
            '清空已有的数据，防止出错
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(187, i).Value = 0
            Next
            '———————————————————————————————————————————————————————————————————————————————————————— 
            '逐年材料费和其它费写入Excel
            For i = 3 To 33
                '设备材料费和其它费率
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(187, i).Value = clfl_qtfl_fd_list(i - 2) / 100
            Next
            '————————————————————————————————————————————————————————————————————————————————————————        
            '计算一次工作簿
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————     
            '读取并显示计算出的逐年材料费和其它费率
            Dim SJ As Double
            Dim nf
            For i = 1 To 31  '根据数组中的元素数量循环
                nf = i '年份序号
                SJ = Math.Round(clfl_qtfl_fd_list(i), 2)
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "元/kW(" & nf & ") " '输出到RichTextBox1
            Next
            Me.RichTextBox1.Text = "风电设备逐年材料费率：" & Me.RichTextBox1.Text
        ElseIf ComboBox1.Text = "其它费率" Then
            '清空已有的数据，防止出错
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(193, i).Value = 0
            Next
            '————————————————————————————————————————————————————————————————————————————————————————    
            '逐年材料费和其它费写入Excel
            For i = 3 To 33
                '设备材料费和其它费率
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(193, i).Value = clfl_qtfl_fd_list(i - 2)
            Next
            '————————————————————————————————————————————————————————————————————————————————————————        
            '计算一次工作簿
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————     
            '读取并显示计算出的逐年材料费和其它费率
            Dim SJ As Double
            Dim nf
            For i = 1 To 31  '根据数组中的元素数量循环
                nf = i '年份序号
                SJ = Math.Round(clfl_qtfl_fd_list(i), 2)
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "元/kW(" & nf & ") " '输出到RichTextBox1
            Next
            Me.RichTextBox1.Text = "风电设备逐年其它费率：" & Me.RichTextBox1.Text
        End If
    End Sub

    Private Sub 燃机_Click(sender As Object, e As EventArgs) Handles 燃机.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '清空窗体
        Me.RichTextBox1.Clear()
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '计算逐年材料费和其它费率，燃机设备只可以以投资百分比计算，选择无效
        Dim clfl_qtfl_rj_list = 逐年材料费率和其它费率计算_base(ExcelApp)
        '———————————————————————————————————————————————————————————————————————————————————————— 
        If ComboBox1.Text = "材料费率" Then
            '清空已有的数据，防止出错
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(182, i).Value = 0
            Next
            '———————————————————————————————————————————————————————————————————————————————————————— 
            '逐年材料费和其它费写入Excel
            For i = 3 To 33
                '设备材料费和其它费率
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(182, i).Value = clfl_qtfl_rj_list(i - 2) / 100
            Next
            '————————————————————————————————————————————————————————————————————————————————————————        
            '计算一次工作簿
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————     
            '读取并显示计算出的逐年材料费和其它费率
            Dim SJ As Double
            Dim nf
            For i = 1 To 31  '根据数组中的元素数量循环
                nf = i '年份序号
                SJ = Math.Round(clfl_qtfl_rj_list(i), 2)
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "元/MWh(" & nf & ") " '输出到RichTextBox1
            Next
            Me.RichTextBox1.Text = "燃机设备逐年材料费率：" & Me.RichTextBox1.Text
        ElseIf ComboBox1.Text = "其它费率" Then
            '清空已有的数据，防止出错
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(188, i).Value = 0
            Next
            '————————————————————————————————————————————————————————————————————————————————————————    
            '逐年材料费和其它费写入Excel
            For i = 3 To 33
                '设备材料费和其它费率
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(188, i).Value = clfl_qtfl_rj_list(i - 2)
            Next
            '————————————————————————————————————————————————————————————————————————————————————————        
            '计算一次工作簿
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————     
            '读取并显示计算出的逐年材料费和其它费率
            Dim SJ As Double
            Dim nf
            For i = 1 To 31  '根据数组中的元素数量循环
                nf = i '年份序号
                SJ = Math.Round(clfl_qtfl_rj_list(i), 2)
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "元/MWh(" & nf & ") " '输出到RichTextBox1
            Next
            Me.RichTextBox1.Text = "燃机设备逐年其它费率：" & Me.RichTextBox1.Text
        End If
    End Sub

    Private Sub 燃煤_Click(sender As Object, e As EventArgs) Handles 燃煤.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '清空窗体
        Me.RichTextBox1.Clear()
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '计算逐年材料费和其它费率，燃煤设备只可以以投资百分比计算，选择无效
        Dim clfl_qtfl_rm_list = 逐年材料费率和其它费率计算_base(ExcelApp)
        '———————————————————————————————————————————————————————————————————————————————————————— 
        If ComboBox1.Text = "材料费率" Then
            '清空已有的数据，防止出错
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(183, i).Value = 0
            Next
            '———————————————————————————————————————————————————————————————————————————————————————— 
            '逐年材料费和其它费写入Excel
            For i = 3 To 33
                '设备材料费和其它费率
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(183, i).Value = clfl_qtfl_rm_list(i - 2) / 100
            Next
            '————————————————————————————————————————————————————————————————————————————————————————        
            '计算一次工作簿
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————     
            '读取并显示计算出的逐年材料费和其它费率
            Dim SJ As Double
            Dim nf
            For i = 1 To 31  '根据数组中的元素数量循环
                nf = i '年份序号
                SJ = Math.Round(clfl_qtfl_rm_list(i), 2)
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "元/MWh(" & nf & ") " '输出到RichTextBox1
            Next
            Me.RichTextBox1.Text = "燃煤设备逐年材料费率：" & Me.RichTextBox1.Text
        ElseIf ComboBox1.Text = "其它费率" Then
            '清空已有的数据，防止出错
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(189, i).Value = 0
            Next
            '————————————————————————————————————————————————————————————————————————————————————————    
            '逐年材料费和其它费写入Excel
            For i = 3 To 33
                '设备材料费和其它费率
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(189, i).Value = clfl_qtfl_rm_list(i - 2)
            Next
            '————————————————————————————————————————————————————————————————————————————————————————        
            '计算一次工作簿
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————     
            '读取并显示计算出的逐年材料费和其它费率
            Dim SJ As Double
            Dim nf
            For i = 1 To 31  '根据数组中的元素数量循环
                nf = i '年份序号
                SJ = Math.Round(clfl_qtfl_rm_list(i), 2)
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "元/MWh(" & nf & ") " '输出到RichTextBox1
            Next
            Me.RichTextBox1.Text = "燃煤设备逐年其它费率：" & Me.RichTextBox1.Text
        End If
    End Sub

    Private Sub 暖通_Click(sender As Object, e As EventArgs) Handles 暖通.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '清空窗体
        Me.RichTextBox1.Clear()
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '计算逐年材料费和其它费率，暖通设备只可以以投资百分比计算，选择无效
        Dim clfl_qtfl_nt_list = 逐年材料费率和其它费率计算_base(ExcelApp)
        '———————————————————————————————————————————————————————————————————————————————————————— 
        If ComboBox1.Text = "材料费率" Then
            '清空已有的数据，防止出错
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(185, i).Value = 0
            Next
            '———————————————————————————————————————————————————————————————————————————————————————— 
            '逐年材料费和其它费写入Excel
            For i = 3 To 33
                '设备材料费和其它费率
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(185, i).Value = clfl_qtfl_nt_list(i - 2) / 100
            Next
            '————————————————————————————————————————————————————————————————————————————————————————        
            '计算一次工作簿
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————     
            '读取并显示计算出的逐年材料费和其它费率
            Dim SJ As Double
            Dim nf
            For i = 1 To 31  '根据数组中的元素数量循环
                nf = i '年份序号
                SJ = Math.Round(clfl_qtfl_nt_list(i), 2)
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "元/MWh(" & nf & ") " '输出到RichTextBox1
            Next
            Me.RichTextBox1.Text = "暖通设备逐年材料费率：" & Me.RichTextBox1.Text
        ElseIf ComboBox1.Text = "其它费率" Then
            '清空已有的数据，防止出错
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(191, i).Value = 0
            Next
            '————————————————————————————————————————————————————————————————————————————————————————    
            '逐年材料费和其它费写入Excel
            For i = 3 To 33
                '设备材料费和其它费率
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(191, i).Value = clfl_qtfl_nt_list(i - 2)
            Next
            '————————————————————————————————————————————————————————————————————————————————————————        
            '计算一次工作簿
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————     
            '读取并显示计算出的逐年材料费和其它费率
            Dim SJ As Double
            Dim nf
            For i = 1 To 31  '根据数组中的元素数量循环
                nf = i '年份序号
                SJ = Math.Round(clfl_qtfl_nt_list(i), 2)
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "元/MWh(" & nf & ") " '输出到RichTextBox1
            Next
            Me.RichTextBox1.Text = "暖通设备逐年其它费率：" & Me.RichTextBox1.Text
        End If
    End Sub

    Private Sub 垃圾发电_Click(sender As Object, e As EventArgs) Handles 垃圾发电.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '清空窗体
        Me.RichTextBox1.Clear()
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '计算逐年材料费和其它费率，垃圾发电设备只可以以投资百分比计算，选择无效
        Dim clfl_qtfl_ljfd_list = 逐年材料费率和其它费率计算_base(ExcelApp)
        '———————————————————————————————————————————————————————————————————————————————————————— 
        If ComboBox1.Text = "材料费率" Then
            '清空已有的数据，防止出错
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(184, i).Value = 0
            Next
            '———————————————————————————————————————————————————————————————————————————————————————— 
            '逐年材料费和其它费写入Excel
            For i = 3 To 33
                '设备材料费和其它费率
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(184, i).Value = clfl_qtfl_ljfd_list(i - 2) / 100
            Next
            '————————————————————————————————————————————————————————————————————————————————————————        
            '计算一次工作簿
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————     
            '读取并显示计算出的逐年材料费和其它费率
            Dim SJ As Double
            Dim nf
            For i = 1 To 31  '根据数组中的元素数量循环
                nf = i '年份序号
                SJ = Math.Round(clfl_qtfl_ljfd_list(i), 2)
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "元/MWh(" & nf & ") " '输出到RichTextBox1
            Next
            Me.RichTextBox1.Text = "垃圾发电设备逐年材料费率：" & Me.RichTextBox1.Text
        ElseIf ComboBox1.Text = "其它费率" Then
            '清空已有的数据，防止出错
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(190, i).Value = 0
            Next
            '————————————————————————————————————————————————————————————————————————————————————————    
            '逐年材料费和其它费写入Excel
            For i = 3 To 33
                '设备材料费和其它费率
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(190, i).Value = clfl_qtfl_ljfd_list(i - 2)
            Next
            '————————————————————————————————————————————————————————————————————————————————————————        
            '计算一次工作簿
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————     
            '读取并显示计算出的逐年材料费和其它费率
            Dim SJ As Double
            Dim nf
            For i = 1 To 31  '根据数组中的元素数量循环
                nf = i '年份序号
                SJ = Math.Round(clfl_qtfl_ljfd_list(i), 2)
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "元/MWh(" & nf & ") " '输出到RichTextBox1
            Next
            Me.RichTextBox1.Text = "垃圾发电设备逐年其它费率：" & Me.RichTextBox1.Text
        End If
    End Sub

    Private Sub 计算_Click(sender As Object, e As EventArgs) Handles 计算.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否确定输入的全部材料费和其它费计算参数？", vbOKCancel)
        If XZ = vbOK Then
            '计算一次工作簿
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————    
            'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
            Dim clfl_qtfl_model As Integer = 1
            'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
            Dim kcje_clf_qtf_model As Integer = 1
            '材料费其它费计算
            Call 材料费其它费计算.材料费其它费计算(ExcelApp, clfl_qtfl_model, kcje_clf_qtf_model)
            '————————————————————————————————————————————————————————————————————————————————————————        
            Me.RichTextBox1.Text = "设备逐年材料费和其它费计算完成！"
        End If
    End Sub
End Class