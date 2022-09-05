Public Class 设置材料费和其它费计算方式
    '开始年份和结束年份的列表
    Public ksnf_list As New List(Of Integer)
    Public jsnf_list As New List(Of Integer)
    '开始年份和结束年份的费率
    Public ksfl_list As New List(Of Double)
    Public jsfl_list As New List(Of Double)
    Function 逐年材料费率和其它费率计算_base(ExcelApp As Object)
        On Error Resume Next
        '———————————————————————————————————————————————————————————————————————————————————————— 
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        Dim ZDJSNF = jsnf_list.Max
        '各种年限系数的计算开始年份（补贴收入、销售收入和成本）
        Dim JSKSNF As Integer = 0
        For i = 1 To 15
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(5, 4 + i).Value > 0 Then
                JSKSNF = i
                Exit For
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '输入的年份总数量
        Dim n_nf As Integer = ksnf_list.LongCount
        '添加报错功能
        For i = 0 To n_nf - 1
            If ksnf_list(i) <> 0 Or jsnf_list(i) <> 0 Then
                If ksnf_list(i) < 1 Or jsnf_list(i) < 1 Then
                    MsgBox("开始年份或者结束年份不可以存在小于1的情况，请重新输入！")
                    Exit Function
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
                    Exit Function
                End If
            Next
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
        '————————————————————————————————————————————————————————————————————————————————————————
        '逐年变化率
        Dim znbhl_list As New List(Of Double)
        For i = 0 To n_nf - 1
            Dim znbhl As Double
            If jsfl_list(i) - ksfl_list(i) = 0 Then
                znbhl = 0
            Else
                znbhl = (jsfl_list(i) - ksfl_list(i)) / (jsnf_list(i) - ksnf_list(i))
            End If
            znbhl_list.Add(znbhl)
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '设备材料费和其它费率列表
        Dim clfl_qtfl_list(31) As Double
        '需要计算费率的年份
        For j = 0 To n_nf - 1
            Dim js As Integer = 0
            For i = 1 To 31
                If i >= ksnf_list(j) And i <= jsnf_list(j) Then
                    js = js + 1
                    clfl_qtfl_list(i) = (ksfl_list(j) + (js - 1) * znbhl_list(j))
                End If
            Next
        Next
        '其他年份
        For Each qtnf In qtnf_list
            For i = 1 To 31
                If i = qtnf Or i > jsnx Then
                    clfl_qtfl_list(i) = 0
                End If
            Next
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
        Me.开始年份tmp.Clear()
        Me.结束年份tmp.Clear()
        Me.开始费率tmp.Clear()
        Me.结束费率tmp.Clear()
        Me.开始年份列表.Items.Clear()
        Me.结束年份列表.Items.Clear()
        Me.开始费率列表.Items.Clear()
        Me.结束费率列表.Items.Clear()
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
        Me.开始年份tmp.Text = JSKSNF
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        Me.结束年份tmp.Text = jsnx
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
            Me.YYNX_RJ.Clear()
            Me.KCBL_RJ.Enabled = False
            Me.KCBL_RJ.Clear()
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
            Me.YYNX_RM.Clear()
            Me.KCBL_RM.Enabled = False
            Me.KCBL_RM.Clear()
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
            Me.YYNX_NT.Clear()
            Me.KCBL_NT.Enabled = False
            Me.KCBL_NT.Clear()
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
            Me.YYNX_GF.Clear()
            Me.KCBL_GF.Enabled = False
            Me.KCBL_GF.Clear()
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
            Me.YYNX_FD.Clear()
            Me.KCBL_FD.Enabled = False
            Me.KCBL_FD.Clear()
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(47, 1).Value > 0 Then
            Me.xudianchi.Enabled = True
            Me.xudianchi.Checked = False
            Me.YYNX_XDC.Enabled = True
            Me.YYNX_XDC.Text = jsnx - 1
            Me.KCBL_XDC.Enabled = True
            Me.KCBL_XDC.Text = 100
        Else
            Me.xudianchi.Enabled = False
            Me.xudianchi.Checked = False
            Me.YYNX_XDC.Enabled = False
            Me.YYNX_XDC.Clear()
            Me.KCBL_XDC.Enabled = False
            Me.KCBL_XDC.Clear()
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
            Me.YYNX_LJFD.Clear()
            Me.KCBL_LJFD.Enabled = False
            Me.KCBL_LJFD.Clear()
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
        '蓄电池
        If xudianchi.Checked = True Then
            '运营年限
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(106, 4).Value = CInt(YYNX_XDC.Text)
            '扣除比例
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(107, 4).Value = CDbl(KCBL_XDC.Text) / 100
            '显示
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<蓄电池设备>材料费和其它费计算参数设置写入完成！"
        End If
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
            Me.开始年份tmp.Clear()
            Me.结束年份tmp.Clear()
            Me.开始费率tmp.Clear()
            Me.结束费率tmp.Clear()
            Me.开始年份列表.Items.Clear()
            Me.结束年份列表.Items.Clear()
            Me.开始费率列表.Items.Clear()
            Me.结束费率列表.Items.Clear()
            ksnf_list.Clear()
            jsnf_list.Clear()
            ksfl_list.Clear()
            jsfl_list.Clear()
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
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(186, i).Value = clfl_qtfl_gf_list(i - 2)
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
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(187, i).Value = clfl_qtfl_fd_list(i - 2)
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
    Private Sub 蓄电池_Click(sender As Object, e As EventArgs) Handles 蓄电池.Click
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
        Dim clfl_qtfl_xdc_list = 逐年材料费率和其它费率计算_base(ExcelApp)
        '———————————————————————————————————————————————————————————————————————————————————————— 
        If ComboBox1.Text = "材料费率" Then
            '清空已有的数据，防止出错
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(194, i).Value = 0
            Next
            '———————————————————————————————————————————————————————————————————————————————————————— 
            '逐年材料费和其它费写入Excel
            For i = 3 To 33
                '设备材料费和其它费率
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(194, i).Value = clfl_qtfl_xdc_list(i - 2)
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
                SJ = Math.Round(clfl_qtfl_xdc_list(i), 2)
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "元/kW(" & nf & ") " '输出到RichTextBox1
            Next
            Me.RichTextBox1.Text = "蓄电池设备逐年材料费率：" & Me.RichTextBox1.Text
        ElseIf ComboBox1.Text = "其它费率" Then
            '清空已有的数据，防止出错
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(195, i).Value = 0
            Next
            '————————————————————————————————————————————————————————————————————————————————————————    
            '逐年材料费和其它费写入Excel
            For i = 3 To 33
                '设备材料费和其它费率
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(195, i).Value = clfl_qtfl_xdc_list(i - 2)
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
                SJ = Math.Round(clfl_qtfl_xdc_list(i), 2)
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "元/kW(" & nf & ") " '输出到RichTextBox1
            Next
            Me.RichTextBox1.Text = "蓄电池设备逐年其它费率：" & Me.RichTextBox1.Text
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
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(182, i).Value = clfl_qtfl_rj_list(i - 2)
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
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(183, i).Value = clfl_qtfl_rm_list(i - 2)
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
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(185, i).Value = clfl_qtfl_nt_list(i - 2)
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
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(184, i).Value = clfl_qtfl_ljfd_list(i - 2)
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

    Private Sub 添加输入_Click(sender As Object, e As EventArgs) Handles 添加输入.Click
        On Error Resume Next
        '判断1
        If 开始年份tmp.Text = Nothing Or 结束年份tmp.Text = Nothing Then
            MsgBox("输入的开始年份和结束年份都必须不能为空！，请重新输入")
            Exit Sub
        End If
        If 开始费率tmp.Text = Nothing Or 结束费率tmp.Text = Nothing Then
            MsgBox("输入的开始年份费率和结束年份费率都必须不能为空！，请重新输入")
            Exit Sub
        End If
        '开始年份
        Dim ksnf_text As String = 开始年份tmp.Text
        Dim ksnf As Integer = CType(ksnf_text, Integer)
        '结束年份
        Dim jsnf_text As String = 结束年份tmp.Text
        Dim jsnf As Integer = CType(jsnf_text, Integer)
        '开始年份费率
        Dim ksfl_text As String = 开始费率tmp.Text
        Dim ksfl As Double = CType(ksfl_text, Double)
        '结束年份费率
        Dim jsfl_text As String = 结束费率tmp.Text
        Dim jsfl As Double = CType(jsfl_text, Double)
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
        开始费率列表.Items.Add(ksfl_text)
        ksfl_list.Add(ksfl)
        结束费率列表.Items.Add(jsfl_text)
        jsfl_list.Add(jsfl)
    End Sub

    Private Sub 清空输入_Click(sender As Object, e As EventArgs) Handles 清空输入.Click
        On Error Resume Next
        Me.开始年份tmp.Clear()
        Me.结束年份tmp.Clear()
        Me.开始费率tmp.Clear()
        Me.结束费率tmp.Clear()
        Me.开始年份列表.Items.Clear()
        Me.结束年份列表.Items.Clear()
        Me.开始费率列表.Items.Clear()
        Me.结束费率列表.Items.Clear()
        ksnf_list.Clear()
        jsnf_list.Clear()
        ksfl_list.Clear()
        jsfl_list.Clear()
    End Sub
End Class