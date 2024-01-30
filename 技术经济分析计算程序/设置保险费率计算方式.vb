Public Class 设置保险费率计算方式
    '开始年份和结束年份的列表
    Public ksnf_list As New List(Of Integer)
    Public jsnf_list As New List(Of Integer)
    '开始年份和结束年份的费率
    Public ksfl_list As New List(Of Double)
    Public jsfl_list As New List(Of Double)
    Private Sub 设置保险费率计算方式_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        Me.开始费率tmp.Text = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 5).Value * 100
        Me.结束费率tmp.Text = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 5).Value * 100
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '根据输入的投资情况，确定每个checkbox和投资比例输入是否可以选择和输入，并载入默认值
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(44, 1).Value > 0 Then
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
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(46, 1).Value > 0 Then
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
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(48, 1).Value > 0 Then
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
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(50, 1).Value > 0 Then
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
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(53, 1).Value > 0 Then
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
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
            Me.changgui.Enabled = True
            Me.changgui.Checked = False
            Me.YYNX_CG.Enabled = True
            Me.YYNX_CG.Text = jsnx - 1
            Me.KCBL_CG.Enabled = True
            Me.KCBL_CG.Text = 100
        Else
            Me.changgui.Enabled = False
            Me.changgui.Checked = False
            Me.YYNX_CG.Enabled = False
            Me.YYNX_CG.Clear()
            Me.KCBL_CG.Enabled = False
            Me.KCBL_CG.Clear()
        End If
    End Sub
    Private Sub 添加输入_Click(sender As Object, e As EventArgs) Handles 添加输入.Click
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
            Me.xudianchi.Checked = False
            Me.nuantong.Checked = False
            Me.guangfu.Checked = False
            Me.fengdian.Checked = False
            Me.changgui.Checked = False
            Me.YYNX_RJ.Clear()
            Me.YYNX_XDC.Clear()
            Me.YYNX_NT.Clear()
            Me.YYNX_GF.Clear()
            Me.YYNX_FD.Clear()
            Me.YYNX_CG.Clear()
            Me.KCBL_RJ.Clear()
            Me.KCBL_XDC.Clear()
            Me.KCBL_NT.Clear()
            Me.KCBL_GF.Clear()
            Me.KCBL_FD.Clear()
            Me.KCBL_CG.Clear()
        End If
    End Sub
    Private Sub 常规设备_Click(sender As Object, e As EventArgs) Handles 常规设备.Click
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '清空已有的数据，防止出错
        For i = 3 To 33
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(140, i).Value = 0
        Next
        '清空窗体
        Me.RichTextBox1.Clear()
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '计算逐年保险费率，常规设备只可以以投资百分比计算，选择无效
        Dim bxfl_cg_list = 逐年费率计算_base(ExcelApp, ksnf_list, jsnf_list, ksfl_list, jsfl_list)
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '逐年常规保险费写入Excel
        For i = 3 To 33
            '常规设备保险费率（%）
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(140, i).Value = bxfl_cg_list(i - 2) / 100
        Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '计算一次工作簿
        ExcelApp.Calculate()
        '————————————————————————————————————————————————————————————————————————————————————————     
        '读取并显示计算出的逐年保险费率
        Dim SJ As Double
        Dim nf
        For i = 1 To 31  '根据数组中的元素数量循环
            nf = i '年份序号
            SJ = Math.Round(bxfl_cg_list(i), 2)
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "%(" & nf & ") " '输出到RichTextBox1
        Next
        Me.RichTextBox1.Text = "常规设备逐年保险费率：" & Me.RichTextBox1.Text
    End Sub
    Private Sub 燃机_Click(sender As Object, e As EventArgs) Handles 燃机.Click
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '清空已有的数据，防止出错
        For i = 3 To 33
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(135, i).Value = 0
        Next
        '清空窗体
        Me.RichTextBox1.Clear()
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '计算逐年保险费率，燃机设备只可以以投资百分比计算，选择无效
        Dim bxfl_rj_list = 逐年费率计算_base(ExcelApp, ksnf_list, jsnf_list, ksfl_list, jsfl_list)
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '逐年保险费写入Excel
        For i = 3 To 33
            '设备保险费率
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(135, i).Value = bxfl_rj_list(i - 2) / 100
        Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '计算一次工作簿
        ExcelApp.Calculate()
        '————————————————————————————————————————————————————————————————————————————————————————     
        '读取并显示计算出的逐年保险费率
        Dim SJ As Double
        Dim nf
        For i = 1 To 31  '根据数组中的元素数量循环
            nf = i '年份序号
            SJ = Math.Round(bxfl_rj_list(i), 2)
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "%(" & nf & ") " '输出到RichTextBox1
        Next
        Me.RichTextBox1.Text = "燃机设备逐年保险费率：" & Me.RichTextBox1.Text
    End Sub

    Private Sub 蓄电池_Click(sender As Object, e As EventArgs) Handles 蓄电池.Click
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '清空已有的数据，防止出错
        For i = 3 To 33
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(136, i).Value = 0
        Next
        '清空窗体
        Me.RichTextBox1.Clear()
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '计算逐年保险费率，蓄电池设备只可以以投资百分比计算，选择无效
        Dim bxfl_xdc_list = 逐年费率计算_base(ExcelApp, ksnf_list, jsnf_list, ksfl_list, jsfl_list)
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '逐年保险费写入Excel
        For i = 3 To 33
            '设备保险费率
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(136, i).Value = bxfl_xdc_list(i - 2) / 100
        Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '计算一次工作簿
        ExcelApp.Calculate()
        '————————————————————————————————————————————————————————————————————————————————————————     
        '读取并显示计算出的逐年保险费率
        Dim SJ As Double
        Dim nf
        For i = 1 To 31  '根据数组中的元素数量循环
            nf = i '年份序号
            SJ = Math.Round(bxfl_xdc_list(i), 2)
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "%(" & nf & ") " '输出到RichTextBox1
        Next
        Me.RichTextBox1.Text = "蓄电池设备逐年保险费率：" & Me.RichTextBox1.Text
    End Sub
    Private Sub 暖通_Click(sender As Object, e As EventArgs) Handles 暖通.Click
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '清空已有的数据，防止出错
        For i = 3 To 33
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(138, i).Value = 0
        Next
        '清空窗体
        Me.RichTextBox1.Clear()
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '计算逐年保险费率，暖通设备只可以以投资百分比计算，选择无效
        Dim bxfl_nt_list = 逐年费率计算_base(ExcelApp, ksnf_list, jsnf_list, ksfl_list, jsfl_list)
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '逐年保险费写入Excel
        For i = 3 To 33
            '设备保险费率
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(138, i).Value = bxfl_nt_list(i - 2) / 100
        Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '计算一次工作簿
        ExcelApp.Calculate()
        '————————————————————————————————————————————————————————————————————————————————————————     
        '读取并显示计算出的逐年保险费率
        Dim SJ As Double
        Dim nf
        For i = 1 To 31  '根据数组中的元素数量循环
            nf = i '年份序号
            SJ = Math.Round(bxfl_nt_list(i), 2)
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "%(" & nf & ") " '输出到RichTextBox1
        Next
        Me.RichTextBox1.Text = "暖通设备逐年保险费率：" & Me.RichTextBox1.Text
    End Sub
    Private Sub 计算_Click(sender As Object, e As EventArgs) Handles 计算.Click
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否确定输入的全部保险费计算参数？", vbOKCancel)
        If XZ = vbOK Then
            '计算一次工作簿
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
            '保险费计算
            Call 保险费相关计算(ExcelApp, xlfl_cg_model, xlfl_qt_model, sdsl_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            '————————————————————————————————————————————————————————————————————————————————————————        
            Me.RichTextBox1.Text = "设备逐年保险费计算完成！"
        End If
    End Sub
    Private Sub 光伏_Click(sender As Object, e As EventArgs) Handles 光伏.Click
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '清空已有的数据，防止出错
        For i = 3 To 33
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(137, i).Value = 0
        Next
        '清空窗体
        Me.RichTextBox1.Clear()
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '计算逐年保险费率，光伏设备只可以以投资百分比计算，选择无效
        Dim bxfl_gf_list = 逐年费率计算_base(ExcelApp, ksnf_list, jsnf_list, ksfl_list, jsfl_list)
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '逐年保险费写入Excel
        For i = 3 To 33
            '设备保险费率
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(137, i).Value = bxfl_gf_list(i - 2) / 100
        Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '计算一次工作簿
        ExcelApp.Calculate()
        '————————————————————————————————————————————————————————————————————————————————————————     
        '读取并显示计算出的逐年保险费率
        Dim SJ As Double
        Dim nf
        For i = 1 To 31  '根据数组中的元素数量循环
            nf = i '年份序号
            SJ = Math.Round(bxfl_gf_list(i), 2)
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "%(" & nf & ") " '输出到RichTextBox1
        Next
        Me.RichTextBox1.Text = "光伏设备逐年保险费率：" & Me.RichTextBox1.Text
    End Sub
    Private Sub 风电_Click(sender As Object, e As EventArgs) Handles 风电.Click
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '清空已有的数据，防止出错
        For i = 3 To 33
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(139, i).Value = 0
        Next
        '清空窗体
        Me.RichTextBox1.Clear()
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '计算逐年保险费率，风电设备只可以以投资百分比计算，选择无效
        Dim bxfl_fd_list = 逐年费率计算_base(ExcelApp, ksnf_list, jsnf_list, ksfl_list, jsfl_list)
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '逐年保险费写入Excel
        For i = 3 To 33
            '设备保险费率
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(139, i).Value = bxfl_fd_list(i - 2) / 100
        Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '计算一次工作簿
        ExcelApp.Calculate()
        '————————————————————————————————————————————————————————————————————————————————————————     
        '读取并显示计算出的逐年保险费率
        Dim SJ As Double
        Dim nf
        For i = 1 To 31  '根据数组中的元素数量循环
            nf = i '年份序号
            SJ = Math.Round(bxfl_fd_list(i), 2)
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "%(" & nf & ") " '输出到RichTextBox1
        Next
        Me.RichTextBox1.Text = "风电设备逐年保险费率：" & Me.RichTextBox1.Text
    End Sub
    Private Sub 设置剔除_Click(sender As Object, e As EventArgs) Handles 设置剔除.Click
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '燃机
        If ranji.Checked = True Then
            '运营年限
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(109, 3).Value = CInt(YYNX_RJ.Text)
            '扣除比例
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(110, 3).Value = CDbl(KCBL_RJ.Text) / 100
            '显示
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<燃机设备>保险费计算参数设置写入完成！"
        End If
        '蓄电池
        If xudianchi.Checked = True Then
            '运营年限
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(109, 4).Value = CInt(YYNX_XDC.Text)
            '扣除比例
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(110, 4).Value = CDbl(KCBL_XDC.Text) / 100
            '显示
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<蓄电池设备>保险费计算参数设置写入完成！"
        End If
        '暖通
        If nuantong.Checked = True Then
            '运营年限
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(109, 5).Value = CInt(YYNX_NT.Text)
            '扣除比例
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(110, 5).Value = CDbl(KCBL_NT.Text) / 100
            '显示
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<暖通设备>保险费计算参数设置写入完成！"
        End If
        '光伏
        If guangfu.Checked = True Then
            '运营年限
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(109, 6).Value = CInt(YYNX_GF.Text)
            '扣除比例
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(110, 6).Value = CDbl(KCBL_GF.Text) / 100
            '显示
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<光伏设备>保险费计算参数设置写入完成！"
        End If
        '风电
        If fengdian.Checked = True Then
            '运营年限
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(109, 7).Value = CInt(YYNX_FD.Text)
            '扣除比例
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(110, 7).Value = CDbl(KCBL_FD.Text) / 100
            '显示
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<风电设备>保险费计算参数设置写入完成！"
        End If
        '常规设备
        If changgui.Checked = True Then
            '运营年限
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(109, 8).Value = CInt(YYNX_CG.Text)
            '扣除比例
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(110, 8).Value = CDbl(KCBL_CG.Text) / 100
            '显示
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<常规设备>保险费计算参数设置写入完成！"
        End If
        '————————————————————————————————————————————————————————————————————————————————————————        
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
End Class