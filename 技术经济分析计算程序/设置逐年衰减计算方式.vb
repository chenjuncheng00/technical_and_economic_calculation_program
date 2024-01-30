Public Class 设置逐年衰减计算方式
    Public ksnf_list As New List(Of Integer)
    Public sjl_list As New List(Of Double)
    Private Sub 光伏发电_Click(sender As Object, e As EventArgs) Handles 光伏发电.Click
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '定义局部变量
        Dim jsksnf As Integer '计算开始年份
        Dim jsjsnf As Integer '计算结束年份
        '光伏计算开始年份
        If Me.gfksnf.Text = "" Then
            jsksnf = 0
        Else
            jsksnf = CType(Me.gfksnf.Text, Integer)
        End If
        '光伏计算结束年份
        If Me.gfjsnf.Text = "" Then
            jsjsnf = 0
        Else
            jsjsnf = CType(Me.gfjsnf.Text, Integer)
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(51, 1).Value > 0 Then '当光伏总发电量大于0，激活功能的条件
            Me.RichTextBox1.Clear()
            For i = 3 To 33 '清空已有的负荷率，防止出错
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(100, i).Value = 0
            Next
            '————————————————————————————————————————————————————————————————————————————————————————        
            '光伏发电的基础负荷率
            Dim fhl_base_gf = 逐年衰减率计算_base(ExcelApp, jsksnf, jsjsnf, ksnf_list, sjl_list)
            '————————————————————————————————————————————————————————————————————————————————————————
            ExcelApp.Calculate()
            '光伏逐年衰减系数计算方式
            Dim tcyfzs As Boolean
            If Me.CheckBox1.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(46, 18).Value = "逐年投产月份比例"
                tcyfzs = True
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(46, 18).Value = "常规设置"
                tcyfzs = False
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            ExcelApp.Calculate()
            '计算光伏发电逐年综合系数
            Call 光伏逐年综合达产率计算(ExcelApp, fhl_base_gf, tcyfzs)
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
            '显示此时的光伏逐年负荷率
            Dim SJ As Double
            Dim nf
            For i = 1 To 31  '根据数组中的元素数量循环
                nf = i '年份序号
                SJ = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(100, i + 2).Value) * 100, 2) '读取综合负荷率
                Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "%(" & nf & ") " '输出到RichTextBox1
            Next
            Me.RichTextBox1.Text = "光伏发电逐年负荷率：" & Me.RichTextBox1.Text
        Else
            MsgBox("估算表中输入的光伏装机规模为0，请检查！")
        End If
    End Sub
    Protected Overrides Sub OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim key As String
        key = e.KeyChar
        '检验按键是否为回车键，如果是就把焦点附给按钮1，并执行Click命令
        If key = Microsoft.VisualBasic.ChrW(13) Then
            光伏发电.Focus()
            光伏发电.PerformClick()
        End If
    End Sub
    Private Sub 设置逐年衰减计算方式_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————        
        MyBase.KeyPreview = True
        '清空文本框中已有的各种数据
        Me.开始年份tmp.Clear()
        Me.衰减率tmp.Clear()
        Me.开始年份列表.Items.Clear()
        Me.衰减率列表.Items.Clear()
        Me.gfksnf.Clear()
        Me.xdcksnf.Clear()
        Me.gfjsnf.Clear()
        Me.xdcjsnf.Clear()
        Me.RichTextBox1.Rtf = Nothing
        Me.RichTextBox1.Clear()
        '是否勾选按照逐年投产月份数折算
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(9, 11).Value = "折算" Then
            Me.CheckBox1.Checked = True
        Else
            Me.CheckBox1.Checked = False
        End If
    End Sub
    Private Sub 蓄电池供电_Click(sender As Object, e As EventArgs) Handles 蓄电池供电.Click
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————        
        '定义局部变量
        Dim jsksnf As Integer '计算开始年份
        Dim jsjsnf As Integer '计算结束年份
        '蓄电池计算开始年份
        If Me.xdcksnf.Text = "" Then
            jsksnf = 0
        Else
            jsksnf = CType(Me.xdcksnf.Text, Integer)
        End If
        '蓄电池计算年限
        If Me.xdcjsnf.Text = "" Then
            jsjsnf = 0
        Else
            jsjsnf = CType(Me.xdcjsnf.Text, Integer)
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Me.RichTextBox1.Clear()
        For i = 3 To 33 '清空已有的负荷率，防止出错
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(96, i).Value = 0
        Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '蓄电池供电的基础负荷率
        Dim fhl_base_xdc = 逐年衰减率计算_base(ExcelApp, jsksnf, jsjsnf, ksnf_list, sjl_list)
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '计算模式
        Dim jsms As String = "供电"
        '是否需要根据逐年投产月份数量折算
        Dim tcyfzs As Boolean
        If Me.CheckBox1.Checked = True Then
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(47, 18).Value = "逐年投产月份比例"
            tcyfzs = True
        Else
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(47, 18).Value = "常规设置"
            tcyfzs = False
        End If
        '————————————————————————————————————————————————————————————————————————————————————————        
        '计算一次Excle
        ExcelApp.Calculate()
        '计算逐年负荷率
        Call 蓄电池逐年综合达产率计算(ExcelApp, fhl_base_xdc, jsms, tcyfzs)
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
        '显示此时的蓄电池逐年负荷率
        Dim SJ As Double
        Dim nf
        For i = 1 To 31  '根据数组中的元素数量循环
            SJ = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(96, i + 2).Value) * 100, 2) '读取综合负荷率
            'SJ = Math.Round((fhl_base_xdc(i)) * 100, 2)
            nf = i
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "%(" & nf & ") " '输出到RichTextBox1
        Next
        Me.RichTextBox1.Text = "蓄电池供电逐年负荷率：" & Me.RichTextBox1.Text
    End Sub
    Private Sub 清空窗体_Click(sender As Object, e As EventArgs) Handles 清空窗体.Click
        Dim XZ = MsgBox("是否清空窗体中的全部内容？", vbOKCancel)
        If XZ = vbOK Then
            Me.开始年份tmp.Clear()
            Me.衰减率tmp.Clear()
            Me.开始年份列表.Items.Clear()
            Me.衰减率列表.Items.Clear()
            ksnf_list.Clear()
            sjl_list.Clear()
            Me.gfjsnf.Clear()
            Me.xdcjsnf.Clear()
            Me.RichTextBox1.Rtf = Nothing
            Me.RichTextBox1.Clear()
        End If
    End Sub
    Private Sub 光伏发电默认系数_Click(sender As Object, e As EventArgs) Handles 光伏发电默认系数.Click
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————               
        '读取项目计算年限
        Dim xmjsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '————————————————————————————————————————————————————————————————————————————————————————————     
        Dim XZ = MsgBox("是否加载光伏发电的默认逐年衰减系数和计算年限？", vbOKCancel)
        If XZ = vbOK Then
            '清空
            Me.开始年份tmp.Clear()
            Me.衰减率tmp.Clear()
            Me.开始年份列表.Items.Clear()
            Me.衰减率列表.Items.Clear()
            ksnf_list.Clear()
            sjl_list.Clear()
            Me.RichTextBox1.Rtf = Nothing
            Me.RichTextBox1.Clear()
            Me.xdcksnf.Clear()
            Me.xdcjsnf.Clear()
            '默认值
            开始年份列表.Items.Add(2)
            开始年份列表.Items.Add(3)
            衰减率列表.Items.Add(2)
            衰减率列表.Items.Add(0.55)
            ksnf_list.Add(2)
            ksnf_list.Add(3)
            sjl_list.Add(2)
            sjl_list.Add(0.55)
            Me.gfjsnf.Text = CType(xmjsnx, String)
            '——————————————————————————————————————————————————————————————————————————————————————————————
            '各种年限系数的计算开始年份（补贴收入、销售收入和成本）
            Dim JSKSNF As Integer = 0
            For i = 1 To 15
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(5, 4 + i).Value > 0 Then
                    JSKSNF = i
                    Exit For
                End If
            Next
            Me.gfksnf.Text = CType(JSKSNF, String)
        End If
    End Sub
    Private Sub 蓄电池默认系数_Click(sender As Object, e As EventArgs) Handles 蓄电池默认系数.Click
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————               
        Dim XZ = MsgBox("是否加载蓄电池的默认逐年衰减系数和计算年限？", vbOKCancel)
        If XZ = vbOK Then
            '清空
            Me.开始年份tmp.Clear()
            Me.衰减率tmp.Clear()
            Me.开始年份列表.Items.Clear()
            Me.衰减率列表.Items.Clear()
            ksnf_list.Clear()
            sjl_list.Clear()
            Me.RichTextBox1.Rtf = Nothing
            Me.RichTextBox1.Clear()
            Me.gfksnf.Clear()
            Me.gfjsnf.Clear()
            '默认值
            开始年份列表.Items.Add(3)
            衰减率列表.Items.Add(1)
            ksnf_list.Add(3)
            sjl_list.Add(1)
            Me.xdcjsnf.Text = 10
            '——————————————————————————————————————————————————————————————————————————————————————————————
            Dim JSKSNF As Integer = 0
            For i = 1 To 15
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(5, 4 + i).Value > 0 Then
                    JSKSNF = i
                    Exit For
                End If
            Next
            Me.xdcksnf.Text = CType(JSKSNF, String)
        End If
    End Sub

    Private Sub 蓄电池购电_Click(sender As Object, e As EventArgs) Handles 蓄电池购电.Click
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————        
        '定义局部变量
        Dim jsksnf As Integer '计算开始年份
        Dim jsjsnf As Integer '计算结束年份
        '蓄电池计算开始年份
        If Me.xdcksnf.Text = "" Then
            jsksnf = 0
        Else
            jsksnf = CType(Me.xdcksnf.Text, Integer)
        End If
        '蓄电池计算年限
        If Me.xdcjsnf.Text = "" Then
            jsjsnf = 0
        Else
            jsjsnf = CType(Me.xdcjsnf.Text, Integer)
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        Me.RichTextBox1.Clear()
        For i = 3 To 33 '清空已有的负荷率，防止出错
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(97, i).Value = 0
        Next
        '蓄电池购电基础负荷率
        Dim fhl_base_xdc = 逐年衰减率计算_base(ExcelApp, jsksnf, jsjsnf, ksnf_list, sjl_list)
        '————————————————————————————————————————————————————————————————————————————————————————  
        '计算模式
        Dim jsms As String = "购电"
        '是否需要根据逐年投产月份数量折算
        Dim tcyfzs As Boolean
        If Me.CheckBox1.Checked = True Then
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(47, 18).Value = "逐年投产月份比例"
            tcyfzs = True
        Else
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(47, 18).Value = "常规设置"
            tcyfzs = False
        End If
        '————————————————————————————————————————————————————————————————————————————————————————        
        '计算一次Excle
        ExcelApp.Calculate()
        '计算逐年负荷率
        Call 蓄电池逐年综合达产率计算(ExcelApp, fhl_base_xdc, jsms, tcyfzs)
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
        '显示此时的蓄电池逐年负荷率
        Dim SJ As Double
        Dim nf
        For i = 1 To 31  '根据数组中的元素数量循环
            SJ = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(97, i + 2).Value) * 100, 2) '读取综合负荷率
            'SJ = Math.Round((fhl_base_xdc(i)) * 100, 2)
            nf = i
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "%(" & nf & ") " '输出到RichTextBox1
        Next
        Me.RichTextBox1.Text = "蓄电池购电逐年负荷率：" & Me.RichTextBox1.Text
    End Sub

    Private Sub 清空输入_Click(sender As Object, e As EventArgs) Handles 清空输入.Click
        Me.开始年份tmp.Clear()
        Me.衰减率tmp.Clear()
        Me.开始年份列表.Items.Clear()
        Me.衰减率列表.Items.Clear()
        ksnf_list.Clear()
        sjl_list.Clear()
        Me.gfksnf.Clear()
        Me.xdcksnf.Clear()
        Me.gfjsnf.Clear()
        Me.xdcjsnf.Clear()
    End Sub

    Private Sub 添加输入_Click(sender As Object, e As EventArgs) Handles 添加输入.Click
        '判断1
        If 开始年份tmp.Text = Nothing Or 衰减率tmp.Text = Nothing Then
            MsgBox("输入的开始年份和衰减率都必须不能为空！，请重新输入")
            Exit Sub
        End If
        '开始年份
        Dim ksnf_text As String = 开始年份tmp.Text
        Dim ksnf As Integer = CType(ksnf_text, Integer)
        '衰减率
        Dim sjl_text As String = 衰减率tmp.Text
        Dim sjl As Double = CType(sjl_text, Double)
        '添加数据
        开始年份列表.Items.Add(ksnf_text)
        ksnf_list.Add(ksnf)
        开始年份tmp.Clear()
        衰减率列表.Items.Add(sjl_text)
        sjl_list.Add(sjl)
        衰减率tmp.Clear()
    End Sub
End Class