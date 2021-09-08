Public Class 逐年衰减系数设置
    Private Sub 光伏发电_Click(sender As Object, e As EventArgs) Handles 光伏发电.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp       
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '定义局部变量
        Dim SJKSNF1, SJKSNF2, SJKSNF3, SJKSNF4, SJKSNF5 As Integer '衰减开始年份
        Dim SJL1, SJL2, SJL3, SJL4, SJL5 As Double '衰减率
        Dim jsnx = CType(Me.gfjsnf.Text, Integer) '光伏计算结束年份
        Dim ksnf = CType(Me.gfksnf.Text, Integer) '光伏计算开始年份
        '读取项目计算年限
        Dim xmjsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(51, 1).Value > 0 Then '当光伏总发电量大于0，激活功能的条件
            SJKSNF1 = CType(Me.TextBox1.Text, Integer)
            SJKSNF2 = CType(Me.TextBox2.Text, Integer)
            SJKSNF3 = CType(Me.TextBox3.Text, Integer)
            SJKSNF4 = CType(Me.TextBox7.Text, Integer)
            SJKSNF5 = CType(Me.TextBox9.Text, Integer)
            SJL1 = CType(Me.TextBox4.Text, Double)
            SJL2 = CType(Me.TextBox5.Text, Double)
            SJL3 = CType(Me.TextBox6.Text, Double)
            SJL4 = CType(Me.TextBox8.Text, Double)
            SJL5 = CType(Me.TextBox10.Text, Double)
            '————————————————————————————————————————————————————————————————————————————————————————  
            '输入的光伏计算年限不可以大于项目计算年限
            If jsnx > xmjsnx Then
                MsgBox("窗口中输入的光伏计算年限不可以超过项目计算年限，请重新输入！")
                Exit Sub
            End If
            '检查输入的衰减开始年份
            If jsnx > 31 Then
                MsgBox("输入的项目计算年限不可以大于31年，请重新输入！")
                Exit Sub
            End If
            If SJKSNF1 > jsnx Or SJKSNF2 > jsnx Or SJKSNF3 > jsnx Or SJKSNF4 > jsnx Or SJKSNF5 > jsnx Then
                MsgBox("输入的衰减年份不可以大于项目计算年限，请重新输入！")
                Exit Sub
            End If
            If SJKSNF1 > 0 And SJKSNF2 > 0 And SJKSNF1 >= SJKSNF2 Then
                MsgBox("输入的第1个开始的衰减年份大于等于了第2个开始的衰减年份，请重新输入！")
                Exit Sub
            End If
            If SJKSNF1 > 0 And SJKSNF3 > 0 And SJKSNF1 >= SJKSNF3 Then
                MsgBox("输入的第1个开始的衰减年份大于等于了第3个开始的衰减年份，请重新输入！")
                Exit Sub
            End If
            If SJKSNF1 > 0 And SJKSNF4 > 0 And SJKSNF1 >= SJKSNF4 Then
                MsgBox("输入的第1个开始的衰减年份大于等于了第4个开始的衰减年份，请重新输入！")
                Exit Sub
            End If
            If SJKSNF1 > 0 And SJKSNF5 > 0 And SJKSNF1 >= SJKSNF5 Then
                MsgBox("输入的第1个开始的衰减年份大于等于了第5个开始的衰减年份，请重新输入！")
                Exit Sub
            End If
            If SJKSNF2 > 0 And SJKSNF3 > 0 And SJKSNF2 >= SJKSNF3 Then
                MsgBox("输入的第2个开始的衰减年份大于等于了第3个开始的衰减年份，请重新输入！")
                Exit Sub
            End If
            If SJKSNF2 > 0 And SJKSNF4 > 0 And SJKSNF2 >= SJKSNF4 Then
                MsgBox("输入的第2个开始的衰减年份大于等于了第4个开始的衰减年份，请重新输入！")
                Exit Sub
            End If
            If SJKSNF2 > 0 And SJKSNF5 > 0 And SJKSNF2 >= SJKSNF5 Then
                MsgBox("输入的第2个开始的衰减年份大于等于了第5个开始的衰减年份，请重新输入！")
                Exit Sub
            End If
            If SJKSNF2 > 0 And SJKSNF1 = 0 Then
                MsgBox("衰减开始年份必需从第1个年份开始输入，请重新输入！")
                Exit Sub
            End If
            If SJKSNF3 > 0 And SJKSNF1 = 0 Then
                MsgBox("衰减开始年份必需从第1个年份开始输入，请重新输入！")
                Exit Sub
            End If
            If SJKSNF4 > 0 And SJKSNF1 = 0 Then
                MsgBox("衰减开始年份必需从第1个年份开始输入，请重新输入！")
                Exit Sub
            End If
            If SJKSNF5 > 0 And SJKSNF1 = 0 Then
                MsgBox("衰减开始年份必需从第1个年份开始输入，请重新输入！")
                Exit Sub
            End If
            If SJKSNF3 > 0 And SJKSNF2 = 0 Then
                MsgBox("衰减开始年份必需从第1个和第2个年份开始输入，请重新输入！")
                Exit Sub
            End If
            If SJKSNF4 > 0 And SJKSNF2 = 0 Then
                MsgBox("衰减开始年份必需从第1个和第2个年份开始输入，请重新输入！")
                Exit Sub
            End If
            If SJKSNF5 > 0 And SJKSNF2 = 0 Then
                MsgBox("衰减开始年份必需从第1个和第2个年份开始输入，请重新输入！")
                Exit Sub
            End If
            If SJKSNF4 > 0 And SJKSNF3 = 0 Then
                MsgBox("衰减开始年份必需从第1个、第2个和第3个年份开始输入，请重新输入！")
                Exit Sub
            End If
            If SJKSNF5 > 0 And SJKSNF3 = 0 Then
                MsgBox("衰减开始年份必需从第1个、第2个和第3个年份开始输入，请重新输入！")
                Exit Sub
            End If
            If SJKSNF5 > 0 And SJKSNF4 = 0 Then
                MsgBox("衰减开始年份必需从第1个、第2个、第3个和第4个年份开始输入，请重新输入！")
                Exit Sub
            End If
            If SJKSNF1 < 0 Or SJKSNF2 < 0 Or SJKSNF3 < 0 Or SJKSNF4 < 0 Or SJKSNF5 < 0 Then
                MsgBox("输入的第1个、第2个、第3个、第4个和第5个开始的衰减年份不能小于0，请重新输入！")
                Exit Sub
            End If
            If SJKSNF1 < 1 Then
                MsgBox("输入的第1个开始的衰减年份最小只能是1，请重新输入！")
                Exit Sub
            End If
            '————————————————————————————————————————————————————————————————————————————————————————        
            Me.RichTextBox1.Clear()
            For i = 3 To 33 '清空已有的负荷率，防止出错
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(100, i).Value = 0
            Next
            '光伏发电的基础负荷率
            Dim fhl_base_gf(31) As Double
            '输入了5种不同的衰减系数
            If SJKSNF1 > 0 And SJKSNF2 > 0 And SJKSNF3 > 0 And SJKSNF4 > 0 And SJKSNF5 > 0 Then
                Dim YJSJ1 As Double = 0 '第一种衰减率已经衰减了多少，初始值为0
                Dim JS1 As Integer = 0 '计数，用于记录第一种衰减率已经衰减的年份数量，初始值为0
                For i = 1 To 31 '小于第一个衰减开始年份的负荷率设置为1
                    If i >= ksnf And i < SJKSNF1 Then
                        fhl_base_gf(i) = 1
                    End If
                Next
                For i = 1 To 31 '收入税收表的列，第1种衰减率
                    If i >= SJKSNF1 And i < SJKSNF2 Then
                        JS1 = JS1 + 1
                        YJSJ1 = SJL1 * JS1
                        fhl_base_gf(i) = (100 - YJSJ1) / 100
                    End If
                Next
                Dim YJSJ2 As Double = 0 '第2种衰减率已经衰减了多少，初始值为0
                Dim JS2 As Integer = 0 '计数，用于记录第2种衰减率已经衰减的年份数量，初始值为0
                For i = 1 To 31 '收入税收表的列，第2种衰减率
                    If i >= SJKSNF2 And i < SJKSNF3 Then
                        JS2 = JS2 + 1
                        YJSJ2 = SJL2 * JS2
                        fhl_base_gf(i) = (100 - YJSJ1 - YJSJ2) / 100
                    End If
                Next
                Dim YJSJ3 As Double = 0 '第3种衰减率已经衰减了多少，初始值为0
                Dim JS3 As Integer = 0 '计数，用于记录第3种衰减率已经衰减的年份数量，初始值为0
                For i = 1 To 31 '收入税收表的列，第3种衰减率
                    If i >= SJKSNF3 And i < SJKSNF4 Then
                        JS3 = JS3 + 1
                        YJSJ3 = SJL3 * JS3
                        fhl_base_gf(i) = (100 - YJSJ1 - YJSJ2 - YJSJ3) / 100
                    End If
                Next
                Dim YJSJ4 As Double = 0 '第4种衰减率已经衰减了多少，初始值为0
                Dim JS4 As Integer = 0 '计数，用于记录第4种衰减率已经衰减的年份数量，初始值为0
                For i = 1 To 31 '收入税收表的列，第4种衰减率
                    If i >= SJKSNF4 And i < SJKSNF5 Then
                        JS4 = JS4 + 1
                        YJSJ4 = SJL4 * JS4
                        fhl_base_gf(i) = (100 - YJSJ1 - YJSJ2 - YJSJ3 - YJSJ4) / 100
                    End If
                Next
                Dim YJSJ5 As Double = 0 '第5种衰减率已经衰减了多少，初始值为0
                Dim JS5 As Integer = 0 '计数，用于记录第5种衰减率已经衰减的年份数量，初始值为0
                For i = 1 To 31 '收入税收表的列，第4种衰减率
                    If i >= SJKSNF5 And i <= jsnx Then
                        JS5 = JS5 + 1
                        YJSJ5 = SJL5 * JS5
                        fhl_base_gf(i) = (100 - YJSJ1 - YJSJ2 - YJSJ3 - YJSJ4 - YJSJ5) / 100
                    End If
                Next
            End If
            '输入了4种不同的衰减系数
            If SJKSNF1 > 0 And SJKSNF2 > 0 And SJKSNF3 > 0 And SJKSNF4 > 0 And SJKSNF5 = 0 Then
                Dim YJSJ1 As Double = 0 '第一种衰减率已经衰减了多少，初始值为0
                Dim JS1 As Integer = 0 '计数，用于记录第一种衰减率已经衰减的年份数量，初始值为0
                For i = 1 To 31 '小于第一个衰减开始年份的负荷率设置为1
                    If i >= ksnf And i < SJKSNF1 Then
                        fhl_base_gf(i) = 1
                    End If
                Next
                For i = 1 To 31 '收入税收表的列，第1种衰减率
                    If i >= SJKSNF1 And i < SJKSNF2 Then
                        JS1 = JS1 + 1
                        YJSJ1 = SJL1 * JS1
                        fhl_base_gf(i) = (100 - YJSJ1) / 100
                    End If
                Next
                Dim YJSJ2 As Double = 0 '第2种衰减率已经衰减了多少，初始值为0
                Dim JS2 As Integer = 0 '计数，用于记录第2种衰减率已经衰减的年份数量，初始值为0
                For i = 1 To 31 '收入税收表的列，第2种衰减率
                    If i >= SJKSNF2 And i < SJKSNF3 Then
                        JS2 = JS2 + 1
                        YJSJ2 = SJL2 * JS2
                        fhl_base_gf(i) = (100 - YJSJ1 - YJSJ2) / 100
                    End If
                Next
                Dim YJSJ3 As Double = 0 '第3种衰减率已经衰减了多少，初始值为0
                Dim JS3 As Integer = 0 '计数，用于记录第3种衰减率已经衰减的年份数量，初始值为0
                For i = 1 To 31 '收入税收表的列，第3种衰减率
                    If i >= SJKSNF3 And i < SJKSNF4 Then
                        JS3 = JS3 + 1
                        YJSJ3 = SJL3 * JS3
                        fhl_base_gf(i) = (100 - YJSJ1 - YJSJ2 - YJSJ3) / 100
                    End If
                Next
                Dim YJSJ4 As Double = 0 '第4种衰减率已经衰减了多少，初始值为0
                Dim JS4 As Integer = 0 '计数，用于记录第4种衰减率已经衰减的年份数量，初始值为0
                For i = 1 To 31 '收入税收表的列，第4种衰减率
                    If i >= SJKSNF4 And i <= jsnx Then
                        JS4 = JS4 + 1
                        YJSJ4 = SJL4 * JS4
                        fhl_base_gf(i) = (100 - YJSJ1 - YJSJ2 - YJSJ3 - YJSJ4) / 100
                    End If
                Next
            End If
            '输入了3种不同的衰减系数
            If SJKSNF1 > 0 And SJKSNF2 > 0 And SJKSNF3 > 0 And SJKSNF4 = 0 And SJKSNF5 = 0 Then
                Dim YJSJ1 As Double = 0 '第一种衰减率已经衰减了多少，初始值为0
                Dim JS1 As Integer = 0 '计数，用于记录第一种衰减率已经衰减的年份数量，初始值为0
                For i = 1 To 31 '小于第一个衰减开始年份的负荷率设置为1
                    If i >= ksnf And i < SJKSNF1 Then
                        fhl_base_gf(i) = 1
                    End If
                Next
                For i = 1 To 31 '收入税收表的列，第1种衰减率
                    If i >= SJKSNF1 And i < SJKSNF2 Then
                        JS1 = JS1 + 1
                        YJSJ1 = SJL1 * JS1
                        fhl_base_gf(i) = (100 - YJSJ1) / 100
                    End If
                Next
                Dim YJSJ2 As Double = 0 '第2种衰减率已经衰减了多少，初始值为0
                Dim JS2 As Integer = 0 '计数，用于记录第2种衰减率已经衰减的年份数量，初始值为0
                For i = 1 To 31 '收入税收表的列，第2种衰减率
                    If i >= SJKSNF2 And i < SJKSNF3 Then
                        JS2 = JS2 + 1
                        YJSJ2 = SJL2 * JS2
                        fhl_base_gf(i) = (100 - YJSJ1 - YJSJ2) / 100
                    End If
                Next
                Dim YJSJ3 As Double = 0 '第3种衰减率已经衰减了多少，初始值为0
                Dim JS3 As Integer = 0 '计数，用于记录第3种衰减率已经衰减的年份数量，初始值为0
                For i = 1 To 31 '收入税收表的列，第2种衰减率
                    If i >= SJKSNF3 And i <= jsnx Then
                        JS3 = JS3 + 1
                        YJSJ3 = SJL3 * JS3
                        fhl_base_gf(i) = (100 - YJSJ1 - YJSJ2 - YJSJ3) / 100
                    End If
                Next
            End If
            '输入了2种不同的衰减系数
            If SJKSNF1 > 0 And SJKSNF2 > 0 And SJKSNF3 = 0 And SJKSNF4 = 0 And SJKSNF5 = 0 Then
                Dim YJSJ1 As Double = 0 '第一种衰减率已经衰减了多少，初始值为0
                Dim JS1 As Integer = 0 '计数，用于记录第一种衰减率已经衰减的年份数量，初始值为0
                For i = 1 To 31 '小于第一个衰减开始年份的负荷率设置为1
                    If i >= ksnf And i < SJKSNF1 Then
                        fhl_base_gf(i) = 1
                    End If
                Next
                For i = 1 To 31 '收入税收表的列，第1种衰减率
                    If i >= SJKSNF1 And i < SJKSNF2 Then
                        JS1 = JS1 + 1
                        YJSJ1 = SJL1 * JS1
                        fhl_base_gf(i) = (100 - YJSJ1) / 100
                    End If
                Next
                Dim YJSJ2 As Double = 0 '第2种衰减率已经衰减了多少，初始值为0
                Dim JS2 As Integer = 0 '计数，用于记录第2种衰减率已经衰减的年份数量，初始值为0
                For i = 1 To 31 '收入税收表的列，第2种衰减率
                    If i >= SJKSNF2 And i <= jsnx Then
                        JS2 = JS2 + 1
                        YJSJ2 = SJL2 * JS2
                        fhl_base_gf(i) = (100 - YJSJ1 - YJSJ2) / 100
                    End If
                Next
            End If
            '输入了1种不同的衰减系数
            If SJKSNF1 > 0 And SJKSNF2 = 0 And SJKSNF3 = 0 And SJKSNF4 = 0 And SJKSNF5 = 0 Then
                Dim YJSJ1 As Double = 0 '第一种衰减率已经衰减了多少，初始值为0
                Dim JS1 As Integer = 0 '计数，用于记录第一种衰减率已经衰减的年份数量，初始值为0
                For i = 1 To 31 '小于第一个衰减开始年份的负荷率设置为1
                    If i >= ksnf And i < SJKSNF1 Then
                        fhl_base_gf(i) = 1
                    End If
                Next
                For i = 1 To 31 '收入税收表的列，第1种衰减率
                    If i >= SJKSNF1 And i <= jsnx Then
                        JS1 = JS1 + 1
                        YJSJ1 = SJL1 * JS1
                        fhl_base_gf(i) = (100 - YJSJ1) / 100
                    End If
                Next
            End If
            '将大于计算年限的负荷率设置为0
            For i = 1 To 31 '收入税收表的列，第1种衰减率
                If i > jsnx Then
                    fhl_base_gf(i) = 0
                End If
            Next
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
            'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
            Dim sdsl_model As Integer = 1
            '收入和成本变化后相关计算
            Call 计算功能合并整理.收入成本相关计算(ExcelApp, sdsl_model)
            '———————————————————————————————————————————————————————————————————————————————————————— 
            '清空已有内容
            Me.RichTextBox1.Rtf = Nothing
            Me.RichTextBox1.Clear()
            '显示此时的光伏逐年负荷率
            Dim SJ As Double
            Dim nf
            For i = 1 To 31  '根据数组中的元素数量循环
                nf = i '年份序号
                SJ = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(100, i).Value) * 100, 2) '读取综合负荷率
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
    Private Sub 逐年衰减系数设置_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————        
        MyBase.KeyPreview = True
        '清空文本框中已有的各种数据
        Me.TextBox1.Text = Nothing
        Me.TextBox2.Text = Nothing
        Me.TextBox3.Text = Nothing
        Me.TextBox4.Text = Nothing
        Me.TextBox5.Text = Nothing
        Me.TextBox6.Text = Nothing
        Me.TextBox7.Text = Nothing
        Me.TextBox8.Text = Nothing
        Me.TextBox9.Text = Nothing
        Me.TextBox10.Text = Nothing
        Me.gfksnf.Text = Nothing
        Me.xdcksnf.Text = Nothing
        Me.gfjsnf.Text = Nothing
        Me.xdcjsnf.Text = Nothing
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
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————        
        '定义局部变量
        Dim SJKSNF1, SJKSNF2, SJKSNF3, SJKSNF4, SJKSNF5 As Integer '衰减开始年份
        Dim SJL1, SJL2, SJL3, SJL4, SJL5 As Double '衰减率
        Dim jsnx = CType(Me.xdcjsnf.Text, Integer) '蓄电池计算年限
        Dim ksnf = CType(Me.xdcksnf.Text, Integer) '蓄电池计算开始年份
        '读取项目计算年限
        Dim xmjsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        SJKSNF1 = CType(Me.TextBox1.Text, Integer)
        SJKSNF2 = CType(Me.TextBox2.Text, Integer)
        SJKSNF3 = CType(Me.TextBox3.Text, Integer)
        SJKSNF4 = CType(Me.TextBox7.Text, Integer)
        SJKSNF5 = CType(Me.TextBox9.Text, Integer)
        SJL1 = CType(Me.TextBox4.Text, Double)
        SJL2 = CType(Me.TextBox5.Text, Double)
        SJL3 = CType(Me.TextBox6.Text, Double)
        SJL4 = CType(Me.TextBox8.Text, Double)
        SJL5 = CType(Me.TextBox10.Text, Double)
        '————————————————————————————————————————————————————————————————————————————————————————  
        '输入的蓄电池计算年限不可以大于项目计算年限
        If jsnx > xmjsnx Then
            MsgBox("窗口中输入的蓄电池计算年限不可以超过项目计算年限，请重新输入！")
            Exit Sub
        End If
        '检查输入的衰减开始年份
        If jsnx > 31 Then
            MsgBox("输入的项目计算年限不可以大于31年，请重新输入！")
            Exit Sub
        End If
        If SJKSNF1 > jsnx Or SJKSNF2 > jsnx Or SJKSNF3 > jsnx Or SJKSNF4 > jsnx Or SJKSNF5 > jsnx Then
            MsgBox("输入的衰减年份不可以大于项目计算年限，请重新输入！")
            Exit Sub
        End If
        If SJKSNF1 > 0 And SJKSNF2 > 0 And SJKSNF1 >= SJKSNF2 Then
            MsgBox("输入的第1个开始的衰减年份大于等于了第2个开始的衰减年份，请重新输入！")
            Exit Sub
        End If
        If SJKSNF1 > 0 And SJKSNF3 > 0 And SJKSNF1 >= SJKSNF3 Then
            MsgBox("输入的第1个开始的衰减年份大于等于了第3个开始的衰减年份，请重新输入！")
            Exit Sub
        End If
        If SJKSNF1 > 0 And SJKSNF4 > 0 And SJKSNF1 >= SJKSNF4 Then
            MsgBox("输入的第1个开始的衰减年份大于等于了第4个开始的衰减年份，请重新输入！")
            Exit Sub
        End If
        If SJKSNF1 > 0 And SJKSNF5 > 0 And SJKSNF1 >= SJKSNF5 Then
            MsgBox("输入的第1个开始的衰减年份大于等于了第5个开始的衰减年份，请重新输入！")
            Exit Sub
        End If
        If SJKSNF2 > 0 And SJKSNF3 > 0 And SJKSNF2 >= SJKSNF3 Then
            MsgBox("输入的第2个开始的衰减年份大于等于了第3个开始的衰减年份，请重新输入！")
            Exit Sub
        End If
        If SJKSNF2 > 0 And SJKSNF4 > 0 And SJKSNF2 >= SJKSNF4 Then
            MsgBox("输入的第2个开始的衰减年份大于等于了第4个开始的衰减年份，请重新输入！")
            Exit Sub
        End If
        If SJKSNF2 > 0 And SJKSNF5 > 0 And SJKSNF2 >= SJKSNF5 Then
            MsgBox("输入的第2个开始的衰减年份大于等于了第5个开始的衰减年份，请重新输入！")
            Exit Sub
        End If
        If SJKSNF2 > 0 And SJKSNF1 = 0 Then
            MsgBox("衰减开始年份必需从第1个年份开始输入，请重新输入！")
            Exit Sub
        End If
        If SJKSNF3 > 0 And SJKSNF1 = 0 Then
            MsgBox("衰减开始年份必需从第1个年份开始输入，请重新输入！")
            Exit Sub
        End If
        If SJKSNF4 > 0 And SJKSNF1 = 0 Then
            MsgBox("衰减开始年份必需从第1个年份开始输入，请重新输入！")
            Exit Sub
        End If
        If SJKSNF5 > 0 And SJKSNF1 = 0 Then
            MsgBox("衰减开始年份必需从第1个年份开始输入，请重新输入！")
            Exit Sub
        End If
        If SJKSNF3 > 0 And SJKSNF2 = 0 Then
            MsgBox("衰减开始年份必需从第1个和第2个年份开始输入，请重新输入！")
            Exit Sub
        End If
        If SJKSNF4 > 0 And SJKSNF2 = 0 Then
            MsgBox("衰减开始年份必需从第1个和第2个年份开始输入，请重新输入！")
            Exit Sub
        End If
        If SJKSNF5 > 0 And SJKSNF2 = 0 Then
            MsgBox("衰减开始年份必需从第1个和第2个年份开始输入，请重新输入！")
            Exit Sub
        End If
        If SJKSNF4 > 0 And SJKSNF3 = 0 Then
            MsgBox("衰减开始年份必需从第1个、第2个和第3个年份开始输入，请重新输入！")
            Exit Sub
        End If
        If SJKSNF5 > 0 And SJKSNF3 = 0 Then
            MsgBox("衰减开始年份必需从第1个、第2个和第3个年份开始输入，请重新输入！")
            Exit Sub
        End If
        If SJKSNF5 > 0 And SJKSNF4 = 0 Then
            MsgBox("衰减开始年份必需从第1个、第2个、第3个和第4个年份开始输入，请重新输入！")
            Exit Sub
        End If
        If SJKSNF1 < 0 Or SJKSNF2 < 0 Or SJKSNF3 < 0 Or SJKSNF4 < 0 Or SJKSNF5 < 0 Then
            MsgBox("输入的第1个、第2个、第3个、第4个和第5个开始的衰减年份不能小于0，请重新输入！")
            Exit Sub
        End If
        If SJKSNF1 < 1 Then
            MsgBox("输入的第1个开始的衰减年份最小只能是1，请重新输入！")
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————               
        Me.RichTextBox1.Clear()
        For i = 3 To 33 '清空已有的负荷率，防止出错
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(96, i).Value = 0
        Next
        '蓄电池供电的基础负荷率
        Dim fhl_base_xdc(31) As Double
        '输入了5种不同的衰减系数
        If SJKSNF1 > 0 And SJKSNF2 > 0 And SJKSNF3 > 0 And SJKSNF4 > 0 And SJKSNF5 > 0 Then
            Dim YJSJ1 As Double = 0 '第一种衰减率已经衰减了多少，初始值为0
            Dim JS1 As Integer = 0 '计数，用于记录第一种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '小于第一个衰减开始年份的负荷率设置为1
                If i >= ksnf And i < SJKSNF1 Then
                    fhl_base_xdc(i) = 1
                End If
            Next
            For i = 1 To 31 '收入税收表的列，第1种衰减率
                If i >= SJKSNF1 And i < SJKSNF2 Then
                    JS1 = JS1 + 1
                    YJSJ1 = SJL1 * JS1
                    fhl_base_xdc(i) = (100 - YJSJ1) / 100
                End If
            Next
            Dim YJSJ2 As Double = 0 '第2种衰减率已经衰减了多少，初始值为0
            Dim JS2 As Integer = 0 '计数，用于记录第2种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '收入税收表的列，第2种衰减率
                If i >= SJKSNF2 And i < SJKSNF3 Then
                    JS2 = JS2 + 1
                    YJSJ2 = SJL2 * JS2
                    fhl_base_xdc(i) = (100 - YJSJ1 - YJSJ2) / 100
                End If
            Next
            Dim YJSJ3 As Double = 0 '第3种衰减率已经衰减了多少，初始值为0
            Dim JS3 As Integer = 0 '计数，用于记录第3种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '收入税收表的列，第3种衰减率
                If i >= SJKSNF3 And i < SJKSNF4 Then
                    JS3 = JS3 + 1
                    YJSJ3 = SJL3 * JS3
                    fhl_base_xdc(i) = (100 - YJSJ1 - YJSJ2 - YJSJ3) / 100
                End If
            Next
            Dim YJSJ4 As Double = 0 '第4种衰减率已经衰减了多少，初始值为0
            Dim JS4 As Integer = 0 '计数，用于记录第4种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '收入税收表的列，第4种衰减率
                If i >= SJKSNF4 And i < SJKSNF5 Then
                    JS4 = JS4 + 1
                    YJSJ4 = SJL4 * JS4
                    fhl_base_xdc(i) = (100 - YJSJ1 - YJSJ2 - YJSJ3 - YJSJ4) / 100
                End If
            Next
            Dim YJSJ5 As Double = 0 '第5种衰减率已经衰减了多少，初始值为0
            Dim JS5 As Integer = 0 '计数，用于记录第5种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '收入税收表的列，第4种衰减率
                If i >= SJKSNF5 And i <= jsnx Then
                    JS5 = JS5 + 1
                    YJSJ5 = SJL5 * JS5
                    fhl_base_xdc(i) = (100 - YJSJ1 - YJSJ2 - YJSJ3 - YJSJ4 - YJSJ5) / 100
                End If
            Next
        End If
        '输入了4种不同的衰减系数
        If SJKSNF1 > 0 And SJKSNF2 > 0 And SJKSNF3 > 0 And SJKSNF4 > 0 And SJKSNF5 = 0 Then
            Dim YJSJ1 As Double = 0 '第一种衰减率已经衰减了多少，初始值为0
            Dim JS1 As Integer = 0 '计数，用于记录第一种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '小于第一个衰减开始年份的负荷率设置为1
                If i >= ksnf And i < SJKSNF1 Then
                    fhl_base_xdc(i) = 1
                End If
            Next
            For i = 1 To 31 '收入税收表的列，第1种衰减率
                If i >= SJKSNF1 And i < SJKSNF2 Then
                    JS1 = JS1 + 1
                    YJSJ1 = SJL1 * JS1
                    fhl_base_xdc(i) = (100 - YJSJ1) / 100
                End If
            Next
            Dim YJSJ2 As Double = 0 '第2种衰减率已经衰减了多少，初始值为0
            Dim JS2 As Integer = 0 '计数，用于记录第2种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '收入税收表的列，第2种衰减率
                If i >= SJKSNF2 And i < SJKSNF3 Then
                    JS2 = JS2 + 1
                    YJSJ2 = SJL2 * JS2
                    fhl_base_xdc(i) = (100 - YJSJ1 - YJSJ2) / 100
                End If
            Next
            Dim YJSJ3 As Double = 0 '第3种衰减率已经衰减了多少，初始值为0
            Dim JS3 As Integer = 0 '计数，用于记录第3种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '收入税收表的列，第3种衰减率
                If i >= SJKSNF3 And i < SJKSNF4 Then
                    JS3 = JS3 + 1
                    YJSJ3 = SJL3 * JS3
                    fhl_base_xdc(i) = (100 - YJSJ1 - YJSJ2 - YJSJ3) / 100
                End If
            Next
            Dim YJSJ4 As Double = 0 '第4种衰减率已经衰减了多少，初始值为0
            Dim JS4 As Integer = 0 '计数，用于记录第4种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '收入税收表的列，第4种衰减率
                If i >= SJKSNF4 And i <= jsnx Then
                    JS4 = JS4 + 1
                    YJSJ4 = SJL4 * JS4
                    fhl_base_xdc(i) = (100 - YJSJ1 - YJSJ2 - YJSJ3 - YJSJ4) / 100
                End If
            Next
        End If
        '输入了3种不同的衰减系数
        If SJKSNF1 > 0 And SJKSNF2 > 0 And SJKSNF3 > 0 And SJKSNF4 = 0 And SJKSNF5 = 0 Then
            Dim YJSJ1 As Double = 0 '第一种衰减率已经衰减了多少，初始值为0
            Dim JS1 As Integer = 0 '计数，用于记录第一种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '小于第一个衰减开始年份的负荷率设置为1
                If i >= ksnf And i < SJKSNF1 Then
                    fhl_base_xdc(i) = 1
                End If
            Next
            For i = 1 To 31 '收入税收表的列，第1种衰减率
                If i >= SJKSNF1 And i < SJKSNF2 Then
                    JS1 = JS1 + 1
                    YJSJ1 = SJL1 * JS1
                    fhl_base_xdc(i) = (100 - YJSJ1) / 100
                End If
            Next
            Dim YJSJ2 As Double = 0 '第2种衰减率已经衰减了多少，初始值为0
            Dim JS2 As Integer = 0 '计数，用于记录第2种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '收入税收表的列，第2种衰减率
                If i >= SJKSNF2 And i < SJKSNF3 Then
                    JS2 = JS2 + 1
                    YJSJ2 = SJL2 * JS2
                    fhl_base_xdc(i) = (100 - YJSJ1 - YJSJ2) / 100
                End If
            Next
            Dim YJSJ3 As Double = 0 '第3种衰减率已经衰减了多少，初始值为0
            Dim JS3 As Integer = 0 '计数，用于记录第3种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '收入税收表的列，第2种衰减率
                If i >= SJKSNF3 And i <= jsnx Then
                    JS3 = JS3 + 1
                    YJSJ3 = SJL3 * JS3
                    fhl_base_xdc(i) = (100 - YJSJ1 - YJSJ2 - YJSJ3) / 100
                End If
            Next
        End If
        '输入了2种不同的衰减系数
        If SJKSNF1 > 0 And SJKSNF2 > 0 And SJKSNF3 = 0 And SJKSNF4 = 0 And SJKSNF5 = 0 Then
            Dim YJSJ1 As Double = 0 '第一种衰减率已经衰减了多少，初始值为0
            Dim JS1 As Integer = 0 '计数，用于记录第一种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '小于第一个衰减开始年份的负荷率设置为1
                If i >= ksnf And i < SJKSNF1 Then
                    fhl_base_xdc(i) = 1
                End If
            Next
            For i = 1 To 31 '收入税收表的列，第1种衰减率
                If i >= SJKSNF1 And i < SJKSNF2 Then
                    JS1 = JS1 + 1
                    YJSJ1 = SJL1 * JS1
                    fhl_base_xdc(i) = (100 - YJSJ1) / 100
                End If
            Next
            Dim YJSJ2 As Double = 0 '第2种衰减率已经衰减了多少，初始值为0
            Dim JS2 As Integer = 0 '计数，用于记录第2种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '收入税收表的列，第2种衰减率
                If i >= SJKSNF2 And i <= jsnx Then
                    JS2 = JS2 + 1
                    YJSJ2 = SJL2 * JS2
                    fhl_base_xdc(i) = (100 - YJSJ1 - YJSJ2) / 100
                End If
            Next
        End If
        '输入了1种不同的衰减系数
        If SJKSNF1 > 0 And SJKSNF2 = 0 And SJKSNF3 = 0 And SJKSNF4 = 0 And SJKSNF5 = 0 Then
            Dim YJSJ1 As Double = 0 '第一种衰减率已经衰减了多少，初始值为0
            Dim JS1 As Integer = 0 '计数，用于记录第一种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '小于第一个衰减开始年份的负荷率设置为1
                If i >= ksnf And i < SJKSNF1 Then
                    fhl_base_xdc(i) = 1
                End If
            Next
            For i = 1 To 31 '收入税收表的列，第1种衰减率
                If i >= SJKSNF1 And i <= jsnx Then
                    JS1 = JS1 + 1
                    YJSJ1 = SJL1 * JS1
                    fhl_base_xdc(i) = (100 - YJSJ1) / 100
                End If
            Next
        End If
        '将大于计算年限的负荷率设置为0
        For i = 1 To 31 '收入税收表的列，第1种衰减率
            If i > jsnx Then
                fhl_base_xdc(i) = 0
            End If
        Next
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
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim sdsl_model As Integer = 1
        '收入和成本变化后相关计算
        Call 计算功能合并整理.收入成本相关计算(ExcelApp, sdsl_model)
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '清空已有内容
        Me.RichTextBox1.Rtf = Nothing
        Me.RichTextBox1.Clear()
        '显示此时的蓄电池逐年负荷率
        Dim SJ As Double
        Dim nf
        For i = 1 To 31  '根据数组中的元素数量循环
            SJ = Math.Round((fhl_base_xdc(i)) * 100, 2)
            nf = i
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "%(" & nf & ") " '输出到RichTextBox1
        Next
        Me.RichTextBox1.Text = "蓄电池供电逐年负荷率：" & Me.RichTextBox1.Text
    End Sub
    Private Sub 清空窗体_Click(sender As Object, e As EventArgs) Handles 清空窗体.Click
        Dim XZ = MsgBox("是否清空窗体中的全部内容？", vbOKCancel)
        If XZ = vbOK Then
            Me.TextBox1.Clear()
            Me.TextBox2.Clear()
            Me.TextBox3.Clear()
            Me.TextBox4.Clear()
            Me.TextBox5.Clear()
            Me.TextBox6.Clear()
            Me.TextBox7.Clear()
            Me.TextBox8.Clear()
            Me.TextBox9.Clear()
            Me.TextBox10.Clear()
            Me.gfjsnf.Clear()
            Me.xdcjsnf.Clear()
            Me.RichTextBox1.Rtf = Nothing
            Me.RichTextBox1.Clear()
        End If
    End Sub
    Private Sub 光伏发电默认系数_Click(sender As Object, e As EventArgs) Handles 光伏发电默认系数.Click
        On Error Resume Next
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
            Me.TextBox1.Clear()
            Me.TextBox2.Clear()
            Me.TextBox3.Clear()
            Me.TextBox4.Clear()
            Me.TextBox5.Clear()
            Me.TextBox6.Clear()
            Me.TextBox7.Clear()
            Me.TextBox8.Clear()
            Me.TextBox9.Clear()
            Me.TextBox10.Clear()
            Me.RichTextBox1.Rtf = Nothing
            Me.RichTextBox1.Clear()
            Me.TextBox1.Text = 2
            Me.TextBox2.Text = 3
            Me.TextBox4.Text = 2.5
            Me.TextBox5.Text = 0.7
            Me.gfjsnf.Text = CType(xmjsnx, String)
            Me.xdcksnf.Clear()
            Me.xdcjsnf.Clear()
            '——————————————————————————————————————————————————————————————————————————————————————————————
            '各种年限系数的计算开始年份（补贴收入、销售收入和成本）
            Dim JSKSNF As Integer = 0
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value > 0 Then
                    JSKSNF = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(156, i).Value
                    Exit For
                End If
            Next
            Me.gfksnf.Text = CType(JSKSNF, String)
        End If
    End Sub
    Private Sub 蓄电池默认系数_Click(sender As Object, e As EventArgs) Handles 蓄电池默认系数.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————               
        Dim XZ = MsgBox("是否加载蓄电池的默认逐年衰减系数和计算年限？", vbOKCancel)
        If XZ = vbOK Then
            Me.TextBox1.Clear()
            Me.TextBox2.Clear()
            Me.TextBox3.Clear()
            Me.TextBox4.Clear()
            Me.TextBox5.Clear()
            Me.TextBox6.Clear()
            Me.TextBox7.Clear()
            Me.TextBox8.Clear()
            Me.TextBox9.Clear()
            Me.TextBox10.Clear()
            Me.RichTextBox1.Rtf = Nothing
            Me.RichTextBox1.Clear()
            Me.TextBox1.Text = 3
            Me.TextBox4.Text = 1
            Me.xdcjsnf.Text = 11
            Me.gfksnf.Clear()
            Me.gfjsnf.Clear()
            '——————————————————————————————————————————————————————————————————————————————————————————————
            Dim JSKSNF As Integer = 0
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value > 0 Then
                    JSKSNF = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(156, i).Value
                    Exit For
                End If
            Next
            Me.xdcksnf.Text = CType(JSKSNF, String)
        End If
    End Sub

    Private Sub 蓄电池购电_Click(sender As Object, e As EventArgs) Handles 蓄电池购电.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————        
        '定义局部变量
        Dim SJKSNF1, SJKSNF2, SJKSNF3, SJKSNF4, SJKSNF5 As Integer '衰减开始年份
        Dim SJL1, SJL2, SJL3, SJL4, SJL5 As Double '衰减率
        Dim jsnx = CType(Me.xdcjsnf.Text, Integer) '蓄电池计算年限
        Dim ksnf = CType(Me.xdcksnf.Text, Integer) '蓄电池计算开始年份
        '读取项目计算年限
        Dim xmjsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————        
        SJKSNF1 = CType(Me.TextBox1.Text, Integer)
        SJKSNF2 = CType(Me.TextBox2.Text, Integer)
        SJKSNF3 = CType(Me.TextBox3.Text, Integer)
        SJKSNF4 = CType(Me.TextBox7.Text, Integer)
        SJKSNF5 = CType(Me.TextBox9.Text, Integer)
        SJL1 = CType(Me.TextBox4.Text, Double)
        SJL2 = CType(Me.TextBox5.Text, Double)
        SJL3 = CType(Me.TextBox6.Text, Double)
        SJL4 = CType(Me.TextBox8.Text, Double)
        SJL5 = CType(Me.TextBox10.Text, Double)
        '————————————————————————————————————————————————————————————————————————————————————————  
        '输入的蓄电池计算年限不可以大于项目计算年限
        If jsnx > xmjsnx Then
            MsgBox("窗口中输入的蓄电池计算年限不可以超过项目计算年限，请重新输入！")
            Exit Sub
        End If
        '检查输入的衰减开始年份
        If jsnx > 31 Then
            MsgBox("输入的项目计算年限不可以大于31年，请重新输入！")
            Exit Sub
        End If
        If SJKSNF1 > jsnx Or SJKSNF2 > jsnx Or SJKSNF3 > jsnx Or SJKSNF4 > jsnx Or SJKSNF5 > jsnx Then
            MsgBox("输入的衰减年份不可以大于项目计算年限，请重新输入！")
            Exit Sub
        End If
        If SJKSNF1 > 0 And SJKSNF2 > 0 And SJKSNF1 >= SJKSNF2 Then
            MsgBox("输入的第1个开始的衰减年份大于等于了第2个开始的衰减年份，请重新输入！")
            Exit Sub
        End If
        If SJKSNF1 > 0 And SJKSNF3 > 0 And SJKSNF1 >= SJKSNF3 Then
            MsgBox("输入的第1个开始的衰减年份大于等于了第3个开始的衰减年份，请重新输入！")
            Exit Sub
        End If
        If SJKSNF1 > 0 And SJKSNF4 > 0 And SJKSNF1 >= SJKSNF4 Then
            MsgBox("输入的第1个开始的衰减年份大于等于了第4个开始的衰减年份，请重新输入！")
            Exit Sub
        End If
        If SJKSNF1 > 0 And SJKSNF5 > 0 And SJKSNF1 >= SJKSNF5 Then
            MsgBox("输入的第1个开始的衰减年份大于等于了第5个开始的衰减年份，请重新输入！")
            Exit Sub
        End If
        If SJKSNF2 > 0 And SJKSNF3 > 0 And SJKSNF2 >= SJKSNF3 Then
            MsgBox("输入的第2个开始的衰减年份大于等于了第3个开始的衰减年份，请重新输入！")
            Exit Sub
        End If
        If SJKSNF2 > 0 And SJKSNF4 > 0 And SJKSNF2 >= SJKSNF4 Then
            MsgBox("输入的第2个开始的衰减年份大于等于了第4个开始的衰减年份，请重新输入！")
            Exit Sub
        End If
        If SJKSNF2 > 0 And SJKSNF5 > 0 And SJKSNF2 >= SJKSNF5 Then
            MsgBox("输入的第2个开始的衰减年份大于等于了第5个开始的衰减年份，请重新输入！")
            Exit Sub
        End If
        If SJKSNF2 > 0 And SJKSNF1 = 0 Then
            MsgBox("衰减开始年份必需从第1个年份开始输入，请重新输入！")
            Exit Sub
        End If
        If SJKSNF3 > 0 And SJKSNF1 = 0 Then
            MsgBox("衰减开始年份必需从第1个年份开始输入，请重新输入！")
            Exit Sub
        End If
        If SJKSNF4 > 0 And SJKSNF1 = 0 Then
            MsgBox("衰减开始年份必需从第1个年份开始输入，请重新输入！")
            Exit Sub
        End If
        If SJKSNF5 > 0 And SJKSNF1 = 0 Then
            MsgBox("衰减开始年份必需从第1个年份开始输入，请重新输入！")
            Exit Sub
        End If
        If SJKSNF3 > 0 And SJKSNF2 = 0 Then
            MsgBox("衰减开始年份必需从第1个和第2个年份开始输入，请重新输入！")
            Exit Sub
        End If
        If SJKSNF4 > 0 And SJKSNF2 = 0 Then
            MsgBox("衰减开始年份必需从第1个和第2个年份开始输入，请重新输入！")
            Exit Sub
        End If
        If SJKSNF5 > 0 And SJKSNF2 = 0 Then
            MsgBox("衰减开始年份必需从第1个和第2个年份开始输入，请重新输入！")
            Exit Sub
        End If
        If SJKSNF4 > 0 And SJKSNF3 = 0 Then
            MsgBox("衰减开始年份必需从第1个、第2个和第3个年份开始输入，请重新输入！")
            Exit Sub
        End If
        If SJKSNF5 > 0 And SJKSNF3 = 0 Then
            MsgBox("衰减开始年份必需从第1个、第2个和第3个年份开始输入，请重新输入！")
            Exit Sub
        End If
        If SJKSNF5 > 0 And SJKSNF4 = 0 Then
            MsgBox("衰减开始年份必需从第1个、第2个、第3个和第4个年份开始输入，请重新输入！")
            Exit Sub
        End If
        If SJKSNF1 < 0 Or SJKSNF2 < 0 Or SJKSNF3 < 0 Or SJKSNF4 < 0 Or SJKSNF5 < 0 Then
            MsgBox("输入的第1个、第2个、第3个、第4个和第5个开始的衰减年份不能小于0，请重新输入！")
            Exit Sub
        End If
        If SJKSNF1 < 1 Then
            MsgBox("输入的第1个开始的衰减年份最小只能是1，请重新输入！")
            Exit Sub
        End If
        '————————————————————————————————————————————————————————————————————————————————————————               
        Me.RichTextBox1.Clear()
        For i = 3 To 33 '清空已有的负荷率，防止出错
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(97, i).Value = 0
        Next
        '蓄电池购电基础负荷率
        Dim fhl_base_xdc(31) As Double
        '输入了5种不同的衰减系数
        If SJKSNF1 > 0 And SJKSNF2 > 0 And SJKSNF3 > 0 And SJKSNF4 > 0 And SJKSNF5 > 0 Then
            Dim YJSJ1 As Double = 0 '第一种衰减率已经衰减了多少，初始值为0
            Dim JS1 As Integer = 0 '计数，用于记录第一种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '小于第一个衰减开始年份的负荷率设置为1
                If i >= ksnf And i < SJKSNF1 Then
                    fhl_base_xdc(i) = 1
                End If
            Next
            For i = 1 To 31 '收入税收表的列，第1种衰减率
                If i >= SJKSNF1 And i < SJKSNF2 Then
                    JS1 = JS1 + 1
                    YJSJ1 = SJL1 * JS1
                    fhl_base_xdc(i) = (100 - YJSJ1) / 100
                End If
            Next
            Dim YJSJ2 As Double = 0 '第2种衰减率已经衰减了多少，初始值为0
            Dim JS2 As Integer = 0 '计数，用于记录第2种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '收入税收表的列，第2种衰减率
                If i >= SJKSNF2 And i < SJKSNF3 Then
                    JS2 = JS2 + 1
                    YJSJ2 = SJL2 * JS2
                    fhl_base_xdc(i) = (100 - YJSJ1 - YJSJ2) / 100
                End If
            Next
            Dim YJSJ3 As Double = 0 '第3种衰减率已经衰减了多少，初始值为0
            Dim JS3 As Integer = 0 '计数，用于记录第3种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '收入税收表的列，第3种衰减率
                If i >= SJKSNF3 And i < SJKSNF4 Then
                    JS3 = JS3 + 1
                    YJSJ3 = SJL3 * JS3
                    fhl_base_xdc(i) = (100 - YJSJ1 - YJSJ2 - YJSJ3) / 100
                End If
            Next
            Dim YJSJ4 As Double = 0 '第4种衰减率已经衰减了多少，初始值为0
            Dim JS4 As Integer = 0 '计数，用于记录第4种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '收入税收表的列，第4种衰减率
                If i >= SJKSNF4 And i < SJKSNF5 Then
                    JS4 = JS4 + 1
                    YJSJ4 = SJL4 * JS4
                    fhl_base_xdc(i) = (100 - YJSJ1 - YJSJ2 - YJSJ3 - YJSJ4) / 100
                End If
            Next
            Dim YJSJ5 As Double = 0 '第5种衰减率已经衰减了多少，初始值为0
            Dim JS5 As Integer = 0 '计数，用于记录第5种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '收入税收表的列，第4种衰减率
                If i >= SJKSNF5 And i <= jsnx Then
                    JS5 = JS5 + 1
                    YJSJ5 = SJL5 * JS5
                    fhl_base_xdc(i) = (100 - YJSJ1 - YJSJ2 - YJSJ3 - YJSJ4 - YJSJ5) / 100
                End If
            Next
        End If
        '输入了4种不同的衰减系数
        If SJKSNF1 > 0 And SJKSNF2 > 0 And SJKSNF3 > 0 And SJKSNF4 > 0 And SJKSNF5 = 0 Then
            Dim YJSJ1 As Double = 0 '第一种衰减率已经衰减了多少，初始值为0
            Dim JS1 As Integer = 0 '计数，用于记录第一种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '小于第一个衰减开始年份的负荷率设置为1
                If i >= ksnf And i < SJKSNF1 Then
                    fhl_base_xdc(i) = 1
                End If
            Next
            For i = 1 To 31 '收入税收表的列，第1种衰减率
                If i >= SJKSNF1 And i < SJKSNF2 Then
                    JS1 = JS1 + 1
                    YJSJ1 = SJL1 * JS1
                    fhl_base_xdc(i) = (100 - YJSJ1) / 100
                End If
            Next
            Dim YJSJ2 As Double = 0 '第2种衰减率已经衰减了多少，初始值为0
            Dim JS2 As Integer = 0 '计数，用于记录第2种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '收入税收表的列，第2种衰减率
                If i >= SJKSNF2 And i < SJKSNF3 Then
                    JS2 = JS2 + 1
                    YJSJ2 = SJL2 * JS2
                    fhl_base_xdc(i) = (100 - YJSJ1 - YJSJ2) / 100
                End If
            Next
            Dim YJSJ3 As Double = 0 '第3种衰减率已经衰减了多少，初始值为0
            Dim JS3 As Integer = 0 '计数，用于记录第3种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '收入税收表的列，第3种衰减率
                If i >= SJKSNF3 And i < SJKSNF4 Then
                    JS3 = JS3 + 1
                    YJSJ3 = SJL3 * JS3
                    fhl_base_xdc(i) = (100 - YJSJ1 - YJSJ2 - YJSJ3) / 100
                End If
            Next
            Dim YJSJ4 As Double = 0 '第4种衰减率已经衰减了多少，初始值为0
            Dim JS4 As Integer = 0 '计数，用于记录第4种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '收入税收表的列，第4种衰减率
                If i >= SJKSNF4 And i <= jsnx Then
                    JS4 = JS4 + 1
                    YJSJ4 = SJL4 * JS4
                    fhl_base_xdc(i) = (100 - YJSJ1 - YJSJ2 - YJSJ3 - YJSJ4) / 100
                End If
            Next
        End If
        '输入了3种不同的衰减系数
        If SJKSNF1 > 0 And SJKSNF2 > 0 And SJKSNF3 > 0 And SJKSNF4 = 0 And SJKSNF5 = 0 Then
            Dim YJSJ1 As Double = 0 '第一种衰减率已经衰减了多少，初始值为0
            Dim JS1 As Integer = 0 '计数，用于记录第一种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '小于第一个衰减开始年份的负荷率设置为1
                If i >= ksnf And i < SJKSNF1 Then
                    fhl_base_xdc(i) = 1
                End If
            Next
            For i = 1 To 31 '收入税收表的列，第1种衰减率
                If i >= SJKSNF1 And i < SJKSNF2 Then
                    JS1 = JS1 + 1
                    YJSJ1 = SJL1 * JS1
                    fhl_base_xdc(i) = (100 - YJSJ1) / 100
                End If
            Next
            Dim YJSJ2 As Double = 0 '第2种衰减率已经衰减了多少，初始值为0
            Dim JS2 As Integer = 0 '计数，用于记录第2种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '收入税收表的列，第2种衰减率
                If i >= SJKSNF2 And i < SJKSNF3 Then
                    JS2 = JS2 + 1
                    YJSJ2 = SJL2 * JS2
                    fhl_base_xdc(i) = (100 - YJSJ1 - YJSJ2) / 100
                End If
            Next
            Dim YJSJ3 As Double = 0 '第3种衰减率已经衰减了多少，初始值为0
            Dim JS3 As Integer = 0 '计数，用于记录第3种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '收入税收表的列，第2种衰减率
                If i >= SJKSNF3 And i <= jsnx Then
                    JS3 = JS3 + 1
                    YJSJ3 = SJL3 * JS3
                    fhl_base_xdc(i) = (100 - YJSJ1 - YJSJ2 - YJSJ3) / 100
                End If
            Next
        End If
        '输入了2种不同的衰减系数
        If SJKSNF1 > 0 And SJKSNF2 > 0 And SJKSNF3 = 0 And SJKSNF4 = 0 And SJKSNF5 = 0 Then
            Dim YJSJ1 As Double = 0 '第一种衰减率已经衰减了多少，初始值为0
            Dim JS1 As Integer = 0 '计数，用于记录第一种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '小于第一个衰减开始年份的负荷率设置为1
                If i >= ksnf And i < SJKSNF1 Then
                    fhl_base_xdc(i) = 1
                End If
            Next
            For i = 1 To 31 '收入税收表的列，第1种衰减率
                If i >= SJKSNF1 And i < SJKSNF2 Then
                    JS1 = JS1 + 1
                    YJSJ1 = SJL1 * JS1
                    fhl_base_xdc(i) = (100 - YJSJ1) / 100
                End If
            Next
            Dim YJSJ2 As Double = 0 '第2种衰减率已经衰减了多少，初始值为0
            Dim JS2 As Integer = 0 '计数，用于记录第2种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '收入税收表的列，第2种衰减率
                If i >= SJKSNF2 And i <= jsnx Then
                    JS2 = JS2 + 1
                    YJSJ2 = SJL2 * JS2
                    fhl_base_xdc(i) = (100 - YJSJ1 - YJSJ2) / 100
                End If
            Next
        End If
        '输入了1种不同的衰减系数
        If SJKSNF1 > 0 And SJKSNF2 = 0 And SJKSNF3 = 0 And SJKSNF4 = 0 And SJKSNF5 = 0 Then
            Dim YJSJ1 As Double = 0 '第一种衰减率已经衰减了多少，初始值为0
            Dim JS1 As Integer = 0 '计数，用于记录第一种衰减率已经衰减的年份数量，初始值为0
            For i = 1 To 31 '小于第一个衰减开始年份的负荷率设置为1
                If i >= ksnf And i < SJKSNF1 Then
                    fhl_base_xdc(i) = 1
                End If
            Next
            For i = 1 To 31 '收入税收表的列，第1种衰减率
                If i >= SJKSNF1 And i <= jsnx Then
                    JS1 = JS1 + 1
                    YJSJ1 = SJL1 * JS1
                    fhl_base_xdc(i) = (100 - YJSJ1) / 100
                End If
            Next
        End If
        '将大于计算年限的负荷率设置为0
        For i = 1 To 31 '收入税收表的列，第1种衰减率
            If i > jsnx Then
                fhl_base_xdc(i) = 0
            End If
        Next
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
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim sdsl_model As Integer = 1
        '收入和成本变化后相关计算
        Call 计算功能合并整理.收入成本相关计算(ExcelApp, sdsl_model)
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '清空已有内容
        Me.RichTextBox1.Rtf = Nothing
        Me.RichTextBox1.Clear()
        '显示此时的蓄电池逐年负荷率
        Dim SJ As Double
        Dim nf
        For i = 1 To 31  '根据数组中的元素数量循环
            SJ = Math.Round((fhl_base_xdc(i)) * 100, 2)
            nf = i
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & SJ & "%(" & nf & ") " '输出到RichTextBox1
        Next
        Me.RichTextBox1.Text = "蓄电池购电逐年负荷率：" & Me.RichTextBox1.Text
    End Sub
End Class