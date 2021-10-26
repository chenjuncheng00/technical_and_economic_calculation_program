Imports Microsoft.Office.Interop.Excel

Public Class 设置长期贷款计算方式
    Private Sub 设置长期贷款计算方式_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————
        '读取输入的默认值
        Dim cqdknx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value '长期贷款还款年限
        Dim kxnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value '长期贷款宽限年限
        '清空已有的全部数据
        '长期贷款还款年限
        Me.cqdkhknx1.Clear()
        Me.cqdkhknx2.Clear()
        Me.cqdkhknx3.Clear()
        Me.cqdkhknx4.Clear()
        Me.cqdkhknx5.Clear()
        Me.cqdkhknx6.Clear()
        Me.cqdkhknx7.Clear()
        Me.cqdkhknx8.Clear()
        Me.cqdkhknx9.Clear()
        Me.cqdkhknx10.Clear()
        '长期贷款宽限年限
        Me.kxnx1.Clear()
        Me.kxnx2.Clear()
        Me.kxnx3.Clear()
        Me.kxnx4.Clear()
        Me.kxnx5.Clear()
        Me.kxnx6.Clear()
        Me.kxnx7.Clear()
        Me.kxnx8.Clear()
        Me.kxnx9.Clear()
        Me.kxnx10.Clear()
        '显示
        Me.RichTextBox1.Clear()
        '——————————————————————————————————————————————————————————————————————————————
        '载入默认值
        '根据已经输入的投资情况，载入默认值
        '第1次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 3).Value > 0 Then
            '长期贷款还款年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(23, 39).Value > 0 Then
                Me.cqdkhknx1.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(23, 39).Value, String)
            Else
                Me.cqdkhknx1.Text = CType(cqdknx, String)
            End If
            '长期贷款宽限年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(34, 39).Value > 0 Then
                Me.kxnx1.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(34, 39).Value, String)
            Else
                Me.kxnx1.Text = CType(kxnx, String)
            End If
        End If
        '第2次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value > 0 Then
            '长期贷款还款年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(24, 39).Value > 0 Then
                Me.cqdkhknx2.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(24, 39).Value, String)
            Else
                Me.cqdkhknx2.Text = CType(cqdknx, String)
            End If
            '长期贷款宽限年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(35, 39).Value > 0 Then
                Me.kxnx2.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(35, 39).Value, String)
            Else
                Me.kxnx2.Text = CType(kxnx, String)
            End If
        End If
        '第3次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value > 0 Then
            '长期贷款还款年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(25, 39).Value > 0 Then
                Me.cqdkhknx3.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(25, 39).Value, String)
            Else
                Me.cqdkhknx3.Text = CType(cqdknx, String)
            End If
            '长期贷款宽限年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(36, 39).Value > 0 Then
                Me.kxnx3.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(36, 39).Value, String)
            Else
                Me.kxnx3.Text = CType(kxnx, String)
            End If
        End If
        '第4次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value > 0 Then
            '长期贷款还款年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(26, 39).Value > 0 Then
                Me.cqdkhknx4.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(26, 39).Value, String)
            Else
                Me.cqdkhknx4.Text = CType(cqdknx, String)
            End If
            '长期贷款宽限年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(37, 39).Value > 0 Then
                Me.kxnx4.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(37, 39).Value, String)
            Else
                Me.kxnx4.Text = CType(kxnx, String)
            End If
        End If
        '第5次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value > 0 Then
            '长期贷款还款年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(27, 39).Value > 0 Then
                Me.cqdkhknx5.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(27, 39).Value, String)
            Else
                Me.cqdkhknx5.Text = CType(cqdknx, String)
            End If
            '长期贷款宽限年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(38, 39).Value > 0 Then
                Me.kxnx5.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(38, 39).Value, String)
            Else
                Me.kxnx5.Text = CType(kxnx, String)
            End If
        End If
        '第6次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value > 0 Then
            '长期贷款还款年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(28, 39).Value > 0 Then
                Me.cqdkhknx6.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(28, 39).Value, String)
            Else
                Me.cqdkhknx6.Text = CType(cqdknx, String)
            End If
            '长期贷款宽限年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(39, 39).Value > 0 Then
                Me.kxnx6.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(39, 39).Value, String)
            Else
                Me.kxnx6.Text = CType(kxnx, String)
            End If
        End If
        '第7次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value > 0 Then
            '长期贷款还款年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(29, 39).Value > 0 Then
                Me.cqdkhknx7.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(29, 39).Value, String)
            Else
                Me.cqdkhknx7.Text = CType(cqdknx, String)
            End If
            '长期贷款宽限年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(40, 39).Value > 0 Then
                Me.kxnx7.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(40, 39).Value, String)
            Else
                Me.kxnx7.Text = CType(kxnx, String)
            End If
        End If
        '第8次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value > 0 Then
            '长期贷款还款年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(30, 39).Value > 0 Then
                Me.cqdkhknx8.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(30, 39).Value, String)
            Else
                Me.cqdkhknx8.Text = CType(cqdknx, String)
            End If
            '长期贷款宽限年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(41, 39).Value > 0 Then
                Me.kxnx8.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(41, 39).Value, String)
            Else
                Me.kxnx8.Text = CType(kxnx, String)
            End If
        End If
        '第9次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value > 0 Then
            '长期贷款还款年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(31, 39).Value > 0 Then
                Me.cqdkhknx9.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(31, 39).Value, String)
            Else
                Me.cqdkhknx9.Text = CType(cqdknx, String)
            End If
            '长期贷款宽限年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(42, 39).Value > 0 Then
                Me.kxnx9.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(42, 39).Value, String)
            Else
                Me.kxnx9.Text = CType(kxnx, String)
            End If
        End If
        '第10次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value > 0 Then
            '长期贷款还款年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(32, 39).Value > 0 Then
                Me.cqdkhknx10.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(32, 39).Value, String)
            Else
                Me.cqdkhknx10.Text = CType(cqdknx, String)
            End If
            '长期贷款宽限年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(43, 39).Value > 0 Then
                Me.kxnx10.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(43, 39).Value, String)
            Else
                Me.kxnx10.Text = CType(kxnx, String)
            End If
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '根据输入的投资情况，确定每个checkbox和投资比例输入是否可以选择和输入，并载入默认值
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(44, 1).Value > 0 Then
            Me.ranji.Enabled = True
            Me.ranji.Checked = False
        Else
            Me.ranji.Enabled = False
            Me.ranji.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(46, 1).Value > 0 Then
            Me.xudianchi.Enabled = True
            Me.xudianchi.Checked = False
        Else
            Me.xudianchi.Enabled = False
            Me.xudianchi.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(48, 1).Value > 0 Then
            Me.nuantong.Enabled = True
            Me.nuantong.Checked = False
        Else
            Me.nuantong.Enabled = False
            Me.nuantong.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(50, 1).Value > 0 Then
            Me.guangfu.Enabled = True
            Me.guangfu.Checked = False
        Else
            Me.guangfu.Enabled = False
            Me.guangfu.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(53, 1).Value > 0 Then
            Me.fengdian.Enabled = True
            Me.fengdian.Checked = False
        Else
            Me.fengdian.Enabled = False
            Me.fengdian.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
            Me.changgui.Enabled = True
            Me.changgui.Checked = False
        Else
            Me.changgui.Enabled = False
            Me.changgui.Checked = False
        End If
    End Sub

    Private Sub 清空窗体数据_Click(sender As Object, e As EventArgs) Handles 清空窗体.Click
        Dim XZ = MsgBox("确定要清空本窗体输入的的全部内容？", vbOKCancel)
        If XZ = vbOK Then
            '长期贷款还款年限
            Me.cqdkhknx1.Clear()
            Me.cqdkhknx2.Clear()
            Me.cqdkhknx3.Clear()
            Me.cqdkhknx4.Clear()
            Me.cqdkhknx5.Clear()
            Me.cqdkhknx6.Clear()
            Me.cqdkhknx7.Clear()
            Me.cqdkhknx8.Clear()
            Me.cqdkhknx9.Clear()
            Me.cqdkhknx10.Clear()
            '长期贷款宽限年限
            Me.kxnx1.Clear()
            Me.kxnx2.Clear()
            Me.kxnx3.Clear()
            Me.kxnx4.Clear()
            Me.kxnx5.Clear()
            Me.kxnx6.Clear()
            Me.kxnx7.Clear()
            Me.kxnx8.Clear()
            Me.kxnx9.Clear()
            Me.kxnx10.Clear()
            '选择
            Me.changgui.Checked = False
            Me.ranji.Checked = False
            Me.xudianchi.Checked = False
            Me.nuantong.Checked = False
            Me.guangfu.Checked = False
            Me.fengdian.Checked = False
            '显示
            Me.RichTextBox1.Clear()
        End If
    End Sub

    Private Sub 确定计算方式_Click(sender As Object, e As EventArgs) Handles 开始计算.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————
        ''长期贷款计算方法需要采用方法三或者方法四
        'If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法一" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法二" Then
        '    MsgBox("长期借款计算方式，请选择计算方法三或者方法四，计算终止！")
        '    Exit Sub
        'End If
        '——————————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否确认输入的逐次投资的长期贷款还款和宽限年限？？", vbOKCancel)
        If XZ = vbOK Then
            '计算一次工作簿
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
            Dim sdsl_model As Integer = 1
            '计算长期贷款
            Call 长期贷款相关计算(ExcelApp, sdsl_model)
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(167, 7).Value = "不相同"
            '显示
            Me.RichTextBox1.Text = "长期贷款计算完成！"
        End If
    End Sub

    Private Sub 重置回默认方式_Click(sender As Object, e As EventArgs) Handles 重置默认.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp      
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否将逐次投资的长期贷款还款年限和宽限年限重置回默认值？？", vbOKCancel)
        If XZ = vbOK Then
            '常规设备
            For i = 23 To 32
                '长期贷款还款年限
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 39).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value
            Next
            For i = 34 To 43
                '长期贷款宽限年限
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 39).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value
            Next
            '其它设备
            For i = 16 To 25
                '长期贷款年限
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 42).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 45).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 48).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 51).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 54).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value
            Next
            For i = 27 To 36
                '宽限年限
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 42).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 45).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 48).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 51).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 54).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value
            Next
            '————————————————————————————————————————————————————————————————————————————————————————
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(167, 7).Value = "相同"
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算一次工作簿
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
            Dim sdsl_model As Integer = 1
            '计算长期贷款
            Call 长期贷款相关计算(ExcelApp, sdsl_model)
            '显示
            Me.RichTextBox1.Text = "设置完成，每次投资的长期贷款还款年限和宽限年限均相同！"
        End If
    End Sub

    Private Sub 写入参数_Click(sender As Object, e As EventArgs) Handles 写入参数.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————
        ''长期贷款计算方法需要采用方法三或者方法四
        'If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法一" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法二" Then
        '    MsgBox("长期借款计算方式，请选择计算方法三或者方法四，计算终止！")
        '    Exit Sub
        'End If
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '读取输入的数据
        '长期贷款还款年限
        Dim cqdkhknx1 = CType(Me.cqdkhknx1.Text, Integer)
        Dim cqdkhknx2 = CType(Me.cqdkhknx2.Text, Integer)
        Dim cqdkhknx3 = CType(Me.cqdkhknx3.Text, Integer)
        Dim cqdkhknx4 = CType(Me.cqdkhknx4.Text, Integer)
        Dim cqdkhknx5 = CType(Me.cqdkhknx5.Text, Integer)
        Dim cqdkhknx6 = CType(Me.cqdkhknx6.Text, Integer)
        Dim cqdkhknx7 = CType(Me.cqdkhknx7.Text, Integer)
        Dim cqdkhknx8 = CType(Me.cqdkhknx8.Text, Integer)
        Dim cqdkhknx9 = CType(Me.cqdkhknx9.Text, Integer)
        Dim cqdkhknx10 = CType(Me.cqdkhknx10.Text, Integer)
        '长期贷款宽限年限
        Dim kxnx1 = CType(Me.kxnx1.Text, Integer)
        Dim kxnx2 = CType(Me.kxnx2.Text, Integer)
        Dim kxnx3 = CType(Me.kxnx3.Text, Integer)
        Dim kxnx4 = CType(Me.kxnx4.Text, Integer)
        Dim kxnx5 = CType(Me.kxnx5.Text, Integer)
        Dim kxnx6 = CType(Me.kxnx6.Text, Integer)
        Dim kxnx7 = CType(Me.kxnx7.Text, Integer)
        Dim kxnx8 = CType(Me.kxnx8.Text, Integer)
        Dim kxnx9 = CType(Me.kxnx9.Text, Integer)
        Dim kxnx10 = CType(Me.kxnx10.Text, Integer)
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '读取项目总的计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '将已有的数据清零
        For i = 23 To 32
            '长期贷款还款年限
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 39).Value = 0
        Next
        For i = 34 To 43
            '长期贷款宽限年限
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 39).Value = 0
        Next
        '确定每次的数据
        '第1次投资
        Dim tznf1 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 3).Value '投资年份1
        If tznf1 > 0 Then
            If tznf1 + cqdkhknx1 > jsnx Then
                MsgBox("输入的第1次投资的长期贷款还款年限超出允许范围，请检查并重新输入！")
                Exit Sub
            ElseIf cqdkhknx1 <= 0 Then
                MsgBox("输入的第1次投资的长期贷款还款年限不可以小于等于0，请检查并重新输入！")
                Exit Sub
            ElseIf kxnx1 <= 0 Then
                MsgBox("输入的第1次投资的长期贷款宽限年限不可以小于等于0，请检查并重新输入！")
                Exit Sub
            End If
        End If
        '第2次投资
        Dim tznf2 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value '投资年份2
        If tznf2 > 0 Then
            If tznf2 + cqdkhknx2 > jsnx Then
                MsgBox("输入的第2次投资的长期贷款还款年限超出允许范围，请检查并重新输入！")
                Exit Sub
            ElseIf cqdkhknx2 <= 0 Then
                MsgBox("输入的第2次投资的长期贷款还款年限不可以小于等于0，请检查并重新输入！")
                Exit Sub
            ElseIf kxnx2 <= 0 Then
                MsgBox("输入的第2次投资的长期贷款宽限年限不可以小于等于0，请检查并重新输入！")
                Exit Sub
            End If
        End If
        '第3次投资
        Dim tznf3 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value '投资年份3
        If tznf3 > 0 Then
            If tznf3 + cqdkhknx3 > jsnx Then
                MsgBox("输入的第3次投资的长期贷款还款年限超出允许范围，请检查并重新输入！")
                Exit Sub
            ElseIf cqdkhknx3 <= 0 Then
                MsgBox("输入的第3次投资的长期贷款还款年限不可以小于等于0，请检查并重新输入！")
                Exit Sub
            ElseIf kxnx3 <= 0 Then
                MsgBox("输入的第3次投资的长期贷款宽限年限不可以小于等于0，请检查并重新输入！")
                Exit Sub
            End If
        End If
        '第4次投资
        Dim tznf4 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value '投资年份4
        If tznf4 > 0 Then
            If tznf4 + cqdkhknx4 > jsnx Then
                MsgBox("输入的第4次投资的长期贷款还款年限超出允许范围，请检查并重新输入！")
                Exit Sub
            ElseIf cqdkhknx4 <= 0 Then
                MsgBox("输入的第4次投资的长期贷款还款年限不可以小于等于0，请检查并重新输入！")
                Exit Sub
            ElseIf kxnx4 <= 0 Then
                MsgBox("输入的第4次投资的长期贷款宽限年限不可以小于等于0，请检查并重新输入！")
                Exit Sub
            End If
        End If
        '第5次投资
        Dim tznf5 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value '投资年份5
        If tznf5 > 0 Then
            If tznf5 + cqdkhknx5 > jsnx Then
                MsgBox("输入的第5次投资的长期贷款还款年限超出允许范围，请检查并重新输入！")
                Exit Sub
            ElseIf cqdkhknx5 <= 0 Then
                MsgBox("输入的第5次投资的长期贷款还款年限不可以小于等于0，请检查并重新输入！")
                Exit Sub
            ElseIf kxnx5 <= 0 Then
                MsgBox("输入的第5次投资的长期贷款宽限年限不可以小于等于0，请检查并重新输入！")
                Exit Sub
            End If
        End If
        '第6次投资
        Dim tznf6 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value '投资年份6
        If tznf6 > 0 Then
            If tznf6 + cqdkhknx6 > jsnx Then
                MsgBox("输入的第6次投资的长期贷款还款年限超出允许范围，请检查并重新输入！")
                Exit Sub
            ElseIf cqdkhknx6 <= 0 Then
                MsgBox("输入的第6次投资的长期贷款还款年限不可以小于等于0，请检查并重新输入！")
                Exit Sub
            ElseIf kxnx6 <= 0 Then
                MsgBox("输入的第6次投资的长期贷款宽限年限不可以小于等于0，请检查并重新输入！")
                Exit Sub
            End If
        End If
        '第7次投资
        Dim tznf7 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value '投资年份7
        If tznf7 > 0 Then
            If tznf7 + cqdkhknx7 > jsnx Then
                MsgBox("输入的第7次投资的长期贷款还款年限超出允许范围，请检查并重新输入！")
                Exit Sub
            ElseIf cqdkhknx7 <= 0 Then
                MsgBox("输入的第7次投资的长期贷款还款年限不可以小于等于0，请检查并重新输入！")
                Exit Sub
            ElseIf kxnx7 <= 0 Then
                MsgBox("输入的第7次投资的长期贷款宽限年限不可以小于等于0，请检查并重新输入！")
                Exit Sub
            End If
        End If
        '第8次投资
        Dim tznf8 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value '投资年份8
        If tznf8 > 0 Then
            If tznf8 + cqdkhknx8 > jsnx Then
                MsgBox("输入的第8次投资的长期贷款还款年限超出允许范围，请检查并重新输入！")
                Exit Sub
            ElseIf cqdkhknx8 <= 0 Then
                MsgBox("输入的第8次投资的长期贷款还款年限不可以小于等于0，请检查并重新输入！")
                Exit Sub
            ElseIf kxnx8 <= 0 Then
                MsgBox("输入的第8次投资的长期贷款宽限年限不可以小于等于0，请检查并重新输入！")
                Exit Sub
            End If
        End If
        '第9次投资
        Dim tznf9 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value '投资年份9
        If tznf9 > 0 Then
            If tznf9 + cqdkhknx9 > jsnx Then
                MsgBox("输入的第9次投资的长期贷款还款年限超出允许范围，请检查并重新输入！")
                Exit Sub
            ElseIf cqdkhknx9 <= 0 Then
                MsgBox("输入的第9次投资的长期贷款还款年限不可以小于等于0，请检查并重新输入！")
                Exit Sub
            ElseIf kxnx9 <= 0 Then
                MsgBox("输入的第9次投资的长期贷款宽限年限不可以小于等于0，请检查并重新输入！")
                Exit Sub
            End If
        End If
        '第10次投资
        Dim tznf10 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value '投资年份10
        If tznf10 > 0 Then
            If tznf10 + cqdkhknx10 > jsnx Then
                MsgBox("输入的第10次投资的长期贷款还款年限超出允许范围，请检查并重新输入！")
                Exit Sub
            ElseIf cqdkhknx10 <= 0 Then
                MsgBox("输入的第10次投资的长期贷款还款年限不可以小于等于0，请检查并重新输入！")
                Exit Sub
            ElseIf kxnx10 <= 0 Then
                MsgBox("输入的第10次投资的长期贷款宽限年限不可以小于等于0，请检查并重新输入！")
                Exit Sub
            End If
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '常规设备
        If Me.changgui.Checked = True Then
            '第1次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(23, 39).Value = cqdkhknx1
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(34, 39).Value = kxnx1
            '第2次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(24, 39).Value = cqdkhknx2
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(35, 39).Value = kxnx2
            '第3次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(25, 39).Value = cqdkhknx3
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(36, 39).Value = kxnx3
            '第4次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(26, 39).Value = cqdkhknx4
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(37, 39).Value = kxnx4
            '第5次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(27, 39).Value = cqdkhknx5
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(38, 39).Value = kxnx5
            '第6次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(28, 39).Value = cqdkhknx6
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(39, 39).Value = kxnx6
            '第7次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(29, 39).Value = cqdkhknx7
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(40, 39).Value = kxnx7
            '第8次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(30, 39).Value = cqdkhknx8
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(41, 39).Value = kxnx8
            '第9次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(31, 39).Value = cqdkhknx9
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(42, 39).Value = kxnx9
            '第10次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(32, 39).Value = cqdkhknx10
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(43, 39).Value = kxnx10
            '显示
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<常规设备>长期贷款计算参数设置写入完成！"
        End If
        '燃机
        If Me.ranji.Checked = True Then
            '第1次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(16, 42).Value = cqdkhknx1
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(27, 42).Value = kxnx1
            '第2次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(17, 42).Value = cqdkhknx2
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(28, 42).Value = kxnx2
            '第3次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(18, 42).Value = cqdkhknx3
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(29, 42).Value = kxnx3
            '第4次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(19, 42).Value = cqdkhknx4
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(30, 42).Value = kxnx4
            '第5次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(20, 42).Value = cqdkhknx5
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(31, 42).Value = kxnx5
            '第6次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(21, 42).Value = cqdkhknx6
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(32, 42).Value = kxnx6
            '第7次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(22, 42).Value = cqdkhknx7
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(33, 42).Value = kxnx7
            '第8次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(23, 42).Value = cqdkhknx8
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(34, 42).Value = kxnx8
            '第9次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(24, 42).Value = cqdkhknx9
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(35, 42).Value = kxnx9
            '第10次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(25, 42).Value = cqdkhknx10
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(36, 42).Value = kxnx10
            '显示
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<燃机设备>长期贷款计算参数设置写入完成！"
        End If
        '蓄电池
        If Me.xudianchi.Checked = True Then
            '第1次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(16, 45).Value = cqdkhknx1
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(27, 45).Value = kxnx1
            '第2次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(17, 45).Value = cqdkhknx2
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(28, 45).Value = kxnx2
            '第3次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(18, 45).Value = cqdkhknx3
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(29, 45).Value = kxnx3
            '第4次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(19, 45).Value = cqdkhknx4
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(30, 45).Value = kxnx4
            '第5次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(20, 45).Value = cqdkhknx5
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(31, 45).Value = kxnx5
            '第6次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(21, 45).Value = cqdkhknx6
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(32, 45).Value = kxnx6
            '第7次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(22, 45).Value = cqdkhknx7
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(33, 45).Value = kxnx7
            '第8次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(23, 45).Value = cqdkhknx8
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(34, 45).Value = kxnx8
            '第9次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(24, 45).Value = cqdkhknx9
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(35, 45).Value = kxnx9
            '第10次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(25, 45).Value = cqdkhknx10
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(36, 45).Value = kxnx10
            '显示
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<蓄电池设备>长期贷款计算参数设置写入完成！"
        End If
        '暖通
        If Me.nuantong.Checked = True Then
            '第1次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(16, 48).Value = cqdkhknx1
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(27, 48).Value = kxnx1
            '第2次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(17, 48).Value = cqdkhknx2
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(28, 48).Value = kxnx2
            '第3次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(18, 48).Value = cqdkhknx3
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(29, 48).Value = kxnx3
            '第4次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(19, 48).Value = cqdkhknx4
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(30, 48).Value = kxnx4
            '第5次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(20, 48).Value = cqdkhknx5
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(31, 48).Value = kxnx5
            '第6次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(21, 48).Value = cqdkhknx6
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(32, 48).Value = kxnx6
            '第7次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(22, 48).Value = cqdkhknx7
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(33, 48).Value = kxnx7
            '第8次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(23, 48).Value = cqdkhknx8
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(34, 48).Value = kxnx8
            '第9次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(24, 48).Value = cqdkhknx9
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(35, 48).Value = kxnx9
            '第10次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(25, 48).Value = cqdkhknx10
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(36, 48).Value = kxnx10
            '显示
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<暖通设备>长期贷款计算参数设置写入完成！"
        End If
        '光伏
        If Me.guangfu.Checked = True Then
            '第1次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(16, 51).Value = cqdkhknx1
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(27, 51).Value = kxnx1
            '第2次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(17, 51).Value = cqdkhknx2
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(28, 51).Value = kxnx2
            '第3次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(18, 51).Value = cqdkhknx3
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(29, 51).Value = kxnx3
            '第4次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(19, 51).Value = cqdkhknx4
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(30, 51).Value = kxnx4
            '第5次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(20, 51).Value = cqdkhknx5
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(31, 51).Value = kxnx5
            '第6次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(21, 51).Value = cqdkhknx6
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(32, 51).Value = kxnx6
            '第7次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(22, 51).Value = cqdkhknx7
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(33, 51).Value = kxnx7
            '第8次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(23, 51).Value = cqdkhknx8
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(34, 51).Value = kxnx8
            '第9次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(24, 51).Value = cqdkhknx9
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(35, 51).Value = kxnx9
            '第10次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(25, 51).Value = cqdkhknx10
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(36, 51).Value = kxnx10
            '显示
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<光伏设备>长期贷款计算参数设置写入完成！"
        End If
        '风电
        If Me.fengdian.Checked = True Then
            '第1次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(16, 54).Value = cqdkhknx1
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(27, 54).Value = kxnx1
            '第2次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(17, 54).Value = cqdkhknx2
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(28, 54).Value = kxnx2
            '第3次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(18, 54).Value = cqdkhknx3
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(29, 54).Value = kxnx3
            '第4次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(19, 54).Value = cqdkhknx4
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(30, 54).Value = kxnx4
            '第5次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(20, 54).Value = cqdkhknx5
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(31, 54).Value = kxnx5
            '第6次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(21, 54).Value = cqdkhknx6
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(32, 54).Value = kxnx6
            '第7次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(22, 54).Value = cqdkhknx7
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(33, 54).Value = kxnx7
            '第8次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(23, 54).Value = cqdkhknx8
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(34, 54).Value = kxnx8
            '第9次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(24, 54).Value = cqdkhknx9
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(35, 54).Value = kxnx9
            '第10次投资
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(25, 54).Value = cqdkhknx10
            ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(36, 54).Value = kxnx10
            '显示
            Me.RichTextBox1.Text = Me.RichTextBox1.Text & "<风电设备>长期贷款计算参数设置写入完成！"
        End If
        '常规设备
        For i = 23 To 32
            '长期贷款还款年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 39).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 39).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value
            End If
        Next
        For i = 34 To 43
            '长期贷款宽限年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 39).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 39).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value
            End If
        Next
        '其它设备
        For i = 16 To 25
            '长期贷款年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 42).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 42).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value
            End If
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 45).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 45).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value
            End If
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 48).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 48).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value
            End If
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 51).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 51).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value
            End If
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 54).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 54).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value
            End If
        Next
        For i = 27 To 36
            '宽限年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 42).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 42).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value
            End If
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 45).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 45).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value
            End If
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 48).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 48).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value
            End If
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 51).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 51).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value
            End If
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 54).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 54).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value
            End If
        Next
    End Sub
End Class