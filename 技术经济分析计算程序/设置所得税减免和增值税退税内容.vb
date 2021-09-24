Imports Microsoft.Office.Interop.Excel

Public Class 设置所得税减免和增值税退税内容
    Private Sub 设置所得税减免和增值税退税内容_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '清空全部已选
        '收入
        Me.收入1.Checked = False
        Me.收入2.Checked = False
        Me.收入3.Checked = False
        Me.收入4.Checked = False
        Me.收入5.Checked = False
        Me.收入6.Checked = False
        Me.收入7.Checked = False
        Me.收入8.Checked = False
        Me.收入9.Checked = False
        '清空已有投资
        Me.投资1.Text = Nothing
        Me.投资2.Text = Nothing
        Me.投资3.Text = Nothing
        Me.投资4.Text = Nothing
        Me.投资5.Text = Nothing
        Me.投资6.Text = Nothing
        Me.投资7.Text = Nothing
        Me.投资8.Text = Nothing
        Me.投资9.Text = Nothing
        '成本
        Me.成本1.Checked = False
        Me.成本2.Checked = False
        Me.成本3.Checked = False
        Me.成本4.Checked = False
        Me.成本5.Checked = False
        Me.成本6.Checked = False
        Me.成本7.Checked = False
        Me.成本8.Checked = False
        Me.成本9.Checked = False
        Me.成本10.Checked = False
        Me.成本11.Checked = False
        '清空已有收入比例
        Me.收入比例1.Text = Nothing
        Me.收入比例2.Text = Nothing
        Me.收入比例3.Text = Nothing
        Me.收入比例4.Text = Nothing
        Me.收入比例5.Text = Nothing
        Me.收入比例6.Text = Nothing
        Me.收入比例7.Text = Nothing
        Me.收入比例8.Text = Nothing
        Me.收入比例9.Text = Nothing
        '清空已有成本比例
        Me.成本比例1.Text = Nothing
        Me.成本比例2.Text = Nothing
        Me.成本比例3.Text = Nothing
        Me.成本比例4.Text = Nothing
        Me.成本比例5.Text = Nothing
        Me.成本比例6.Text = Nothing
        Me.成本比例7.Text = Nothing
        Me.成本比例8.Text = Nothing
        Me.成本比例9.Text = Nothing
        Me.成本比例10.Text = Nothing
        Me.成本比例11.Text = Nothing
        '专项投资和年量
        Me.专项投资1.Checked = False
        Me.专项投资2.Checked = False
        Me.专项投资3.Checked = False
        Me.专项年量1.Checked = False
        Me.专项年量2.Checked = False
        Me.专项年量3.Checked = False
        Me.专项年量4.Checked = False
        Me.专项年量5.Checked = False
        '专项投资和年量减去比例
        Me.专项投资减去比例1.Text = Nothing
        Me.专项投资减去比例2.Text = Nothing
        Me.专项投资减去比例3.Text = Nothing
        Me.专项年量减去比例1.Text = Nothing
        Me.专项年量减去比例2.Text = Nothing
        Me.专项年量减去比例3.Text = Nothing
        Me.专项年量减去比例4.Text = Nothing
        Me.专项年量减去比例5.Text = Nothing
        '光伏和风电所得税减免参数
        Me.光伏所得税免征年限.Text = Nothing
        Me.风电所得税免征年限.Text = Nothing
        Me.光伏所得税减少年限.Text = Nothing
        Me.风电所得税减少年限.Text = Nothing
        Me.光伏所得税减少比例.Text = Nothing
        Me.风电所得税减少比例.Text = Nothing
        '光伏和风电增值税退税参数
        Me.光伏增值税退税年限.Text = Nothing
        Me.风电增值税退税年限.Text = Nothing
        Me.光伏增值税退税比例.Text = Nothing
        Me.风电增值税退税比例.Text = Nothing
        '光伏和风电人员工资
        Me.光伏人员工资.Text = Nothing
        Me.风电人员工资.Text = Nothing
        '其它所得税减免和增值税退税参数
        Me.其它所得税免征年限.Text = Nothing
        Me.其它所得税减少年限.Text = Nothing
        Me.其它所得税减少比例.Text = Nothing
        Me.其它增值税退税年限.Text = Nothing
        Me.其它增值税退税比例.Text = Nothing
        '清空计算结果
        Me.RichTextBox1.Rtf = Nothing
        Me.RichTextBox1.Clear()
        '载入文字
        '收入
        Me.收入4.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(13, 2).Value, String) '工业蒸汽收入
        Me.收入1.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 2).Value, String) '供电收入
        Me.收入2.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 2).Value, String) '供冷收入
        Me.收入3.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 2).Value, String) '供热收入        
        Me.收入5.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 2).Value, String) '光伏发电收入
        Me.收入6.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(20, 2).Value, String) '光伏补贴收入
        Me.收入7.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 2).Value, String) '风力发电收入
        Me.收入8.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 2).Value, String) '垃圾处理补贴收入
        Me.收入9.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 2).Value, String) '生物质补贴收入
        '成本
        Me.成本1.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 9).Value, String) '天然气成本
        Me.成本2.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 9).Value, String) '煤炭成本
        Me.成本3.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 9).Value, String) '购电成本
        Me.成本10.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 9).Value, String) '外购蒸汽
        Me.成本11.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 9).Value, String) '购电容量费
        Me.成本4.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 9).Value, String) '补水成本
        Me.成本5.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 9).Value, String) '石灰石
        Me.成本6.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 9).Value, String) '生石灰
        Me.成本7.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 9).Value, String) '尿素
        Me.成本9.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 9).Value, String) '城市管廊
        Me.成本8.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 9).Value, String) '人员工资
        '载入光伏发电和风力发电投资（如果有的话）
        Me.投资5.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(50, 1).Value, String) '光伏发电
        Me.投资6.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(50, 1).Value, String) '光伏发电
        Me.投资7.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(53, 1).Value, String) '风力发电
        '载入光伏风电所得税减免参数默认值
        Me.光伏所得税免征年限.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(11, 7).Value, String)
        Me.风电所得税免征年限.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(11, 7).Value, String)
        Me.光伏所得税减少年限.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(12, 7).Value, String)
        Me.风电所得税减少年限.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(12, 7).Value, String)
        Me.光伏所得税减少比例.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(14, 7).Value, String) * 100
        Me.风电所得税减少比例.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(14, 7).Value, String) * 100
        '载入光伏和风电增值税退税参数默认值
        Me.光伏增值税退税年限.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value, String)
        Me.风电增值税退税年限.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value, String)
        Me.光伏增值税退税比例.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(11, 5).Value, String) * 100
        Me.风电增值税退税比例.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(11, 5).Value, String) * 100
        '载入其它所得税减免和增值税退税参数默认值
        Me.其它所得税免征年限.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(11, 7).Value, String)
        Me.其它所得税减少年限.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(12, 7).Value, String)
        Me.其它所得税减少比例.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(14, 7).Value, String) * 100
        Me.其它增值税退税年限.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value, String)
        Me.其它增值税退税比例.Text = CType(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(11, 5).Value, String) * 100
        '设置存在收入的可以选择
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(13, 7).Value > 0 Then
            Me.收入4.Enabled = True
            Me.投资4.Enabled = True
            Me.收入比例4.Enabled = True
            Me.收入比例4.Text = 100 '比例默认值
        Else
            Me.收入4.Enabled = False
            Me.投资4.Enabled = False
            Me.收入比例4.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 7).Value > 0 Then
            Me.收入1.Enabled = True
            Me.投资1.Enabled = True
            Me.收入比例1.Enabled = True
            Me.收入比例1.Text = 100 '比例默认值
        Else
            Me.收入1.Enabled = False
            Me.投资1.Enabled = False
            Me.收入比例1.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 7).Value > 0 Then
            Me.收入2.Enabled = True
            Me.投资2.Enabled = True
            Me.收入比例2.Enabled = True
            Me.收入比例2.Text = 100 '比例默认值
        Else
            Me.收入2.Enabled = False
            Me.投资2.Enabled = False
            Me.收入比例2.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 7).Value > 0 Then
            Me.收入3.Enabled = True
            Me.投资3.Enabled = True
            Me.收入比例3.Enabled = True
            Me.收入比例3.Text = 100 '比例默认值
        Else
            Me.收入3.Enabled = False
            Me.投资3.Enabled = False
            Me.收入比例3.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 7).Value > 0 Then
            Me.收入5.Enabled = True
            Me.投资5.Enabled = True
            Me.收入比例5.Enabled = True
            Me.收入比例5.Text = 100 '比例默认值
        Else
            Me.收入5.Enabled = False
            Me.投资5.Enabled = False
            Me.投资5.Text = Nothing
            Me.收入比例5.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(20, 7).Value > 0 Then
            Me.收入6.Enabled = True
            Me.投资6.Enabled = True
            Me.收入比例6.Enabled = True
            Me.收入比例6.Text = 100 '比例默认值
        Else
            Me.收入6.Enabled = False
            Me.投资6.Enabled = False
            Me.投资6.Text = Nothing
            Me.收入比例6.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 7).Value > 0 Or ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(20, 7).Value > 0 Then
            Me.光伏所得税免征年限.Enabled = True
            Me.光伏所得税减少年限.Enabled = True
            Me.光伏所得税减少比例.Enabled = True
            Me.光伏增值税退税年限.Enabled = True
            Me.光伏增值税退税比例.Enabled = True
            Me.光伏人员工资.Enabled = True
        Else
            Me.光伏所得税免征年限.Enabled = False
            Me.光伏所得税减少年限.Enabled = False
            Me.光伏所得税减少比例.Enabled = False
            Me.光伏增值税退税年限.Enabled = False
            Me.光伏增值税退税比例.Enabled = False
            Me.光伏人员工资.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 7).Value > 0 Then
            Me.收入7.Enabled = True
            Me.投资7.Enabled = True
            Me.收入比例7.Enabled = True
            Me.收入比例7.Text = 100 '比例默认值
            Me.风电所得税免征年限.Enabled = True
            Me.风电所得税减少年限.Enabled = True
            Me.风电所得税减少比例.Enabled = True
            Me.风电增值税退税年限.Enabled = True
            Me.风电增值税退税比例.Enabled = True
            Me.风电人员工资.Enabled = True
        Else
            Me.收入7.Enabled = False
            Me.投资7.Enabled = False
            Me.收入比例7.Enabled = False
            Me.风电所得税免征年限.Enabled = False
            Me.风电所得税减少年限.Enabled = False
            Me.风电所得税减少比例.Enabled = False
            Me.风电增值税退税年限.Enabled = False
            Me.风电增值税退税比例.Enabled = False
            Me.风电人员工资.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 7).Value > 0 Then
            Me.收入8.Enabled = True
            Me.投资8.Enabled = True
            Me.收入比例8.Enabled = True
            Me.收入比例8.Text = 100 '比例默认值
        Else
            Me.收入8.Enabled = False
            Me.投资8.Enabled = False
            Me.收入比例8.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 7).Value > 0 Then
            Me.收入9.Enabled = True
            Me.投资9.Enabled = True
            Me.收入比例9.Enabled = True
            Me.收入比例9.Text = 100 '比例默认值
        Else
            Me.收入9.Enabled = False
            Me.投资9.Enabled = False
            Me.收入比例9.Enabled = False
        End If
        '设置存在成本的可以选择
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 14).Value > 0 Then
            Me.成本1.Enabled = True
            Me.成本比例1.Enabled = True
            Me.成本比例1.Text = 100 '比例默认值
        Else
            Me.成本1.Enabled = False
            Me.成本比例1.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 14).Value > 0 Then
            Me.成本2.Enabled = True
            Me.成本比例2.Enabled = True
            Me.成本比例2.Text = 100 '比例默认值
        Else
            Me.成本2.Enabled = False
            Me.成本比例2.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 14).Value > 0 Then
            Me.成本3.Enabled = True
            Me.成本比例3.Enabled = True
            Me.成本比例3.Text = 100 '比例默认值
        Else
            Me.成本3.Enabled = False
            Me.成本比例3.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 14).Value > 0 Then
            Me.成本10.Enabled = True
            Me.成本比例10.Enabled = True
            Me.成本比例10.Text = 100 '比例默认值
        Else
            Me.成本10.Enabled = False
            Me.成本比例10.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 14).Value > 0 Then
            Me.成本11.Enabled = True
            Me.成本比例11.Enabled = True
            Me.成本比例11.Text = 100 '比例默认值
        Else
            Me.成本11.Enabled = False
            Me.成本比例11.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 14).Value > 0 Then
            Me.成本4.Enabled = True
            Me.成本比例4.Enabled = True
            Me.成本比例4.Text = 100 '比例默认值
        Else
            Me.成本4.Enabled = False
            Me.成本比例4.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 14).Value > 0 Then
            Me.成本5.Enabled = True
            Me.成本比例5.Enabled = True
            Me.成本比例5.Text = 100 '比例默认值
        Else
            Me.成本5.Enabled = False
            Me.成本比例5.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 14).Value > 0 Then
            Me.成本6.Enabled = True
            Me.成本比例6.Enabled = True
            Me.成本比例6.Text = 100 '比例默认值
        Else
            Me.成本6.Enabled = False
            Me.成本比例6.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 14).Value > 0 Then
            Me.成本7.Enabled = True
            Me.成本比例7.Enabled = True
            Me.成本比例7.Text = 100 '比例默认值
        Else
            Me.成本7.Enabled = False
            Me.成本比例7.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 14).Value > 0 Then
            Me.成本9.Enabled = True
            Me.成本比例9.Enabled = True
            Me.成本比例9.Text = 100 '比例默认值
        Else
            Me.成本9.Enabled = False
            Me.成本比例9.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 14).Value > 0 Then
            Me.成本8.Enabled = True
            Me.成本比例8.Enabled = True
            Me.成本比例8.Text = 100 '比例默认值
        Else
            Me.成本8.Enabled = False
            Me.成本比例8.Enabled = False
        End If
        '设置专项投资是否可以勾选
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(46, 1).Value > 0 Then
            Me.专项投资1.Enabled = True
            Me.专项投资减去比例1.Enabled = True
            Me.专项投资减去比例1.Text = 100 '比例默认值
        Else
            Me.专项投资1.Enabled = False
            Me.专项投资减去比例1.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(44, 1).Value > 0 Then
            Me.专项投资2.Enabled = True
            Me.专项投资减去比例2.Enabled = True
            Me.专项投资减去比例2.Text = 100 '比例默认值
        Else
            Me.专项投资2.Enabled = False
            Me.专项投资减去比例2.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(48, 1).Value > 0 Then
            Me.专项投资3.Enabled = True
            Me.专项投资减去比例3.Enabled = True
            Me.专项投资减去比例3.Text = 100 '比例默认值
        Else
            Me.专项投资3.Enabled = False
            Me.专项投资减去比例3.Enabled = False
        End If
        '设置专项年量是否可选
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(47, 1).Value > 0 Then
            Me.专项年量1.Enabled = True
            Me.专项年量减去比例1.Enabled = True
            Me.专项年量减去比例1.Text = 100 '比例默认值
        Else
            Me.专项年量1.Enabled = False
            Me.专项年量减去比例1.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(45, 1).Value > 0 Then
            Me.专项年量2.Enabled = True
            Me.专项年量减去比例2.Enabled = True
            Me.专项年量减去比例2.Text = 100 '比例默认值
        Else
            Me.专项年量2.Enabled = False
            Me.专项年量减去比例2.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(49, 1).Value > 0 Then
            Me.专项年量3.Enabled = True
            Me.专项年量减去比例3.Enabled = True
            Me.专项年量减去比例3.Text = 100 '比例默认值
        Else
            Me.专项年量3.Enabled = False
            Me.专项年量减去比例3.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(52, 1).Value > 0 Then
            Me.专项年量4.Enabled = True
            Me.专项年量减去比例4.Enabled = True
            Me.专项年量减去比例4.Text = 100 '比例默认值
        Else
            Me.专项年量4.Enabled = False
            Me.专项年量减去比例4.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(55, 1).Value > 0 Then
            Me.专项年量5.Enabled = True
            Me.专项年量减去比例5.Enabled = True
            Me.专项年量减去比例5.Text = 100 '比例默认值
        Else
            Me.专项年量5.Enabled = False
            Me.专项年量减去比例5.Enabled = False
        End If
    End Sub

    Private Sub 所得税减免_Click(sender As Object, e As EventArgs) Handles 所得税减免.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '读取输入的投资
        Dim srtz1 As Double = CType(Me.投资1.Text, Double)
        Dim srtz2 As Double = CType(Me.投资2.Text, Double)
        Dim srtz3 As Double = CType(Me.投资3.Text, Double)
        Dim srtz4 As Double = CType(Me.投资4.Text, Double)
        Dim srtz5 As Double = CType(Me.投资5.Text, Double)
        Dim srtz6 As Double = CType(Me.投资6.Text, Double)
        Dim srtz7 As Double = CType(Me.投资7.Text, Double)
        Dim srtz8 As Double = CType(Me.投资8.Text, Double)
        Dim srtz9 As Double = CType(Me.投资9.Text, Double)
        '读取输入的专项投资和年量减去比例
        Dim zxtzjqbl1 As Double = CType(Me.专项投资减去比例1.Text, Double) / 100
        Dim zxtzjqbl2 As Double = CType(Me.专项投资减去比例2.Text, Double) / 100
        Dim zxtzjqbl3 As Double = CType(Me.专项投资减去比例3.Text, Double) / 100
        Dim zxzljqbl1 As Double = CType(Me.专项年量减去比例1.Text, Double) / 100
        Dim zxzljqbl2 As Double = CType(Me.专项年量减去比例2.Text, Double) / 100
        Dim zxzljqbl3 As Double = CType(Me.专项年量减去比例3.Text, Double) / 100
        Dim zxzljqbl4 As Double = CType(Me.专项年量减去比例4.Text, Double) / 100
        Dim zxzljqbl5 As Double = CType(Me.专项年量减去比例5.Text, Double) / 100
        '读取输入的各个收入占比比例
        Dim srbl1 As Double = CType(Me.收入比例1.Text, Double) / 100
        Dim srbl2 As Double = CType(Me.收入比例2.Text, Double) / 100
        Dim srbl3 As Double = CType(Me.收入比例3.Text, Double) / 100
        Dim srbl4 As Double = CType(Me.收入比例4.Text, Double) / 100
        Dim srbl5 As Double = CType(Me.收入比例5.Text, Double) / 100
        Dim srbl6 As Double = CType(Me.收入比例6.Text, Double) / 100
        Dim srbl7 As Double = CType(Me.收入比例7.Text, Double) / 100
        Dim srbl8 As Double = CType(Me.收入比例8.Text, Double) / 100
        Dim srbl9 As Double = CType(Me.收入比例9.Text, Double) / 100
        '读取输入的各个成本占比比例
        Dim cbbl1 As Double = CType(Me.成本比例1.Text, Double) / 100
        Dim cbbl2 As Double = CType(Me.成本比例2.Text, Double) / 100
        Dim cbbl3 As Double = CType(Me.成本比例3.Text, Double) / 100
        Dim cbbl4 As Double = CType(Me.成本比例4.Text, Double) / 100
        Dim cbbl5 As Double = CType(Me.成本比例5.Text, Double) / 100
        Dim cbbl6 As Double = CType(Me.成本比例6.Text, Double) / 100
        Dim cbbl7 As Double = CType(Me.成本比例7.Text, Double) / 100
        Dim cbbl8 As Double = CType(Me.成本比例8.Text, Double) / 100
        Dim cbbl9 As Double = CType(Me.成本比例9.Text, Double) / 100
        Dim cbbl10 As Double = CType(Me.成本比例10.Text, Double) / 100
        Dim cbbl11 As Double = CType(Me.成本比例11.Text, Double) / 100
        '读取设置的光伏和风力发电所得税减免参数、增值税退税参数等
        Dim gfsdsmznx As Integer = CType(Me.光伏所得税免征年限.Text, Integer)
        Dim gfsdsjsnx As Integer = CType(Me.光伏所得税减少年限.Text, Integer)
        Dim gfsdsjsbl As Double = CType(Me.光伏所得税减少比例.Text, Double) / 100
        Dim gfzzstsnx As Integer = CType(Me.光伏增值税退税年限.Text, Integer)
        Dim gfzzstsbl As Double = CType(Me.光伏增值税退税比例.Text, Double) / 100
        Dim gfrygz As Double = CType(Me.光伏人员工资.Text, Double)
        Dim fdsdsmznx As Integer = CType(Me.风电所得税免征年限.Text, Integer)
        Dim fdsdsjsnx As Integer = CType(Me.风电所得税减少年限.Text, Integer)
        Dim fdsdsjsbl As Double = CType(Me.风电所得税减少比例.Text, Double) / 100
        Dim fdzzstsnx As Integer = CType(Me.风电增值税退税年限.Text, Integer)
        Dim fdzzstsbl As Double = CType(Me.风电增值税退税比例.Text, Double) / 100
        Dim fdrygz As Double = CType(Me.风电人员工资.Text, Double)
        '读取输入的其它部分所得税减免和增值税退税参数
        Dim qtsdsmznx As Integer = CType(Me.其它所得税免征年限.Text, Integer)
        Dim qtsdsjsnx As Integer = CType(Me.其它所得税减少年限.Text, Integer)
        Dim qtsdsjsbl As Double = CType(Me.其它所得税减少比例.Text, Double) / 100
        Dim qtzzstsnx As Integer = CType(Me.其它增值税退税年限.Text, Integer)
        Dim qtzzstsbl As Double = CType(Me.其它增值税退税比例.Text, Double) / 100
        '记录目前各个收入的单价
        Dim srdj4 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(13, 5).Value, Double)
        Dim srdj1 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 5).Value, Double)
        Dim srdj2 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 5).Value, Double)
        Dim srdj3 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 5).Value, Double)
        Dim srdj5 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 5).Value, Double)
        Dim srdj6 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(20, 5).Value, Double)
        Dim srdj7 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 5).Value, Double)
        Dim srdj8 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 5).Value, Double)
        Dim srdj9 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(29, 5).Value, Double)
        '记录目前各个成本的单价
        Dim cbdj1 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 12).Value, Double)
        Dim cbdj2 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 12).Value, Double)
        Dim cbdj3 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 12).Value, Double)
        Dim cbdj10 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 12).Value, Double)
        Dim cbdj11 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 12).Value, Double)
        Dim cbdj4 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 12).Value, Double)
        Dim cbdj5 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 12).Value, Double)
        Dim cbdj6 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 12).Value, Double)
        Dim cbdj7 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 12).Value, Double)
        Dim cbdj9 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 12).Value, Double)
        Dim cbdj8 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 12).Value, Double)
        '————————————————————————————————————————————————————————————————————————————————————————————
        '针对输入和选择的内容，添加报错功能
        '光伏
        If Me.收入5.Checked = True Or Me.收入6.Checked = True Then
            If gfsdsmznx + gfsdsjsnx = 0 Then
                MsgBox("光伏所得税免征年限和光伏所得税减征年限不可以都是0，请重新输入！")
                Exit Sub
            End If
            If gfsdsjsnx > 0 And gfsdsjsbl = 0 Then
                MsgBox("光伏所得税减征年限不是0，但光伏所得税减征比例是0，请重新输入！")
                Exit Sub
            End If
        End If
        '风电
        If Me.收入7.Checked = True Then
            If fdsdsmznx + fdsdsjsnx = 0 Then
                MsgBox("风电所得税免征年限和风电所得税减征年限不可以都是0，请重新输入！")
                Exit Sub
            End If
            If fdsdsjsnx > 0 And fdsdsjsbl = 0 Then
                MsgBox("风电所得税减征年限不是0，但风电所得税减征比例是0，请重新输入！")
                Exit Sub
            End If
        End If
        '其它
        If Me.收入5.Checked = False And Me.收入6.Checked = False And Me.收入7.Checked = False Then
            If qtsdsmznx + qtsdsjsnx = 0 Then
                MsgBox("其它所得税免征年限和其它所得税减征年限不可以都是0，请重新输入！")
                Exit Sub
            End If
            If qtsdsjsnx > 0 And qtsdsjsbl = 0 Then
                MsgBox("其它所得税减征年限不是0，但其它所得税减征比例是0，请重新输入！")
                Exit Sub
            End If
        End If
        '————————————————————————————————————————————————————————————————————————————————————————————
        'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
        Dim zbj_model As Integer
        If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 7).Value = "动态" Then
            zbj_model = 0
        Else
            zbj_model = 1
        End If
        'hscy：计算期末，是否回收资产残值
        Dim hscz As Boolean
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(53, 18).Value = "期末不回收残值" Then
            hscz = False
        Else
            hscz = True
        End If
        'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim xlfl_cg_model As Integer = 1
        'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim xlfl_qt_model As Integer = 1
        'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim sdsl_model As Integer = 1
        'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
        Dim clfl_qtfl_model As Integer = 1
        'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_xlf_model As Integer = 1
        'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
        Dim kcje_clf_qtf_model As Integer = 1
        '———————————————————————————————————————————————————————————————————————————————————————— 
        Dim XZ = MsgBox("是否确定设置的各个参数？", vbOKCancel)
        If XZ = vbOK Then
            '所得税免征年限和减少年限设置为0
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(11, 7).Value = 0
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(12, 7).Value = 0
            '将所得税征收比例设置重置回0
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(14, 7).Value = 0
            '逐年所得税率重置回默认值
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(49, i).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(13, 7).Value
            Next
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算一次工作簿
            ExcelApp.Calculate()
            '计算税收相关计算
            Call 税收相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model)
            '在已经设置好的项目方案中，根据选择的收入内容，记录下不进行任何所得税减免时，实际的逐年所得税，然后计算将需要计算的部分的投资和收入都去掉后，剩下的所得税，差值就是需要计算的部分的所得税，然后将这部分按比例减免           
            '记录此时的逐年所得税值（原始值）
            Dim ZNSDS_YSZ(50) As Double
            '前15年（1-15）
            For i = 1 To 15
                ZNSDS_YSZ(i) = ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(12, i + 4).Value
            Next
            '后16年（16-31）
            For i = 16 To 31
                ZNSDS_YSZ(i) = ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(32, i - 12).Value
            Next
            '————————————————————————————————————————————————————————————————————————————————————————————
            '10次投资金额、可抵扣增值税额
            Dim TZJE(50) As Double
            Dim KDKZZS(50) As Double
            '第一到第十次光伏和风电投资金额，光伏装机规模
            Dim GFTZ(50) As Double '光伏投资
            Dim GFZJ(50) As Double '光伏装机
            Dim FDTZ(50) As Double '风电投资
            Dim FDZJ(50) As Double '风电装机
            '第一到第十次投资专项投资和年量
            Dim LHXHTZ(50) As Double '联合循环投资
            Dim LHXHZL(50) As Double '联合循环总量
            Dim NRJTZ(50) As Double '内燃机投资
            Dim NRJZL(50) As Double '内燃机总量
            Dim LRGWTZ(50) As Double '冷热管网投资
            Dim GLGRZL(50) As Double '供冷供热总量
            Dim RMFDZL(50) As Double '燃煤发电总量
            Dim LJFDZL(50) As Double '垃圾发电总量
            '第一到第十次投资
            Dim JS As Integer = 0 '投资年份计数
            For i = 30 To 73 Step 43 '估算表行号
                For j = 3 To 11 Step 2 '估算表列号
                    JS = JS + 1
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                        TZJE(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value
                        KDKZZS(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 8, j).Value
                        GFTZ(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 20, j).Value
                        GFZJ(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 21, j).Value
                        FDTZ(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 23, j).Value
                        FDZJ(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 24, j).Value
                        LHXHTZ(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 14, j).Value
                        LHXHZL(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 15, j).Value
                        NRJTZ(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 16, j).Value
                        NRJZL(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 17, j).Value
                        LRGWTZ(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 18, j).Value
                        GLGRZL(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 19, j).Value
                        RMFDZL(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 22, j).Value
                        LJFDZL(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 25, j).Value
                    Else
                        TZJE(JS) = 0
                        KDKZZS(JS) = 0
                        GFTZ(JS) = 0
                        GFZJ(JS) = 0
                        FDTZ(JS) = 0
                        FDZJ(JS) = 0
                        LHXHTZ(JS) = 0
                        LHXHZL(JS) = 0
                        NRJTZ(JS) = 0
                        NRJZL(JS) = 0
                        LRGWTZ(JS) = 0
                        GLGRZL(JS) = 0
                        RMFDZL(JS) = 0
                        LJFDZL(JS) = 0
                    End If
                Next
            Next
            '——————————————————————————————————————————————————————————————————————————————————————————
            '——————————————————————————————————————————————————————————————————————————————————————————
            '第一步，仅减去光伏发电的部分
            '减去投资：光伏发电和光伏补贴的投资金额仅减去一次（减去较大的那个）
            Dim GFTZ_max = Math.Max(srtz5, srtz6) '光伏发电和光伏补贴投资中较大的值
            Dim GFTZBL = GFTZ_max / ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value
            '投资和可抵扣增值税中减去这部分（按照逐次投资占静态投资的比例进行减去）
            If GFTZ_max > 0 Then
                Dim JS_gftz As Integer = 0 '投资年份计数
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_gftz = JS_gftz + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value = TZJE(JS_gftz) * (1 - GFTZBL) '静态投资
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 8, j).Value = KDKZZS(JS_gftz) * (1 - GFTZBL) '可抵扣增值税
                        End If
                    Next
                Next
            End If
            '————————————————————————————————————————————————————————————————————————————————————
            '减去光伏发电和光伏补贴的收入
            If Me.收入5.Checked = True Then
                '收入减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 5).Value = srdj5 * (1 - srbl5)
            End If
            If Me.收入6.Checked = True Then
                '收入减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(20, 5).Value = srdj6 * (1 - srbl6)
            End If
            '减去光伏专项人员工资
            If Me.成本8.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 12).Value = cbdj8 - gfrygz
            End If
            '————————————————————————————————————————————————————————————————————————————————————
            '减去光伏专项投资和光伏专项年量
            '记录逐年所得税金额减少值
            Dim ZNSDS_gf_js(50) As Double
            '收入5或者收入6被勾选时
            If Me.收入5.Checked = True Or Me.收入6.Checked = True Then
                '将估算表中输入的光伏投资和装机规模减小
                Dim JS_gftz As Integer = 0 '投资年份计数
                Dim GFTZBL_1 As Double = Math.Max(srtz5, srtz6) / ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(50, 1).Value
                Dim GFSRBL As Double = Math.Max(srbl5, srbl6)
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 20, j).Value = GFTZ(JS_gftz） * (1 - GFTZBL_1)  '光伏投资
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 21, j).Value = GFZJ(JS_gftz) * (1 - GFSRBL) '光伏装机规模
                        End If
                    Next
                Next
                '————————————————————————————————————————————————————————————————————————————————————————
                '计算一次工作簿
                ExcelApp.Calculate()
                '计算税收相关计算
                Call 税收相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model)
                '————————————————————————————————————————————————————————————————————————————————————
                '记录此时的逐年所得税金额
                Dim ZNSDS_gf(50) As Double
                '前15年（1-15）
                For i = 1 To 15
                    ZNSDS_gf(i) = ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(12, i + 4).Value
                Next
                '后16年（16-31）
                For i = 16 To 31
                    ZNSDS_gf(i) = ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(32, i - 12).Value
                Next
                '记录逐年所得税金额减少值
                For i = 1 To 31
                    ZNSDS_gf_js(i) = ZNSDS_YSZ(i) - ZNSDS_gf(i)
                Next
                '————————————————————————————————————————————————————————————————————————————————————
                '返回投资、收入、专项投资和年量默认值
                Dim JS_gf As Integer = 0 '投资年份计数
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_gf = JS_gf + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value = TZJE(JS_gf)  '静态投资
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 8, j).Value = KDKZZS(JS_gf)  '可抵扣增值税
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 20, j).Value = GFTZ(JS_gf) '光伏投资
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 21, j).Value = GFZJ(JS_gf) '光伏装机规模
                        End If
                    Next
                Next
                '收入
                If Me.收入5.Checked = True Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 5).Value = srdj5
                End If
                If Me.收入6.Checked = True Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(20, 5).Value = srdj6
                End If
                '人员工资
                If Me.成本8.Checked = True Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 12).Value = cbdj8
                End If
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算一次工作簿
            ExcelApp.Calculate()
            '计算税收相关计算
            Call 税收相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model)
            '——————————————————————————————————————————————————————————————————————————————————————————
            '——————————————————————————————————————————————————————————————————————————————————————————
            '第二步，仅减去风电的部分        
            '投资和可抵扣增值税中减去这部分（按照逐次投资占静态投资的比例进行减去）
            Dim FDTZBL As Double = srtz7 / ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value '风电投资占总投资比例
            If srtz7 > 0 Then
                Dim JS_fdtz As Integer = 0 '投资年份计数
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_fdtz = JS_fdtz + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value = TZJE(JS_fdtz) * (1 - FDTZBL) '静态投资
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 8, j).Value = KDKZZS(JS_fdtz) * (1 - FDTZBL) '可抵扣增值税
                        End If
                    Next
                Next
            End If
            '————————————————————————————————————————————————————————————————————————————————————
            '减去风力发电的收入
            If Me.收入7.Checked = True Then
                '收入减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 5).Value = srdj7 * (1 - srbl7)
            End If
            '减去风电专项人员工资
            If Me.成本8.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 12).Value = cbdj8 - fdrygz
            End If
            '————————————————————————————————————————————————————————————————————————————————————
            '减去风电专项投资和风电专项年量
            '记录逐年所得税金额减少值
            Dim ZNSDS_fd_js(50) As Double
            '收入7被勾选时
            If Me.收入7.Checked = True Then
                '将估算表中输入的风电投资和装机规模减少
                Dim FDTZBL_1 As Double = srtz7 / ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(53, 1).Value
                Dim JS_fdtz As Integer = 0 '投资年份计数
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 23, j).Value = FDTZ(JS_fdtz) * (1 - FDTZBL_1) '风电投资
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 24, j).Value = FDZJ(JS_fdtz) * (1 - srbl7) '风电装机规模
                        End If
                    Next
                Next
                '————————————————————————————————————————————————————————————————————————————————————————
                '计算一次工作簿
                ExcelApp.Calculate()
                '计算税收相关计算
                Call 税收相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model)
                '————————————————————————————————————————————————————————————————————————————————————
                '记录此时的逐年所得税金额
                Dim ZNSDS_fd(50) As Double
                '前15年（1-15）
                For i = 1 To 15
                    ZNSDS_fd(i) = ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(12, i + 4).Value
                Next
                '后16年（16-31）
                For i = 16 To 31
                    ZNSDS_fd(i) = ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(32, i - 12).Value
                Next
                '记录逐年所得税金额减少值
                For i = 1 To 31
                    ZNSDS_fd_js(i) = ZNSDS_YSZ(i) - ZNSDS_fd(i)
                Next
                '————————————————————————————————————————————————————————————————————————————————————
                '返回投资、收入、专项投资和年量默认值
                Dim JS_fd As Integer = 0 '投资年份计数
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_fd = JS_fd + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value = TZJE(JS_fd)  '静态投资
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 8, j).Value = KDKZZS(JS_fd)  '可抵扣增值税
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 23, j).Value = FDTZ(JS_fd) '风电投资
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 24, j).Value = FDZJ(JS_fd) '风电装机规模
                        End If
                    Next
                Next
                '收入
                If Me.收入7.Checked = True Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 5).Value = srdj7
                End If
                '人员工资
                If Me.成本8.Checked = True Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 12).Value = cbdj8
                End If
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算一次工作簿
            ExcelApp.Calculate()
            '计算税收相关计算
            Call 税收相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model)
            '——————————————————————————————————————————————————————————————————————————————————————————
            '——————————————————————————————————————————————————————————————————————————————————————————
            '第三步，减去除了光伏和风电外的被勾选的部分
            '在静态投资中减去选择的部分的投资
            Dim SRTZ_sum = srtz1 + srtz2 + srtz3 + srtz4 + srtz8 + srtz9 '(不包括光伏和风电投资)
            Dim SRTZBL_sum = SRTZ_sum / ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value '需要减去的投资占总投资比例，可抵扣增值税也按照这个比例减去
            '投资和可抵扣增值税中减去这部分（按照逐次投资占静态投资的比例进行减去）
            If SRTZ_sum > 0 Then
                Dim JS_tz As Integer = 0 '投资年份计数
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_tz = JS_tz + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value = TZJE(JS_tz) * (1 - SRTZBL_sum) '静态投资
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 8, j).Value = KDKZZS(JS_tz) * (1 - SRTZBL_sum) '可抵扣增值税
                        End If
                    Next
                Next
            End If
            '————————————————————————————————————————————————————————————————————————————————————
            '收入1
            If Me.收入1.Checked = True Then
                '收入减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 5).Value = srdj1 * (1 - srbl1)
            End If
            '收入2
            If Me.收入2.Checked = True Then
                '收入减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 5).Value = srdj2 * (1 - srbl2)
            End If
            '收入3
            If Me.收入3.Checked = True Then
                '收入减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 5).Value = srdj3 * (1 - srbl3)
            End If
            '收入4
            If Me.收入4.Checked = True Then
                '收入减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(13, 5).Value = srdj4 * (1 - srbl4)
            End If
            '收入8
            If Me.收入8.Checked = True Then
                '收入减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 5).Value = srdj8 * (1 - srbl8)
            End If
            '收入9
            If Me.收入9.Checked = True Then
                '收入减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 5).Value = srdj9 * (1 - srbl9)
            End If
            '————————————————————————————————————————————————————————————————————————————————————
            '成本1
            If Me.成本1.Checked = True Then
                '成本减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 12).Value = cbdj1 * (1 - cbbl1)
            End If
            '成本2
            If Me.成本2.Checked = True Then
                '成本减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 12).Value = cbdj2 * (1 - cbbl2)
            End If
            '成本3
            If Me.成本3.Checked = True Then
                '成本减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 12).Value = cbdj3 * (1 - cbbl3)
            End If
            '成本4
            If Me.成本4.Checked = True Then
                '成本减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 12).Value = cbdj4 * (1 - cbbl4)
            End If
            '成本5
            If Me.成本5.Checked = True Then
                '成本减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 12).Value = cbdj5 * (1 - cbbl5)
            End If
            '成本6
            If Me.成本6.Checked = True Then
                '成本减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 12).Value = cbdj6 * (1 - cbbl6)
            End If
            '成本7
            If Me.成本7.Checked = True Then
                '成本减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 12).Value = cbdj7 * (1 - cbbl7)
            End If
            '成本8
            If Me.成本8.Checked = True Then
                '成本减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 12).Value = cbdj8 * (1 - cbbl8)
            End If
            '成本9
            If Me.成本9.Checked = True Then
                '成本减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 12).Value = cbdj9 * (1 - cbbl9)
            End If
            '成本10
            If Me.成本10.Checked = True Then
                '成本减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 12).Value = cbdj10 * (1 - cbbl10)
            End If
            '成本11
            If Me.成本11.Checked = True Then
                '成本减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 12).Value = cbdj11 * (1 - cbbl11)
            End If
            '——————————————————————————————————————————————————————————————————————————————————————————————
            '减去专项投资和年总量
            '投资
            Dim JS_zxtz1 As Integer = 0
            If Me.专项投资1.Checked = True Then
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_zxtz1 = JS_zxtz1 + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 16, j).Value = NRJTZ(JS_zxtz1) * (1 - zxtzjqbl1)
                        End If
                    Next
                Next
            End If
            Dim JS_zxtz2 As Integer = 0
            If Me.专项投资2.Checked = True Then
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_zxtz2 = JS_zxtz2 + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 14, j).Value = LHXHTZ(JS_zxtz2) * (1 - zxtzjqbl2)
                        End If
                    Next
                Next
            End If
            Dim JS_zxtz3 As Integer = 0
            If Me.专项投资3.Checked = True Then
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_zxtz3 = JS_zxtz3 + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 18, j).Value = LRGWTZ(JS_zxtz3) * (1 - zxtzjqbl3)
                        End If
                    Next
                Next
            End If
            '年量
            Dim JS_zxzl1 As Integer = 0
            If Me.专项年量1.Checked = True Then
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_zxzl1 = JS_zxzl1 + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 17, j).Value = NRJZL(JS_zxzl1) * (1 - zxzljqbl1)
                        End If
                    Next
                Next
            End If
            Dim JS_zxzl2 As Integer = 0
            If Me.专项年量2.Checked = True Then
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_zxzl2 = JS_zxzl2 + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 15, j).Value = LHXHZL(JS_zxzl2) * (1 - zxzljqbl2)
                        End If
                    Next
                Next
            End If
            Dim JS_zxzl3 As Integer = 0
            If Me.专项年量3.Checked = True Then
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_zxzl3 = JS_zxzl3 + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 19, j).Value = GLGRZL(JS_zxzl3) * (1 - zxzljqbl3)
                        End If
                    Next
                Next
            End If
            Dim JS_zxzl4 As Integer = 0
            If Me.专项年量4.Checked = True Then
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_zxzl4 = JS_zxzl4 + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 22, j).Value = RMFDZL(JS_zxzl4) * (1 - zxzljqbl4)
                        End If
                    Next
                Next
            End If
            Dim JS_zxzl5 As Integer = 0
            If Me.专项年量5.Checked = True Then
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_zxzl5 = JS_zxzl5 + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 25, j).Value = LJFDZL(JS_zxzl5) * (1 - zxzljqbl5)
                        End If
                    Next
                Next
            End If
            '——————————————————————————————————————————————————————————————————————————————————————————————
            '计算一次工作簿
            ExcelApp.Calculate()
            '计算税收相关计算
            Call 税收相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model)
            '记录此时的逐年所得税金额
            Dim ZNSDS_qt(50) As Double
            '前15年（1-15）
            For i = 1 To 15
                ZNSDS_qt(i) = ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(12, i + 4).Value
            Next
            '后16年（16-31）
            For i = 16 To 31
                ZNSDS_qt(i) = ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(32, i - 12).Value
            Next
            '计算逐年所得税差值（需要减免所得税的部分所缴纳的所得税金额）
            Dim ZNSDS_qt_js(50) As Double
            For i = 1 To 31
                ZNSDS_qt_js(i) = ZNSDS_YSZ(i) - ZNSDS_qt(i)
            Next
            '————————————————————————————————————————————————————————————————————————————————————
            '返回所得默认值
            '投资金额，可抵扣增值税金额，专项投资和年量
            Dim JS_a As Integer = 0 '投资年份计数
            For i = 30 To 73 Step 43 '估算表行号
                For j = 3 To 11 Step 2 '估算表列号
                    JS_a = JS_a + 1
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value = TZJE(JS_a)
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 8, j).Value = KDKZZS(JS_a)
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 14, j).Value = LHXHTZ(JS_a)
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 15, j).Value = LHXHZL(JS_a)
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 16, j).Value = NRJTZ(JS_a)
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 17, j).Value = NRJZL(JS_a)
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 18, j).Value = LRGWTZ(JS_a)
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 19, j).Value = GLGRZL(JS_a)
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 22, j).Value = RMFDZL(JS_a)
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 25, j).Value = LJFDZL(JS_a)
                    End If
                Next
            Next
            '各种收入单价
            If Me.收入4.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(13, 5).Value = srdj4
            End If
            If Me.收入1.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 5).Value = srdj1
            End If
            If Me.收入2.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 5).Value = srdj2
            End If
            If Me.收入3.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 5).Value = srdj3
            End If
            If Me.收入8.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 5).Value = srdj8
            End If

            If Me.收入9.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(29, 5).Value = srdj9
            End If
            '各项成本单价
            '成本1
            If Me.成本1.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 12).Value = cbdj1
            End If
            '成本2
            If Me.成本2.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 12).Value = cbdj2
            End If
            '成本3
            If Me.成本3.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 12).Value = cbdj3
            End If
            '成本4
            If Me.成本4.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 12).Value = cbdj4
            End If
            '成本5
            If Me.成本5.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 12).Value = cbdj5
            End If
            '成本6
            If Me.成本6.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 12).Value = cbdj6
            End If
            '成本7
            If Me.成本7.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 12).Value = cbdj7
            End If
            '成本8
            If Me.成本8.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 12).Value = cbdj8
            End If
            '成本9
            If Me.成本9.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 12).Value = cbdj9
            End If
            '成本10
            If Me.成本10.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 12).Value = cbdj10
            End If
            '成本11
            If Me.成本11.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 12).Value = cbdj11
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算一次工作簿
            ExcelApp.Calculate()
            '计算税收相关计算
            Call 税收相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model)
            '——————————————————————————————————————————————————————————————————————————————————————————————
            '——————————————————————————————————————————————————————————————————————————————————————————————
            '计算此时实际的所得税征收金额
            Dim SDS_JS(50) As Double
            For i = 0 To 49
                SDS_JS(i) = 0 '初始化数组，都是0
            Next
            '光伏每年减少的所得税金额
            If Me.收入5.Checked = True Or Me.收入6.Checked = True Then
                '寻找有光伏收入的第一年的年份序号
                Dim GFSRNF_No1 As Integer = 0
                '前15年（1-15）
                For i = 5 To 35
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(15, i).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(15, i - 1).Value = 0 Then
                        GFSRNF_No1 = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(4, i).Value
                        Exit For
                    End If
                Next
                '后16年（16-31）
                For i = 6 To 35
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(38, i).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(38, i - 1).Value = 0 Then
                        GFSRNF_No1 = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(27, i).Value
                        Exit For
                    End If
                Next
                '计算逐年光伏发电所得税减少金额
                For i = 1 To 31
                    If i >= GFSRNF_No1 And i <= GFSRNF_No1 + gfsdsmznx - 1 Then
                        SDS_JS(i) = SDS_JS(i) + ZNSDS_gf_js(i) '光伏所得税全部免除
                    ElseIf i >= GFSRNF_No1 + gfsdsmznx And i <= GFSRNF_No1 + gfsdsmznx + gfsdsjsnx - 1 Then
                        SDS_JS(i) = SDS_JS(i) + ZNSDS_gf_js(i) * gfsdsjsbl
                    End If
                Next
            End If
            '——————————————————————————————————————————————————————————————————————————————————————————————
            '风电每年减少的所得税金额
            If Me.收入7.Checked = True Then
                '寻找有风电收入的第一年的年份序号
                Dim FDSRNF_No1 As Integer = 0
                '前15年（1-15）
                For i = 5 To 35
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(14, i).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(14, i - 1).Value = 0 Then
                        FDSRNF_No1 = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(4, i).Value
                        Exit For
                    End If
                Next
                '后16年（16-31）
                For i = 6 To 35
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(37, i).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(37, i - 1).Value = 0 Then
                        FDSRNF_No1 = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(27, i).Value
                        Exit For
                    End If
                Next
                '计算逐年风力发电发电所得税减少金额
                For i = 1 To 31
                    If i >= FDSRNF_No1 And i <= FDSRNF_No1 + fdsdsmznx - 1 Then
                        SDS_JS(i) = SDS_JS(i) + ZNSDS_fd_js(i) '风电所得税全部免除
                    ElseIf i >= FDSRNF_No1 + fdsdsmznx And i <= FDSRNF_No1 + fdsdsmznx + fdsdsjsnx - 1 Then
                        SDS_JS(i) = SDS_JS(i) + ZNSDS_fd_js(i) * fdsdsjsbl
                    End If
                Next
            End If
            '——————————————————————————————————————————————————————————————————————————————————————————————
            '其它部分每年减少的所得税金额
            If SRTZ_sum > 0 Then
                '将光伏发电收入、光伏补贴收入、风电收入归零，从而找出其它收入第一年的年份序号
                '光伏发电和光伏补贴收入
                If Me.收入5.Checked = True Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 5).Value = 0
                End If
                If Me.收入6.Checked = True Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(20, 5).Value = 0
                End If
                '风力发电收入
                If Me.收入7.Checked = True Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 5).Value = 0
                End If
                '————————————————————————————————————————————————————————————————————————————————————————
                '计算一次工作簿
                ExcelApp.Calculate()
                '计算税收相关计算
                Call 税收相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model)
                '寻找有其它收入的第一年的年份序号
                Dim QTSRNF_No1 As Integer = 0
                '前15年（1-15）
                For i = 5 To 35
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(5, i).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(5, i - 1).Value = 0 Then
                        QTSRNF_No1 = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(4, i).Value
                        Exit For
                    End If
                Next
                '后16年（16-31）
                For i = 6 To 35
                    If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(28, i).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(28, i - 1).Value = 0 Then
                        QTSRNF_No1 = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(27, i).Value
                        Exit For
                    End If
                Next
                '返回光伏收入、光伏补贴、风电收入原始值
                If Me.收入5.Checked = True Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 5).Value = srdj5
                End If
                If Me.收入6.Checked = True Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(20, 5).Value = srdj6
                End If
                '风力发电收入
                If Me.收入7.Checked = True Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 5).Value = srdj7
                End If
                '————————————————————————————————————————————————————————————————————————————————————————
                '计算一次工作簿
                ExcelApp.Calculate()
                '计算税收相关计算
                Call 税收相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model)
                '计算逐年其它所得税减少金额
                For i = 1 To 31
                    If i >= QTSRNF_No1 And i <= QTSRNF_No1 + qtsdsmznx - 1 Then
                        SDS_JS(i) = SDS_JS(i) + ZNSDS_qt_js(i) '其它所得税全部免除
                    ElseIf i >= QTSRNF_No1 + qtsdsmznx And i <= QTSRNF_No1 + qtsdsmznx + qtsdsjsnx - 1 Then
                        SDS_JS(i) = SDS_JS(i) + ZNSDS_qt_js(i) * qtsdsjsbl
                    End If
                Next
            End If
            '——————————————————————————————————————————————————————————————————————————————————————————————
            '——————————————————————————————————————————————————————————————————————————————————————————————
            '计算所得税逐年征收比例
            Dim SDYZN_BL(31) As Double
            Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
            For i = 1 To 31
                If i <= jsnx And ZNSDS_YSZ(i) > 0 Then
                    SDYZN_BL(i) = (ZNSDS_YSZ(i) - SDS_JS(i)) / ZNSDS_YSZ(i)
                Else
                    SDYZN_BL(i) = 0
                End If
            Next
            '将小于0的情况设置为0，大于1的情况下设置为1
            For i = 1 To 31
                If SDYZN_BL(i) < 0 Then
                    SDYZN_BL(i) = 0
                ElseIf SDYZN_BL(i) > 1 Then
                    SDYZN_BL(i) = 1
                End If
            Next
            '写入实际的逐年所得税金额
            '前15年
            For i = 1 To 15
                ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(12, i + 4).Value = ZNSDS_YSZ(i) * SDYZN_BL(i)
            Next
            '16-31年
            For i = 16 To 31
                ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(32, i - 12).Value = ZNSDS_YSZ(i) * SDYZN_BL(i)
            Next
            '写入实际的逐年所得税率
            For i = 3 To 33
                ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(49, i).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(13, 7).Value * SDYZN_BL(i - 2)
            Next
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算一次工作簿
            ExcelApp.Calculate()
            '计算税收相关计算
            Call 税收相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model)
            '写入计算模式
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(44, 18).Value = "特殊设置"
            '——————————————————————————————————————————————————————————————————————————————
            '计算一次Excel
            ExcelApp.Calculate()
            '清空已有内容
            Me.RichTextBox1.Rtf = Nothing
            Me.RichTextBox1.Clear()
            '显示补贴收入计算结果
            Dim sdszsbl As Double '所得税征收比例
            Dim sdsl As Double '实际所得税率
            Dim nf As Integer
            For i = 1 To 31  '根据数组中的元素数量循环
                If i <= jsnx Then
                    nf = i '年份序号
                    sdszsbl = Math.Round(SDYZN_BL(i) * 100, 2) '读取所得税征收比例
                    Me.RichTextBox1.Text = Me.RichTextBox1.Text & sdszsbl & "%(" & nf & ") "  '输出到RichTextBox1
                End If
            Next
            Me.RichTextBox1.Text = "逐年所得税实际征收比例：" & Me.RichTextBox1.Text & vbCrLf & "逐年实际所得税率："
            For i = 1 To 31  '根据数组中的元素数量循环
                If i <= jsnx Then
                    nf = i '年份序号
                    sdsl = Math.Round((ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(49, i).Value) * 100, 2) '读取所得税率
                    Me.RichTextBox1.Text = Me.RichTextBox1.Text & sdsl & "%(" & nf & ") " '输出到RichTextBox1
                End If
            Next
            Me.RichTextBox1.Text = Me.RichTextBox1.Text
        End If
    End Sub

    Private Sub 清空选择_Click(sender As Object, e As EventArgs) Handles 清空选择.Click
        Dim XZ = MsgBox("是否清除所有已勾选内容？", vbOKCancel)
        If XZ = vbOK Then
            '清空全部已选
            '收入
            Me.收入1.Checked = False
            Me.收入2.Checked = False
            Me.收入3.Checked = False
            Me.收入4.Checked = False
            Me.收入5.Checked = False
            Me.收入6.Checked = False
            Me.收入7.Checked = False
            Me.收入8.Checked = False
            Me.收入9.Checked = False
            '成本
            Me.成本1.Checked = False
            Me.成本2.Checked = False
            Me.成本3.Checked = False
            Me.成本4.Checked = False
            Me.成本5.Checked = False
            Me.成本6.Checked = False
            Me.成本7.Checked = False
            Me.成本8.Checked = False
            Me.成本9.Checked = False
            Me.成本10.Checked = False
            Me.成本11.Checked = False
            '专项投资和年量
            Me.专项投资1.Checked = False
            Me.专项投资2.Checked = False
            Me.专项投资3.Checked = False
            Me.专项年量1.Checked = False
            Me.专项年量2.Checked = False
            Me.专项年量3.Checked = False
            Me.专项年量4.Checked = False
            Me.专项年量5.Checked = False
            '专项投资和年量减去比例
            Me.专项投资减去比例1.Text = Nothing
            Me.专项投资减去比例2.Text = Nothing
            Me.专项投资减去比例3.Text = Nothing
            Me.专项年量减去比例1.Text = Nothing
            Me.专项年量减去比例2.Text = Nothing
            Me.专项年量减去比例3.Text = Nothing
            Me.专项年量减去比例4.Text = Nothing
            Me.专项年量减去比例5.Text = Nothing
            '清空已有投资
            Me.投资1.Text = Nothing
            Me.投资2.Text = Nothing
            Me.投资3.Text = Nothing
            Me.投资4.Text = Nothing
            Me.投资5.Text = Nothing
            Me.投资6.Text = Nothing
            Me.投资7.Text = Nothing
            Me.投资8.Text = Nothing
            Me.投资9.Text = Nothing
            '清空已有收入比例
            Me.收入比例1.Text = Nothing
            Me.收入比例2.Text = Nothing
            Me.收入比例3.Text = Nothing
            Me.收入比例4.Text = Nothing
            Me.收入比例5.Text = Nothing
            Me.收入比例6.Text = Nothing
            Me.收入比例7.Text = Nothing
            Me.收入比例8.Text = Nothing
            Me.收入比例9.Text = Nothing
            '清空已有成本比例
            Me.成本比例1.Text = Nothing
            Me.成本比例2.Text = Nothing
            Me.成本比例3.Text = Nothing
            Me.成本比例4.Text = Nothing
            Me.成本比例5.Text = Nothing
            Me.成本比例6.Text = Nothing
            Me.成本比例7.Text = Nothing
            Me.成本比例8.Text = Nothing
            Me.成本比例9.Text = Nothing
            Me.成本比例10.Text = Nothing
            Me.成本比例11.Text = Nothing
            '光伏和风电所得税减免参数
            Me.光伏所得税免征年限.Text = Nothing
            Me.风电所得税免征年限.Text = Nothing
            Me.光伏所得税减少年限.Text = Nothing
            Me.风电所得税减少年限.Text = Nothing
            Me.光伏所得税减少比例.Text = Nothing
            Me.风电所得税减少比例.Text = Nothing
            '光伏和风电增值税退税参数
            Me.光伏增值税退税年限.Text = Nothing
            Me.风电增值税退税年限.Text = Nothing
            Me.光伏增值税退税比例.Text = Nothing
            Me.风电增值税退税比例.Text = Nothing
            '光伏和风电人员工资
            Me.光伏人员工资.Text = Nothing
            Me.风电人员工资.Text = Nothing
            '其它所得税减免和增值税退税参数
            Me.其它所得税免征年限.Text = Nothing
            Me.其它所得税减少年限.Text = Nothing
            Me.其它所得税减少比例.Text = Nothing
            Me.其它增值税退税年限.Text = Nothing
            Me.其它增值税退税比例.Text = Nothing
            '清空计算结果
            Me.RichTextBox1.Rtf = Nothing
            Me.RichTextBox1.Clear()
        End If
    End Sub

    Private Sub 增值税退税_Click(sender As Object, e As EventArgs) Handles 增值税退税.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '读取输入的投资
        Dim srtz1 As Double = CType(Me.投资1.Text, Double)
        Dim srtz2 As Double = CType(Me.投资2.Text, Double)
        Dim srtz3 As Double = CType(Me.投资3.Text, Double)
        Dim srtz4 As Double = CType(Me.投资4.Text, Double)
        Dim srtz5 As Double = CType(Me.投资5.Text, Double)
        Dim srtz6 As Double = CType(Me.投资6.Text, Double)
        Dim srtz7 As Double = CType(Me.投资7.Text, Double)
        Dim srtz8 As Double = CType(Me.投资8.Text, Double)
        Dim srtz9 As Double = CType(Me.投资9.Text, Double)
        '读取输入的专项投资和年量减去比例
        Dim zxtzjqbl1 As Double = CType(Me.专项投资减去比例1.Text, Double) / 100
        Dim zxtzjqbl2 As Double = CType(Me.专项投资减去比例2.Text, Double) / 100
        Dim zxtzjqbl3 As Double = CType(Me.专项投资减去比例3.Text, Double) / 100
        Dim zxzljqbl1 As Double = CType(Me.专项年量减去比例1.Text, Double) / 100
        Dim zxzljqbl2 As Double = CType(Me.专项年量减去比例2.Text, Double) / 100
        Dim zxzljqbl3 As Double = CType(Me.专项年量减去比例3.Text, Double) / 100
        Dim zxzljqbl4 As Double = CType(Me.专项年量减去比例4.Text, Double) / 100
        Dim zxzljqbl5 As Double = CType(Me.专项年量减去比例5.Text, Double) / 100
        '读取输入的各个收入占比比例
        Dim srbl1 As Double = CType(Me.收入比例1.Text, Double) / 100
        Dim srbl2 As Double = CType(Me.收入比例2.Text, Double) / 100
        Dim srbl3 As Double = CType(Me.收入比例3.Text, Double) / 100
        Dim srbl4 As Double = CType(Me.收入比例4.Text, Double) / 100
        Dim srbl5 As Double = CType(Me.收入比例5.Text, Double) / 100
        Dim srbl6 As Double = CType(Me.收入比例6.Text, Double) / 100
        Dim srbl7 As Double = CType(Me.收入比例7.Text, Double) / 100
        Dim srbl8 As Double = CType(Me.收入比例8.Text, Double) / 100
        Dim srbl9 As Double = CType(Me.收入比例9.Text, Double) / 100
        '读取输入的各个成本占比比例
        Dim cbbl1 As Double = CType(Me.成本比例1.Text, Double) / 100
        Dim cbbl2 As Double = CType(Me.成本比例2.Text, Double) / 100
        Dim cbbl3 As Double = CType(Me.成本比例3.Text, Double) / 100
        Dim cbbl4 As Double = CType(Me.成本比例4.Text, Double) / 100
        Dim cbbl5 As Double = CType(Me.成本比例5.Text, Double) / 100
        Dim cbbl6 As Double = CType(Me.成本比例6.Text, Double) / 100
        Dim cbbl7 As Double = CType(Me.成本比例7.Text, Double) / 100
        Dim cbbl8 As Double = CType(Me.成本比例8.Text, Double) / 100
        Dim cbbl9 As Double = CType(Me.成本比例9.Text, Double) / 100
        Dim cbbl10 As Double = CType(Me.成本比例10.Text, Double) / 100
        Dim cbbl11 As Double = CType(Me.成本比例11.Text, Double) / 100
        '读取设置的光伏和风力发电所得税减免参数、增值税退税参数等
        Dim gfsdsmznx As Integer = CType(Me.光伏所得税免征年限.Text, Integer)
        Dim gfsdsjsnx As Integer = CType(Me.光伏所得税减少年限.Text, Integer)
        Dim gfsdsjsbl As Double = CType(Me.光伏所得税减少比例.Text, Double) / 100
        Dim gfzzstsnx As Integer = CType(Me.光伏增值税退税年限.Text, Integer)
        Dim gfzzstsbl As Double = CType(Me.光伏增值税退税比例.Text, Double) / 100
        Dim gfrygz As Double = CType(Me.光伏人员工资.Text, Double)
        Dim fdsdsmznx As Integer = CType(Me.风电所得税免征年限.Text, Integer)
        Dim fdsdsjsnx As Integer = CType(Me.风电所得税减少年限.Text, Integer)
        Dim fdsdsjsbl As Double = CType(Me.风电所得税减少比例.Text, Double) / 100
        Dim fdzzstsnx As Integer = CType(Me.风电增值税退税年限.Text, Integer)
        Dim fdzzstsbl As Double = CType(Me.风电增值税退税比例.Text, Double) / 100
        Dim fdrygz As Double = CType(Me.风电人员工资.Text, Double)
        '读取输入的其它部分所得税减免和增值税退税参数
        Dim qtsdsmznx As Integer = CType(Me.其它所得税免征年限.Text, Integer)
        Dim qtsdsjsnx As Integer = CType(Me.其它所得税减少年限.Text, Integer)
        Dim qtsdsjsbl As Double = CType(Me.其它所得税减少比例.Text, Double) / 100
        Dim qtzzstsnx As Integer = CType(Me.其它增值税退税年限.Text, Integer)
        Dim qtzzstsbl As Double = CType(Me.其它增值税退税比例.Text, Double) / 100
        '记录目前各个收入的单价
        Dim srdj4 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(13, 5).Value, Double)
        Dim srdj1 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 5).Value, Double)
        Dim srdj2 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 5).Value, Double)
        Dim srdj3 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 5).Value, Double)
        Dim srdj5 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 5).Value, Double)
        Dim srdj6 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(20, 5).Value, Double)
        Dim srdj7 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 5).Value, Double)
        Dim srdj8 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 5).Value, Double)
        Dim srdj9 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(29, 5).Value, Double)
        '记录目前各个成本的单价
        Dim cbdj1 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 12).Value, Double)
        Dim cbdj2 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 12).Value, Double)
        Dim cbdj3 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 12).Value, Double)
        Dim cbdj10 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 12).Value, Double)
        Dim cbdj11 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 12).Value, Double)
        Dim cbdj4 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 12).Value, Double)
        Dim cbdj5 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 12).Value, Double)
        Dim cbdj6 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 12).Value, Double)
        Dim cbdj7 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 12).Value, Double)
        Dim cbdj9 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 12).Value, Double)
        Dim cbdj8 As Double = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 12).Value, Double)
        '————————————————————————————————————————————————————————————————————————————————————————————
        '针对输入和选择的内容，添加报错功能
        '光伏
        If Me.收入5.Checked = True Or Me.收入6.Checked = True Then
            If gfzzstsnx = 0 Then
                MsgBox("光伏增值税退税年限不可以是0，请重新输入！")
                Exit Sub
            End If
            If gfzzstsnx > 0 And gfzzstsbl = 0 Then
                MsgBox("光伏增值税退税年限不是0，但光伏增值税退税比例是0，请重新输入！")
                Exit Sub
            End If
        End If
        '风电
        If Me.收入7.Checked = True Then
            If fdzzstsnx = 0 Then
                MsgBox("风电增值税退税年限不可以是0，请重新输入！")
                Exit Sub
            End If
            If fdzzstsnx > 0 And fdzzstsbl = 0 Then
                MsgBox("风电增值税退税年限不是0，但风电增值税退税比例是0，请重新输入！")
                Exit Sub
            End If
        End If
        '其它
        If Me.收入5.Checked = False And Me.收入6.Checked = False And Me.收入7.Checked = False Then
            If qtzzstsnx = 0 Then
                MsgBox("其它增值税退税年限不可以是0，请重新输入！")
                Exit Sub
            End If
            If qtzzstsnx > 0 And qtzzstsbl = 0 Then
                MsgBox("其它增值税退税年限不是0，但其它增值税退税比例是0，请重新输入！")
                Exit Sub
            End If
        End If
        '————————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否确定设置的各个参数？", vbOKCancel)
        If XZ = vbOK Then
            'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
            Dim zbj_model As Integer
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 7).Value = "动态" Then
                zbj_model = 0
            Else
                zbj_model = 1
            End If
            'hscy：计算期末，是否回收资产残值
            Dim hscz As Boolean
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(53, 18).Value = "期末不回收残值" Then
                hscz = False
            Else
                hscz = True
            End If
            'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
            Dim xlfl_cg_model As Integer = 1
            'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
            Dim xlfl_qt_model As Integer = 1
            'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
            Dim sdsl_model As Integer = 1
            'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
            Dim clfl_qtfl_model As Integer = 1
            'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
            Dim kcje_xlf_model As Integer = 1
            'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
            Dim kcje_clf_qtf_model As Integer = 1
            '————————————————————————————————————————————————————————————————————————————————————————————
            '记录此时的逐年增值税税值（原始值）
            Dim ZZS_YSZ(50) As Double
            '前15年（1-15）
            For i = 1 To 15
                ZZS_YSZ(i) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(20, i + 4).Value
            Next
            '后16年（16-31）
            For i = 16 To 31
                ZZS_YSZ(i) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(43, i - 12).Value
            Next
            '————————————————————————————————————————————————————————————————————————————————————————————
            '10次投资金额、可抵扣增值税额
            Dim TZJE(50) As Double
            Dim KDKZZS(50) As Double
            '第一到第十次光伏和风电投资金额，光伏装机规模
            Dim GFTZ(50) As Double '光伏投资
            Dim GFZJ(50) As Double '光伏装机
            Dim FDTZ(50) As Double '风电投资
            Dim FDZJ(50) As Double '风电装机
            '第一到第十次投资专项投资和年量
            Dim LHXHTZ(50) As Double '联合循环投资
            Dim LHXHZL(50) As Double '联合循环总量
            Dim NRJTZ(50) As Double '内燃机投资
            Dim NRJZL(50) As Double '内燃机总量
            Dim LRGWTZ(50) As Double '冷热管网投资
            Dim GLGRZL(50) As Double '供冷供热总量
            Dim RMFDZL(50) As Double '燃煤发电总量
            Dim LJFDZL(50) As Double '垃圾发电总量
            '第一到第十次投资
            Dim JS As Integer = 0 '投资年份计数
            For i = 30 To 73 Step 43 '估算表行号
                For j = 3 To 11 Step 2 '估算表列号
                    JS = JS + 1
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                        TZJE(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value
                        KDKZZS(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 8, j).Value
                        GFTZ(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 20, j).Value
                        GFZJ(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 21, j).Value
                        FDTZ(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 23, j).Value
                        FDZJ(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 24, j).Value
                        LHXHTZ(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 14, j).Value
                        LHXHZL(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 15, j).Value
                        NRJTZ(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 16, j).Value
                        NRJZL(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 17, j).Value
                        LRGWTZ(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 18, j).Value
                        GLGRZL(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 19, j).Value
                        RMFDZL(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 22, j).Value
                        LJFDZL(JS) = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 25, j).Value
                    Else
                        TZJE(JS) = 0
                        KDKZZS(JS) = 0
                        GFTZ(JS) = 0
                        GFZJ(JS) = 0
                        FDTZ(JS) = 0
                        FDZJ(JS) = 0
                        LHXHTZ(JS) = 0
                        LHXHZL(JS) = 0
                        NRJTZ(JS) = 0
                        NRJZL(JS) = 0
                        LRGWTZ(JS) = 0
                        GLGRZL(JS) = 0
                        RMFDZL(JS) = 0
                        LJFDZL(JS) = 0
                    End If
                Next
            Next
            '——————————————————————————————————————————————————————————————————————————————————————————
            '——————————————————————————————————————————————————————————————————————————————————————————
            '第一步，仅减去光伏发电的部分
            '减去投资：光伏发电和光伏补贴的投资金额仅减去一次（减去较大的那个）
            Dim GFTZ_max = Math.Max(srtz5, srtz6) '光伏发电和光伏补贴投资中较大的值
            Dim GFTZBL = GFTZ_max / ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value
            '投资和可抵扣增值税中减去这部分（按照逐次投资占静态投资的比例进行减去）
            If GFTZ_max > 0 Then
                Dim JS_gftz As Integer = 0 '投资年份计数
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_gftz = JS_gftz + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value = TZJE(JS_gftz) * (1 - GFTZBL) '静态投资
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 8, j).Value = KDKZZS(JS_gftz) * (1 - GFTZBL) '可抵扣增值税
                        End If
                    Next
                Next
            End If
            '————————————————————————————————————————————————————————————————————————————————————
            '减去光伏发电和光伏补贴的收入
            If Me.收入5.Checked = True Then
                '收入减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 5).Value = srdj5 * (1 - srbl5)
            End If
            If Me.收入6.Checked = True Then
                '收入减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(20, 5).Value = srdj6 * (1 - srbl6)
            End If
            '减去光伏专项人员工资
            If Me.成本8.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 12).Value = cbdj8 - gfrygz
            End If
            '————————————————————————————————————————————————————————————————————————————————————
            '减去光伏专项投资和光伏专项年量
            '记录逐年增值税金额减少值
            Dim ZZS_gf_js(50) As Double
            '收入5或者收入6被勾选时
            If Me.收入5.Checked = True Or Me.收入6.Checked = True Then
                '将估算表中输入的光伏投资和装机规模减小
                Dim JS_gftz As Integer = 0 '投资年份计数
                Dim GFTZBL_1 As Double = Math.Max(srtz5, srtz6) / ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(50, 1).Value
                Dim GFSRBL As Double = Math.Max(srbl5, srbl6)
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 20, j).Value = GFTZ(JS_gftz） * (1 - GFTZBL_1)  '光伏投资
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 21, j).Value = GFZJ(JS_gftz) * (1 - GFSRBL) '光伏装机规模
                        End If
                    Next
                Next
                '————————————————————————————————————————————————————————————————————————————————————————
                '计算一次工作簿
                ExcelApp.Calculate()
                '计算税收相关计算
                Call 税收相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model)
                '————————————————————————————————————————————————————————————————————————————————————
                '记录此时的逐年增值税金额
                Dim ZZS_gf(50) As Double
                '前15年（1-15）
                For i = 1 To 15
                    ZZS_gf(i) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(20, i + 4).Value
                Next
                '后16年（16-31）
                For i = 16 To 31
                    ZZS_gf(i) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(43, i - 12).Value
                Next
                '记录逐年增值税金额减少值
                For i = 1 To 31
                    ZZS_gf_js(i) = ZZS_YSZ(i) - ZZS_gf(i)
                Next
                '————————————————————————————————————————————————————————————————————————————————————
                '返回投资、收入、专项投资和年量默认值
                Dim JS_gf As Integer = 0 '投资年份计数
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_gf = JS_gf + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value = TZJE(JS_gf)  '静态投资
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 8, j).Value = KDKZZS(JS_gf)  '可抵扣增值税
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 20, j).Value = GFTZ(JS_gf) '光伏投资
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 21, j).Value = GFZJ(JS_gf) '光伏装机规模
                        End If
                    Next
                Next
                '收入
                If Me.收入5.Checked = True Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 5).Value = srdj5
                End If
                If Me.收入6.Checked = True Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(20, 5).Value = srdj6
                End If
                '人员工资
                If Me.成本8.Checked = True Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 12).Value = cbdj8
                End If
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算一次工作簿
            ExcelApp.Calculate()
            '计算税收相关计算
            Call 税收相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model)
            '——————————————————————————————————————————————————————————————————————————————————————————
            '——————————————————————————————————————————————————————————————————————————————————————————
            '第二步，仅减去风电的部分        
            '投资和可抵扣增值税中减去这部分（按照逐次投资占静态投资的比例进行减去）
            Dim FDTZBL As Double = srtz7 / ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value '风电投资占总投资比例
            If srtz7 > 0 Then
                Dim JS_fdtz As Integer = 0 '投资年份计数
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_fdtz = JS_fdtz + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value = TZJE(JS_fdtz) * (1 - FDTZBL) '静态投资
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 8, j).Value = KDKZZS(JS_fdtz) * (1 - FDTZBL) '可抵扣增值税
                        End If
                    Next
                Next
            End If
            '————————————————————————————————————————————————————————————————————————————————————
            '减去风力发电的收入
            If Me.收入7.Checked = True Then
                '收入减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 5).Value = srdj7 * (1 - srbl7)
            End If
            '减去风电专项人员工资
            If Me.成本8.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 12).Value = cbdj8 - fdrygz
            End If
            '————————————————————————————————————————————————————————————————————————————————————
            '减去风电专项投资和风电专项年量
            '记录逐年增值税金额减少值
            Dim ZZS_fd_js(50) As Double
            '收入7被勾选时
            If Me.收入7.Checked = True Then
                '将估算表中输入的风电投资和装机规模减少
                Dim FDTZBL_1 As Double = srtz7 / ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(53, 1).Value
                Dim JS_fdtz As Integer = 0 '投资年份计数
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 23, j).Value = FDTZ(JS_fdtz) * (1 - FDTZBL_1) '风电投资
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 24, j).Value = FDZJ(JS_fdtz) * (1 - srbl7) '风电装机规模
                        End If
                    Next
                Next
                '————————————————————————————————————————————————————————————————————————————————————————
                '计算一次工作簿
                ExcelApp.Calculate()
                '计算税收相关计算
                Call 税收相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model)
                '————————————————————————————————————————————————————————————————————————————————————
                '记录此时的逐年增值税金额
                Dim ZZS_fd(50) As Double
                '前15年（1-15）
                For i = 1 To 15
                    ZZS_fd(i) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(20, i + 4).Value
                Next
                '后16年（16-31）
                For i = 16 To 31
                    ZZS_fd(i) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(43, i - 12).Value
                Next
                '记录逐年增值税金额减少值
                For i = 1 To 31
                    ZZS_fd_js(i) = ZZS_YSZ(i) - ZZS_fd(i)
                Next
                '————————————————————————————————————————————————————————————————————————————————————
                '返回投资、收入、专项投资和年量默认值
                Dim JS_fd As Integer = 0 '投资年份计数
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_fd = JS_fd + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value = TZJE(JS_fd)  '静态投资
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 8, j).Value = KDKZZS(JS_fd)  '可抵扣增值税
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 23, j).Value = FDTZ(JS_fd) '风电投资
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 24, j).Value = FDZJ(JS_fd) '风电装机规模
                        End If
                    Next
                Next
                '收入
                If Me.收入7.Checked = True Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 5).Value = srdj7
                End If
                '人员工资
                If Me.成本8.Checked = True Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 12).Value = cbdj8
                End If
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算一次工作簿
            ExcelApp.Calculate()
            '计算税收相关计算
            Call 税收相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model)
            '——————————————————————————————————————————————————————————————————————————————————————————
            '——————————————————————————————————————————————————————————————————————————————————————————
            '第三步，减去除了光伏和风电外的被勾选的部分
            '在静态投资中减去选择的部分的投资
            Dim SRTZ_sum = srtz1 + srtz2 + srtz3 + srtz4 + srtz8 + srtz9 '(不包括光伏和风电投资)
            Dim SRTZBL_sum = SRTZ_sum / ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value '需要减去的投资占总投资比例，可抵扣增值税也按照这个比例减去
            '投资和可抵扣增值税中减去这部分（按照逐次投资占静态投资的比例进行减去）
            If SRTZ_sum > 0 Then
                Dim JS_tz As Integer = 0 '投资年份计数
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_tz = JS_tz + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value = TZJE(JS_tz) * (1 - SRTZBL_sum) '静态投资
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 8, j).Value = KDKZZS(JS_tz) * (1 - SRTZBL_sum) '可抵扣增值税
                        End If
                    Next
                Next
            End If
            '————————————————————————————————————————————————————————————————————————————————————
            '收入1
            If Me.收入1.Checked = True Then
                '收入减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 5).Value = srdj1 * (1 - srbl1)
            End If
            '收入2
            If Me.收入2.Checked = True Then
                '收入减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 5).Value = srdj2 * (1 - srbl2)
            End If
            '收入3
            If Me.收入3.Checked = True Then
                '收入减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 5).Value = srdj3 * (1 - srbl3)
            End If
            '收入4
            If Me.收入4.Checked = True Then
                '收入减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(13, 5).Value = srdj4 * (1 - srbl4)
            End If
            '收入8
            If Me.收入8.Checked = True Then
                '收入减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 5).Value = srdj8 * (1 - srbl8)
            End If
            '收入9
            If Me.收入9.Checked = True Then
                '收入减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 5).Value = srdj9 * (1 - srbl9)
            End If
            '————————————————————————————————————————————————————————————————————————————————————
            '成本1
            If Me.成本1.Checked = True Then
                '成本减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 12).Value = cbdj1 * (1 - cbbl1)
            End If
            '成本2
            If Me.成本2.Checked = True Then
                '成本减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 12).Value = cbdj2 * (1 - cbbl2)
            End If
            '成本3
            If Me.成本3.Checked = True Then
                '成本减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 12).Value = cbdj3 * (1 - cbbl3)
            End If
            '成本4
            If Me.成本4.Checked = True Then
                '成本减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 12).Value = cbdj4 * (1 - cbbl4)
            End If
            '成本5
            If Me.成本5.Checked = True Then
                '成本减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 12).Value = cbdj5 * (1 - cbbl5)
            End If
            '成本6
            If Me.成本6.Checked = True Then
                '成本减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 12).Value = cbdj6 * (1 - cbbl6)
            End If
            '成本7
            If Me.成本7.Checked = True Then
                '成本减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 12).Value = cbdj7 * (1 - cbbl7)
            End If
            '成本8
            If Me.成本8.Checked = True Then
                '成本减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 12).Value = cbdj8 * (1 - cbbl8)
            End If
            '成本9
            If Me.成本9.Checked = True Then
                '成本减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 12).Value = cbdj9 * (1 - cbbl9)
            End If
            '成本10
            If Me.成本10.Checked = True Then
                '成本减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 12).Value = cbdj10 * (1 - cbbl10)
            End If
            '成本11
            If Me.成本11.Checked = True Then
                '成本减小
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 12).Value = cbdj11 * (1 - cbbl11)
            End If
            '——————————————————————————————————————————————————————————————————————————————————————————————
            '减去专项投资和年总量
            '投资
            Dim JS_zxtz1 As Integer = 0
            If Me.专项投资1.Checked = True Then
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_zxtz1 = JS_zxtz1 + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 16, j).Value = NRJTZ(JS_zxtz1) * (1 - zxtzjqbl1)
                        End If
                    Next
                Next
            End If
            Dim JS_zxtz2 As Integer = 0
            If Me.专项投资2.Checked = True Then
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_zxtz2 = JS_zxtz2 + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 14, j).Value = LHXHTZ(JS_zxtz2) * (1 - zxtzjqbl2)
                        End If
                    Next
                Next
            End If
            Dim JS_zxtz3 As Integer = 0
            If Me.专项投资3.Checked = True Then
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_zxtz3 = JS_zxtz3 + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 18, j).Value = LRGWTZ(JS_zxtz3) * (1 - zxtzjqbl3)
                        End If
                    Next
                Next
            End If
            '年量
            Dim JS_zxzl1 As Integer = 0
            If Me.专项年量1.Checked = True Then
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_zxzl1 = JS_zxzl1 + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 17, j).Value = NRJZL(JS_zxzl1) * (1 - zxzljqbl1)
                        End If
                    Next
                Next
            End If
            Dim JS_zxzl2 As Integer = 0
            If Me.专项年量2.Checked = True Then
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_zxzl2 = JS_zxzl2 + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 15, j).Value = LHXHZL(JS_zxzl2) * (1 - zxzljqbl2)
                        End If
                    Next
                Next
            End If
            Dim JS_zxzl3 As Integer = 0
            If Me.专项年量3.Checked = True Then
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_zxzl3 = JS_zxzl3 + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 19, j).Value = GLGRZL(JS_zxzl3) * (1 - zxzljqbl3)
                        End If
                    Next
                Next
            End If
            Dim JS_zxzl4 As Integer = 0
            If Me.专项年量4.Checked = True Then
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_zxzl4 = JS_zxzl4 + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 22, j).Value = RMFDZL(JS_zxzl4) * (1 - zxzljqbl4)
                        End If
                    Next
                Next
            End If
            Dim JS_zxzl5 As Integer = 0
            If Me.专项年量5.Checked = True Then
                For i = 30 To 73 Step 43 '估算表行号
                    For j = 3 To 11 Step 2 '估算表列号
                        JS_zxzl5 = JS_zxzl5 + 1
                        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 25, j).Value = LJFDZL(JS_zxzl5) * (1 - zxzljqbl5)
                        End If
                    Next
                Next
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算一次工作簿
            ExcelApp.Calculate()
            '计算税收相关计算
            Call 税收相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model)
            '——————————————————————————————————————————————————————————————————————————————————————————————
            '记录此时的逐年增值税金额
            Dim ZZS_qt(50) As Double
            '前15年（1-15）
            For i = 1 To 15
                ZZS_qt(i) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(20, i + 4).Value
            Next
            '后16年（16-31）
            For i = 16 To 31
                ZZS_qt(i) = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(43, i - 12).Value
            Next
            '计算逐年增值税差值（需要减免增值税的部分所缴纳的增值税金额）
            Dim ZZS_qt_js(50) As Double
            For i = 1 To 31
                ZZS_qt_js(i) = ZZS_YSZ(i) - ZZS_qt(i)
            Next
            '————————————————————————————————————————————————————————————————————————————————————
            '返回所得默认值
            '投资金额，可抵扣增值税金额，专项投资和年量
            Dim JS_a As Integer = 0 '投资年份计数
            For i = 30 To 73 Step 43 '估算表行号
                For j = 3 To 11 Step 2 '估算表列号
                    JS_a = JS_a + 1
                    If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i, j).Value = TZJE(JS_a)
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 8, j).Value = KDKZZS(JS_a)
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 14, j).Value = LHXHTZ(JS_a)
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 15, j).Value = LHXHZL(JS_a)
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 16, j).Value = NRJTZ(JS_a)
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 17, j).Value = NRJZL(JS_a)
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 18, j).Value = LRGWTZ(JS_a)
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 19, j).Value = GLGRZL(JS_a)
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 22, j).Value = RMFDZL(JS_a)
                        ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(i + 25, j).Value = LJFDZL(JS_a)
                    End If
                Next
            Next
            '各种收入单价
            If Me.收入4.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(13, 5).Value = srdj4
            End If
            If Me.收入1.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 5).Value = srdj1
            End If
            If Me.收入2.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 5).Value = srdj2
            End If
            If Me.收入3.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 5).Value = srdj3
            End If
            If Me.收入8.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 5).Value = srdj8
            End If

            If Me.收入9.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(29, 5).Value = srdj9
            End If
            '各项成本单价
            '成本1
            If Me.成本1.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 12).Value = cbdj1
            End If
            '成本2
            If Me.成本2.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 12).Value = cbdj2
            End If
            '成本3
            If Me.成本3.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 12).Value = cbdj3
            End If
            '成本4
            If Me.成本4.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 12).Value = cbdj4
            End If
            '成本5
            If Me.成本5.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 12).Value = cbdj5
            End If
            '成本6
            If Me.成本6.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 12).Value = cbdj6
            End If
            '成本7
            If Me.成本7.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 12).Value = cbdj7
            End If
            '成本8
            If Me.成本8.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 12).Value = cbdj8
            End If
            '成本9
            If Me.成本9.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 12).Value = cbdj9
            End If
            '成本10
            If Me.成本10.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 12).Value = cbdj10
            End If
            '成本11
            If Me.成本11.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 12).Value = cbdj11
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算一次工作簿
            ExcelApp.Calculate()
            '计算税收相关计算
            Call 税收相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model)
            '——————————————————————————————————————————————————————————————————————————————————————————————
            '——————————————————————————————————————————————————————————————————————————————————————————————
            '计算此时实际的增值税退税比例
            Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
            Dim ZZS_TS(50) As Double
            For i = 0 To 49
                ZZS_TS(i) = 0 '初始化数组，都是0
            Next
            '光伏每年减少的增值税金额
            If Me.收入5.Checked = True Or Me.收入6.Checked = True Then
                '计算逐年光伏发电增值税退税金额
                For i = 1 To 31
                    If i <= jsnx And i <= gfzzstsnx Then
                        ZZS_TS(i) = ZZS_TS(i) + ZZS_gf_js(i) * gfzzstsbl
                    End If
                Next
            End If
            '——————————————————————————————————————————————————————————————————————————————————————————————
            '风电每年减少的增值税金额
            If Me.收入7.Checked = True Then
                '计算逐年风力发电发电增值税减少金额
                For i = 1 To 31
                    If i <= jsnx And i <= fdzzstsnx Then
                        ZZS_TS(i) = ZZS_TS(i) + ZZS_fd_js(i) * fdzzstsbl
                    End If
                Next
            End If
            '——————————————————————————————————————————————————————————————————————————————————————————————
            '其它部分每年减少的增值税金额
            If SRTZ_sum > 0 Then
                '计算逐年其它增值税减少金额
                For i = 1 To 31
                    If i <= jsnx And i <= qtzzstsnx Then
                        ZZS_TS(i) = ZZS_TS(i) + ZZS_qt_js(i) * qtzzstsbl
                    End If
                Next
            End If
            '——————————————————————————————————————————————————————————————————————————————————————————————
            '——————————————————————————————————————————————————————————————————————————————————————————————
            '计算增值税逐年退税比例
            For i = 1 To 31
                If i <= jsnx And ZZS_YSZ(i) > 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(172, i + 2).Value = ZZS_TS(i) / ZZS_YSZ(i)
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(172, i + 2).Value = 0
                End If
            Next
            '将小于0的情况设置为0
            For i = 1 To 31
                If ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(172, i + 2).Value < 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(172, i + 2).Value = 0
                ElseIf ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(172, i + 2).Value > 1 Then
                    ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(172, i + 2).Value = 1
                End If
            Next
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算一次工作簿
            ExcelApp.Calculate()
            '计算税收相关计算
            Call 税收相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model)
            '——————————————————————————————————————————————————————————————————————————————————————————————
            '设置表格中的收入行隐藏或者显示
            '如果增值税退税比例为0，则隐藏增值税退税收入，前15年表格
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(172, 34).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(16).EntireRow.Hidden = True
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(16).EntireRow.Hidden = False
            End If
            '如果增值税退税比例为0，则隐藏增值税退税收入，后15年表格
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(172, 34).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(39).EntireRow.Hidden = True
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(39).EntireRow.Hidden = False
            End If
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(45, 18).Value = "特殊设置"
            '——————————————————————————————————————————————————————————————————————————————————————————————
            '计算一次Excel
            ExcelApp.Calculate()
            '清空已有内容
            Me.RichTextBox1.Rtf = Nothing
            Me.RichTextBox1.Clear()
            '显示增值税退税计算结果
            Dim zzstsbl As Double '增值税退税比例
            Dim nf As Integer
            For i = 3 To 33  '根据数组中的元素数量循环
                If i - 2 <= jsnx Then
                    nf = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(171, i).Value '年份序号
                    zzstsbl = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(172, i).Value) * 100, 2) '读取增值税退税比例
                    Me.RichTextBox1.Text = Me.RichTextBox1.Text & zzstsbl & "%(" & nf & ") " '输出到RichTextBox1
                End If
            Next
            Me.RichTextBox1.Text = "逐年增值税实际退税比例：" & Me.RichTextBox1.Text
        End If
    End Sub

    Private Sub 重置默认_Click(sender As Object, e As EventArgs) Handles 重置默认.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否将所得税减免和增值税退税计算模式均重置回默认方式？", vbOKCancel)
        If XZ = vbOK Then
            Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
            'zbj_model：资本金计算模式，0：以动态投资为计算基础，1：以静态投资为计算基础，数据来自用户设置
            Dim zbj_model As Integer
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 7).Value = "动态" Then
                zbj_model = 0
            Else
                zbj_model = 1
            End If
            'hscy：计算期末，是否回收资产残值
            Dim hscz As Boolean
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(53, 18).Value = "期末不回收残值" Then
                hscz = False
            Else
                hscz = True
            End If
            'xlfl_cg_model：常规设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
            Dim xlfl_cg_model As Integer = 1
            'xlfl_qt_model：其它设备修理费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
            Dim xlfl_qt_model As Integer = 1
            'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
            Dim sdsl_model As Integer = 0
            'clfl_qtfl_model：材料费率、其它费率的计算方式，0：使用默认值，1：从Excel中读取已有的值
            Dim clfl_qtfl_model As Integer = 1
            'kcje_xlf_model：设备修理费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
            Dim kcje_xlf_model As Integer = 1
            'kcje_clf_qtf_model：材料费、其它费计算基数扣除计算模式，0：使用默认值，1：从Excel中读取已有的值
            Dim kcje_clf_qtf_model As Integer = 1
            '————————————————————————————————————————————————————————————
            '所得税减免
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(44, 18).Value = "常规设置"
            '————————————————————————————————————————————————————————————
            '将增值税退税计算重置回默认
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(171, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(172, i).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(11, 5).Value
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(172, i).Value = 0
                End If
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(45, 18).Value = "常规设置"
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算一次工作簿
            ExcelApp.Calculate()
            '计算税收相关计算
            Call 税收相关计算(ExcelApp, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model, kcje_xlf_model, kcje_clf_qtf_model)
            '—————————————————————————————————————————————————————————————
            '设置表格中的收入行隐藏或者显示
            '如果增值税退税比例为0，则隐藏增值税退税收入，前15年表格
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(172, 34).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(16).EntireRow.Hidden = True
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(16).EntireRow.Hidden = False
            End If
            '如果增值税退税比例为0，则隐藏增值税退税收入，后15年表格
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(172, 34).Value = 0 Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(39).EntireRow.Hidden = True
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(39).EntireRow.Hidden = False
            End If
            '————————————————————————————————————————————————————————————
            '计算一次Excel
            ExcelApp.Calculate()
            '清空已有内容
            Me.RichTextBox1.Rtf = Nothing
            Me.RichTextBox1.Clear()
            '显示所得税减免计算结果
            Dim sdsl As Double '实际所得税率
            Dim nf As Integer
            For i = 3 To 33  '根据数组中的元素数量循环
                If i - 2 <= jsnx Then
                    nf = ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(48, i).Value '年份序号
                    sdsl = Math.Round((ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(49, i).Value) * 100, 2) '读取所得税率
                    Me.RichTextBox1.Text = Me.RichTextBox1.Text & sdsl & "%(" & nf & ") "  '输出到RichTextBox1
                End If
            Next
            Me.RichTextBox1.Text = "逐年实际所得税率：" & Me.RichTextBox1.Text & vbCrLf & "逐年增值税退税比例："
            '显示增值税退税计算结果
            Dim zzstsbl As Double '增值税退税比例
            For i = 3 To 33  '根据数组中的元素数量循环
                If i - 2 <= jsnx Then
                    nf = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(171, i).Value '年份序号
                    zzstsbl = Math.Round((ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(172, i).Value) * 100, 2) '读取增值税退税比例
                    Me.RichTextBox1.Text = Me.RichTextBox1.Text & zzstsbl & "%(" & nf & ") " '输出到RichTextBox1
                End If
            Next
            Me.RichTextBox1.Text = Me.RichTextBox1.Text
        End If
    End Sub

End Class