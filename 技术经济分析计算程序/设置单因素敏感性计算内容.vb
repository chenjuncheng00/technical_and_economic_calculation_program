Imports Microsoft.Office.Interop.Excel

Public Class 设置单因素敏感性计算内容
    Private Sub 设置单因素敏感性计算内容_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————
        'Checkbox载入收入成本文字
        '收入
        Me.sr1.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(12, 18).Value, String) & "敏感性分析计算"
        Me.sr2.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(17, 18).Value, String) & "敏感性分析计算"
        Me.sr3.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(22, 18).Value, String) & "敏感性分析计算"
        Me.sr4.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(27, 18).Value, String) & "敏感性分析计算"
        Me.sr5.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(32, 18).Value, String) & "敏感性分析计算"
        Me.sr6.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(37, 18).Value, String) & "敏感性分析计算"
        Me.sr7.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(42, 18).Value, String) & "敏感性分析计算"
        Me.sr8.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(47, 18).Value, String) & "敏感性分析计算"
        Me.sr9.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(52, 18).Value, String) & "敏感性分析计算"
        Me.sr10.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(57, 18).Value, String) & "敏感性分析计算"
        Me.sr11.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(62, 18).Value, String) & "敏感性分析计算"
        Me.sr12.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(67, 18).Value, String) & "敏感性分析计算"
        Me.sr13.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(72, 18).Value, String) & "敏感性分析计算"
        '成本
        Me.cb1.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(77, 18).Value, String) & "敏感性分析计算"
        Me.cb2.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(82, 18).Value, String) & "敏感性分析计算"
        Me.cb3.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(87, 18).Value, String) & "敏感性分析计算"
        Me.cb4.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(92, 18).Value, String) & "敏感性分析计算"
        Me.cb5.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(97, 18).Value, String) & "敏感性分析计算"
        Me.cb6.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(102, 18).Value, String) & "敏感性分析计算"
        Me.cb7.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(107, 18).Value, String) & "敏感性分析计算"
        Me.cb8.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(112, 18).Value, String) & "敏感性分析计算"
        Me.cb9.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(117, 18).Value, String) & "敏感性分析计算"
        Me.cb10.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(122, 18).Value, String) & "敏感性分析计算"
        Me.cb11.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(127, 18).Value, String) & "敏感性分析计算"
        Me.cb12.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(132, 18).Value, String) & "敏感性分析计算"
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '根据输入的收入成本等内容，判断checkbox是否可以被勾选
        '静态投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
            Me.jttz.Enabled = True
        Else
            Me.jttz.Enabled = False
        End If
        '收入
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(13, 7).Value > 0 Then
            Me.sr1.Enabled = True
        Else
            Me.sr1.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 7).Value > 0 Then
            Me.sr2.Enabled = True
        Else
            Me.sr2.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 7).Value > 0 Then
            Me.sr3.Enabled = True
        Else
            Me.sr3.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 7).Value > 0 Then
            Me.sr4.Enabled = True
        Else
            Me.sr4.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 7).Value > 0 Then
            Me.sr5.Enabled = True
        Else
            Me.sr5.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(18, 7).Value > 0 Then
            Me.sr6.Enabled = True
        Else
            Me.sr6.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 7).Value > 0 Then
            Me.sr7.Enabled = True
        Else
            Me.sr7.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(20, 7).Value > 0 Then
            Me.sr8.Enabled = True
        Else
            Me.sr8.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 7).Value > 0 Then
            Me.sr9.Enabled = True
        Else
            Me.sr9.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 7).Value > 0 Then
            Me.sr10.Enabled = True
        Else
            Me.sr10.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 7).Value > 0 Then
            Me.sr11.Enabled = True
        Else
            Me.sr11.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 7).Value > 0 Then
            Me.sr12.Enabled = True
        Else
            Me.sr12.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 7).Value > 0 Then
            Me.sr13.Enabled = True
        Else
            Me.sr13.Enabled = False
        End If
        '成本
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 14).Value > 0 Then
            Me.cb1.Enabled = True
        Else
            Me.cb1.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 14).Value > 0 Then
            Me.cb2.Enabled = True
        Else
            Me.cb2.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 14).Value > 0 Then
            Me.cb3.Enabled = True
        Else
            Me.cb3.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 14).Value > 0 Then
            Me.cb4.Enabled = True
        Else
            Me.cb4.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(18, 14).Value > 0 Then
            Me.cb5.Enabled = True
        Else
            Me.cb5.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 14).Value > 0 Then
            Me.cb6.Enabled = True
        Else
            Me.cb6.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 14).Value > 0 Then
            Me.cb7.Enabled = True
        Else
            Me.cb7.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 14).Value > 0 Then
            Me.cb8.Enabled = True
        Else
            Me.cb8.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 14).Value > 0 Then
            Me.cb9.Enabled = True
        Else
            Me.cb9.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 14).Value > 0 Then
            Me.cb10.Enabled = True
        Else
            Me.cb10.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 14).Value > 0 Then
            Me.cb11.Enabled = True
        Else
            Me.cb11.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 14).Value > 0 Then
            Me.cb12.Enabled = True
        Else
            Me.cb12.Enabled = False
        End If
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '根据之前的结果勾选
        '静态投资
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(7, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(30, 1).Value > 0 Then
            Me.jttz.Checked = True
        Else
            Me.jttz.Checked = False
        End If
        '年运行小时数
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(137, 29).Value = 1 Then
            Me.nyxxss_check.Checked = True
            '载入之前输入的年利用小时数
            Me.nyxxss_text.Text = CType(ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(137, 8).Value, String)
        Else
            Me.nyxxss_check.Checked = False
            '年利用小时数为空
            Me.nyxxss_text.Text = Nothing
        End If
        '收入
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(12, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(13, 7).Value > 0 Then
            Me.sr1.Checked = True
        Else
            Me.sr1.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(17, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 7).Value > 0 Then
            Me.sr2.Checked = True
        Else
            Me.sr2.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(22, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 7).Value > 0 Then
            Me.sr3.Checked = True
        Else
            Me.sr3.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(27, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 7).Value > 0 Then
            Me.sr4.Checked = True
        Else
            Me.sr4.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(32, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 7).Value > 0 Then
            Me.sr5.Checked = True
        Else
            Me.sr5.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(37, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(18, 7).Value > 0 Then
            Me.sr6.Checked = True
        Else
            Me.sr6.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(42, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 7).Value > 0 Then
            Me.sr7.Checked = True
        Else
            Me.sr7.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(47, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(20, 7).Value > 0 Then
            Me.sr8.Checked = True
        Else
            Me.sr8.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(52, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 7).Value > 0 Then
            Me.sr9.Checked = True
        Else
            Me.sr9.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(57, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 7).Value > 0 Then
            Me.sr10.Checked = True
        Else
            Me.sr10.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(62, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 7).Value > 0 Then
            Me.sr11.Checked = True
        Else
            Me.sr11.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(67, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 7).Value > 0 Then
            Me.sr12.Checked = True
        Else
            Me.sr12.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(72, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 7).Value > 0 Then
            Me.sr13.Checked = True
        Else
            Me.sr13.Checked = False
        End If
        '成本
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(77, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 14).Value > 0 Then
            Me.cb1.Checked = True
        Else
            Me.cb1.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(82, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 14).Value > 0 Then
            Me.cb2.Checked = True
        Else
            Me.cb2.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(87, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 14).Value > 0 Then
            Me.cb3.Checked = True
        Else
            Me.cb3.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(92, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 14).Value > 0 Then
            Me.cb4.Checked = True
        Else
            Me.cb4.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(97, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(18, 14).Value > 0 Then
            Me.cb5.Checked = True
        Else
            Me.cb5.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(102, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 14).Value > 0 Then
            Me.cb6.Checked = True
        Else
            Me.cb6.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(107, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 14).Value > 0 Then
            Me.cb7.Checked = True
        Else
            Me.cb7.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(112, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 14).Value > 0 Then
            Me.cb8.Checked = True
        Else
            Me.cb8.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(117, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 14).Value > 0 Then
            Me.cb9.Checked = True
        Else
            Me.cb9.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(122, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 14).Value > 0 Then
            Me.cb10.Checked = True
        Else
            Me.cb10.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(127, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 14).Value > 0 Then
            Me.cb11.Checked = True
        Else
            Me.cb11.Checked = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(132, 29).Value = 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 14).Value > 0 Then
            Me.cb12.Checked = True
        Else
            Me.cb12.Checked = False
        End If
        '绘制敏感性分析图
        Me.hzzxt.Checked = True
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '敏感性分析变化率
        Dim bhl = ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(10, 9).Value * 100
        Me.mgxfxbhl.Text = CType(bhl, String)
    End Sub

    Private Sub 开始计算_Click(sender As Object, e As EventArgs) Handles 开始计算.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否确认选择的各项内容？", vbOKCancel)
        If XZ = vbOK Then
            '隐藏窗体
            Me.Hide()
            '清空Excel内已有的输入
            For i = 7 To 137 Step 5
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i, 29).Value = Nothing
            Next
            '————————————————————————————————————————————————————————————————————           
            '根据复选框的内容，将系数写入Excel，1代表计算，0代表不计算
            '静态投资
            If Me.jttz.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(7, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(7, 29).Value = 0
            End If
            '年利用小时数
            If Me.nyxxss_text.Text <> Nothing And Me.nyxxss_check.Checked = True Then '写入年利用小时数
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(137, 8).Value = CType(Me.nyxxss_text.Text, Double)
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(137, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(137, 8).Value = Nothing
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(137, 29).Value = 0
            End If
            '收入
            If Me.sr1.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(12, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(12, 29).Value = 0
            End If
            If Me.sr2.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(17, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(17, 29).Value = 0
            End If
            If Me.sr3.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(22, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(22, 29).Value = 0
            End If
            If Me.sr4.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(27, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(27, 29).Value = 0
            End If
            If Me.sr5.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(32, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(32, 29).Value = 0
            End If
            If Me.sr6.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(37, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(37, 29).Value = 0
            End If
            If Me.sr7.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(42, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(42, 29).Value = 0
            End If
            If Me.sr8.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(47, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(47, 29).Value = 0
            End If
            If Me.sr9.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(52, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(52, 29).Value = 0
            End If
            If Me.sr10.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(57, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(57, 29).Value = 0
            End If
            If Me.sr11.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(62, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(62, 29).Value = 0
            End If
            If Me.sr12.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(67, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(67, 29).Value = 0
            End If
            If Me.sr13.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(72, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(72, 29).Value = 0
            End If
            '成本
            If Me.cb1.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(77, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(77, 29).Value = 0
            End If
            If Me.cb2.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(82, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(82, 29).Value = 0
            End If
            If Me.cb3.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(87, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(87, 29).Value = 0
            End If
            If Me.cb4.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(92, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(92, 29).Value = 0
            End If
            If Me.cb5.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(97, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(97, 29).Value = 0
            End If
            If Me.cb6.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(102, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(102, 29).Value = 0
            End If
            If Me.cb7.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(107, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(107, 29).Value = 0
            End If
            If Me.cb8.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(112, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(112, 29).Value = 0
            End If
            If Me.cb9.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(117, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(117, 29).Value = 0
            End If
            If Me.cb10.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(122, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(122, 29).Value = 0
            End If
            If Me.cb11.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(127, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(127, 29).Value = 0
            End If
            If Me.cb12.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(132, 29).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(132, 29).Value = 0
            End If
            '关闭窗体
            Me.Close()
            '————————————————————————————————————————————————————————————————————————————————————————————————
            '实例化计算进度显示窗体
            Dim Form1 As New 计算进度显示
            '清空已有的敏感性分析计算数据
            '敏感性分析计算结果
            ExcelApp.ThisWorkbook.Worksheets("指标数据").Range("K7:L141").ClearContents
            ExcelApp.ThisWorkbook.Worksheets("指标数据").Range("H7:H136").ClearContents
            '删除敏感性分析图表
            ExcelApp.ThisWorkbook.Worksheets("指标数据").ChartObjects("单因素敏感性分析图").Activate
            ExcelApp.ActiveChart.Parent.Delete
            '输入敏感性分析每次的计算步长
            Dim BHL = CType(Me.mgxfxbhl.Text, Double)
            Dim MGXFXBHL As Double = BHL / 100 '敏感性分析变化率
            '判断输入的变化率是否为整数
            Dim ZSJC1 As Integer = 0 '整数检测1
            If BHL = Int(BHL) Then
                ZSJC1 = 1
            End If
            Dim ZSJC2 As Integer = 0 '整数检测2
            If BHL * 2 = Int(BHL * 2) Then
                ZSJC2 = 1
            End If
            '将敏感性分析变化率写入表格中
            For i = 7 To 137 Step 5
                For j = 1 To 5
                    ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i + j - 1, 9).Value = MGXFXBHL * j - 3 * MGXFXBHL '从小往大排列
                    If j = 2 Or j = 4 Then '变化率的1倍
                        If ZSJC1 = 1 Then '1倍是整数
                            ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i + j - 1, 9).NumberFormatLocal = "0%"
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i + j - 1, 9).NumberFormatLocal = "0.0%"
                        End If
                    End If
                    If j = 1 Or j = 5 Then '变化率的2倍
                        If ZSJC2 = 1 Then '2倍是整数
                            ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i + j - 1, 9).NumberFormatLocal = "0%"
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i + j - 1, 9).NumberFormatLocal = "0.0%"
                        End If
                    End If
                Next
            Next
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
            '敏感性计算
            Call 静态投资敏感性分析计算(ExcelApp, MGXFXBHL, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model,
                                        kcje_xlf_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            Call 收入敏感性分析计算(ExcelApp, MGXFXBHL, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model,
                                    kcje_xlf_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            Call 成本敏感性分析计算(ExcelApp, MGXFXBHL, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model,
                                    kcje_xlf_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            Call 年运行小时数敏感性分析(ExcelApp, MGXFXBHL, zbj_model, hscz, xlfl_cg_model, xlfl_qt_model, clfl_qtfl_model, sdsl_model,
                                        kcje_xlf_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            '绘制敏感性分析图并设置格式
            If Me.hzzxt.Checked = True Then
                Call 绘制单因素敏感性分析图(ExcelApp)
                Call 设置敏感性分析图格式(ExcelApp)
            End If
            '选中工作表
            ExcelApp.ThisWorkbook.Worksheets("指标数据").Activate()
            '提醒计算完成
            Form1.Show()
            Form1.Label1.Text = "敏感性分析已计算完成！"
            Form1.TopMost = True
            System.Windows.Forms.Application.DoEvents()
        End If
    End Sub

    Private Sub 清空输入_Click(sender As Object, e As EventArgs) Handles 清空输入.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否清空已选择的各项内容？", vbOKCancel)
        If XZ = vbOK Then
            '清空Excel内已有的输入
            For i = 7 To 137 Step 5
                ExcelApp.ThisWorkbook.Worksheets("指标数据").Cells(i, 29).Value = Nothing
            Next
            '清空Excel内已有的输入(年运行小时数敏感性分析内容)
            '收入
            For i = 13 To 25
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 30).Value = Nothing
            Next
            '成本
            For i = 14 To 19
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 31).Value = Nothing
            Next
            For i = 21 To 25
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 31).Value = Nothing
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 31).Value = Nothing
            '静态投资
            Me.jttz.Checked = False
            '收入
            Me.sr1.Checked = False
            Me.sr2.Checked = False
            Me.sr3.Checked = False
            Me.sr4.Checked = False
            Me.sr5.Checked = False
            Me.sr6.Checked = False
            Me.sr7.Checked = False
            Me.sr8.Checked = False
            Me.sr9.Checked = False
            Me.sr10.Checked = False
            Me.sr11.Checked = False
            Me.sr12.Checked = False
            Me.sr13.Checked = False
            '成本
            Me.cb1.Checked = False
            Me.cb2.Checked = False
            Me.cb3.Checked = False
            Me.cb4.Checked = False
            Me.cb5.Checked = False
            Me.cb6.Checked = False
            Me.cb7.Checked = False
            Me.cb8.Checked = False
            Me.cb9.Checked = False
            Me.cb10.Checked = False
            Me.cb11.Checked = False
            Me.cb12.Checked = False
            '绘制敏感性分析图
            Me.hzzxt.Checked = False
            '敏感性分析变化率
            Me.mgxfxbhl.Text = Nothing
            '年利用小时数
            Me.nyxxss_text.Text = Nothing
        End If
    End Sub

    Private Sub 运行小时数敏感性分析设置_Click(sender As Object, e As EventArgs) Handles 运行小时数敏感性分析设置.Click
        On Error Resume Next
        If Me.nyxxss_text.Text <> Nothing And Me.nyxxss_check.Checked = True Then
            Dim nyxxss_set As New 设置年运行小时数敏感性分析内容
            nyxxss_set.Show()
            nyxxss_set.TopMost = True
        Else
            MsgBox("没有勾选<项目年运行小时数敏感性分析计算>同时输入<项目年运行小时数(h/a)>，请重新设置！")
        End If
    End Sub
End Class