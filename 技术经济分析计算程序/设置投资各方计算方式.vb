Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Public Class 设置投资各方计算方式
    Private Sub 确定参数_Click(sender As Object, e As EventArgs) Handles 确定参数.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否确定输入的投资各方收益计算参数？", vbOKCancel)
        If XZ = vbOK Then
            '读取输入的参数
            '出资比例
            Dim czbl1 = CType(Me.czbl1.Text, Double)
            Dim czbl2 = CType(Me.czbl2.Text, Double)
            Dim czbl3 = CType(Me.czbl3.Text, Double)
            Dim czbl4 = CType(Me.czbl4.Text, Double)
            Dim czbl5 = CType(Me.czbl5.Text, Double)
            '资产处置比例
            Dim zcczbl1 = CType(Me.zcczbl1.Text, Double)
            Dim zcczbl2 = CType(Me.zcczbl2.Text, Double)
            Dim zcczbl3 = CType(Me.zcczbl3.Text, Double)
            Dim zcczbl4 = CType(Me.zcczbl4.Text, Double)
            Dim zcczbl5 = CType(Me.zcczbl5.Text, Double)
            '利润分配比例
            Dim lrfpbl1 = CType(Me.lrfpbl1.Text, Double)
            Dim lrfpbl2 = CType(Me.lrfpbl2.Text, Double)
            Dim lrfpbl3 = CType(Me.lrfpbl3.Text, Double)
            Dim lrfpbl4 = CType(Me.lrfpbl4.Text, Double)
            Dim lrfpbl5 = CType(Me.lrfpbl5.Text, Double)
            '检查输入的各种参数是否正确
            '输入的参数之和必须等于100
            If czbl1 + czbl2 + czbl3 + czbl4 + czbl5 <> 100 Then
                MsgBox("输入的投资各方出资比例之和必须等于100%，请检查！")
                Exit Sub
            End If
            If zcczbl1 + zcczbl2 + zcczbl3 + zcczbl4 + zcczbl5 <> 100 Then
                MsgBox("输入的投资各方资产处置比例之和必须等于100%，请检查！")
                Exit Sub
            End If
            If lrfpbl1 + lrfpbl2 + lrfpbl3 + lrfpbl4 + lrfpbl5 <> 100 Then
                MsgBox("输入的投资各方利润分配比例之和必须等于100%，请检查！")
                Exit Sub
            End If
            '针对读取的投资各方出资比例添加报错功能
            Dim TZFCS1 = czbl1 + zcczbl1 + lrfpbl1
            Dim TZFCS2 = czbl2 + zcczbl2 + lrfpbl2
            Dim TZFCS3 = czbl3 + zcczbl3 + lrfpbl3
            Dim TZFCS4 = czbl4 + zcczbl4 + lrfpbl4
            Dim TZFCS5 = czbl5 + zcczbl5 + lrfpbl5
            '形成数组，方便直接移植代码
            Dim TZGFCZBL(5) As Double
            '初始化
            TZGFCZBL(0) = 0
            TZGFCZBL(1) = 0
            TZGFCZBL(2) = 0
            TZGFCZBL(3) = 0
            TZGFCZBL(4) = 0
            TZGFCZBL(5) = 0
            '写入数值
            TZGFCZBL(0) = TZFCS1
            TZGFCZBL(1) = TZFCS2
            TZGFCZBL(2) = TZFCS3
            TZGFCZBL(3) = TZFCS4
            TZGFCZBL(4) = TZFCS5
            '不能隔行输入（不能是前一个投资方各项参数都是0，后面一个有参数）
            If TZGFCZBL(0) = 0 And (TZGFCZBL(1) > 0 Or TZGFCZBL(2) > 0 Or TZGFCZBL(3) > 0 Or TZGFCZBL(4) > 0) Then
                MsgBox("不能隔行收入，请检查！")
                Exit Sub
            End If
            If TZGFCZBL(1) = 0 And (TZGFCZBL(2) > 0 Or TZGFCZBL(3) > 0 Or TZGFCZBL(4) > 0) Then
                MsgBox("不能隔行收入，请检查！")
                Exit Sub
            End If
            If TZGFCZBL(2) = 0 And (TZGFCZBL(3) > 0 Or TZGFCZBL(4) > 0) Then
                MsgBox("不能隔行收入，请检查！")
                Exit Sub
            End If
            If TZGFCZBL(3) = 0 And TZGFCZBL(4) > 0 Then
                MsgBox("不能隔行收入，请检查！")
                Exit Sub
            End If
            '将输入写入Excel
            '投资方1
            ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(4, 38).Value = czbl1 / 100
            ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(5, 38).Value = zcczbl1 / 100
            ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(6, 38).Value = lrfpbl1 / 100
            '投资方2
            ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(4, 38).Value = czbl2 / 100
            ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(5, 38).Value = zcczbl2 / 100
            ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(6, 38).Value = lrfpbl2 / 100
            '投资方3
            ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(4, 38).Value = czbl3 / 100
            ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(5, 38).Value = zcczbl3 / 100
            ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(6, 38).Value = lrfpbl3 / 100
            '投资方4
            ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(4, 38).Value = czbl4 / 100
            ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(5, 38).Value = zcczbl4 / 100
            ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(6, 38).Value = lrfpbl4 / 100
            '投资方5
            ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(4, 38).Value = czbl5 / 100
            ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(5, 38).Value = zcczbl5 / 100
            ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(6, 38).Value = lrfpbl5 / 100
            '设置投资方1~5现金流量表的隐藏和显示
            '出资比例、资产处置比例、利润分配比例只要有一个大于0，就显示；只有都为0，才隐藏
            '先将5个表格都彻底隐藏
            ExcelApp.Worksheets("投资方1现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
            ExcelApp.Worksheets("投资方2现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
            ExcelApp.Worksheets("投资方3现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
            ExcelApp.Worksheets("投资方4现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
            ExcelApp.Worksheets("投资方5现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
            '设置需要显示的表格
            If czbl1 > 0 Or zcczbl1 > 0 Or lrfpbl1 > 0 Then
                ExcelApp.Worksheets("投资方1现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
            End If
            If czbl2 > 0 Or zcczbl2 > 0 Or lrfpbl2 > 0 Then
                ExcelApp.Worksheets("投资方2现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
            End If
            If czbl3 > 0 Or zcczbl3 > 0 Or lrfpbl3 > 0 Then
                ExcelApp.Worksheets("投资方3现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
            End If
            If czbl4 > 0 Or zcczbl4 > 0 Or lrfpbl4 > 0 Then
                ExcelApp.Worksheets("投资方4现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
            End If
            If czbl5 > 0 Or zcczbl5 > 0 Or lrfpbl5 > 0 Then
                ExcelApp.Worksheets("投资方5现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
            End If
            '写入计算模式
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(168, 7).Value = "不相同"
            MsgBox("设置完成！")
        End If
        '计算一次Excel
        ExcelApp.Calculate()
        Me.Close()
    End Sub

    Private Sub 清空窗体_Click(sender As Object, e As EventArgs) Handles 清空窗体.Click
        Dim XZ = MsgBox("确定要清空本窗体输入的的全部内容？", vbOKCancel)
        If XZ = vbOK Then
            '清空已有的全部数据
            '清空已有的全部数据
            '出资比例
            Me.czbl1.Clear()
            Me.czbl2.Clear()
            Me.czbl3.Clear()
            Me.czbl4.Clear()
            Me.czbl5.Clear()
            '资产处置比例
            Me.zcczbl1.Clear()
            Me.zcczbl2.Clear()
            Me.zcczbl3.Clear()
            Me.zcczbl4.Clear()
            Me.zcczbl5.Clear()
            '利润分配比例
            Me.lrfpbl1.Clear()
            Me.lrfpbl2.Clear()
            Me.lrfpbl3.Clear()
            Me.lrfpbl4.Clear()
            Me.lrfpbl5.Clear()
        End If
    End Sub

    Private Sub 重置默认_Click(sender As Object, e As EventArgs) Handles 重置默认.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否将投资各方收益计算参数重置回默认状态？", vbOKCancel)
        If XZ = vbOK Then
            '根据在<建设期时间计划表>中输入的投资各方出资比例，将投资各方的出资比例、利润分配比例、资产处置比例设置为一样的值           
            '将输入写入Excel
            '投资方1
            ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(4, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(163, 2).Value / 100
            ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(5, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(163, 2).Value / 100
            ExcelApp.ThisWorkbook.Worksheets("投资方1现金流量表").Cells(6, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(163, 2).Value / 100
            '投资方2
            ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(4, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 2).Value / 100
            ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(5, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 2).Value / 100
            ExcelApp.ThisWorkbook.Worksheets("投资方2现金流量表").Cells(6, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 2).Value / 100
            '投资方3
            ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(4, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 2).Value / 100
            ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(5, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 2).Value / 100
            ExcelApp.ThisWorkbook.Worksheets("投资方3现金流量表").Cells(6, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 2).Value / 100
            '投资方4
            ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(4, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(166, 2).Value / 100
            ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(5, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(166, 2).Value / 100
            ExcelApp.ThisWorkbook.Worksheets("投资方4现金流量表").Cells(6, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(166, 2).Value / 100
            '投资方5
            ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(4, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(167, 2).Value / 100
            ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(5, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(167, 2).Value / 100
            ExcelApp.ThisWorkbook.Worksheets("投资方5现金流量表").Cells(6, 38).Value = ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(167, 2).Value / 100
            '设置投资方1~5现金流量表的隐藏和显示         
            '先将5个表格都彻底隐藏
            ExcelApp.Worksheets("投资方1现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
            ExcelApp.Worksheets("投资方2现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
            ExcelApp.Worksheets("投资方3现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
            ExcelApp.Worksheets("投资方4现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
            ExcelApp.Worksheets("投资方5现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVeryHidden
            '设置需要显示的表格
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(163, 2).Value > 0 Then
                ExcelApp.Worksheets("投资方1现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
            End If
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 2).Value > 0 Then
                ExcelApp.Worksheets("投资方2现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
            End If
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 2).Value > 0 Then
                ExcelApp.Worksheets("投资方3现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
            End If
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(166, 2).Value > 0 Then
                ExcelApp.Worksheets("投资方4现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
            End If
            If ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(167, 2).Value > 0 Then
                ExcelApp.Worksheets("投资方5现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
            End If
            '写入计算模式
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(168, 7).Value = "相同"
            MsgBox("计算模式重置回默认状态设置完成！")
        End If
        '计算一次Excel
        ExcelApp.Calculate()
        Me.Close()
    End Sub

    Private Sub 设置投资各方计算方式_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '清空已有的全部数据
        '出资比例
        Me.czbl1.Clear()
        Me.czbl2.Clear()
        Me.czbl3.Clear()
        Me.czbl4.Clear()
        Me.czbl5.Clear()
        '资产处置比例
        Me.zcczbl1.Clear()
        Me.zcczbl2.Clear()
        Me.zcczbl3.Clear()
        Me.zcczbl4.Clear()
        Me.zcczbl5.Clear()
        '利润分配比例
        Me.lrfpbl1.Clear()
        Me.lrfpbl2.Clear()
        Me.lrfpbl3.Clear()
        Me.lrfpbl4.Clear()
        Me.lrfpbl5.Clear()
        '载入默认值
        '投资方1
        Me.czbl1.Text = CType(ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(163, 2).Value, String)
        Me.zcczbl1.Text = CType(ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(163, 2).Value, String)
        Me.lrfpbl1.Text = CType(ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(163, 2).Value, String)
        '投资方2
        Me.czbl2.Text = CType(ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 2).Value, String)
        Me.zcczbl2.Text = CType(ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 2).Value, String)
        Me.lrfpbl2.Text = CType(ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(164, 2).Value, String)
        '投资方3
        Me.czbl3.Text = CType(ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 2).Value, String)
        Me.zcczbl3.Text = CType(ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 2).Value, String)
        Me.lrfpbl3.Text = CType(ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(165, 2).Value, String)
        '投资方4
        Me.czbl4.Text = CType(ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(166, 2).Value, String)
        Me.zcczbl4.Text = CType(ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(166, 2).Value, String)
        Me.lrfpbl4.Text = CType(ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(166, 2).Value, String)
        '投资方5
        Me.czbl5.Text = CType(ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(167, 2).Value, String)
        Me.zcczbl5.Text = CType(ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(167, 2).Value, String)
        Me.lrfpbl5.Text = CType(ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(167, 2).Value, String)
    End Sub

End Class