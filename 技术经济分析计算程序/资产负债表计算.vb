Module 资产负债表计算
    Sub 资产负债表计算(ExcelApp As Object)
        '如果是建设期的年份，则需要一些特殊计算
        '1.2：处于建设期，在建工程 = 静态投资 + 建设期贷款利息 - 可抵扣增值税 - 固定资产折旧费 - 无形资产摊销费
        '1.4：处于建设期，固定资产净值=0
        '1.5：处于建设期，无形资产净值=0
        '计算期最后一年，需要一些特殊计算
        '1.1.4累计盈余资金

        '项目计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '计算一次Excel
        ExcelApp.Calculate()
        Dim GSBSJ = 读取估算表数据(ExcelApp)
        '投资年份，10次投资的情况
        Dim tznf_list = GSBSJ(0)
        '读取建设期标记，1：处于建设期，0：不处于建设期
        Dim jsq_index(31) As Integer
        For i = 1 To 31
            jsq_index(i) = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(162, i + 2).Value
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '读取计算用到的数据
        Dim zjf_list(31) As Double '固定资产折旧费
        Dim txf_list(31) As Double '无形资产摊销费
        Dim zjgc_raw_list(31) As Double '在建工程，原始数据
        Dim gdzcjz_raw_list(31) As Double '固定资产净值，原始数据
        Dim wxzcjz_raw_list(31) As Double '无形资产净值，原始数据
        '前15年
        For i = 1 To 15
            zjf_list(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(8, i + 4).Value
            txf_list(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(14, i + 4).Value
            zjgc_raw_list(i) = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(6, i + 4).Value +
                               ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(7, i + 4).Value -
                               ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(8, i + 4).Value
            gdzcjz_raw_list(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(10, i + 4).Value -
                                 ExcelApp.ThisWorkbook.Worksheets("财务计划现金流量表").Cells(23, i + 4).Value
            wxzcjz_raw_list(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(16, i + 4).Value
        Next
        '16~31年
        For i = 16 To 31
            zjf_list(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(25, i - 12).Value
            txf_list(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(31, i - 12).Value
            zjgc_raw_list(i) = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(27, i - 12).Value +
                               ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(28, i - 12).Value -
                               ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(29, i - 12).Value
            gdzcjz_raw_list(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(27, i - 12).Value -
                                 ExcelApp.ThisWorkbook.Worksheets("财务计划现金流量表").Cells(56, i - 12).Value
            wxzcjz_raw_list(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(33, i - 12).Value
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算逐年在建工程、固定资产净值、无形资产净值
        Dim zjgc_list(31) As Double
        Dim gdzcjz_list(31) As Double
        Dim wxzcjz_list(31) As Double
        For i = 1 To 31
            '在建工程中需要减去的折旧摊销费（第1个建设年份：减去全部；之后的建设年份：减去增量）
            Dim zjtxf_jq As Double = 0
            If i = tznf_list(1) Then
                zjtxf_jq = zjf_list(i) + txf_list(i)
            Else
                zjtxf_jq = (zjf_list(i) + txf_list(i)) - (zjf_list(i - 1) + txf_list(i - 1))
            End If
            '固定资产净值需要减去的部分（第1个建设年份：减去全部；之后的建设年份：减去增量）
            Dim gdzcjz_jq As Double
            If i = tznf_list(1) Then
                gdzcjz_jq = gdzcjz_raw_list(i)
            Else
                gdzcjz_jq = gdzcjz_raw_list(i) - gdzcjz_raw_list(i - 1) + zjf_list(i - 1)
            End If
            '无形资产净值需要减去的部分（第1个建设年份：减去全部；之后的建设年份：减去增量）
            Dim wxzcjz_jq As Double
            If i = tznf_list(1) Then
                wxzcjz_jq = wxzcjz_raw_list(i)
            Else
                wxzcjz_jq = wxzcjz_raw_list(i) - wxzcjz_raw_list(i - 1) + txf_list(i - 1)
            End If
            '在建工程
            zjgc_list(i) = (zjgc_raw_list(i) - zjtxf_jq) * jsq_index(i)
            '固定资产净值
            gdzcjz_list(i) = gdzcjz_raw_list(i) - gdzcjz_jq * jsq_index(i)
            '无形资产净值
            wxzcjz_list(i) = wxzcjz_raw_list(i) - wxzcjz_jq * jsq_index(i)
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算期末最后一年减去自有流动资金，1.1.4累计盈余资金
        Dim ljyyzj_list(31) As Double
        '前15年
        For i = 1 To 15
            If i = jsnx Then
                ljyyzj_list(i) = ExcelApp.ThisWorkbook.Worksheets("财务计划现金流量表").Cells(32, i + 4).Value - ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(14, i + 5).Value
            Else
                ljyyzj_list(i) = ExcelApp.ThisWorkbook.Worksheets("财务计划现金流量表").Cells(32, i + 4).Value
            End If
        Next
        '16~31年
        For i = 16 To 31
            If i = jsnx Then
                ljyyzj_list(i) = ExcelApp.ThisWorkbook.Worksheets("财务计划现金流量表").Cells(65, i - 12).Value - ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(32, i - 12).Value
            Else
                ljyyzj_list(i) = ExcelApp.ThisWorkbook.Worksheets("财务计划现金流量表").Cells(65, i - 12).Value
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '数据写入Excel
        '前15年
        For i = 1 To 15
            ExcelApp.ThisWorkbook.Worksheets("资产负债表").Cells(10, i + 4).Value = ljyyzj_list(i)
            ExcelApp.ThisWorkbook.Worksheets("资产负债表").Cells(11, i + 4).Value = zjgc_list(i)
            ExcelApp.ThisWorkbook.Worksheets("资产负债表").Cells(13, i + 4).Value = gdzcjz_list(i)
            ExcelApp.ThisWorkbook.Worksheets("资产负债表").Cells(14, i + 4).Value = wxzcjz_list(i)
        Next
        '16~31年
        For i = 16 To 31
            ExcelApp.ThisWorkbook.Worksheets("资产负债表").Cells(39, i - 12).Value = ljyyzj_list(i)
            ExcelApp.ThisWorkbook.Worksheets("资产负债表").Cells(40, i - 12).Value = zjgc_list(i)
            ExcelApp.ThisWorkbook.Worksheets("资产负债表").Cells(42, i - 12).Value = gdzcjz_list(i)
            ExcelApp.ThisWorkbook.Worksheets("资产负债表").Cells(43, i - 12).Value = wxzcjz_list(i)
        Next
        ExcelApp.Calculate()
        '————————————————————————————————————————————————————————————————————————————————————————
        '会计恒等式准确性验证
        Dim error_list(31) As Double '资产 - 负债 - 所有者权益
        Dim error_sum As Double = 0
        '前15年
        For i = 1 To 15
            error_list(i) = Math.Abs(ExcelApp.ThisWorkbook.Worksheets("资产负债表").Cells(5, i + 4).Value - ExcelApp.ThisWorkbook.Worksheets("资产负债表").Cells(15, i + 4).Value)
            error_sum += error_list(i)
        Next
        '16~31年
        For i = 16 To 31
            error_list(i) = Math.Abs(ExcelApp.ThisWorkbook.Worksheets("资产负债表").Cells(34, i - 12).Value - ExcelApp.ThisWorkbook.Worksheets("资产负债表").Cells(44, i - 12).Value)
            error_sum += error_list(i)
        Next
        If error_sum < 1 Then
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(12, 3).Value = "正确"
        Else
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(12, 3).Value = "不正确"
        End If
    End Sub
End Module
