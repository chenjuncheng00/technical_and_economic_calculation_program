Module 利润相关计算
    Sub 计算还款利润(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '每次计算完折旧摊销和长期贷款后，都要计算一次本SUB
        '当折旧费不够还贷款本金时，用当年利润还

        '读取项目总的计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '长期贷款还本金额
        Dim znhb(31) As Double
        '固定资产折旧+无形资产摊销金额
        Dim zjtx(31) As Double
        '前15年
        For i = 1 To 15
            znhb(i) = ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(9, i + 4).Value
            zjtx(i) = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(22, i + 4).Value
        Next
        '16-31年
        For i = 16 To 31
            znhb(i) = ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(39, i - 12).Value
            zjtx(i) = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(54, i - 12).Value
        Next
        '计算逐年还款利润
        Dim hklr(31) As Double
        For i = 1 To 31
            If znhb(i) - zjtx(i) > 0 And i <= jsnx Then
                hklr(i) = znhb(i) - zjtx(i)
            Else
                hklr(i) = 0
            End If
        Next
        '结果写入Excel
        '前15年
        For i = 1 To 15
            ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(17, i + 4).Value = hklr(i)
        Next
        '16-31年
        For i = 16 To 31
            ExcelApp.ThisWorkbook.Worksheets("利润与利润分配表").Cells(37, i - 12).Value = hklr(i)
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
End Module
