Module 保险费计算
    Sub 保险费计算(ExcelApp As Object)
        '每次计算完折旧摊销和长期贷款后，都要计算一次本SUB
        '保险费金额只与固定资产原值OR固定资产净值有关

        '项目计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '读取保险费率
        Dim bxfl As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 5).Value
        '读取保险费计算基数设置
        Dim bxf_mode As String = ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(211, 22).Value
        '读取逐年固定资产原值
        Dim gdzc_yz(31) As Double
        '读取逐年固定资产净值
        Dim gdzc_jz(31) As Double
        '前15年
        For i = 1 To 15
            If i <= jsnx Then
                gdzc_yz(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(5, 4 + i).Value
                gdzc_jz(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(10, 4 + i).Value
            Else
                gdzc_yz(i) = 0
                gdzc_jz(i) = 0
            End If
        Next
        '16—31年
        For i = 16 To 31
            If i <= jsnx Then
                gdzc_yz(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(22, i - 12).Value
                gdzc_jz(i) = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(27, i - 12).Value
            Else
                gdzc_yz(i) = 0
                gdzc_jz(i) = 0
            End If
        Next

        '计算逐年保险费
        Dim ans_bxf(31) As Double
        If bxf_mode = "净值" Then
            '以固定资产净值为基础进行计算
            For i = 1 To 31
                ans_bxf(i) = gdzc_jz(i) * bxfl
            Next
        Else
            '以固定资产原值为基础进行计算(默认计算方式)
            For i = 1 To 31
                ans_bxf(i) = gdzc_yz(i) * bxfl
            Next
        End If

        '计算结果写入Excel
        '前15年
        For i = 1 To 15
            If i <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(211, 4 + i).Value = ans_bxf(i)
            Else
                ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(211, 4 + i).Value = 0
            End If
        Next
        '16—31年
        For i = 16 To 31
            If i <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(215, i - 12).Value = ans_bxf(i)
            Else
                ExcelApp.ThisWorkbook.Worksheets("成本税收表").Cells(215, i - 12).Value = 0
            End If
        Next
    End Sub
End Module
