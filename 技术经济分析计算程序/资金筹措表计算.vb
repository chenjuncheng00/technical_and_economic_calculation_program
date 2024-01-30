Module 资金筹措表计算
    Sub 投资计划与资金筹措表计算(ExcelApp As Object)
        '读取输入数据
        Dim GSBSJ = 读取估算表数据(ExcelApp)
        'Dim GSBCL = 读取计算常量设置(ExcelApp)
        '项目计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '无形资产所占比例
        Dim wxzcbl As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 5).Value
        'tznf_list：10次投资的年份序号，列表
        Dim tznf_list = GSBSJ(0)
        '建设静态投资
        Dim jttz_list = GSBSJ(1)
        Dim ans_jttz = 基础计算功能_10_to_31(tznf_list, jttz_list)
        '建设期利息
        Dim dklx_list = GSBSJ(5)
        Dim ans_dklx = 基础计算功能_10_to_31(tznf_list, dklx_list)
        '可抵扣增值税总额
        Dim kdkzzs_list = GSBSJ(9)
        Dim ans_kdkzzs = 基础计算功能_10_to_31(tznf_list, kdkzzs_list)
        '流动资金
        Dim ldzj_list = GSBSJ(6)
        Dim ans_ldzj = 基础计算功能_10_to_31(tznf_list, ldzj_list)
        '建设投资资本金
        Dim tzzbj_list = GSBSJ(3)
        Dim ans_tzzbj = 基础计算功能_10_to_31(tznf_list, tzzbj_list)
        '铺底流动资金
        Dim pdldzj_list = GSBSJ(7)
        Dim ans_pdldzj = 基础计算功能_10_to_31(tznf_list, pdldzj_list)
        '建设投资贷款
        Dim dkje_list = GSBSJ(4)
        Dim ans_dkje = 基础计算功能_10_to_31(tznf_list, dkje_list)
        '流动资金贷款
        Dim ldzjdk_list = GSBSJ(8)
        Dim ans_ldzjdk = 基础计算功能_10_to_31(tznf_list, ldzjdk_list)
        '写入Excel
        '1-15年
        For i = 1 To 15
            If i <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(6, 4 + i).Value = ans_jttz(i)
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(7, 4 + i).Value = ans_dklx(i)
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(8, 4 + i).Value = ans_kdkzzs(i)
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(10, 4 + i).Value = ans_ldzj(i)
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(13, 4 + i).Value = ans_tzzbj(i)
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(14, 4 + i).Value = ans_pdldzj(i)
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(17, 4 + i).Value = ans_dkje(i)
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(18, 4 + i).Value = ans_ldzjdk(i)
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(19, 4 + i).Value = (ans_jttz(i) + ans_dklx(i) - ans_kdkzzs(i)) * (1 - wxzcbl)
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(20, 4 + i).Value = (ans_jttz(i) + ans_dklx(i) - ans_kdkzzs(i)) * wxzcbl
            Else
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(6, 4 + i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(7, 4 + i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(8, 4 + i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(10, 4 + i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(13, 4 + i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(14, 4 + i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(17, 4 + i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(18, 4 + i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(19, 4 + i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(20, 4 + i).Value = 0
            End If
        Next
        '16-31年
        For i = 16 To 31
            If i <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(27, i - 12).Value = ans_jttz(i)
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(28, i - 12).Value = ans_dklx(i)
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(29, i - 12).Value = ans_kdkzzs(i)
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(31, i - 12).Value = ans_ldzj(i)
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(34, i - 12).Value = ans_tzzbj(i)
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(35, i - 12).Value = ans_pdldzj(i)
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(38, i - 12).Value = ans_dkje(i)
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(39, i - 12).Value = ans_ldzjdk(i)
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(40, i - 12).Value = (ans_jttz(i) + ans_dklx(i) - ans_kdkzzs(i)) * (1 - wxzcbl)
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(41, i - 12).Value = (ans_jttz(i) + ans_dklx(i) - ans_kdkzzs(i)) * wxzcbl
            Else
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(27, i - 12).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(28, i - 12).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(29, i - 12).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(31, i - 12).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(34, i - 12).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(35, i - 12).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(38, i - 12).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(39, i - 12).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(40, i - 12).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(41, i - 12).Value = 0
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
End Module
