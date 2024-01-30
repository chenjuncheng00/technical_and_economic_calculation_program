Module 收入成本表格操作
    Sub 隐藏收入税收表中收入为0的行(ExcelApp As Object)
        '定义局部变量
        Dim js As Integer
        '第1年到第15年
        For i = 1 To 13
            For j = 1 To 10
                '如果某一项收入为0，并且收入类型与输入的类型匹配，则隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 7).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 2).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(5 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(5 + j).EntireRow.Hidden = True
                End If
                '如果某一项收入大于0，并且收入类型与输入的类型匹配，则不隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 7).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 2).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(5 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(5 + j).EntireRow.Hidden = False
                End If
            Next
        Next
        For i = 1 To 13
            For j = 1 To 13
                '隐藏销项增值税为0的
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 7).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 2).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(64 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(64 + j).EntireRow.Hidden = True
                End If
                '增值税大于0，不隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 7).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 2).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(64 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(64 + j).EntireRow.Hidden = False
                End If
            Next
        Next
        '如果增值税退税比例为0，则隐藏增值税退税收入，前15年表格
        If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(172, 34).Value = 0 Then
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(16).EntireRow.Hidden = True
        Else
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(16).EntireRow.Hidden = False
        End If
        '隐藏补贴收入分项中为0的行，前15年表格
        For i = 1 To 13
            For j = 1 To 3
                '如果某一项收入为0，并且收入类型与输入的类型匹配，则隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 7).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 2).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(52 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(52 + j).EntireRow.Hidden = True
                End If
                '如果某一项收入大于0，并且收入类型与输入的类型匹配，则不隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 7).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 2).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(52 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(52 + j).EntireRow.Hidden = False
                End If
            Next
        Next
        '自动给收入税收表的条目编序号，前15年表格
        js = 0 '计数
        For i = 1 To 11
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(5 + i).Height > 0 Then
                js = js + 1
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(5 + i, 1).Value = 1 + js / 10
            End If
        Next
        '自动给销项增值税表编号，前15年表格
        js = 0 '计数
        For i = 1 To 13
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(64 + i).Height > 0 Then
                js = js + 1
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(64 + i, 1).Value = js
            End If
        Next
        '自动给补贴收入分项计算表编号，前15年表格
        js = 0 '计数
        For i = 1 To 3
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(52 + i).Height > 0 Then
                js = js + 1
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(52 + i, 1).Value = js
            End If
        Next
        '第16年到第30年
        For i = 1 To 13
            For j = 1 To 10
                '如果某一项收入为0，并且收入类型与输入的类型匹配，则隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 7).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 2).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(28 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(28 + j).EntireRow.Hidden = True
                End If
                '如果某一项收入大于0，并且收入类型与输入的类型匹配，则隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 7).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 2).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(28 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(28 + j).EntireRow.Hidden = False
                End If
            Next
        Next
        For i = 1 To 13
            For j = 1 To 13
                '隐藏销项增值税为0的
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 7).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 2).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(79 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(79 + j).EntireRow.Hidden = True
                End If
                '增值税大于0，不隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 7).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 2).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(79 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(79 + j).EntireRow.Hidden = False
                End If
            Next
        Next
        '如果增值税退税比例为0，则隐藏增值税退税收入，后15年表格
        If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(172, 34).Value = 0 Then
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(39).EntireRow.Hidden = True
        Else
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(39).EntireRow.Hidden = False
        End If
        '隐藏补贴收入分项中为0的行，后15年表格
        For i = 1 To 13
            For j = 1 To 3
                '如果某一项收入为0，并且收入类型与输入的类型匹配，则隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 7).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 2).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(58 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(58 + j).EntireRow.Hidden = True
                End If
                '如果某一项收入大于0，并且收入类型与输入的类型匹配，则隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 7).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 2).Value = ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(58 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(58 + j).EntireRow.Hidden = False
                End If
            Next
        Next
        '自动给收入税收表的条目编序号，后15年表格
        js = 0 '计数
        For i = 1 To 11
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(28 + i).Height > 0 Then
                js = js + 1
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(28 + i, 1).Value = 1 + js / 10
            End If
        Next
        '自动给销项增值税表编号，后15年表格
        js = 0 '计数
        For i = 1 To 13
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(79 + i).Height > 0 Then
                js = js + 1
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(79 + i, 1).Value = js
            End If
        Next
        '自动给补贴收入分项计算表编号，后15年表格
        js = 0 '计数
        For i = 1 To 3
            If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Rows(58 + i).Height > 0 Then
                js = js + 1
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(58 + i, 1).Value = js
            End If
        Next
    End Sub
    Sub 隐藏总成本表中成本为0的行(ExcelApp As Object)
        '定义局部变量
        Dim js As Integer
        '第1年到第15年
        For i = 1 To 15
            For j = 1 To 15
                '如果某一项成本为0，并且成本类型与输入的类型匹配，则隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 14).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 9).Value = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(4 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("总成本表").Rows(4 + j).EntireRow.Hidden = True
                End If
                '如果某一项成本大于0，并且成本类型与输入的类型匹配，则隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 14).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 9).Value = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(4 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("总成本表").Rows(4 + j).EntireRow.Hidden = False
                End If
            Next
        Next
        '自动给总成本表的条目编序号
        js = 0 '计数
        For i = 1 To 19
            If ExcelApp.ThisWorkbook.Worksheets("总成本表").Rows(4 + i).Height > 0 Then
                js = js + 1
                ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(4 + i, 1).Value = js
            End If
        Next
        '第16年到第30年
        For i = 1 To 15
            For j = 1 To 15
                '如果某一项成本为0，并且成本类型与输入的类型匹配，则隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 14).Value = 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 9).Value = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(36 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("总成本表").Rows(36 + j).EntireRow.Hidden = True
                End If
                '如果某一项成本大于0，并且成本类型与输入的类型匹配，则隐藏
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 14).Value > 0 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(12 + i, 9).Value = ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(36 + j, 2).Value Then
                    ExcelApp.ThisWorkbook.Worksheets("总成本表").Rows(36 + j).EntireRow.Hidden = False
                End If
            Next
        Next
        '自动给总成本表的条目编序号
        js = 0 '计数
        For i = 1 To 19
            If ExcelApp.ThisWorkbook.Worksheets("总成本表").Rows(36 + i).Height > 0 Then
                js = js + 1
                ExcelApp.ThisWorkbook.Worksheets("总成本表").Cells(36 + i, 1).Value = js
            End If
        Next
    End Sub
End Module
