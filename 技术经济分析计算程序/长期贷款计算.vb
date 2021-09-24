Module 长期贷款计算
    Sub 长期贷款计算(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '读取输入数据
        Dim GSBSJ = 读取估算表数据(ExcelApp)
        '项目计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '宽限期是否付息
        Dim kxqfx As Boolean
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 11).Value = "付息" Then
            kxqfx = True
        Else
            kxqfx = False
        End If
        '贷款年利率
        Dim dkll(10) As Double
        '贷款年限
        Dim dknx(10) As Integer
        '宽限年限
        Dim kxnx(10) As Integer
        For i = 1 To 10
            dkll(i) = ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i + 11, 39).Value
            dknx(i) = ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i + 22, 39).Value
            kxnx(i) = ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i + 33, 39).Value
        Next
        '贷款开始年份
        Dim ksnf(10) As Integer
        '贷款金额
        Dim bj_0(10) As Double
        '每年投产月份数
        Dim tcyfs(31) As Integer
        '读取开始年份和资产原值，前5年
        Dim ksnf_list = GSBSJ(0)
        Dim dkje_list = GSBSJ(4)
        For i = 1 To 5
            ksnf(i) = ksnf_list(i)
            bj_0(i) = dkje_list(i)
        Next
        '读取开始年份和资产原值，后5年
        For i = 6 To 10
            ksnf(i) = ksnf_list(i)
            bj_0(i) = dkje_list(i)
        Next
        '读取每年投产月份数
        For i = 1 To 31
            tcyfs(i) = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 2).Value
        Next
        '长期贷款计算
        Dim ans_DKJS
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法一" Then
            ans_DKJS = 长期贷款计算_等额本息_10次投资合并计算(bj_0, dknx(1), ksnf, jsnx, dkll(1), kxnx(1), tcyfs, kxqfx)
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法二" Then
            ans_DKJS = 长期贷款计算_等额本金_10次投资合并计算(bj_0, dknx(1), ksnf, jsnx, dkll(1), kxnx(1), tcyfs, kxqfx)
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法三" Then
            ans_DKJS = 长期贷款计算_等额本息_10次投资分开计算(bj_0, dknx, ksnf, jsnx, dkll, kxnx, tcyfs, kxqfx)
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法四" Then
            ans_DKJS = 长期贷款计算_等额本金_10次投资分开计算(bj_0, dknx, ksnf, jsnx, dkll, kxnx, tcyfs, kxqfx)
        End If
        '结果写入Excel
        '前15年
        For i = 1 To 15
            If i <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(7, 4 + i).Value = ans_DKJS(3)(i) + ans_DKJS(2)(i)
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(8, 4 + i).Value = ans_DKJS(0)(i)
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(9, 4 + i).Value = ans_DKJS(2)(i)
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(10, 4 + i).Value = ans_DKJS(1)(i)
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(11, 4 + i).Value = ans_DKJS(3)(i)
            Else
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(7, 4 + i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(8, 4 + i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(9, 4 + i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(10, 4 + i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(11, 4 + i).Value = 0
            End If
        Next
        '16—31年
        For i = 16 To 31
            If i <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(37, i - 12).Value = ans_DKJS(3)(i) + ans_DKJS(2)(i)
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(38, i - 12).Value = ans_DKJS(0)(i)
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(39, i - 12).Value = ans_DKJS(2)(i)
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(40, i - 12).Value = ans_DKJS(1)(i)
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(41, i - 12).Value = ans_DKJS(3)(i)
            Else
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(37, i - 12).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(38, i - 12).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(39, i - 12).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(40, i - 12).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(41, i - 12).Value = 0
            End If
        Next
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '计算一次Excel
        ExcelApp.Calculate()
    End Sub
    Function 长期贷款计算_等额本息_10次投资合并计算(bj_0_list As Array, dknx_0 As Integer, ksnf_list As Array, jsnx As Integer,
                                              dkll As Double, kxnx As Integer, tcyfs_list As Array, kxqfx As Boolean)
        'bj_0_list: 整个计算期内，逐年新增的贷款本金初始值，列表，长度10
        'dnkx_0：长期贷款年限初始值
        'ksnf_list：10次投资，贷款计算的开始年份，列表，长度10
        'jsnx：整个项目的最大计算年限(含建设期)
        'dkll：长期贷款年化利率(复利)
        'kxnx：长期贷款宽限期年份数
        'tcyfs_list：逐年投产的月份数，列表，长度31
        'kxqfx：宽限期内是否付息

        '10次投资长期贷款还本付息合并在一起进行计算

        '第1次投资
        Dim bj_0_1 As Double = bj_0_list(1)
        Dim ksnf_1 As Integer = ksnf_list(1)
        Dim dknx_0_1 As Integer
        If dknx_0 + 1 - ksnf_1 > 0 Then
            dknx_0_1 = Math.Min(dknx_0, dknx_0 + 1 - ksnf_1)
        Else
            dknx_0_1 = Math.Min(dknx_0, jsnx - ksnf_1)
        End If
        Dim kxnx_1 As Integer = kxnx
        '第2次投资
        Dim bj_0_2 As Double = bj_0_list(2)
        Dim ksnf_2 As Integer = ksnf_list(2)
        Dim dknx_0_2 As Integer
        If dknx_0 + 1 - ksnf_2 > 0 Then
            dknx_0_2 = Math.Min(dknx_0, dknx_0 + 1 - ksnf_2)
        Else
            dknx_0_2 = Math.Min(dknx_0, jsnx - ksnf_2)
        End If
        Dim kxnx_2 As Integer = Math.Min(Math.Max(kxnx - (ksnf_2 - ksnf_1), 0), kxnx)
        '第3次投资
        Dim bj_0_3 As Double = bj_0_list(3)
        Dim ksnf_3 As Integer = ksnf_list(3)
        Dim dknx_0_3 As Integer
        If dknx_0 + 1 - ksnf_3 > 0 Then
            dknx_0_3 = Math.Min(dknx_0, dknx_0 + 1 - ksnf_3)
        Else
            dknx_0_3 = Math.Min(dknx_0, jsnx - ksnf_3)
        End If
        Dim kxnx_3 As Integer = Math.Min(Math.Max(kxnx - (ksnf_3 - ksnf_1), 0), kxnx)
        '第4次投资
        Dim bj_0_4 As Double = bj_0_list(4)
        Dim ksnf_4 As Integer = ksnf_list(4)
        Dim dknx_0_4 As Integer
        If dknx_0 + 1 - ksnf_4 > 0 Then
            dknx_0_4 = Math.Min(dknx_0, dknx_0 + 1 - ksnf_4)
        Else
            dknx_0_4 = Math.Min(dknx_0, jsnx - ksnf_4)
        End If
        Dim kxnx_4 As Integer = Math.Min(Math.Max(kxnx - (ksnf_4 - ksnf_1), 0), kxnx)
        '第5次投资
        Dim bj_0_5 As Double = bj_0_list(5)
        Dim ksnf_5 As Integer = ksnf_list(5)
        Dim dknx_0_5 As Integer
        If dknx_0 + 1 - ksnf_5 > 0 Then
            dknx_0_5 = Math.Min(dknx_0, dknx_0 + 1 - ksnf_5)
        Else
            dknx_0_5 = Math.Min(dknx_0, jsnx - ksnf_5)
        End If
        Dim kxnx_5 As Integer = Math.Min(Math.Max(kxnx - (ksnf_5 - ksnf_1), 0), kxnx)
        '第6次投资
        Dim bj_0_6 As Double = bj_0_list(6)
        Dim ksnf_6 As Integer = ksnf_list(6)
        Dim dknx_0_6 As Integer
        If dknx_0 + 1 - ksnf_6 > 0 Then
            dknx_0_6 = Math.Min(dknx_0, dknx_0 + 1 - ksnf_6)
        Else
            dknx_0_6 = Math.Min(dknx_0, jsnx - ksnf_6)
        End If
        Dim kxnx_6 As Integer = Math.Min(Math.Max(kxnx - (ksnf_6 - ksnf_1), 0), kxnx)
        '第7次投资
        Dim bj_0_7 As Double = bj_0_list(7)
        Dim ksnf_7 As Integer = ksnf_list(7)
        Dim dknx_0_7 As Integer
        If dknx_0 + 1 - ksnf_7 > 0 Then
            dknx_0_7 = Math.Min(dknx_0, dknx_0 + 1 - ksnf_7)
        Else
            dknx_0_7 = Math.Min(dknx_0, jsnx - ksnf_7)
        End If
        Dim kxnx_7 As Integer = Math.Min(Math.Max(kxnx - (ksnf_7 - ksnf_1), 0), kxnx)
        '第8次投资
        Dim bj_0_8 As Double = bj_0_list(8)
        Dim ksnf_8 As Integer = ksnf_list(8)
        Dim dknx_0_8 As Integer
        If dknx_0 + 1 - ksnf_8 > 0 Then
            dknx_0_8 = Math.Min(dknx_0, dknx_0 + 1 - ksnf_8)
        Else
            dknx_0_8 = Math.Min(dknx_0, jsnx - ksnf_8)
        End If
        Dim kxnx_8 As Integer = Math.Min(Math.Max(kxnx - (ksnf_8 - ksnf_1), 0), kxnx)
        '第9次投资
        Dim bj_0_9 As Double = bj_0_list(9)
        Dim ksnf_9 As Integer = ksnf_list(9)
        Dim dknx_0_9 As Integer
        If dknx_0 + 1 - ksnf_9 > 0 Then
            dknx_0_9 = Math.Min(dknx_0, dknx_0 + 1 - ksnf_9)
        Else
            dknx_0_9 = Math.Min(dknx_0, jsnx - ksnf_9)
        End If
        Dim kxnx_9 As Integer = Math.Min(Math.Max(kxnx - (ksnf_9 - ksnf_1), 0), kxnx)
        '第10次投资
        Dim bj_0_10 As Double = bj_0_list(10)
        Dim ksnf_10 As Integer = ksnf_list(10)
        Dim dknx_0_10 As Integer
        If dknx_0 + 1 - ksnf_10 > 0 Then
            dknx_0_10 = Math.Min(dknx_0, dknx_0 + 1 - ksnf_10)
        Else
            dknx_0_10 = Math.Min(dknx_0, jsnx - ksnf_10)
        End If
        Dim kxnx_10 As Integer = Math.Min(Math.Max(kxnx - (ksnf_10 - ksnf_1), 0), kxnx)

        '列表，储存计算结果
        Dim DNBXZE(31) As Double '当年还本付息总额
        Dim DNFX(31) As Double '当年付息额
        Dim DNHB(31) As Double '当年还本额
        Dim BJYE(31) As Double '当年本金余额

        '分次进行计算
        ' 只有第1次投资考虑宽限期，后面不考虑
        Dim ans_1 = 长期贷款计算_等额本息(bj_0_1, dknx_0_1, ksnf_1, jsnx, dkll, kxnx_1, tcyfs_list, kxqfx)
        Dim ans_2 = 长期贷款计算_等额本息(bj_0_2, dknx_0_2, ksnf_2, jsnx, dkll, kxnx_2, tcyfs_list, kxqfx)
        Dim ans_3 = 长期贷款计算_等额本息(bj_0_3, dknx_0_3, ksnf_3, jsnx, dkll, kxnx_3, tcyfs_list, kxqfx)
        Dim ans_4 = 长期贷款计算_等额本息(bj_0_4, dknx_0_4, ksnf_4, jsnx, dkll, kxnx_4, tcyfs_list, kxqfx)
        Dim ans_5 = 长期贷款计算_等额本息(bj_0_5, dknx_0_5, ksnf_5, jsnx, dkll, kxnx_5, tcyfs_list, kxqfx)
        Dim ans_6 = 长期贷款计算_等额本息(bj_0_6, dknx_0_6, ksnf_6, jsnx, dkll, kxnx_6, tcyfs_list, kxqfx)
        Dim ans_7 = 长期贷款计算_等额本息(bj_0_7, dknx_0_7, ksnf_7, jsnx, dkll, kxnx_7, tcyfs_list, kxqfx)
        Dim ans_8 = 长期贷款计算_等额本息(bj_0_8, dknx_0_8, ksnf_8, jsnx, dkll, kxnx_8, tcyfs_list, kxqfx)
        Dim ans_9 = 长期贷款计算_等额本息(bj_0_9, dknx_0_9, ksnf_9, jsnx, dkll, kxnx_9, tcyfs_list, kxqfx)
        Dim ans_10 = 长期贷款计算_等额本息(bj_0_10, dknx_0_10, ksnf_10, jsnx, dkll, kxnx_10, tcyfs_list, kxqfx)
        '累加
        For i = 1 To 31 '第1年到第31年
            '当年还本付息总额
            DNBXZE(i) = ans_1(0)(i) + ans_2(0)(i) + ans_3(0)(i) + ans_4(0)(i) + ans_5(0)(i) + ans_6(0)(i) +
                        ans_7(0)(i) + ans_8(0)(i) + ans_9(0)(i) + ans_10(0)(i)
            '当年付息
            DNFX(i) = ans_1(1)(i) + ans_2(1)(i) + ans_3(1)(i) + ans_4(1)(i) + ans_5(1)(i) + ans_6(1)(i) +
                      ans_7(1)(i) + ans_8(1)(i) + ans_9(1)(i) + ans_10(1)(i)
            '当年还本
            DNHB(i) = ans_1(2)(i) + ans_2(2)(i) + ans_3(2)(i) + ans_4(2)(i) + ans_5(2)(i) + ans_6(2)(i) +
                      ans_7(2)(i) + ans_8(2)(i) + ans_9(2)(i) + ans_10(2)(i)
            '当年本金余额
            BJYE(i) = ans_1(3)(i) + ans_2(3)(i) + ans_3(3)(i) + ans_4(3)(i) + ans_5(3)(i) + ans_6(3)(i) +
                       ans_7(3)(i) + ans_8(3)(i) + ans_9(3)(i) + ans_10(3)(i)
        Next
        '返回结果
        Dim ans(3)
        ans(0) = DNBXZE
        ans(1) = DNFX
        ans(2) = DNHB
        ans(3) = BJYE
        Return ans
    End Function
    Function 长期贷款计算_等额本金_10次投资合并计算(bj_0_list As Array, dknx_0 As Integer, ksnf_list As Array, jsnx As Integer,
                                              dkll As Double, kxnx As Integer, tcyfs_list As Array, kxqfx As Boolean)
        'bj_0_list: 整个计算期内，逐年新增的贷款本金初始值，列表，长度10
        'dnkx_0：长期贷款年限初始值
        'ksnf_list：10次投资，贷款计算的开始年份，列表，长度10
        'jsnx：整个项目的最大计算年限(含建设期)
        'dkll：长期贷款年化利率(复利)
        'kxnx：长期贷款宽限期年份数
        'tcyfs_list：逐年投产的月份数，列表，长度31
        'kxqfx：宽限期内是否付息

        '10次投资长期贷款还本付息合并在一起进行计算

        '第1次投资
        Dim bj_0_1 As Double = bj_0_list(1)
        Dim ksnf_1 As Integer = ksnf_list(1)
        Dim dknx_0_1 As Integer
        If dknx_0 + 1 - ksnf_1 > 0 Then
            dknx_0_1 = Math.Min(dknx_0, dknx_0 + 1 - ksnf_1)
        Else
            dknx_0_1 = Math.Min(dknx_0, jsnx - ksnf_1)
        End If
        Dim kxnx_1 As Integer = kxnx
        '第2次投资
        Dim bj_0_2 As Double = bj_0_list(2)
        Dim ksnf_2 As Integer = ksnf_list(2)
        Dim dknx_0_2 As Integer
        If dknx_0 + 1 - ksnf_2 > 0 Then
            dknx_0_2 = Math.Min(dknx_0, dknx_0 + 1 - ksnf_2)
        Else
            dknx_0_2 = Math.Min(dknx_0, jsnx - ksnf_2)
        End If
        Dim kxnx_2 As Integer = Math.Min(Math.Max(kxnx - (ksnf_2 - ksnf_1), 0), kxnx)
        '第3次投资
        Dim bj_0_3 As Double = bj_0_list(3)
        Dim ksnf_3 As Integer = ksnf_list(3)
        Dim dknx_0_3 As Integer
        If dknx_0 + 1 - ksnf_3 > 0 Then
            dknx_0_3 = Math.Min(dknx_0, dknx_0 + 1 - ksnf_3)
        Else
            dknx_0_3 = Math.Min(dknx_0, jsnx - ksnf_3)
        End If
        Dim kxnx_3 As Integer = Math.Min(Math.Max(kxnx - (ksnf_3 - ksnf_1), 0), kxnx)
        '第4次投资
        Dim bj_0_4 As Double = bj_0_list(4)
        Dim ksnf_4 As Integer = ksnf_list(4)
        Dim dknx_0_4 As Integer
        If dknx_0 + 1 - ksnf_4 > 0 Then
            dknx_0_4 = Math.Min(dknx_0, dknx_0 + 1 - ksnf_4)
        Else
            dknx_0_4 = Math.Min(dknx_0, jsnx - ksnf_4)
        End If
        Dim kxnx_4 As Integer = Math.Min(Math.Max(kxnx - (ksnf_4 - ksnf_1), 0), kxnx)
        '第5次投资
        Dim bj_0_5 As Double = bj_0_list(5)
        Dim ksnf_5 As Integer = ksnf_list(5)
        Dim dknx_0_5 As Integer
        If dknx_0 + 1 - ksnf_5 > 0 Then
            dknx_0_5 = Math.Min(dknx_0, dknx_0 + 1 - ksnf_5)
        Else
            dknx_0_5 = Math.Min(dknx_0, jsnx - ksnf_5)
        End If
        Dim kxnx_5 As Integer = Math.Min(Math.Max(kxnx - (ksnf_5 - ksnf_1), 0), kxnx)
        '第6次投资
        Dim bj_0_6 As Double = bj_0_list(6)
        Dim ksnf_6 As Integer = ksnf_list(6)
        Dim dknx_0_6 As Integer
        If dknx_0 + 1 - ksnf_6 > 0 Then
            dknx_0_6 = Math.Min(dknx_0, dknx_0 + 1 - ksnf_6)
        Else
            dknx_0_6 = Math.Min(dknx_0, jsnx - ksnf_6)
        End If
        Dim kxnx_6 As Integer = Math.Min(Math.Max(kxnx - (ksnf_6 - ksnf_1), 0), kxnx)
        '第7次投资
        Dim bj_0_7 As Double = bj_0_list(7)
        Dim ksnf_7 As Integer = ksnf_list(7)
        Dim dknx_0_7 As Integer
        If dknx_0 + 1 - ksnf_7 > 0 Then
            dknx_0_7 = Math.Min(dknx_0, dknx_0 + 1 - ksnf_7)
        Else
            dknx_0_7 = Math.Min(dknx_0, jsnx - ksnf_7)
        End If
        Dim kxnx_7 As Integer = Math.Min(Math.Max(kxnx - (ksnf_7 - ksnf_1), 0), kxnx)
        '第8次投资
        Dim bj_0_8 As Double = bj_0_list(8)
        Dim ksnf_8 As Integer = ksnf_list(8)
        Dim dknx_0_8 As Integer
        If dknx_0 + 1 - ksnf_8 > 0 Then
            dknx_0_8 = Math.Min(dknx_0, dknx_0 + 1 - ksnf_8)
        Else
            dknx_0_8 = Math.Min(dknx_0, jsnx - ksnf_8)
        End If
        Dim kxnx_8 As Integer = Math.Min(Math.Max(kxnx - (ksnf_8 - ksnf_1), 0), kxnx)
        '第9次投资
        Dim bj_0_9 As Double = bj_0_list(9)
        Dim ksnf_9 As Integer = ksnf_list(9)
        Dim dknx_0_9 As Integer
        If dknx_0 + 1 - ksnf_9 > 0 Then
            dknx_0_9 = Math.Min(dknx_0, dknx_0 + 1 - ksnf_9)
        Else
            dknx_0_9 = Math.Min(dknx_0, jsnx - ksnf_9)
        End If
        Dim kxnx_9 As Integer = Math.Min(Math.Max(kxnx - (ksnf_9 - ksnf_1), 0), kxnx)
        '第10次投资
        Dim bj_0_10 As Double = bj_0_list(10)
        Dim ksnf_10 As Integer = ksnf_list(10)
        Dim dknx_0_10 As Integer
        If dknx_0 + 1 - ksnf_10 > 0 Then
            dknx_0_10 = Math.Min(dknx_0, dknx_0 + 1 - ksnf_10)
        Else
            dknx_0_10 = Math.Min(dknx_0, jsnx - ksnf_10)
        End If
        Dim kxnx_10 As Integer = Math.Min(Math.Max(kxnx - (ksnf_10 - ksnf_1), 0), kxnx)

        '列表，储存计算结果
        Dim DNBXZE(31) As Double '当年还本付息总额
        Dim DNFX(31) As Double '当年付息额
        Dim DNHB(31) As Double '当年还本额
        Dim BJYE(31) As Double '当年本金余额

        '分次进行计算
        ' 只有第1次投资考虑宽限期，后面不考虑
        Dim ans_1 = 长期贷款计算_等额本金(bj_0_1, dknx_0_1, ksnf_1, jsnx, dkll, kxnx_1, tcyfs_list, kxqfx)
        Dim ans_2 = 长期贷款计算_等额本金(bj_0_2, dknx_0_2, ksnf_2, jsnx, dkll, kxnx_2, tcyfs_list, kxqfx)
        Dim ans_3 = 长期贷款计算_等额本金(bj_0_3, dknx_0_3, ksnf_3, jsnx, dkll, kxnx_3, tcyfs_list, kxqfx)
        Dim ans_4 = 长期贷款计算_等额本金(bj_0_4, dknx_0_4, ksnf_4, jsnx, dkll, kxnx_4, tcyfs_list, kxqfx)
        Dim ans_5 = 长期贷款计算_等额本金(bj_0_5, dknx_0_5, ksnf_5, jsnx, dkll, kxnx_5, tcyfs_list, kxqfx)
        Dim ans_6 = 长期贷款计算_等额本金(bj_0_6, dknx_0_6, ksnf_6, jsnx, dkll, kxnx_6, tcyfs_list, kxqfx)
        Dim ans_7 = 长期贷款计算_等额本金(bj_0_7, dknx_0_7, ksnf_7, jsnx, dkll, kxnx_7, tcyfs_list, kxqfx)
        Dim ans_8 = 长期贷款计算_等额本金(bj_0_8, dknx_0_8, ksnf_8, jsnx, dkll, kxnx_8, tcyfs_list, kxqfx)
        Dim ans_9 = 长期贷款计算_等额本金(bj_0_9, dknx_0_9, ksnf_9, jsnx, dkll, kxnx_9, tcyfs_list, kxqfx)
        Dim ans_10 = 长期贷款计算_等额本金(bj_0_10, dknx_0_10, ksnf_10, jsnx, dkll, kxnx_10, tcyfs_list, kxqfx)
        '累加
        For i = 1 To 31 '第1年到第31年
            '当年还本付息总额
            DNBXZE(i) = ans_1(0)(i) + ans_2(0)(i) + ans_3(0)(i) + ans_4(0)(i) + ans_5(0)(i) + ans_6(0)(i) +
                        ans_7(0)(i) + ans_8(0)(i) + ans_9(0)(i) + ans_10(0)(i)
            '当年付息
            DNFX(i) = ans_1(1)(i) + ans_2(1)(i) + ans_3(1)(i) + ans_4(1)(i) + ans_5(1)(i) + ans_6(1)(i) +
                      ans_7(1)(i) + ans_8(1)(i) + ans_9(1)(i) + ans_10(1)(i)
            '当年还本
            DNHB(i) = ans_1(2)(i) + ans_2(2)(i) + ans_3(2)(i) + ans_4(2)(i) + ans_5(2)(i) + ans_6(2)(i) +
                      ans_7(2)(i) + ans_8(2)(i) + ans_9(2)(i) + ans_10(2)(i)
            '当年本金余额
            BJYE(i) = ans_1(3)(i) + ans_2(3)(i) + ans_3(3)(i) + ans_4(3)(i) + ans_5(3)(i) + ans_6(3)(i) +
                       ans_7(3)(i) + ans_8(3)(i) + ans_9(3)(i) + ans_10(3)(i)
        Next
        '返回结果
        Dim ans(3)
        ans(0) = DNBXZE
        ans(1) = DNFX
        ans(2) = DNHB
        ans(3) = BJYE
        Return ans
    End Function
    Function 长期贷款计算_等额本息_10次投资分开计算(bj_0_list As Array, dknx_0_list As Array, ksnf_list As Array, jsnx As Integer,
                                                    dkll_list As Array, kxnx_list As Array, tcyfs_list As Array, kxqfx As Boolean)
        'bj_0_list: 10次投资，贷款本金初始值，列表，长度10
        'dnkx_0_list：10次投资，长期贷款年限初始值，列表，长度10
        'ksnf_list：10次投资，贷款计算的开始年份，列表，长度10
        'jsnx：整个项目的最大计算年限(含建设期)
        'dkll_list：10次投资，长期贷款年化利率(复利)，列表，长度10
        'kxnx_list：10次投资，长期贷款宽限年限，列表，长度10
        'tcyfs_list：逐年投产的月份数，列表，长度31
        'kxqfx：宽限期内是否付息

        '10次投资分别计算长期贷款还本付息，再累加在一起

        '第1次投资
        Dim bj_0_1 As Double = bj_0_list(1)
        Dim dknx_0_1 As Integer = dknx_0_list(1)
        Dim ksnf_1 As Integer = ksnf_list(1)
        Dim dkll_1 As Double = dkll_list(1)
        Dim kxnx_1 As Integer = kxnx_list(1)
        '第2次投资
        Dim bj_0_2 As Double = bj_0_list(2)
        Dim dknx_0_2 As Integer = dknx_0_list(2)
        Dim ksnf_2 As Integer = ksnf_list(2)
        Dim dkll_2 As Double = dkll_list(2)
        Dim kxnx_2 As Integer = kxnx_list(2)
        '第3次投资
        Dim bj_0_3 As Double = bj_0_list(3)
        Dim dknx_0_3 As Integer = dknx_0_list(3)
        Dim ksnf_3 As Integer = ksnf_list(3)
        Dim dkll_3 As Double = dkll_list(3)
        Dim kxnx_3 As Integer = kxnx_list(3)
        '第4次投资
        Dim bj_0_4 As Double = bj_0_list(4)
        Dim dknx_0_4 As Integer = dknx_0_list(4)
        Dim ksnf_4 As Integer = ksnf_list(4)
        Dim dkll_4 As Double = dkll_list(4)
        Dim kxnx_4 As Integer = kxnx_list(4)
        '第5次投资
        Dim bj_0_5 As Double = bj_0_list(5)
        Dim dknx_0_5 As Integer = dknx_0_list(5)
        Dim ksnf_5 As Integer = ksnf_list(5)
        Dim dkll_5 As Double = dkll_list(5)
        Dim kxnx_5 As Integer = kxnx_list(5)
        '第6次投资
        Dim bj_0_6 As Double = bj_0_list(6)
        Dim dknx_0_6 As Integer = dknx_0_list(6)
        Dim ksnf_6 As Integer = ksnf_list(6)
        Dim dkll_6 As Double = dkll_list(6)
        Dim kxnx_6 As Integer = kxnx_list(6)
        '第7次投资
        Dim bj_0_7 As Double = bj_0_list(7)
        Dim dknx_0_7 As Integer = dknx_0_list(7)
        Dim ksnf_7 As Integer = ksnf_list(7)
        Dim dkll_7 As Double = dkll_list(7)
        Dim kxnx_7 As Integer = kxnx_list(7)
        '第8次投资
        Dim bj_0_8 As Double = bj_0_list(8)
        Dim dknx_0_8 As Integer = dknx_0_list(8)
        Dim ksnf_8 As Integer = ksnf_list(8)
        Dim dkll_8 As Double = dkll_list(8)
        Dim kxnx_8 As Integer = kxnx_list(8)
        '第9次投资
        Dim bj_0_9 As Double = bj_0_list(9)
        Dim dknx_0_9 As Integer = dknx_0_list(9)
        Dim ksnf_9 As Integer = ksnf_list(9)
        Dim dkll_9 As Double = dkll_list(9)
        Dim kxnx_9 As Integer = kxnx_list(9)
        '第10次投资
        Dim bj_0_10 As Double = bj_0_list(10)
        Dim dknx_0_10 As Integer = dknx_0_list(10)
        Dim ksnf_10 As Integer = ksnf_list(10)
        Dim dkll_10 As Double = dkll_list(10)
        Dim kxnx_10 As Integer = kxnx_list(10)

        '列表，储存计算结果
        Dim DNBXZE(31) As Double '当年还本付息总额
        Dim DNFX(31) As Double '当年付息额
        Dim DNHB(31) As Double '当年还本额
        Dim BJYE(31) As Double '当年本金余额

        '分次进行计算
        Dim ans_1 = 长期贷款计算_等额本息(bj_0_1, dknx_0_1, ksnf_1, jsnx, dkll_1, kxnx_1, tcyfs_list, kxqfx)
        Dim ans_2 = 长期贷款计算_等额本息(bj_0_2, dknx_0_2, ksnf_2, jsnx, dkll_2, kxnx_2, tcyfs_list, kxqfx)
        Dim ans_3 = 长期贷款计算_等额本息(bj_0_3, dknx_0_3, ksnf_3, jsnx, dkll_3, kxnx_3, tcyfs_list, kxqfx)
        Dim ans_4 = 长期贷款计算_等额本息(bj_0_4, dknx_0_4, ksnf_4, jsnx, dkll_4, kxnx_4, tcyfs_list, kxqfx)
        Dim ans_5 = 长期贷款计算_等额本息(bj_0_5, dknx_0_5, ksnf_5, jsnx, dkll_5, kxnx_5, tcyfs_list, kxqfx)
        Dim ans_6 = 长期贷款计算_等额本息(bj_0_6, dknx_0_6, ksnf_6, jsnx, dkll_6, kxnx_6, tcyfs_list, kxqfx)
        Dim ans_7 = 长期贷款计算_等额本息(bj_0_7, dknx_0_7, ksnf_7, jsnx, dkll_7, kxnx_7, tcyfs_list, kxqfx)
        Dim ans_8 = 长期贷款计算_等额本息(bj_0_8, dknx_0_8, ksnf_8, jsnx, dkll_8, kxnx_8, tcyfs_list, kxqfx)
        Dim ans_9 = 长期贷款计算_等额本息(bj_0_9, dknx_0_9, ksnf_9, jsnx, dkll_9, kxnx_9, tcyfs_list, kxqfx)
        Dim ans_10 = 长期贷款计算_等额本息(bj_0_10, dknx_0_10, ksnf_10, jsnx, dkll_10, kxnx_10, tcyfs_list, kxqfx)
        '累加
        For i = 1 To 31 '第1年到第31年
            '当年还本付息总额
            DNBXZE(i) = ans_1(0)(i) + ans_2(0)(i) + ans_3(0)(i) + ans_4(0)(i) + ans_5(0)(i) + ans_6(0)(i) +
                        ans_7(0)(i) + ans_8(0)(i) + ans_9(0)(i) + ans_10(0)(i)
            '当年付息
            DNFX(i) = ans_1(1)(i) + ans_2(1)(i) + ans_3(1)(i) + ans_4(1)(i) + ans_5(1)(i) + ans_6(1)(i) +
                      ans_7(1)(i) + ans_8(1)(i) + ans_9(1)(i) + ans_10(1)(i)
            '当年还本
            DNHB(i) = ans_1(2)(i) + ans_2(2)(i) + ans_3(2)(i) + ans_4(2)(i) + ans_5(2)(i) + ans_6(2)(i) +
                      ans_7(2)(i) + ans_8(2)(i) + ans_9(2)(i) + ans_10(2)(i)
            '当年本金余额
            BJYE(i) = ans_1(3)(i) + ans_2(3)(i) + ans_3(3)(i) + ans_4(3)(i) + ans_5(3)(i) + ans_6(3)(i) +
                       ans_7(3)(i) + ans_8(3)(i) + ans_9(3)(i) + ans_10(3)(i)
        Next
        '返回结果
        Dim ans(3)
        ans(0) = DNBXZE
        ans(1) = DNFX
        ans(2) = DNHB
        ans(3) = BJYE
        Return ans
    End Function
    Function 长期贷款计算_等额本金_10次投资分开计算(bj_0_list As Array, dknx_0_list As Array, ksnf_list As Array, jsnx As Integer,
                                                    dkll_list As Array, kxnx_list As Array, tcyfs_list As Array, kxqfx As Boolean)
        'bj_0_list: 贷款本金初始值，列表，长度10
        'dnkx_0_list：长期贷款年限初始值，列表，长度10
        'ksnf_list：贷款计算的开始年份，列表，长度10
        'jsnx：整个项目的最大计算年限(含建设期)
        'dkll_list：长期贷款年化利率(复利)，列表，长度10
        'kxnx_list：10次投资，长期贷款宽限年限，列表，长度10
        'tcyfs_list：逐年投产的月份数，列表，长度31
        'kxqfx：宽限期内是否付息

        '10次投资分别计算长期贷款还本付息，再累加在一起

        '第1次投资
        Dim bj_0_1 As Double = bj_0_list(1)
        Dim dknx_0_1 As Integer = dknx_0_list(1)
        Dim ksnf_1 As Integer = ksnf_list(1)
        Dim dkll_1 As Double = dkll_list(1)
        Dim kxnx_1 As Integer = kxnx_list(1)
        '第2次投资
        Dim bj_0_2 As Double = bj_0_list(2)
        Dim dknx_0_2 As Integer = dknx_0_list(2)
        Dim ksnf_2 As Integer = ksnf_list(2)
        Dim dkll_2 As Double = dkll_list(2)
        Dim kxnx_2 As Integer = kxnx_list(2)
        '第3次投资
        Dim bj_0_3 As Double = bj_0_list(3)
        Dim dknx_0_3 As Integer = dknx_0_list(3)
        Dim ksnf_3 As Integer = ksnf_list(3)
        Dim dkll_3 As Double = dkll_list(3)
        Dim kxnx_3 As Integer = kxnx_list(3)
        '第4次投资
        Dim bj_0_4 As Double = bj_0_list(4)
        Dim dknx_0_4 As Integer = dknx_0_list(4)
        Dim ksnf_4 As Integer = ksnf_list(4)
        Dim dkll_4 As Double = dkll_list(4)
        Dim kxnx_4 As Integer = kxnx_list(4)
        '第5次投资
        Dim bj_0_5 As Double = bj_0_list(5)
        Dim dknx_0_5 As Integer = dknx_0_list(5)
        Dim ksnf_5 As Integer = ksnf_list(5)
        Dim dkll_5 As Double = dkll_list(5)
        Dim kxnx_5 As Integer = kxnx_list(5)
        '第6次投资
        Dim bj_0_6 As Double = bj_0_list(6)
        Dim dknx_0_6 As Integer = dknx_0_list(6)
        Dim ksnf_6 As Integer = ksnf_list(6)
        Dim dkll_6 As Double = dkll_list(6)
        Dim kxnx_6 As Integer = kxnx_list(6)
        '第7次投资
        Dim bj_0_7 As Double = bj_0_list(7)
        Dim dknx_0_7 As Integer = dknx_0_list(7)
        Dim ksnf_7 As Integer = ksnf_list(7)
        Dim dkll_7 As Double = dkll_list(7)
        Dim kxnx_7 As Integer = kxnx_list(7)
        '第8次投资
        Dim bj_0_8 As Double = bj_0_list(8)
        Dim dknx_0_8 As Integer = dknx_0_list(8)
        Dim ksnf_8 As Integer = ksnf_list(8)
        Dim dkll_8 As Double = dkll_list(8)
        Dim kxnx_8 As Integer = kxnx_list(8)
        '第9次投资
        Dim bj_0_9 As Double = bj_0_list(9)
        Dim dknx_0_9 As Integer = dknx_0_list(9)
        Dim ksnf_9 As Integer = ksnf_list(9)
        Dim dkll_9 As Double = dkll_list(9)
        Dim kxnx_9 As Integer = kxnx_list(9)
        '第10次投资
        Dim bj_0_10 As Double = bj_0_list(10)
        Dim dknx_0_10 As Integer = dknx_0_list(10)
        Dim ksnf_10 As Integer = ksnf_list(10)
        Dim dkll_10 As Double = dkll_list(10)
        Dim kxnx_10 As Integer = kxnx_list(10)

        '列表，储存计算结果
        Dim DNBXZE(31) As Double '当年还本付息总额
        Dim DNFX(31) As Double '当年付息额
        Dim DNHB(31) As Double '当年还本额
        Dim BJYE(31) As Double '当年本金余额

        '分次进行计算
        Dim ans_1 = 长期贷款计算_等额本金(bj_0_1, dknx_0_1, ksnf_1, jsnx, dkll_1, kxnx_1, tcyfs_list, kxqfx)
        Dim ans_2 = 长期贷款计算_等额本金(bj_0_2, dknx_0_2, ksnf_2, jsnx, dkll_2, kxnx_2, tcyfs_list, kxqfx)
        Dim ans_3 = 长期贷款计算_等额本金(bj_0_3, dknx_0_3, ksnf_3, jsnx, dkll_3, kxnx_3, tcyfs_list, kxqfx)
        Dim ans_4 = 长期贷款计算_等额本金(bj_0_4, dknx_0_4, ksnf_4, jsnx, dkll_4, kxnx_4, tcyfs_list, kxqfx)
        Dim ans_5 = 长期贷款计算_等额本金(bj_0_5, dknx_0_5, ksnf_5, jsnx, dkll_5, kxnx_5, tcyfs_list, kxqfx)
        Dim ans_6 = 长期贷款计算_等额本金(bj_0_6, dknx_0_6, ksnf_6, jsnx, dkll_6, kxnx_6, tcyfs_list, kxqfx)
        Dim ans_7 = 长期贷款计算_等额本金(bj_0_7, dknx_0_7, ksnf_7, jsnx, dkll_7, kxnx_7, tcyfs_list, kxqfx)
        Dim ans_8 = 长期贷款计算_等额本金(bj_0_8, dknx_0_8, ksnf_8, jsnx, dkll_8, kxnx_8, tcyfs_list, kxqfx)
        Dim ans_9 = 长期贷款计算_等额本金(bj_0_9, dknx_0_9, ksnf_9, jsnx, dkll_9, kxnx_9, tcyfs_list, kxqfx)
        Dim ans_10 = 长期贷款计算_等额本金(bj_0_10, dknx_0_10, ksnf_10, jsnx, dkll_10, kxnx_10, tcyfs_list, kxqfx)
        '累加
        For i = 1 To 31 '第1年到第31年
            '当年还本付息总额
            DNBXZE(i) = ans_1(0)(i) + ans_2(0)(i) + ans_3(0)(i) + ans_4(0)(i) + ans_5(0)(i) + ans_6(0)(i) +
                        ans_7(0)(i) + ans_8(0)(i) + ans_9(0)(i) + ans_10(0)(i)
            '当年付息
            DNFX(i) = ans_1(1)(i) + ans_2(1)(i) + ans_3(1)(i) + ans_4(1)(i) + ans_5(1)(i) + ans_6(1)(i) +
                      ans_7(1)(i) + ans_8(1)(i) + ans_9(1)(i) + ans_10(1)(i)
            '当年还本
            DNHB(i) = ans_1(2)(i) + ans_2(2)(i) + ans_3(2)(i) + ans_4(2)(i) + ans_5(2)(i) + ans_6(2)(i) +
                      ans_7(2)(i) + ans_8(2)(i) + ans_9(2)(i) + ans_10(2)(i)
            '当年本金余额
            BJYE(i) = ans_1(3)(i) + ans_2(3)(i) + ans_3(3)(i) + ans_4(3)(i) + ans_5(3)(i) + ans_6(3)(i) +
                       ans_7(3)(i) + ans_8(3)(i) + ans_9(3)(i) + ans_10(3)(i)
        Next
        '返回结果
        Dim ans(3)
        ans(0) = DNBXZE
        ans(1) = DNFX
        ans(2) = DNHB
        ans(3) = BJYE
        Return ans
    End Function
    Function 长期贷款计算_等额本息(bj_0 As Double, dknx_0 As Integer, ksnf As Integer, jsnx As Integer, dkll As Double,
                                   kxnx As Integer, tcyfs_list As Array, kxqfx As Boolean)
        'bj_0: 贷款本金初始值
        'dnkx_0：长期贷款年限初始值
        'ksnf：贷款计算的开始年份
        'jsnx：整个项目的最大计算年限(含建设期)
        'dkll：长期贷款年化利率(复利)
        'kxnx：长期贷款宽限期年份数
        'tcyfs_list：逐年投产的月份数，列表，长度31
        'kxqfx：宽限期内是否付息

        '某一笔贷款，从贷款发生的第二年才开始计算还本金，还本金不按照投产月份数折算；如果贷款当年存在投产月份数，则计算付息，付息额按照投产月份数折算；
        '某一笔贷款，宽限期从贷款发生当年及以后的有投产月份数的第一个年份开始算；

        '列表，储存计算过程量
        Dim BJYE(31) As Double '当年本金余额
        Dim DNFX(31) As Double '当年付息额
        Dim DNHB(31) As Double '当年还本额
        Dim DNBXZE(31) As Double '当年还本付息总额

        If bj_0 <> 0 And dknx_0 <> 0 And ksnf <> 0 And jsnx <> 0 And dkll <> 0 Then
            '对贷款年限进行修正
            Dim dknx As Integer
            If ksnf + dknx_0 <= jsnx Then
                dknx = dknx_0
            Else
                dknx = jsnx - ksnf
            End If

            '每期还本付息总额：整个还款期内是一个定值
            Dim BXZE = (bj_0 * dkll * (1 + dkll) ^ dknx) / ((1 + dkll) ^ dknx - 1)
            Dim YJHB As Double = 0 '已经还掉的本金，初始值是0
            For i = ksnf To 31 'i表示年份序号
                If YJHB < bj_0 Then
                    '还剩余的本金
                    Dim SYBJ As Double = bj_0 - YJHB
                    '不考虑任何特殊情况的当年足额付息额
                    Dim ZEFX As Double = bj_0 * dkll * (1 + dkll) ^ (i - ksnf) - BXZE * ((1 + dkll) ^ (i - ksnf) - 1)

                    '每期付息，每年付息额按照(投产月份数/12)进行折算
                    '如果处于宽限期内，则判断是否付息
                    If i <= ksnf + kxnx And kxqfx = False Then
                        DNFX(i) = 0
                    Else
                        DNFX(i) = SYBJ * dkll * (tcyfs_list(i) / 12)
                    End If

                    '每期还本付息总额
                    '如果处于宽限期，则不还本，则还本付息总额=付息额
                    If i <= ksnf + kxnx Then
                        DNBXZE(i) = DNFX(i)
                    Else
                        DNBXZE(i) = BXZE - ZEFX + DNFX(i)
                    End If

                    '每期还本
                    '贷款发生当年不还本，宽限期内不还本
                    If i <= ksnf + kxnx Then
                        DNHB(i) = 0
                    Else
                        DNHB(i) = Math.Min((DNBXZE(i) - DNFX(i)), SYBJ)
                    End If

                    '累计已经还本金额度
                    YJHB += DNHB(i)
                    '每期末本金余额
                    BJYE(i) = bj_0 - YJHB
                Else
                    '每期还本付息总额
                    DNBXZE(i) = 0
                    '每期付息
                    DNFX(i) = 0
                    '每期还本
                    DNHB(i) = 0
                    '每期末本金余额
                    BJYE(i) = 0
                End If
            Next
        End If

        '返回结果
        Dim ans(3)
        ans(0) = DNBXZE
        ans(1) = DNFX
        ans(2) = DNHB
        ans(3) = BJYE
        Return ans
    End Function

    Function 长期贷款计算_等额本金(bj_0 As Double, dknx_0 As Integer, ksnf As Integer, jsnx As Integer, dkll As Double,
                                   kxnx As Integer, tcyfs_list As Array, kxqfx As Boolean)
        'bj_0: 贷款本金初始值
        'dnkx_0：长期贷款年限初始值
        'ksnf：贷款计算的开始年份
        'jsnx：整个项目的最大计算年限(含建设期)
        'dkll：长期贷款年化利率(复利)
        'kxnx：长期贷款宽限期年份数
        'tcyfs_list：逐年投产的月份数，列表，长度31
        'kxqfx：宽限期内是否付息

        '某一笔贷款，从贷款发生的第二年才开始计算还本金，还本金不按照投产月份数折算；如果贷款当年存在投产月份数，则计算付息，付息额按照投产月份数折算；
        '某一笔贷款，宽限期从贷款发生当年及以后的有投产月份数的第一个年份开始算；

        '列表，储存计算过程量
        Dim BJYE(31) As Double '当年本金余额
        Dim DNFX(31) As Double '当年付息额
        Dim DNHB(31) As Double '当年还本额
        Dim DNBXZE(31) As Double '当年还本付息总额

        If bj_0 <> 0 And dknx_0 <> 0 And ksnf <> 0 And jsnx <> 0 And dkll <> 0 Then
            '对贷款年限进行修正
            Dim dknx As Integer
            If ksnf + dknx_0 <= jsnx Then
                dknx = dknx_0
            Else
                dknx = jsnx - ksnf
            End If

            '每期还本金额：整个还款期内是一个定值
            Dim HBJE = bj_0 / dknx
            Dim YJHB As Double = 0 '已经还掉的本金，初始值是0
            For i = ksnf To 31
                If YJHB < bj_0 Then
                    '还剩余的本金
                    Dim SYBJ As Double = bj_0 - YJHB

                    '每期付息，每年付息额按照(投产月份数/12)进行折算
                    '如果处于宽限期内，则判断是否付息
                    If i <= ksnf + kxnx And kxqfx = False Then
                        DNFX(i) = 0
                    Else
                        DNFX(i) = SYBJ * dkll * (tcyfs_list(i) / 12)
                    End If

                    '每期还本,整个还款期内是一个定值
                    '每期还本
                    '贷款发生当年不还本，宽限期内不还本
                    If i <= ksnf + kxnx Then
                        DNHB(i) = 0
                    Else
                        DNHB(i) = Math.Min(HBJE, SYBJ)
                    End If
                    '累计已经还本金额度
                    YJHB += DNHB(i)
                    '每期末本金余额
                    BJYE(i) = bj_0 - YJHB

                    '每期还本付息总额
                    '如果处于宽限期，则不还本，则还本付息总额=付息额
                    If i <= ksnf + kxnx Then
                        DNBXZE(i) = DNFX(i)
                    Else
                        DNBXZE(i) = DNFX(i) + DNHB(i)
                    End If
                Else
                    '每期还本付息总额
                    DNBXZE(i) = 0
                    '每期付息
                    DNFX(i) = 0
                    '每期还本
                    DNHB(i) = 0
                    '每期末本金余额
                    BJYE(i) = 0
                End If
            Next
        End If

        '返回结果
        Dim ans(3)
        ans(0) = DNBXZE
        ans(1) = DNFX
        ans(2) = DNHB
        ans(3) = BJYE
        Return ans
    End Function
End Module
