Module 逐年达产率计算
    Sub 直接输入综合达产率(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        For i = 3 To 17
            '综合负荷率，前15年
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(102, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(5, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value = 0
            End If
        Next
        For i = 2 To 17
            '综合负荷率，后16年        
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(105, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(8, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i).Value = 0
            End If
        Next
        '检测建设期，将没有投产月份数量的年份负荷率设置为0
        For i = 3 To 17 '投资计划与资金筹措表列号
            '前15年
            If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value = 0 Then '如果当年投产的月份数=0                
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value = 0 '负荷率设置为0
            End If
        Next
        For i = 2 To 17 '投资计划与资金筹措表列号           
            '后16年
            If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 16).Value = 0 Then '如果当年投产的月份数=0                
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i).Value = 0 '负荷率设置为0
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
        '如果此时接入费是按照逐年达产率增加值计算的，则重新计算此时的接入费逐年系数
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(41, 18).Value = "逐年达产率增加值" Then
            '接入费计算模式（第2年到计算期最后一年，根据逐年达产率增加值计算）
            Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
            '第1年
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, 3).Value = 0
            '第1-15年
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= 2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i - 1).Value)
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
                End If
            Next
            '第16年
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, 18).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, 2).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, 17).Value)
            '第17-31年
            For i = 19 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= 2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15 - 1).Value)
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
                End If
            Next
            '将小于0的结果设置为0
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value < 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
                End If
            Next
        End If
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '计算一次Excel
        ExcelApp.Calculate()
    End Sub
    Sub 分投资逐次输入达产率(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '10次投资综合达产率计算
        Call 十次投资综合达产率计算(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————        
        '前15年
        For i = 3 To 17
            '综合负荷率
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(102, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(5, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value = 0
            End If
        Next
        '后16年
        For i = 2 To 17
            '综合负荷率          
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(105, i).Value <= ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(8, i).Value
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i).Value = 0
            End If
        Next
        '检测建设期，将没有投产月份数量的年份负荷率设置为0
        For i = 3 To 17 '投资计划与资金筹措表列号
            '前15年
            If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i).Value = 0 Then '如果当年投产的月份数=0                
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value = 0 '负荷率设置为0
            End If
        Next
        For i = 2 To 17 '投资计划与资金筹措表列号          
            '后16年
            If ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 16).Value = 0 Then '如果当年投产的月份数=0                
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i).Value = 0 '负荷率设置为0
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
        '如果此时接入费是按照逐年达产率增加值计算的，则重新计算此时的接入费逐年系数
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(41, 18).Value = "逐年达产率增加值" Then
            '接入费计算模式（第2年到计算期最后一年，根据逐年达产率增加值计算）
            Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
            '第1年
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, 3).Value = 0
            '第1-15年
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= 2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i - 1).Value)
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
                End If
            Next
            '第16年
            ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, 18).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, 2).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, 17).Value)
            '第17-31年
            For i = 19 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value >= 2 And ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(168, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 1 * (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15).Value - ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i - 15 - 1).Value)
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
                End If
            Next
            '将小于0的结果设置为0
            For i = 3 To 33
                If ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value < 0 Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(169, i).Value = 0
                End If
            Next
        End If
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '计算一次Excel
        ExcelApp.Calculate()
    End Sub
    Sub 光伏逐年综合达产率计算(ExcelApp As Object, fhl_base As Array, tcyfzs As Boolean)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————       
        '项目计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        'tcyfzs：逐年负荷率是否按照逐年投产月份数折算
        '读取输入数据
        Dim GSBSJ = 读取估算表数据(ExcelApp)
        '投资年份
        Dim tznf_list = GSBSJ(0)
        '10次投资的装机功率
        Dim gfzjgl = GSBSJ(17)
        '装机功率换算到31年的列表
        Dim gfzjgl_list = 基础计算功能_10_to_31(tznf_list, gfzjgl)
        '计算逐年综合负荷率
        Dim fhl_zh_gf = 逐年综合达产率计算基础功能(fhl_base, gfzjgl_list)
        '读取逐年投产月份数
        Dim tcyfs_list(31) As Integer
        For i = 1 To 31
            tcyfs_list(i) = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 2).Value
        Next
        '结果写入Excel
        '1-31年
        For i = 1 To 31
            If i <= jsnx Then
                If tcyfzs = True Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(100, 2 + i).Value = fhl_zh_gf(i) * tcyfs_list(i) / 12
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(100, 2 + i).Value = fhl_zh_gf(i)
                End If
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(100, 2 + i).Value = 0
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
    Sub 风电逐年综合达产率计算(ExcelApp As Object, fhl_base As Array, tcyfzs As Boolean)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '项目计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '读取输入数据
        Dim GSBSJ = 读取估算表数据(ExcelApp)
        '投资年份
        Dim tznf_list = GSBSJ(0)
        '10次投资的装机功率
        Dim fdzjgl = GSBSJ(19)
        '装机功率换算到31年的列表
        Dim fdzjgl_list = 基础计算功能_10_to_31(tznf_list, fdzjgl)
        '计算逐年综合负荷率
        Dim fhl_zh_fd = 逐年综合达产率计算基础功能(fhl_base, fdzjgl_list)
        '读取逐年投产月份数
        Dim tcyfs_list(31) As Integer
        For i = 1 To 31
            tcyfs_list(i) = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 2).Value
        Next
        '结果写入Excel
        '1-31年
        For i = 1 To 31
            If i <= jsnx Then
                If tcyfzs = True Then
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(178, 2 + i).Value = fhl_zh_fd(i) * tcyfs_list(i) / 12
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(178, 2 + i).Value = fhl_zh_fd(i)
                End If
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(178, 2 + i).Value = 0
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
    Sub 蓄电池逐年综合达产率计算(ExcelApp As Object, fhl_base As Array, jsms As String, tcyfzs As Boolean)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '项目计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '读取输入数据
        Dim GSBSJ = 读取估算表数据(ExcelApp)
        '投资年份
        Dim tznf_list = GSBSJ(0)
        '10次投资的装机功率
        Dim xdczjgl = GSBSJ(13)
        '装机功率换算到31年的列表
        Dim xdczjgl_list = 基础计算功能_10_to_31(tznf_list, xdczjgl)
        '计算逐年综合负荷率
        Dim fhl_zh_xdc = 逐年综合达产率计算基础功能(fhl_base, xdczjgl_list)
        '读取逐年投产月份数
        Dim tcyfs_list(31) As Integer
        For i = 1 To 31
            tcyfs_list(i) = ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(157, i + 2).Value
        Next
        '结果写入Excel
        '1-31年
        For i = 1 To 31
            If i <= jsnx Then
                If tcyfzs = True Then
                    If jsms = "供电" Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(96, 2 + i).Value = fhl_zh_xdc(i) * tcyfs_list(i) / 12
                    End If
                    If jsms = "购电" Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(97, 2 + i).Value = fhl_zh_xdc(i) * tcyfs_list(i) / 12
                    End If
                    If jsms = "供电和购电" Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(96, 2 + i).Value = fhl_zh_xdc(i) * tcyfs_list(i) / 12
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(97, 2 + i).Value = fhl_zh_xdc(i) * tcyfs_list(i) / 12
                    End If
                Else
                    If jsms = "供电" Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(96, 2 + i).Value = fhl_zh_xdc(i)
                    End If
                    If jsms = "购电" Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(97, 2 + i).Value = fhl_zh_xdc(i)
                    End If
                    If jsms = "供电和购电" Then
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(96, 2 + i).Value = fhl_zh_xdc(i)
                        ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(97, 2 + i).Value = fhl_zh_xdc(i)
                    End If
                End If
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(96, 2 + i).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(97, 2 + i).Value = 0
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
    Sub 十次投资综合达产率计算(ExcelApp As Object)
        On Error Resume Next
        '————————————————————————————————————————————————————————————————————————————————————————        
        '项目计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '读取1-10次的燃料费额
        Dim rlf(10) As Double
        For i = 1 To 10
            rlf(i) = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(119, 5 + i).Value
        Next
        '读取输入数据
        Dim GSBSJ = 读取估算表数据(ExcelApp)
        '投资年份
        Dim tznf_list = GSBSJ(0)
        '燃料费换算到31年的列表
        Dim rlf_list = 基础计算功能_10_to_31(tznf_list, rlf)
        '读取基础负荷率
        Dim fhl_base = 读取逐年负荷率(ExcelApp)
        '计算逐年综合负荷率
        Dim ans_fhl_zh = 逐年综合达产率计算基础功能(fhl_base, rlf_list)
        '结果写入Excel
        '前15年
        For i = 1 To 15
            If i <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(123, 2 + i).Value = ans_fhl_zh(i)
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(123, 2 + i).Value = 0
            End If
        Next
        '16—31年
        For i = 16 To 31
            If i <= jsnx Then
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(125, i - 14).Value = ans_fhl_zh(i)
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入税收表").Cells(125, i - 14).Value = 0
            End If
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
    Function 逐年综合达产率计算基础功能(fhl_base As Array, value_list As Array)
        'fhl_base：负荷率(达产率)计算基础值，从第2年开始的，列表，长度31
        'value_list：31年逐年计算用的数值，用估算表数据计算出的，数值对应的年份是建设年份，计算投产从后一年开始，列表，长度31

        '逐年value累计值
        Dim value_sum As Double = 0
        For i = 1 To 31
            value_sum += value_list(i)
        Next
        '逐年综合负荷率计算结果
        Dim ans_fhl(31)
        For i = 1 To 31
            '逐年fhl乘以value的值
            Dim fhl_value_list(31) As Double
            'i：value_list
            For j = 2 To 32 - i
                'j：fhl_base
                '从i+j-2年开始
                fhl_value_list(i + j - 2) = fhl_base(j) * value_list(i - 1)
            Next
            '累加
            For k = 1 To 31
                If value_sum > 0 Then
                    ans_fhl(k) += fhl_value_list(k) / value_sum
                Else
                    ans_fhl(k) = 0
                End If
            Next
        Next
        '返回结果
        Return ans_fhl
    End Function

    Function 默认逐年达产率()
        '风电逐年达产率
        Dim fhl_fd(31) As Double
        '光伏逐年达产率
        Dim fhl_gf(31) As Double
        '蓄电池逐年衰减率
        Dim fhl_xdc(31) As Double
        '风电
        For i = 2 To 21 '从第二年开始
            fhl_fd(i) = 1
        Next
        '光伏
        For i = 2 To 26 '从第二年开始
            fhl_gf(i) = 0.98 - 0.0055 * i
        Next
        '蓄电池
        For i = 2 To 11 '从第二年开始
            fhl_xdc(i) = 1.02 - 0.01 * i
        Next
        '返回结果
        Dim ans(2)
        ans(0) = fhl_fd
        ans(1) = fhl_gf
        ans(2) = fhl_xdc
        Return ans
    End Function
End Module
