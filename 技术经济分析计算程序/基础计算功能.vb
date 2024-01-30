Module 基础计算功能
    Function Array数据转换为Double(value_list As Array, n_value As Integer)
        'value_list：需要被转换的Array
        'n_value：Array的数据数量
        '———————————————————————————————————————————————————————————————————————————————————————— 
        Dim new_list(n_value) As Double
        For i = 0 To n_value
            new_list(i) = CDbl(value_list(i))
        Next
        Return new_list
    End Function
    Function Array数据转换为Integer(value_list As Array, n_value As Integer)
        'value_list：需要被转换的Array
        'n_value：Array的数据数量
        '———————————————————————————————————————————————————————————————————————————————————————— 
        Dim new_list(n_value) As Integer
        For i = 0 To n_value
            new_list(i) = CInt(value_list(i))
        Next
        Return new_list
    End Function
    Function 基础计算功能_10_to_31(tznf_list As Array, value_10_list As Array)
        '将<估算表>中的10次投资情况生成成31年的列表
        '将长度10的列表转为长度31的列表
        'tznf_list：10次投资的年份序号，列表，长度10
        'value_10_list：10次投资的各种金额数值(静态投资，建设期利息，可抵扣增值税，资本金，建设期贷款)，列表，长度10
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '储存计算结果，长度为31的列表
        Dim ans(31) As Double
        '10次投资发生的年份序号
        For i = 1 To 10
            If tznf_list(i) > 0 Then
                ans(tznf_list(i)) = value_10_list(i)
            End If
        Next
        '返回结果
        Return ans
    End Function
    Function 基础计算功能_31_to_10(tznf_list As Array, value_31_list As Array)
        '将长度31的列表转为长度10的列表
        'tznf_list：10次投资的年份序号，列表，长度10
        'value_31_list：31年各种金额数值(静态投资，建设期利息，可抵扣增值税，资本金，建设期贷款)，列表，长度31
        '———————————————————————————————————————————————————————————————————————————————————————— 
        '储存计算结果，长度为10的列表
        Dim ans(10) As Double
        '遍历10次投资的年份序号
        For i = 1 To 10
            If tznf_list(i) > 0 Then
                ans(i) = value_31_list(tznf_list(i))
            End If
        Next
        '返回结果
        Return ans
    End Function
    Function 逐年费率计算_base(ExcelApp As Object, ksnf_list As List(Of Integer), jsnf_list As List(Of Integer),
                               ksfl_list As List(Of Double), jsfl_list As List(Of Double))
        '设置设备逐年费率（修理费、保险费、材料费其它费），费率可以逐年变化
        'ksnf_list：开始年份列表
        'jsnf_list：结束年份列表
        'ksfl_list：开始年份费率
        'jsfl_list：结束年份费率
        '———————————————————————————————————————————————————————————————————————————————————————— 
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        Dim ZDJSNF = jsnf_list.Max
        '输入的年份总数量
        Dim n_nf As Integer = ksnf_list.LongCount
        '添加报错功能
        For i = 0 To n_nf - 1
            If ksnf_list(i) <> 0 Or jsnf_list(i) <> 0 Then
                If ksnf_list(i) < 1 Or jsnf_list(i) < 1 Then
                    MsgBox("开始年份或者结束年份不可以存在小于1的情况，请重新输入！")
                    Exit Function
                End If
            End If
        Next
        For i = 0 To n_nf - 1
            If ksnf_list(i) > jsnx Or jsnf_list(i) > jsnx Then
                MsgBox("开始年份或者结束年份存在大于计算年限的情况，程序会继续计算，但会自动忽略大于计算年限的年份的值！")
                Exit For
            End If
        Next
        If n_nf > 1 Then
            For i = 1 To n_nf - 1
                If jsnf_list(i - 1) > ksnf_list(i) Then
                    MsgBox("输入的后一个开始年份不可以小于上一个结束年份，请重新输入！")
                    Exit Function
                End If
            Next
        End If
        If ZDJSNF < jsnx Then
            MsgBox("输入的结束年份均小于项目计算年限，请重新输入！")
            Exit Function
        End If
        If ZDJSNF > jsnx Then
            MsgBox("输入的结束年份存在大于项目计算年限的情况，请重新输入！")
            Exit Function
        End If
        '————————————————————————————————————————————————————————————————————————————————————————   
        '计算在输入的开始年份和结束年份之外的年份序号
        Dim qtnf_tmp_list As New List(Of Integer)
        For i = 1 To jsnx
            qtnf_tmp_list.Add(i)
        Next
        Dim tmp_list As New List(Of Integer)
        For i = 0 To n_nf - 1
            For j = ksnf_list(i) To jsnf_list(i)
                tmp_list.Add(j)
            Next
        Next
        Dim qtnf_list = qtnf_tmp_list.Except(tmp_list).ToList()
        '————————————————————————————————————————————————————————————————————————————————————————
        '逐年变化率
        Dim znbhl_list As New List(Of Double)
        For i = 0 To n_nf - 1
            Dim znbhl As Double
            If jsfl_list(i) - ksfl_list(i) = 0 Then
                znbhl = 0
            Else
                znbhl = (jsfl_list(i) - ksfl_list(i)) / (jsnf_list(i) - ksnf_list(i))
            End If
            znbhl_list.Add(znbhl)
        Next
        '————————————————————————————————————————————————————————————————————————————————————————
        '设备费率列表
        Dim ans_fl_list(31) As Double
        '需要计算费率的年份
        For j = 0 To n_nf - 1
            Dim js As Integer = 0
            For i = 1 To 31
                If i >= ksnf_list(j) And i <= jsnf_list(j) Then
                    js = js + 1
                    ans_fl_list(i) = (ksfl_list(j) + (js - 1) * znbhl_list(j))
                End If
            Next
        Next
        '其他年份
        For Each qtnf In qtnf_list
            For i = 1 To 31
                If i = qtnf Or i > jsnx Then
                    ans_fl_list(i) = 0
                End If
            Next
        Next
        '————————————————————————————————————————————————————————————————————————————————————————     
        '返回计算结果
        Return ans_fl_list
    End Function
    Function 逐年衰减率计算_base(ExcelApp As Object, jsksnf As Integer, jsjsnf As Integer, ksnf_list As List(Of Integer), sjl_list As List(Of Double))
        '光伏、蓄电池逐年衰减率计算
        'jsksnf：计算衰减率的项目开始计算衰减率的年份序号
        'jsjsnf：计算衰减率的项目结束计算衰减率的年份序号
        'ksnf_list：开始年份列表
        'sjl_list：逐年衰减率
        '————————————————————————————————————————————————————————————————————————————————————————     
        '读取项目计算年限
        Dim jsnx = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value '项目计算年限
        '输入的年份总数量
        Dim n_nf As Integer = ksnf_list.LongCount
        '输入的计算结束年份不可以大于项目计算年限
        If jsjsnf > jsnx Then
            MsgBox("窗口中输入的计算结束年份不可以超过项目计算年限，请重新输入！")
            Exit Function
        End If
        '检查输入的衰减开始年份
        If jsnx > 31 Then
            MsgBox("输入的项目计算年限不可以大于31年，请重新输入！")
            Exit Function
        End If
        For i = 0 To n_nf - 1
            If ksnf_list(i) > jsnx Then
                MsgBox("输入的衰减年份不可以大于项目计算年限，请重新输入！")
                Exit Function
            End If
        Next
        If n_nf > 1 Then
            For i = 1 To n_nf - 1
                If ksnf_list(i - 1) > ksnf_list(i) Then
                    MsgBox("输入的后一个开始年份不可以小于上一个开始年份，请重新输入！")
                    Exit Function
                End If
            Next
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '根据逐年衰减率计算出的基础负荷率
        Dim fhl_base(31) As Double
        '计算开始年份之前
        For i = 1 To 31
            If i < jsksnf Then
                fhl_base(i) = 0
            End If
        Next
        '计算开始年份到第一个开始衰减的年份
        For i = 1 To 31
            If i >= jsksnf And i < ksnf_list(0) Then
                '小于第一个衰减开始年份的负荷率设置为1
                fhl_base(i) = 1
            End If
        Next
        '衰减年份的之间的年份
        Dim yjsj As Double = 0 '已经衰减的系数
        If n_nf > 1 Then
            For j = 1 To n_nf - 1
                For i = 1 To 31
                    If i >= ksnf_list(j - 1) And i < ksnf_list(j) Then
                        yjsj += sjl_list(j - 1)
                        fhl_base(i) = (100 - yjsj) / 100
                    End If
                Next
            Next
        End If
        '最后一个衰减年份到计算衰减结束的年份
        For i = 1 To 31
            If i >= ksnf_list(n_nf - 1) And i <= jsjsnf Then
                yjsj += sjl_list(n_nf - 1)
                fhl_base(i) = (100 - yjsj) / 100
            End If
        Next
        '计算衰减结束的年份到最后
        For i = 1 To 31
            If i > jsjsnf Then
                fhl_base(i) = 0
            End If
        Next
        '返回结果
        Return fhl_base
    End Function
End Module
