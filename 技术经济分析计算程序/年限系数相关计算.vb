Module 年限系数相关计算
    Sub 年限系数相关计算(ExcelApp As Object)
        On Error Resume Next
        '项目计算年限、固定资产折旧年限、无形资产摊销年限、建设期可抵扣增值税计算年限、长期贷款宽限年限变化后相关计算
        '————————————————————————————————————————————————————————————————————————————————————————

        '判断输入的年限系数是否合理，并进行修正
        '项目计算年数改变后改变的计算用系数
        '读取输入的10次投资发生年份的最大值
        '如果贷款计算方法是方法一或者方法二，同时折旧和摊销的计算方法为方法一或者方法二，定义投资年份最大值等于1
        '定义局部变量
        '项目计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        '投资年份最大值
        Dim tznfzdz As Integer
        '固定资产折旧年限
        Dim gdzczjnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 7).Value
        '无形资产摊销年限
        Dim wxzctxnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 7).Value
        '长期贷款还款年限
        Dim cqdkhknx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value
        '折旧贷款年限中较大的值        
        Dim zjdknx_1 As Integer = Math.Max(gdzczjnx, wxzctxnx)
        Dim zjdknx As Integer = Math.Max(zjdknx_1, cqdkhknx)
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        If (ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法一" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法二") And (ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法一" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法二") Then
            tznfzdz = 1
            '添加报错
            If tznfzdz + zjdknx > jsnx Then
                '报错（只在这里报错一次，折旧摊销贷款第二到第十次投资年份模块里就不写了，防止重复报错）
                MsgBox("输入的<项目计算年限、第二次到第十次项目投资年份、长期贷款年限、固定资产折旧年限、无形资产摊销年限>之中的某一项或者某几项参数不合理，请重新核对！！！程序将自动重置不合理的数值！！！")
                '固定资产折旧年限
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 7).Value > jsnx - tznfzdz Then
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 7).Value = jsnx - tznfzdz
                End If
                '长期贷款年限
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value > jsnx - tznfzdz Then
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value = jsnx - tznfzdz
                End If
                '无形资产摊销年限
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 7).Value > jsnx - tznfzdz Then
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 7).Value = jsnx - tznfzdz
                End If
            End If
        Else
            Dim TZNF_LIST As New List(Of Integer)
            For i = 3 To 11 Step 2
                TZNF_LIST.Add(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, i).Value)
            Next
            For i = 3 To 11 Step 2
                TZNF_LIST.Add(ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, i).Value)
            Next
            tznfzdz = TZNF_LIST.Max '投资年份最大值
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        '输入的数字必需为从2到31的整数，同时要大于等于项目投资年份最大值与折旧和贷款年限的较大值的和
        If jsnx >= 2 And jsnx <= 31 And jsnx >= tznfzdz + zjdknx And Int(jsnx) = jsnx Then
            '标记处于计算期内的每一年
            '前15年
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(111, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(112, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(112, i).Value = 0
                End If
            Next
            '后16年
            For i = 2 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(114, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(115, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(115, i).Value = 0
                End If
            Next
            '计算逐年达产系数
            '前15年
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(102, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(5, i).Value
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(103, i).Value = 0
                End If
            Next
            '后16年
            For i = 2 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(105, i).Value <= jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i).Value = ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(8, i).Value
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(106, i).Value = 0
                End If
            Next
            '标记计算期结束的最后一年
            '前15年
            For i = 3 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(135, i).Value = jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(136, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(136, i).Value = 0
                End If
            Next
            '后16年
            For i = 2 To 17
                If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(138, i).Value = jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(139, i).Value = 1
                Else
                    ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(139, i).Value = 0
                End If
            Next
        ElseIf jsnx > 31 Then
            MsgBox("项目计算年限不可以大于31年，请重新输入！")
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value = 31
        ElseIf jsnx < 2 Then
            MsgBox("项目计算年限不可以小于2年，请重新输入！")
            ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value = 2
        Else
            '如果贷款的计算方法中为方法三或者方法四，或者折旧和摊销的计算方法为方法三或者方法四，进行下列计算
            If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法三" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法四" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法三" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法四" Then
                '报错（只在这里报错一次，折旧摊销贷款第二到第十次投资年份模块里就不写了，防止重复报错）
                MsgBox("输入的<项目计算年限、第二次到第十次项目投资年份、长期贷款年限、固定资产折旧年限、无形资产摊销年限>之中的某一项或者某几项参数不合理，请重新核对！！！程序将自动重置不合理的数值！！！")
                '固定资产折旧年限
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 7).Value > jsnx - tznfzdz Then
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 7).Value = jsnx - tznfzdz
                End If
                '长期贷款年限
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value > jsnx - tznfzdz Then
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value = jsnx - tznfzdz
                End If
                '无形资产摊销年限
                If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 7).Value > jsnx - tznfzdz Then
                    ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 7).Value = jsnx - tznfzdz
                End If
            End If
        End If
        '————————————————————————————————————————————————————————————————————————————————————————
        '计算一次工作簿
        ExcelApp.Calculate()
    End Sub
End Module
