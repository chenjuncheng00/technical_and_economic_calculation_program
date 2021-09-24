Imports Microsoft.Office.Interop.Excel

Public Class 设置长期贷款计算方式
    Private Sub 设置长期贷款计算方式_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————
        '读取输入的默认值
        Dim cqdknx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value '长期贷款还款年限
        Dim kxnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value '长期贷款宽限年限
        '清空已有的全部数据
        '长期贷款还款年限
        Me.cqdkhknx1.Clear()
        Me.cqdkhknx2.Clear()
        Me.cqdkhknx3.Clear()
        Me.cqdkhknx4.Clear()
        Me.cqdkhknx5.Clear()
        Me.cqdkhknx6.Clear()
        Me.cqdkhknx7.Clear()
        Me.cqdkhknx8.Clear()
        Me.cqdkhknx9.Clear()
        Me.cqdkhknx10.Clear()
        '长期贷款宽限年限
        Me.kxnx1.Clear()
        Me.kxnx2.Clear()
        Me.kxnx3.Clear()
        Me.kxnx4.Clear()
        Me.kxnx5.Clear()
        Me.kxnx6.Clear()
        Me.kxnx7.Clear()
        Me.kxnx8.Clear()
        Me.kxnx9.Clear()
        Me.kxnx10.Clear()
        '——————————————————————————————————————————————————————————————————————————————
        '载入默认值
        '根据已经输入的投资情况，载入默认值
        '第1次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 3).Value > 0 Then
            '长期贷款还款年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(23, 39).Value > 0 Then
                Me.cqdkhknx1.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(23, 39).Value, String)
            Else
                Me.cqdkhknx1.Text = CType(cqdknx, String)
            End If
            '长期贷款宽限年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(34, 39).Value > 0 Then
                Me.kxnx1.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(34, 39).Value, String)
            Else
                Me.kxnx1.Text = CType(kxnx, String)
            End If
        End If
        '第2次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value > 0 Then
            '长期贷款还款年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(24, 39).Value > 0 Then
                Me.cqdkhknx2.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(24, 39).Value, String)
            Else
                Me.cqdkhknx2.Text = CType(cqdknx, String)
            End If
            '长期贷款宽限年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(35, 39).Value > 0 Then
                Me.kxnx2.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(35, 39).Value, String)
            Else
                Me.kxnx2.Text = CType(kxnx, String)
            End If
        End If
        '第3次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value > 0 Then
            '长期贷款还款年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(25, 39).Value > 0 Then
                Me.cqdkhknx3.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(25, 39).Value, String)
            Else
                Me.cqdkhknx3.Text = CType(cqdknx, String)
            End If
            '长期贷款宽限年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(36, 39).Value > 0 Then
                Me.kxnx3.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(36, 39).Value, String)
            Else
                Me.kxnx3.Text = CType(kxnx, String)
            End If
        End If
        '第4次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value > 0 Then
            '长期贷款还款年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(26, 39).Value > 0 Then
                Me.cqdkhknx4.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(26, 39).Value, String)
            Else
                Me.cqdkhknx4.Text = CType(cqdknx, String)
            End If
            '长期贷款宽限年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(37, 39).Value > 0 Then
                Me.kxnx4.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(37, 39).Value, String)
            Else
                Me.kxnx4.Text = CType(kxnx, String)
            End If
        End If
        '第5次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value > 0 Then
            '长期贷款还款年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(27, 39).Value > 0 Then
                Me.cqdkhknx5.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(27, 39).Value, String)
            Else
                Me.cqdkhknx5.Text = CType(cqdknx, String)
            End If
            '长期贷款宽限年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(38, 39).Value > 0 Then
                Me.kxnx5.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(38, 39).Value, String)
            Else
                Me.kxnx5.Text = CType(kxnx, String)
            End If
        End If
        '第6次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value > 0 Then
            '长期贷款还款年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(28, 39).Value > 0 Then
                Me.cqdkhknx6.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(28, 39).Value, String)
            Else
                Me.cqdkhknx6.Text = CType(cqdknx, String)
            End If
            '长期贷款宽限年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(39, 39).Value > 0 Then
                Me.kxnx6.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(39, 39).Value, String)
            Else
                Me.kxnx6.Text = CType(kxnx, String)
            End If
        End If
        '第7次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value > 0 Then
            '长期贷款还款年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(29, 39).Value > 0 Then
                Me.cqdkhknx7.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(29, 39).Value, String)
            Else
                Me.cqdkhknx7.Text = CType(cqdknx, String)
            End If
            '长期贷款宽限年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(40, 39).Value > 0 Then
                Me.kxnx7.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(40, 39).Value, String)
            Else
                Me.kxnx7.Text = CType(kxnx, String)
            End If
        End If
        '第8次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value > 0 Then
            '长期贷款还款年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(30, 39).Value > 0 Then
                Me.cqdkhknx8.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(30, 39).Value, String)
            Else
                Me.cqdkhknx8.Text = CType(cqdknx, String)
            End If
            '长期贷款宽限年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(41, 39).Value > 0 Then
                Me.kxnx8.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(41, 39).Value, String)
            Else
                Me.kxnx8.Text = CType(kxnx, String)
            End If
        End If
        '第9次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value > 0 Then
            '长期贷款还款年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(31, 39).Value > 0 Then
                Me.cqdkhknx9.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(31, 39).Value, String)
            Else
                Me.cqdkhknx9.Text = CType(cqdknx, String)
            End If
            '长期贷款宽限年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(42, 39).Value > 0 Then
                Me.kxnx9.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(42, 39).Value, String)
            Else
                Me.kxnx9.Text = CType(kxnx, String)
            End If
        End If
        '第10次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value > 0 Then
            '长期贷款还款年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(32, 39).Value > 0 Then
                Me.cqdkhknx10.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(32, 39).Value, String)
            Else
                Me.cqdkhknx10.Text = CType(cqdknx, String)
            End If
            '长期贷款宽限年限
            If ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(43, 39).Value > 0 Then
                Me.kxnx10.Text = CType(ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(43, 39).Value, String)
            Else
                Me.kxnx10.Text = CType(kxnx, String)
            End If
        End If
    End Sub

    Private Sub 清空窗体数据_Click(sender As Object, e As EventArgs) Handles 清空窗体数据.Click
        Dim XZ = MsgBox("确定要清空本窗体输入的的全部内容？", vbOKCancel)
        If XZ = vbOK Then
            '长期贷款还款年限
            Me.cqdkhknx1.Clear()
            Me.cqdkhknx2.Clear()
            Me.cqdkhknx3.Clear()
            Me.cqdkhknx4.Clear()
            Me.cqdkhknx5.Clear()
            Me.cqdkhknx6.Clear()
            Me.cqdkhknx7.Clear()
            Me.cqdkhknx8.Clear()
            Me.cqdkhknx9.Clear()
            Me.cqdkhknx10.Clear()
            '长期贷款宽限年限
            Me.kxnx1.Clear()
            Me.kxnx2.Clear()
            Me.kxnx3.Clear()
            Me.kxnx4.Clear()
            Me.kxnx5.Clear()
            Me.kxnx6.Clear()
            Me.kxnx7.Clear()
            Me.kxnx8.Clear()
            Me.kxnx9.Clear()
            Me.kxnx10.Clear()
        End If
    End Sub

    Private Sub 确定计算方式_Click(sender As Object, e As EventArgs) Handles 确定计算方式.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '长期贷款计算方法需要采用方法三或者方法四
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法一" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 11).Value = "方法二" Then
            MsgBox("长期借款计算方式，请选择计算方法三或者方法四，计算终止！")
            Exit Sub
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '读取输入的数据
        '长期贷款还款年限
        Dim cqdkhknx1 = CType(Me.cqdkhknx1.Text, Integer)
        Dim cqdkhknx2 = CType(Me.cqdkhknx2.Text, Integer)
        Dim cqdkhknx3 = CType(Me.cqdkhknx3.Text, Integer)
        Dim cqdkhknx4 = CType(Me.cqdkhknx4.Text, Integer)
        Dim cqdkhknx5 = CType(Me.cqdkhknx5.Text, Integer)
        Dim cqdkhknx6 = CType(Me.cqdkhknx6.Text, Integer)
        Dim cqdkhknx7 = CType(Me.cqdkhknx7.Text, Integer)
        Dim cqdkhknx8 = CType(Me.cqdkhknx8.Text, Integer)
        Dim cqdkhknx9 = CType(Me.cqdkhknx9.Text, Integer)
        Dim cqdkhknx10 = CType(Me.cqdkhknx10.Text, Integer)
        '长期贷款宽限年限
        Dim kxnx1 = CType(Me.kxnx1.Text, Integer)
        Dim kxnx2 = CType(Me.kxnx2.Text, Integer)
        Dim kxnx3 = CType(Me.kxnx3.Text, Integer)
        Dim kxnx4 = CType(Me.kxnx4.Text, Integer)
        Dim kxnx5 = CType(Me.kxnx5.Text, Integer)
        Dim kxnx6 = CType(Me.kxnx6.Text, Integer)
        Dim kxnx7 = CType(Me.kxnx7.Text, Integer)
        Dim kxnx8 = CType(Me.kxnx8.Text, Integer)
        Dim kxnx9 = CType(Me.kxnx9.Text, Integer)
        Dim kxnx10 = CType(Me.kxnx10.Text, Integer)
        '——————————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否确认输入的逐次投资的长期贷款还款和宽限年限？？", vbOKCancel)
        If XZ = vbOK Then
            '读取项目总的计算年限
            Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
            '将已有的数据清零
            For i = 23 To 32
                '长期贷款还款年限
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 39).Value = 0
            Next
            For i = 34 To 43
                '长期贷款宽限年限
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 39).Value = 0
            Next
            '确定每次的数据
            '第1次投资
            Dim tznf1 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 3).Value '投资年份1
            If tznf1 > 0 Then
                If tznf1 + cqdkhknx1 > jsnx Then
                    MsgBox("输入的第1次投资的长期贷款还款年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf cqdkhknx1 <= 0 Then
                    MsgBox("输入的第1次投资的长期贷款还款年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                ElseIf kxnx1 <= 0 Then
                    MsgBox("输入的第1次投资的长期贷款宽限年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                Else
                    '写入数据
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(23, 39).Value = cqdkhknx1
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(34, 39).Value = kxnx1
                End If
            End If
            '第2次投资
            Dim tznf2 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value '投资年份2
            If tznf2 > 0 Then
                If tznf2 + cqdkhknx2 > jsnx Then
                    MsgBox("输入的第2次投资的长期贷款还款年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf cqdkhknx2 <= 0 Then
                    MsgBox("输入的第2次投资的长期贷款还款年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                ElseIf kxnx2 <= 0 Then
                    MsgBox("输入的第2次投资的长期贷款宽限年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                Else
                    '写入数据
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(24, 39).Value = cqdkhknx2
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(35, 39).Value = kxnx2
                End If
            End If
            '第3次投资
            Dim tznf3 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value '投资年份3
            If tznf3 > 0 Then
                If tznf3 + cqdkhknx3 > jsnx Then
                    MsgBox("输入的第3次投资的长期贷款还款年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf cqdkhknx3 <= 0 Then
                    MsgBox("输入的第3次投资的长期贷款还款年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                ElseIf kxnx3 <= 0 Then
                    MsgBox("输入的第3次投资的长期贷款宽限年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                Else
                    '写入数据
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(25, 39).Value = cqdkhknx3
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(36, 39).Value = kxnx3
                End If
            End If
            '第4次投资
            Dim tznf4 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value '投资年份4
            If tznf4 > 0 Then
                If tznf4 + cqdkhknx4 > jsnx Then
                    MsgBox("输入的第4次投资的长期贷款还款年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf cqdkhknx4 <= 0 Then
                    MsgBox("输入的第4次投资的长期贷款还款年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                ElseIf kxnx4 <= 0 Then
                    MsgBox("输入的第4次投资的长期贷款宽限年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                Else
                    '写入数据
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(26, 39).Value = cqdkhknx4
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(37, 39).Value = kxnx4
                End If
            End If
            '第5次投资
            Dim tznf5 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value '投资年份5
            If tznf5 > 0 Then
                If tznf5 + cqdkhknx5 > jsnx Then
                    MsgBox("输入的第5次投资的长期贷款还款年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf cqdkhknx5 <= 0 Then
                    MsgBox("输入的第5次投资的长期贷款还款年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                ElseIf kxnx5 <= 0 Then
                    MsgBox("输入的第5次投资的长期贷款宽限年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                Else
                    '写入数据
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(27, 39).Value = cqdkhknx5
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(38, 39).Value = kxnx5
                End If
            End If
            '第6次投资
            Dim tznf6 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value '投资年份6
            If tznf6 > 0 Then
                If tznf6 + cqdkhknx6 > jsnx Then
                    MsgBox("输入的第6次投资的长期贷款还款年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf cqdkhknx6 <= 0 Then
                    MsgBox("输入的第6次投资的长期贷款还款年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                ElseIf kxnx6 <= 0 Then
                    MsgBox("输入的第6次投资的长期贷款宽限年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                Else
                    '写入数据
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(28, 39).Value = cqdkhknx6
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(39, 39).Value = kxnx6
                End If
            End If
            '第7次投资
            Dim tznf7 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value '投资年份7
            If tznf7 > 0 Then
                If tznf7 + cqdkhknx7 > jsnx Then
                    MsgBox("输入的第7次投资的长期贷款还款年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf cqdkhknx7 <= 0 Then
                    MsgBox("输入的第7次投资的长期贷款还款年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                ElseIf kxnx7 <= 0 Then
                    MsgBox("输入的第7次投资的长期贷款宽限年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                Else
                    '写入数据
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(29, 39).Value = cqdkhknx7
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(40, 39).Value = kxnx7
                End If
            End If
            '第8次投资
            Dim tznf8 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value '投资年份8
            If tznf8 > 0 Then
                If tznf8 + cqdkhknx8 > jsnx Then
                    MsgBox("输入的第8次投资的长期贷款还款年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf cqdkhknx8 <= 0 Then
                    MsgBox("输入的第8次投资的长期贷款还款年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                ElseIf kxnx8 <= 0 Then
                    MsgBox("输入的第8次投资的长期贷款宽限年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                Else
                    '写入数据
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(30, 39).Value = cqdkhknx8
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(41, 39).Value = kxnx8
                End If
            End If
            '第9次投资
            Dim tznf9 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value '投资年份9
            If tznf9 > 0 Then
                If tznf9 + cqdkhknx9 > jsnx Then
                    MsgBox("输入的第9次投资的长期贷款还款年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf cqdkhknx9 <= 0 Then
                    MsgBox("输入的第9次投资的长期贷款还款年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                ElseIf kxnx9 <= 0 Then
                    MsgBox("输入的第9次投资的长期贷款宽限年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                Else
                    '写入数据
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(31, 39).Value = cqdkhknx9
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(42, 39).Value = kxnx9
                End If
            End If
            '第10次投资
            Dim tznf10 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value '投资年份10
            If tznf10 > 0 Then
                If tznf10 + cqdkhknx10 > jsnx Then
                    MsgBox("输入的第10次投资的长期贷款还款年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf cqdkhknx10 <= 0 Then
                    MsgBox("输入的第10次投资的长期贷款还款年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                ElseIf kxnx10 <= 0 Then
                    MsgBox("输入的第10次投资的长期贷款宽限年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                Else
                    '写入数据
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(32, 39).Value = cqdkhknx10
                    ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(43, 39).Value = kxnx10
                End If
            End If
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算一次工作簿
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
            Dim sdsl_model As Integer = 1
            '计算长期贷款
            Call 长期贷款相关计算(ExcelApp, sdsl_model)
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(167, 7).Value = "不相同"
        End If
        Me.Close()
    End Sub

    Private Sub 重置回默认方式_Click(sender As Object, e As EventArgs) Handles 重置回默认方式.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp      
        '————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否将逐次投资的长期贷款还款年限和宽限年限重置回默认值？？", vbOKCancel)
        If XZ = vbOK Then
            '写入各种系数默认值
            For i = 23 To 32
                '长期贷款还款年限
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 39).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 7).Value
            Next
            For i = 34 To 43
                '长期贷款宽限年限
                ExcelApp.ThisWorkbook.Worksheets("借款还本付息计划表").Cells(i, 39).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(10, 7).Value
            Next
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(167, 7).Value = "相同"
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算一次工作簿
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
            Dim sdsl_model As Integer = 1
            '计算长期贷款
            Call 长期贷款相关计算(ExcelApp, sdsl_model)
            '————————————————————————————————————————————————————————————————————————————————————————————
            MsgBox("设置完成，每次投资的固定资产折旧和无形资产摊销计算系数就均相同！")
        End If
        Me.Close()
    End Sub
End Class