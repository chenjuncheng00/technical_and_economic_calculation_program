Imports Microsoft.Office.Interop.Excel

Public Class 折旧摊销计算方式设置
    Private Sub 每次投资设置不同的折旧摊销计算方式_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————
        '读取输入的默认值
        Dim gdzczjnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 7).Value '固定资产折旧年限
        Dim gdzcczl As Double = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 5).Value * 100 '固定资产残值率
        Dim wxzctxnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 7).Value '无形资产摊销年限
        '清空已有的全部数据
        '固定资产折旧年限
        Me.gdzczjnx1.Clear()
        Me.gdzczjnx2.Clear()
        Me.gdzczjnx3.Clear()
        Me.gdzczjnx4.Clear()
        Me.gdzczjnx5.Clear()
        Me.gdzczjnx6.Clear()
        Me.gdzczjnx7.Clear()
        Me.gdzczjnx8.Clear()
        Me.gdzczjnx9.Clear()
        Me.gdzczjnx10.Clear()
        '固定资产残值率
        Me.gdzcczl1.Clear()
        Me.gdzcczl2.Clear()
        Me.gdzcczl3.Clear()
        Me.gdzcczl4.Clear()
        Me.gdzcczl5.Clear()
        Me.gdzcczl6.Clear()
        Me.gdzcczl7.Clear()
        Me.gdzcczl8.Clear()
        Me.gdzcczl9.Clear()
        Me.gdzcczl10.Clear()
        '无形资产摊销年限
        Me.wxzctxnx1.Clear()
        Me.wxzctxnx2.Clear()
        Me.wxzctxnx3.Clear()
        Me.wxzctxnx4.Clear()
        Me.wxzctxnx5.Clear()
        Me.wxzctxnx6.Clear()
        Me.wxzctxnx7.Clear()
        Me.wxzctxnx8.Clear()
        Me.wxzctxnx9.Clear()
        Me.wxzctxnx10.Clear()
        '——————————————————————————————————————————————————————————————————————————————
        '载入默认值
        '根据已经输入的投资情况，载入默认值
        '第1次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 3).Value > 0 Then
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(4, 22).Value > 0 Then
                Me.gdzczjnx1.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(4, 22).Value, String)
            Else
                Me.gdzczjnx1.Text = CType(gdzczjnx, String)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(4, 23).Value > 0 Then
                Me.gdzcczl1.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(4, 24).Value * 100, String)
            Else
                Me.gdzcczl1.Text = CType(gdzcczl, String)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(4, 24).Value > 0 Then
                Me.wxzctxnx1.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(4, 26).Value, String)
            Else
                Me.wxzctxnx1.Text = CType(wxzctxnx, String)
            End If
        End If
        '第2次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value > 0 Then
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(5, 22).Value > 0 Then
                Me.gdzczjnx2.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(5, 22).Value, String)
            Else
                Me.gdzczjnx2.Text = CType(gdzczjnx, String)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(5, 23).Value > 0 Then
                Me.gdzcczl2.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(5, 24).Value * 100, String)
            Else
                Me.gdzcczl2.Text = CType(gdzcczl, String)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(5, 24).Value > 0 Then
                Me.wxzctxnx2.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(5, 26).Value, String)
            Else
                Me.wxzctxnx2.Text = CType(wxzctxnx, String)
            End If
        End If
        '第3次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value > 0 Then
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(6, 22).Value > 0 Then
                Me.gdzczjnx3.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(6, 22).Value, String)
            Else
                Me.gdzczjnx3.Text = CType(gdzczjnx, String)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(6, 23).Value > 0 Then
                Me.gdzcczl3.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(6, 24).Value * 100, String)
            Else
                Me.gdzcczl3.Text = CType(gdzcczl, String)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(6, 24).Value > 0 Then
                Me.wxzctxnx3.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(6, 26).Value, String)
            Else
                Me.wxzctxnx3.Text = CType(wxzctxnx, String)
            End If
        End If
        '第4次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value > 0 Then
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(7, 22).Value > 0 Then
                Me.gdzczjnx4.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(7, 22).Value, String)
            Else
                Me.gdzczjnx4.Text = CType(gdzczjnx, String)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(7, 23).Value > 0 Then
                Me.gdzcczl4.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(7, 24).Value * 100, String)
            Else
                Me.gdzcczl4.Text = CType(gdzcczl, String)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(7, 24).Value > 0 Then
                Me.wxzctxnx4.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(7, 26).Value, String)
            Else
                Me.wxzctxnx4.Text = CType(wxzctxnx, String)
            End If
        End If
        '第5次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value > 0 Then
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(8, 22).Value > 0 Then
                Me.gdzczjnx5.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(8, 22).Value, String)
            Else
                Me.gdzczjnx5.Text = CType(gdzczjnx, String)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(8, 23).Value > 0 Then
                Me.gdzcczl5.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(8, 24).Value * 100, String)
            Else
                Me.gdzcczl5.Text = CType(gdzcczl, String)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(8, 24).Value > 0 Then
                Me.wxzctxnx5.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(8, 26).Value, String)
            Else
                Me.wxzctxnx5.Text = CType(wxzctxnx, String)
            End If
        End If
        '第6次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value > 0 Then
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(9, 22).Value > 0 Then
                Me.gdzczjnx6.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(9, 22).Value, String)
            Else
                Me.gdzczjnx6.Text = CType(gdzczjnx, String)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(9, 23).Value > 0 Then
                Me.gdzcczl6.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(9, 24).Value * 100, String)
            Else
                Me.gdzcczl6.Text = CType(gdzcczl, String)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(9, 24).Value > 0 Then
                Me.wxzctxnx6.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(9, 26).Value, String)
            Else
                Me.wxzctxnx6.Text = CType(wxzctxnx, String)
            End If
        End If
        '第7次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value > 0 Then
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(10, 22).Value > 0 Then
                Me.gdzczjnx7.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(10, 22).Value, String)
            Else
                Me.gdzczjnx7.Text = CType(gdzczjnx, String)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(10, 23).Value > 0 Then
                Me.gdzcczl7.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(10, 24).Value * 100, String)
            Else
                Me.gdzcczl7.Text = CType(gdzcczl, String)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(10, 24).Value > 0 Then
                Me.wxzctxnx7.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(10, 26).Value, String)
            Else
                Me.wxzctxnx7.Text = CType(wxzctxnx, String)
            End If
        End If
        '第8次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value > 0 Then
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(11, 22).Value > 0 Then
                Me.gdzczjnx8.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(11, 22).Value, String)
            Else
                Me.gdzczjnx8.Text = CType(gdzczjnx, String)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(11, 23).Value > 0 Then
                Me.gdzcczl8.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(11, 24).Value * 100, String)
            Else
                Me.gdzcczl8.Text = CType(gdzcczl, String)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(11, 24).Value > 0 Then
                Me.wxzctxnx8.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(11, 26).Value, String)
            Else
                Me.wxzctxnx8.Text = CType(wxzctxnx, String)
            End If
        End If
        '第9次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value > 0 Then
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(12, 22).Value > 0 Then
                Me.gdzczjnx9.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(12, 22).Value, String)
            Else
                Me.gdzczjnx9.Text = CType(gdzczjnx, String)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(12, 23).Value > 0 Then
                Me.gdzcczl9.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(12, 24).Value * 100, String)
            Else
                Me.gdzcczl9.Text = CType(gdzcczl, String)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(12, 24).Value > 0 Then
                Me.wxzctxnx9.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(12, 26).Value, String)
            Else
                Me.wxzctxnx9.Text = CType(wxzctxnx, String)
            End If
        End If
        '第10次投资
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value > 0 Then
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(13, 22).Value > 0 Then
                Me.gdzczjnx10.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(13, 22).Value, String)
            Else
                Me.gdzczjnx10.Text = CType(gdzczjnx, String)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(13, 23).Value > 0 Then
                Me.gdzcczl10.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(13, 24).Value * 100, String)
            Else
                Me.gdzcczl10.Text = CType(gdzcczl, String)
            End If
            If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(13, 24).Value > 0 Then
                Me.wxzctxnx10.Text = CType(ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(13, 26).Value, String)
            Else
                Me.wxzctxnx10.Text = CType(wxzctxnx, String)
            End If
        End If
    End Sub

    Private Sub 清空窗体数据_Click(sender As Object, e As EventArgs) Handles 清空窗体数据.Click
        Dim XZ = MsgBox("确定要清空本窗体输入的的全部内容？", vbOKCancel)
        If XZ = vbOK Then
            '清空已有的全部数据
            '固定资产折旧年限
            Me.gdzczjnx1.Clear()
            Me.gdzczjnx2.Clear()
            Me.gdzczjnx3.Clear()
            Me.gdzczjnx4.Clear()
            Me.gdzczjnx5.Clear()
            Me.gdzczjnx6.Clear()
            Me.gdzczjnx7.Clear()
            Me.gdzczjnx8.Clear()
            Me.gdzczjnx9.Clear()
            Me.gdzczjnx10.Clear()
            '固定资产残值率
            Me.gdzcczl1.Clear()
            Me.gdzcczl2.Clear()
            Me.gdzcczl3.Clear()
            Me.gdzcczl4.Clear()
            Me.gdzcczl5.Clear()
            Me.gdzcczl6.Clear()
            Me.gdzcczl7.Clear()
            Me.gdzcczl8.Clear()
            Me.gdzcczl9.Clear()
            Me.gdzcczl10.Clear()
            '无形资产摊销年限
            Me.wxzctxnx1.Clear()
            Me.wxzctxnx2.Clear()
            Me.wxzctxnx3.Clear()
            Me.wxzctxnx4.Clear()
            Me.wxzctxnx5.Clear()
            Me.wxzctxnx6.Clear()
            Me.wxzctxnx7.Clear()
            Me.wxzctxnx8.Clear()
            Me.wxzctxnx9.Clear()
            Me.wxzctxnx10.Clear()
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
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法一" Or ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 11).Value = "方法二" Then
            MsgBox("折旧摊销计算方式，请选择计算方法三或者方法四，计算终止！")
            Exit Sub
        End If
        '——————————————————————————————————————————————————————————————————————————————————————————————
        '读取输入的数据
        '固定资产折旧年限
        Dim gdzczjnx1 = CType(Me.gdzczjnx1.Text, Integer)
        Dim gdzczjnx2 = CType(Me.gdzczjnx2.Text, Integer)
        Dim gdzczjnx3 = CType(Me.gdzczjnx3.Text, Integer)
        Dim gdzczjnx4 = CType(Me.gdzczjnx4.Text, Integer)
        Dim gdzczjnx5 = CType(Me.gdzczjnx5.Text, Integer)
        Dim gdzczjnx6 = CType(Me.gdzczjnx6.Text, Integer)
        Dim gdzczjnx7 = CType(Me.gdzczjnx7.Text, Integer)
        Dim gdzczjnx8 = CType(Me.gdzczjnx8.Text, Integer)
        Dim gdzczjnx9 = CType(Me.gdzczjnx9.Text, Integer)
        Dim gdzczjnx10 = CType(Me.gdzczjnx10.Text, Integer)
        '固定资产残值率
        Dim gdzcczl1 = CType(Me.gdzcczl1.Text, Double) / 100
        Dim gdzcczl2 = CType(Me.gdzcczl2.Text, Double) / 100
        Dim gdzcczl3 = CType(Me.gdzcczl3.Text, Double) / 100
        Dim gdzcczl4 = CType(Me.gdzcczl4.Text, Double) / 100
        Dim gdzcczl5 = CType(Me.gdzcczl5.Text, Double) / 100
        Dim gdzcczl6 = CType(Me.gdzcczl6.Text, Double) / 100
        Dim gdzcczl7 = CType(Me.gdzcczl7.Text, Double) / 100
        Dim gdzcczl8 = CType(Me.gdzcczl8.Text, Double) / 100
        Dim gdzcczl9 = CType(Me.gdzcczl9.Text, Double) / 100
        Dim gdzcczl10 = CType(Me.gdzcczl10.Text, Double) / 100
        '无形资产摊销年限
        Dim wxzctxnx1 = CType(Me.wxzctxnx1.Text, Integer)
        Dim wxzctxnx2 = CType(Me.wxzctxnx2.Text, Integer)
        Dim wxzctxnx3 = CType(Me.wxzctxnx3.Text, Integer)
        Dim wxzctxnx4 = CType(Me.wxzctxnx4.Text, Integer)
        Dim wxzctxnx5 = CType(Me.wxzctxnx5.Text, Integer)
        Dim wxzctxnx6 = CType(Me.wxzctxnx6.Text, Integer)
        Dim wxzctxnx7 = CType(Me.wxzctxnx7.Text, Integer)
        Dim wxzctxnx8 = CType(Me.wxzctxnx8.Text, Integer)
        Dim wxzctxnx9 = CType(Me.wxzctxnx9.Text, Integer)
        Dim wxzctxnx10 = CType(Me.wxzctxnx10.Text, Integer)
        '————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————
        '读取项目总的计算年限
        Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
        Dim XZ = MsgBox("是否确认输入的逐次投资的固定资产折旧和无形资产摊销计算系数？", vbOKCancel)
        If XZ = vbOK Then
            '检测输入数据，添加报错功能
            '固定资产残值率不可以大于100%
            If gdzcczl1 > 1 Or gdzcczl2 > 1 Or gdzcczl3 > 1 Or gdzcczl4 > 1 Or gdzcczl5 > 1 Or gdzcczl6 > 1 Or gdzcczl7 > 1 Or gdzcczl8 > 1 Or gdzcczl9 > 1 Or gdzcczl10 > 1 Then
                MsgBox("输入的每次投资的固定资产残值率不可以大于100%，请重新输入！")
                Exit Sub
            End If
            '固定资产残值率不可以小于0%
            If gdzcczl1 < 0 Or gdzcczl2 < 0 Or gdzcczl3 < 0 Or gdzcczl4 < 0 Or gdzcczl5 < 0 Or gdzcczl6 < 0 Or gdzcczl7 < 0 Or gdzcczl8 < 0 Or gdzcczl9 < 0 Or gdzcczl10 < 0 Then
                MsgBox("输入的每次投资的固定资产残值率不可以小于0，请重新输入！")
                Exit Sub
            End If
            '将已有的数据清零
            For i = 4 To 13
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 22).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 24).Value = 0
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 26).Value = 0
            Next
            '确定每次的数据
            '第1次投资
            Dim tznf1 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 3).Value '投资年份1
            If tznf1 > 0 Then
                If tznf1 + gdzczjnx1 > jsnx Then
                    MsgBox("输入的第1次投资固定资产折旧年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf tznf1 + wxzctxnx1 > jsnx Then
                    MsgBox("输入的第1次投资无形资产摊销年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf gdzczjnx1 <= 0 Or wxzctxnx1 <= 0 Then
                    MsgBox("输入的第1次投资固定资产折旧年限或者无形资产摊销年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                Else
                    '写入数据
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(4, 22).Value = gdzczjnx1
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(4, 24).Value = gdzcczl1
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(4, 26).Value = wxzctxnx1
                End If
            End If
            '第2次投资
            Dim tznf2 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 5).Value '投资年份2
            If tznf2 > 0 Then
                If tznf2 + gdzczjnx2 > jsnx Then
                    MsgBox("输入的第2次投资固定资产折旧年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf tznf2 + wxzctxnx2 > jsnx Then
                    MsgBox("输入的第2次投资无形资产摊销年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf gdzczjnx2 <= 0 Or wxzctxnx2 <= 0 Then
                    MsgBox("输入的第2次投资固定资产折旧年限或者无形资产摊销年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                Else
                    '写入数据
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(5, 22).Value = gdzczjnx2
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(5, 24).Value = gdzcczl2
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(5, 26).Value = wxzctxnx2
                End If
            End If
            '第3次投资
            Dim tznf3 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 7).Value '投资年份3
            If tznf3 > 0 Then
                If tznf3 + gdzczjnx3 > jsnx Then
                    MsgBox("输入的第3次投资固定资产折旧年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf tznf3 + wxzctxnx3 > jsnx Then
                    MsgBox("输入的第3次投资无形资产摊销年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf gdzczjnx3 <= 0 Or wxzctxnx3 <= 0 Then
                    MsgBox("输入的第3次投资固定资产折旧年限或者无形资产摊销年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                Else
                    '写入数据
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(6, 22).Value = gdzczjnx3
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(6, 24).Value = gdzcczl3
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(6, 26).Value = wxzctxnx3
                End If
            End If
            '第4次投资
            Dim tznf4 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 9).Value '投资年份4
            If tznf4 > 0 Then
                If tznf4 + gdzczjnx4 > jsnx Then
                    MsgBox("输入的第4次投资固定资产折旧年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf tznf4 + wxzctxnx4 > jsnx Then
                    MsgBox("输入的第4次投资无形资产摊销年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf gdzczjnx4 <= 0 Or wxzctxnx4 <= 0 Then
                    MsgBox("输入的第4次投资固定资产折旧年限或者无形资产摊销年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                Else
                    '写入数据
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(7, 22).Value = gdzczjnx4
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(7, 24).Value = gdzcczl4
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(7, 26).Value = wxzctxnx4
                End If
            End If
            '第5次投资
            Dim tznf5 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(28, 11).Value '投资年份5
            If tznf5 > 0 Then
                If tznf5 + gdzczjnx5 > jsnx Then
                    MsgBox("输入的第5次投资固定资产折旧年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf tznf5 + wxzctxnx5 > jsnx Then
                    MsgBox("输入的第5次投资无形资产摊销年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf gdzczjnx5 <= 0 Or wxzctxnx5 <= 0 Then
                    MsgBox("输入的第5次投资固定资产折旧年限或者无形资产摊销年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                Else
                    '写入数据
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(8, 22).Value = gdzczjnx5
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(8, 24).Value = gdzcczl5
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(8, 26).Value = wxzctxnx5
                End If
            End If
            '第6次投资
            Dim tznf6 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 3).Value '投资年份6
            If tznf6 > 0 Then
                If tznf6 + gdzczjnx6 > jsnx Then
                    MsgBox("输入的第6次投资固定资产折旧年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf tznf6 + wxzctxnx6 > jsnx Then
                    MsgBox("输入的第6次投资无形资产摊销年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf gdzczjnx6 <= 0 Or wxzctxnx6 <= 0 Then
                    MsgBox("输入的第6次投资固定资产折旧年限或者无形资产摊销年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                Else
                    '写入数据
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(9, 22).Value = gdzczjnx6
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(9, 24).Value = gdzcczl6
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(9, 26).Value = wxzctxnx6
                End If
            End If
            '第7次投资
            Dim tznf7 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 5).Value '投资年份7
            If tznf7 > 0 Then
                If tznf7 + gdzczjnx7 > jsnx Then
                    MsgBox("输入的第7次投资固定资产折旧年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf tznf7 + wxzctxnx7 > jsnx Then
                    MsgBox("输入的第7次投资无形资产摊销年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf gdzczjnx7 <= 0 Or wxzctxnx7 <= 0 Then
                    MsgBox("输入的第7次投资固定资产折旧年限或者无形资产摊销年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                Else
                    '写入数据
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(10, 22).Value = gdzczjnx7
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(10, 24).Value = gdzcczl7
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(10, 26).Value = wxzctxnx7
                End If
            End If
            '第8次投资
            Dim tznf8 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 7).Value '投资年份8
            If tznf8 > 0 Then
                If tznf8 + gdzczjnx8 > jsnx Then
                    MsgBox("输入的第8次投资固定资产折旧年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf tznf8 + wxzctxnx8 > jsnx Then
                    MsgBox("输入的第8次投资无形资产摊销年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf gdzczjnx8 <= 0 Or wxzctxnx8 <= 0 Then
                    MsgBox("输入的第8次投资固定资产折旧年限或者无形资产摊销年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                Else
                    '写入数据
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(11, 22).Value = gdzczjnx8
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(11, 24).Value = gdzcczl8
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(11, 26).Value = wxzctxnx8
                End If
            End If
            '第9次投资
            Dim tznf9 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 9).Value '投资年份9
            If tznf9 > 0 Then
                If tznf9 + gdzczjnx9 > jsnx Then
                    MsgBox("输入的第9次投资固定资产折旧年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf tznf9 + wxzctxnx9 > jsnx Then
                    MsgBox("输入的第9次投资无形资产摊销年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf gdzczjnx9 <= 0 Or wxzctxnx9 <= 0 Then
                    MsgBox("输入的第9次投资固定资产折旧年限或者无形资产摊销年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                Else
                    '写入数据
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(12, 22).Value = gdzczjnx9
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(12, 24).Value = gdzcczl9
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(12, 26).Value = wxzctxnx9
                End If
            End If
            '第10次投资
            Dim tznf10 = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(71, 11).Value '投资年份10
            If tznf10 > 0 Then
                If tznf10 + gdzczjnx10 > jsnx Then
                    MsgBox("输入的第10次投资固定资产折旧年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf tznf10 + wxzctxnx10 > jsnx Then
                    MsgBox("输入的第10次投资无形资产摊销年限超出允许范围，请检查并重新输入！")
                    Exit Sub
                ElseIf gdzczjnx10 <= 0 Or wxzctxnx10 <= 0 Then
                    MsgBox("输入的第10次投资固定资产折旧年限或者无形资产摊销年限不可以小于等于0，请检查并重新输入！")
                    Exit Sub
                Else
                    '写入数据
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(13, 22).Value = gdzczjnx10
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(13, 24).Value = gdzcczl10
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(13, 26).Value = wxzctxnx10
                End If
            End If
            '写入计算模式
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(166, 7).Value = "不相同"
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算一次工作簿
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            'hscy：计算期末，是否回收资产残值
            Dim hscz As Boolean
            If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(53, 18).Value = "期末不回收残值" Then
                hscz = False
            Else
                hscz = True
            End If
            'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
            Dim sdsl_model As Integer = 1
            '计算折旧摊销
            Call 折旧摊销相关计算(ExcelApp, hscz, sdsl_model)
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
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
        Dim XZ = MsgBox("是否将逐次投资的固定资产折旧和无形资产摊销计算系数重置回默认值？？", vbOKCancel)
        If XZ = vbOK Then
            '写入各种系数默认值
            For i = 4 To 13
                '固定资产折旧年限
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 22).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 7).Value
                '固定资产残值率
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 24).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 5).Value
                '无形资产摊销年限
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 26).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 7).Value
            Next
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算一次工作簿
            ExcelApp.Calculate()
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            'hscy：计算期末，是否回收资产残值
            Dim hscz As Boolean = True
            'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
            Dim sdsl_model As Integer = 1
            '计算折旧摊销
            Call 折旧摊销相关计算(ExcelApp, hscz, sdsl_model)
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(166, 7).Value = "相同"
            MsgBox("设置完成，每次投资的固定资产折旧和无形资产摊销计算系数就均相同！")
        End If
        Me.Close()
    End Sub

    Private Sub 不回收固定资产残值_Click(sender As Object, e As EventArgs) Handles 不回收固定资产残值.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp      
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        Dim XZ = MsgBox("是否设置计算期末不回收固定资产残值？？请在确认完建设期时间计划，同时在<估算表>中设置完各种计算年限、输入完逐年静态投资金额后进行本设置！！", vbOKCancel)
        If XZ = vbOK Then
            'hscy：计算期末，是否回收资产残值
            Dim hscz As Boolean = False
            'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
            Dim sdsl_model As Integer = 1
            '计算折旧摊销
            Call 折旧摊销相关计算(ExcelApp, hscz, sdsl_model)
            '————————————————————————————————————————————————————————————————————————————————————————————————————————————
            '将设置状态写入表格
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(53, 18).Value = "期末不回收残值"
            MsgBox("设置完毕！")
        End If
        '————————————————————————————————————————————————————————————————————————————————————————————
        Me.Close()
    End Sub

    Private Sub 回收固定资产残值_Click(sender As Object, e As EventArgs) Handles 回收固定资产残值.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp      
        '————————————————————————————————————————————————————————————————————————————————————————
        '———————————————————————————————————————————————————————————————————————————————————————— 
        Dim XZ = MsgBox("是否设置计算期末回收固定资产残值？？", vbOKCancel)
        If XZ = vbOK Then
            '计算折旧摊销
            Dim hscz As Boolean = True
            'sdsl_model：所得税率的计算模式，0：使用默认值，1：从Excel中读取已有的值
            Dim sdsl_model As Integer = 1
            '计算折旧摊销
            Call 折旧摊销相关计算(ExcelApp, hscz, sdsl_model)
            '将设置状态写入表格
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(53, 18).Value = "期末回收残值"
            MsgBox("设置完毕！")
        End If
        '————————————————————————————————————————————————————————————————————————————————————————————
        Me.Close()
    End Sub
End Class