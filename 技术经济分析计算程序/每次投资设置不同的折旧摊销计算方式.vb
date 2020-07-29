Imports Microsoft.Office.Interop.Excel

Public Class 每次投资设置不同的折旧摊销计算方式
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
            '屏蔽屏幕更新，防止屏闪
            ExcelApp.Application.ScreenUpdating = False
            '手动计算，关闭excel的自动计算
            ExcelApp.Application.Calculation = XlCalculation.xlCalculationManual
            '解锁表格
            ExcelApp.ThisWorkbook.Worksheets("估算表").Unprotect(Password:="wscjc")
            ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Unprotect(Password:="wscjc")
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Unprotect(Password:="wscjc")
            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Unprotect(Password:="wscjc")
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").unProtect(Password:="wscjc")
            '———————————————————————————————————————————————————————————————————————————————————
            '———————————————————————————————————————————————————————————————————————————————————
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
            '解锁表格
            ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").unProtect(Password:="wscjc")
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").unProtect(Password:="wscjc")
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
                    '第1次投资发生后，折旧计算系数，默认第一次投资发生在第一年且不可以修改
                    '——————————————————————————————————————————————————————————————————————————————————————————————
                    '直线法计算折旧摊销
                    '根据输入的折旧计算年数，修改相关计算系数
                    For i = 3 To 17
                        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(119, i).Value <= gdzczjnx1 + 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(119, i).Value > 1 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(120, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(120, i).Value = 0
                        End If
                    Next
                    For i = 2 To 17
                        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(122, i).Value <= gdzczjnx1 + 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(122, i).Value > 1 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(123, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(123, i).Value = 0
                        End If
                    Next
                    '第1次投资发生后，无形资产摊销计算系数，默认第一次投资发生在第一年且不可以修改
                    '根据输入的无形资产摊销计算年数，修改相关计算系数
                    For i = 3 To 17
                        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(359, i).Value <= wxzctxnx1 + 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(359, i).Value > 1 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(360, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(360, i).Value = 0
                        End If
                    Next
                    For i = 2 To 17
                        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(362, i).Value <= wxzctxnx1 + 1 And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(362, i).Value > 1 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(363, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(363, i).Value = 0
                        End If
                    Next
                    '——————————————————————————————————————————————————————————————————————————————————————————————
                    '年数总和法计算折旧摊销
                    '求年份之和
                    Dim gdzcnfh1 = gdzczjnx1 * (gdzczjnx1 + 1) / 2 '固定资产年份和
                    Dim wxzcnfh1 = wxzctxnx1 * (wxzctxnx1 + 1) / 2 '无形资产年份和
                    '固定资产折旧逐年系数
                    Dim JS2 As Integer = 0
                    For k = 3 To 33 '投资计划与资金筹措表列号
                        If (ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(166, k).Value >= tznf1 + 1 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(166, k).Value <= tznf1 + gdzczjnx1) Then
                            JS2 = JS2 + 1
                            If JS2 <= gdzczjnx1 Then '仅写入大于0的值
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(168, k).Value = (gdzczjnx1 - (JS2 + 1) + 2) * (1 - gdzcczl1) / gdzcnfh1
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(168, k).Value = Nothing
                            End If
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(168, k).Value = Nothing
                        End If
                    Next
                    '无形资产摊销逐年系数
                    Dim JS3 As Integer = 0
                    For k = 3 To 33 '投资计划与资金筹措表列号
                        If (ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(181, k).Value >= tznf1 + 1 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(181, k).Value <= tznf1 + wxzctxnx1) Then
                            JS3 = JS3 + 1
                            If JS3 <= wxzctxnx1 Then
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(183, k).Value = (wxzctxnx1 - (JS3 + 1) + 2) / wxzcnfh1
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(183, k).Value = Nothing
                            End If
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(183, k).Value = Nothing
                        End If
                    Next

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
                    '——————————————————————————————————————————————————————————————————————————————————————————————
                    '直线法计算折旧摊销
                    '第2次投资发生后，折旧计算系数
                    For i = 3 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(175, i).Value <= (gdzczjnx2 + tznf2)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(175, i).Value > tznf2 And tznf2 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(176, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(176, i).Value = 0
                        End If
                    Next
                    For i = 2 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(178, i).Value <= (gdzczjnx2 + tznf2)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(178, i).Value > tznf2 And tznf2 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(179, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(179, i).Value = 0
                        End If
                    Next
                    '第2次投资发生后，无形资产摊销计算系数
                    For i = 3 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(367, i).Value <= (wxzctxnx2 + tznf2)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(367, i).Value > tznf2 And tznf2 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(368, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(368, i).Value = 0
                        End If
                    Next
                    For i = 2 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(370, i).Value <= (wxzctxnx2 + tznf2)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(370, i).Value > tznf2 And tznf2 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(371, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(371, i).Value = 0
                        End If
                    Next
                    '——————————————————————————————————————————————————————————————————————————————————————————————
                    '年数总和法计算折旧摊销
                    '求年份之和
                    Dim gdzcnfh2 = gdzczjnx2 * (gdzczjnx2 + 1) / 2 '固定资产年份和
                    Dim wxzcnfh2 = wxzctxnx2 * (wxzctxnx2 + 1) / 2 '无形资产年份和
                    '固定资产折旧逐年系数
                    Dim JS2 As Integer = 0
                    For k = 3 To 33 '投资计划与资金筹措表列号
                        If (ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(166, k).Value >= tznf2 + 1 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(166, k).Value <= tznf2 + gdzczjnx2) Then
                            JS2 = JS2 + 1
                            If JS2 <= gdzczjnx2 Then '仅写入大于0的值
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(169, k).Value = (gdzczjnx2 - (JS2 + 1) + 2) * (1 - gdzcczl2) / gdzcnfh2
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(169, k).Value = Nothing
                            End If
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(169, k).Value = Nothing
                        End If
                    Next
                    '无形资产摊销逐年系数
                    Dim JS3 As Integer = 0
                    For k = 3 To 33 '投资计划与资金筹措表列号
                        If (ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(181, k).Value >= tznf2 + 1 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(181, k).Value <= tznf2 + wxzctxnx2) Then
                            JS3 = JS3 + 1
                            If JS3 <= wxzctxnx2 Then
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(184, k).Value = (wxzctxnx2 - (JS3 + 1) + 2) / wxzcnfh2
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(184, k).Value = Nothing
                            End If
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(184, k).Value = Nothing
                        End If
                    Next
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
                    '——————————————————————————————————————————————————————————————————————————————————————————————
                    '直线法计算折旧摊销
                    '第3次投资发生后，折旧计算系数
                    For i = 3 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(183, i).Value <= (gdzczjnx3 + tznf3)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(183, i).Value > tznf3 And tznf3 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(184, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(184, i).Value = 0
                        End If
                    Next
                    For i = 2 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(186, i).Value <= (gdzczjnx3 + tznf3)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(186, i).Value > tznf3 And tznf3 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(187, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(187, i).Value = 0
                        End If
                    Next
                    '第3次投资发生后，无形资产摊销计算系数
                    For i = 3 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(375, i).Value <= (wxzctxnx3 + tznf3)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(375, i).Value > tznf3 And tznf3 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(376, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(376, i).Value = 0
                        End If
                    Next
                    For i = 2 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(378, i).Value <= (wxzctxnx3 + tznf3)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(378, i).Value > tznf3 And tznf3 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(379, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(379, i).Value = 0
                        End If
                    Next
                    '——————————————————————————————————————————————————————————————————————————————————————————————
                    '年数总和法计算折旧摊销
                    '求年份之和
                    Dim gdzcnfh3 = gdzczjnx3 * (gdzczjnx3 + 1) / 2 '固定资产年份和
                    Dim wxzcnfh3 = wxzctxnx3 * (wxzctxnx3 + 1) / 2 '无形资产年份和
                    '固定资产折旧逐年系数
                    Dim JS2 As Integer = 0
                    For k = 3 To 33 '投资计划与资金筹措表列号
                        If (ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(166, k).Value >= tznf3 + 1 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(166, k).Value <= tznf3 + gdzczjnx3) Then
                            JS2 = JS2 + 1
                            If JS2 <= gdzczjnx3 Then '仅写入大于0的值
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(170, k).Value = (gdzczjnx3 - (JS2 + 1) + 2) * (1 - gdzcczl3) / gdzcnfh3
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(170, k).Value = Nothing
                            End If
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(170, k).Value = Nothing
                        End If
                    Next
                    '无形资产摊销逐年系数
                    Dim JS3 As Integer = 0
                    For k = 3 To 33 '投资计划与资金筹措表列号
                        If (ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(181, k).Value >= tznf3 + 1 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(181, k).Value <= tznf3 + wxzctxnx3) Then
                            JS3 = JS3 + 1
                            If JS3 <= wxzctxnx3 Then
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(185, k).Value = (wxzctxnx3 - (JS3 + 1) + 2) / wxzcnfh3
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(185, k).Value = Nothing
                            End If
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(185, k).Value = Nothing
                        End If
                    Next
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
                    '——————————————————————————————————————————————————————————————————————————————————————————————
                    '直线法计算折旧摊销
                    '第4次投资发生后，折旧计算系数
                    For i = 3 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(191, i).Value <= (gdzczjnx4 + tznf4)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(191, i).Value > tznf4 And tznf4 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(192, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(192, i).Value = 0
                        End If
                    Next
                    For i = 2 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(194, i).Value <= (gdzczjnx4 + tznf4)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(194, i).Value > tznf4 And tznf4 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(195, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(195, i).Value = 0
                        End If
                    Next
                    '第4次投资发生后，无形资产摊销计算系数
                    For i = 3 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(383, i).Value <= (wxzctxnx4 + tznf4)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(383, i).Value > tznf4 And tznf4 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(384, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(384, i).Value = 0
                        End If
                    Next
                    For i = 2 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(386, i).Value <= (wxzctxnx4 + tznf4)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(386, i).Value > tznf4 And tznf4 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(387, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(387, i).Value = 0
                        End If
                    Next
                    '——————————————————————————————————————————————————————————————————————————————————————————————
                    '年数总和法计算折旧摊销
                    '求年份之和
                    Dim gdzcnfh4 = gdzczjnx4 * (gdzczjnx4 + 1) / 2 '固定资产年份和
                    Dim wxzcnfh4 = wxzctxnx4 * (wxzctxnx4 + 1) / 2 '无形资产年份和
                    '固定资产折旧逐年系数
                    Dim JS2 As Integer = 0
                    For k = 3 To 33 '投资计划与资金筹措表列号
                        If (ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(166, k).Value >= tznf4 + 1 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(166, k).Value <= tznf4 + gdzczjnx4) Then
                            JS2 = JS2 + 1
                            If JS2 <= gdzczjnx4 Then '仅写入大于0的值
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(171, k).Value = (gdzczjnx4 - (JS2 + 1) + 2) * (1 - gdzcczl4) / gdzcnfh4
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(171, k).Value = Nothing
                            End If
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(171, k).Value = Nothing
                        End If
                    Next
                    '无形资产摊销逐年系数
                    Dim JS3 As Integer = 0
                    For k = 3 To 33 '投资计划与资金筹措表列号
                        If (ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(181, k).Value >= tznf4 + 1 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(181, k).Value <= tznf4 + wxzctxnx4) Then
                            JS3 = JS3 + 1
                            If JS3 <= wxzctxnx4 Then
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(186, k).Value = (wxzctxnx4 - (JS3 + 1) + 2) / wxzcnfh4
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(186, k).Value = Nothing
                            End If
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(186, k).Value = Nothing
                        End If
                    Next
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
                    '——————————————————————————————————————————————————————————————————————————————————————————————
                    '直线法计算折旧摊销
                    '第5次投资发生后，折旧计算系数
                    For i = 3 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(199, i).Value <= (gdzczjnx5 + tznf5)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(199, i).Value > tznf5 And tznf5 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(200, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(200, i).Value = 0
                        End If
                    Next
                    For i = 2 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(202, i).Value <= (gdzczjnx5 + tznf5)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(202, i).Value > tznf5 And tznf5 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(203, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(203, i).Value = 0
                        End If
                    Next
                    '第5次投资发生后，无形资产摊销计算系数
                    For i = 3 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(391, i).Value <= (wxzctxnx5 + tznf5)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(391, i).Value > tznf5 And tznf5 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(392, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(392, i).Value = 0
                        End If
                    Next
                    For i = 2 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(394, i).Value <= (wxzctxnx5 + tznf5)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(394, i).Value > tznf5 And tznf5 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(395, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(395, i).Value = 0
                        End If
                    Next
                    '——————————————————————————————————————————————————————————————————————————————————————————————
                    '年数总和法计算折旧摊销
                    '求年份之和
                    Dim gdzcnfh5 = gdzczjnx5 * (gdzczjnx5 + 1) / 2 '固定资产年份和
                    Dim wxzcnfh5 = wxzctxnx5 * (wxzctxnx5 + 1) / 2 '无形资产年份和
                    '固定资产折旧逐年系数
                    Dim JS2 As Integer = 0
                    For k = 3 To 33 '投资计划与资金筹措表列号
                        If (ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(166, k).Value >= tznf5 + 1 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(166, k).Value <= tznf5 + gdzczjnx5) Then
                            JS2 = JS2 + 1
                            If JS2 <= gdzczjnx5 Then '仅写入大于0的值
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(172, k).Value = (gdzczjnx5 - (JS2 + 1) + 2) * (1 - gdzcczl5) / gdzcnfh5
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(172, k).Value = Nothing
                            End If
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(172, k).Value = Nothing
                        End If
                    Next
                    '无形资产摊销逐年系数
                    Dim JS3 As Integer = 0
                    For k = 3 To 33 '投资计划与资金筹措表列号
                        If (ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(181, k).Value >= tznf5 + 1 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(181, k).Value <= tznf5 + wxzctxnx5) Then
                            JS3 = JS3 + 1
                            If JS3 <= wxzctxnx5 Then
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(187, k).Value = (wxzctxnx5 - (JS3 + 1) + 2) / wxzcnfh5
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(187, k).Value = Nothing
                            End If
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(187, k).Value = Nothing
                        End If
                    Next
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
                    '——————————————————————————————————————————————————————————————————————————————————————————————
                    '直线法计算折旧摊销
                    '第6次投资发生后，折旧计算系数
                    For i = 3 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(279, i).Value <= (gdzczjnx6 + tznf6)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(279, i).Value > tznf6 And tznf6 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(280, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(280, i).Value = 0
                        End If
                    Next
                    For i = 2 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(282, i).Value <= (gdzczjnx6 + tznf6)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(282, i).Value > tznf6 And tznf6 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(283, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(283, i).Value = 0
                        End If
                    Next
                    '第6次投资发生后，无形资产摊销计算系数
                    For i = 3 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(399, i).Value <= (wxzctxnx6 + tznf6)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(399, i).Value > tznf6 And tznf6 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(400, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(400, i).Value = 0
                        End If
                    Next
                    For i = 2 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(402, i).Value <= (wxzctxnx6 + tznf6)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(402, i).Value > tznf6 And tznf6 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(403, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(403, i).Value = 0
                        End If
                    Next
                    '——————————————————————————————————————————————————————————————————————————————————————————————
                    '年数总和法计算折旧摊销
                    '求年份之和
                    Dim gdzcnfh6 = gdzczjnx6 * (gdzczjnx6 + 1) / 2 '固定资产年份和
                    Dim wxzcnfh6 = wxzctxnx6 * (wxzctxnx6 + 1) / 2 '无形资产年份和
                    '固定资产折旧逐年系数
                    Dim JS2 As Integer = 0
                    For k = 3 To 33 '投资计划与资金筹措表列号
                        If (ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(166, k).Value >= tznf6 + 1 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(166, k).Value <= tznf6 + gdzczjnx6) Then
                            JS2 = JS2 + 1
                            If JS2 <= gdzczjnx6 Then '仅写入大于0的值
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(173, k).Value = (gdzczjnx6 - (JS2 + 1) + 2) * (1 - gdzcczl6) / gdzcnfh6
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(173, k).Value = Nothing
                            End If
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(173, k).Value = Nothing
                        End If
                    Next
                    '无形资产摊销逐年系数
                    Dim JS3 As Integer = 0
                    For k = 3 To 33 '投资计划与资金筹措表列号
                        If (ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(181, k).Value >= tznf6 + 1 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(181, k).Value <= tznf6 + wxzctxnx6) Then
                            JS3 = JS3 + 1
                            If JS3 <= wxzctxnx6 Then
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(188, k).Value = (wxzctxnx6 - (JS3 + 1) + 2) / wxzcnfh6
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(188, k).Value = Nothing
                            End If
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(188, k).Value = Nothing
                        End If
                    Next
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
                    '——————————————————————————————————————————————————————————————————————————————————————————————
                    '直线法计算折旧摊销
                    '第7次投资发生后，折旧计算系数
                    For i = 3 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(287, i).Value <= (gdzczjnx7 + tznf7)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(287, i).Value > tznf7 And tznf7 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(288, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(288, i).Value = 0
                        End If
                    Next
                    For i = 2 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(290, i).Value <= (gdzczjnx7 + tznf7)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(290, i).Value > tznf7 And tznf7 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(291, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(291, i).Value = 0
                        End If
                    Next
                    '第7次投资发生后，无形资产摊销计算系数
                    For i = 3 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(407, i).Value <= (wxzctxnx7 + tznf7)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(407, i).Value > tznf7 And tznf7 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(408, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(408, i).Value = 0
                        End If
                    Next
                    For i = 2 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(410, i).Value <= (wxzctxnx7 + tznf7)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(410, i).Value > tznf7 And tznf7 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(411, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(411, i).Value = 0
                        End If
                    Next
                    '——————————————————————————————————————————————————————————————————————————————————————————————
                    '年数总和法计算折旧摊销
                    '求年份之和
                    Dim gdzcnfh7 = gdzczjnx7 * (gdzczjnx7 + 1) / 2 '固定资产年份和
                    Dim wxzcnfh7 = wxzctxnx7 * (wxzctxnx7 + 1) / 2 '无形资产年份和
                    '固定资产折旧逐年系数
                    Dim JS2 As Integer = 0
                    For k = 3 To 33 '投资计划与资金筹措表列号
                        If (ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(166, k).Value >= tznf7 + 1 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(166, k).Value <= tznf7 + gdzczjnx7) Then
                            JS2 = JS2 + 1
                            If JS2 <= gdzczjnx7 Then '仅写入大于0的值
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(174, k).Value = (gdzczjnx7 - (JS2 + 1) + 2) * (1 - gdzcczl7) / gdzcnfh7
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(174, k).Value = Nothing
                            End If
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(174, k).Value = Nothing
                        End If
                    Next
                    '无形资产摊销逐年系数
                    Dim JS3 As Integer = 0
                    For k = 3 To 33 '投资计划与资金筹措表列号
                        If (ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(181, k).Value >= tznf7 + 1 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(181, k).Value <= tznf7 + wxzctxnx7) Then
                            JS3 = JS3 + 1
                            If JS3 <= wxzctxnx7 Then
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(189, k).Value = (wxzctxnx7 - (JS3 + 1) + 2) / wxzcnfh7
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(189, k).Value = Nothing
                            End If
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(189, k).Value = Nothing
                        End If
                    Next
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
                    '——————————————————————————————————————————————————————————————————————————————————————————————
                    '直线法计算折旧摊销
                    '第8次投资发生后，折旧计算系数
                    For i = 3 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(295, i).Value <= (gdzczjnx8 + tznf8)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(295, i).Value > tznf8 And tznf8 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(296, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(296, i).Value = 0
                        End If
                    Next
                    For i = 2 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(298, i).Value <= (gdzczjnx8 + tznf8)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(298, i).Value > tznf8 And tznf8 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(299, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(299, i).Value = 0
                        End If
                    Next
                    '第8次投资发生后，无形资产摊销计算系数
                    For i = 3 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(415, i).Value <= (wxzctxnx8 + tznf8)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(415, i).Value > tznf8 And tznf8 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(416, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(416, i).Value = 0
                        End If
                    Next
                    For i = 2 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(418, i).Value <= (wxzctxnx8 + tznf8)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(418, i).Value > tznf8 And tznf8 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(419, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(419, i).Value = 0
                        End If
                    Next
                    '——————————————————————————————————————————————————————————————————————————————————————————————
                    '年数总和法计算折旧摊销
                    '求年份之和
                    Dim gdzcnfh8 = gdzczjnx8 * (gdzczjnx8 + 1) / 2 '固定资产年份和
                    Dim wxzcnfh8 = wxzctxnx8 * (wxzctxnx8 + 1) / 2 '无形资产年份和
                    '固定资产折旧逐年系数
                    Dim JS2 As Integer = 0
                    For k = 3 To 33 '投资计划与资金筹措表列号
                        If (ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(166, k).Value >= tznf8 + 1 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(166, k).Value <= tznf8 + gdzczjnx8) Then
                            JS2 = JS2 + 1
                            If JS2 <= gdzczjnx8 Then '仅写入大于0的值
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(175, k).Value = (gdzczjnx8 - (JS2 + 1) + 2) * (1 - gdzcczl8) / gdzcnfh8
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(175, k).Value = Nothing
                            End If
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(175, k).Value = Nothing
                        End If
                    Next
                    '无形资产摊销逐年系数
                    Dim JS3 As Integer = 0
                    For k = 3 To 33 '投资计划与资金筹措表列号
                        If (ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(181, k).Value >= tznf8 + 1 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(181, k).Value <= tznf8 + wxzctxnx8) Then
                            JS3 = JS3 + 1
                            If JS3 <= wxzctxnx8 Then
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(190, k).Value = (wxzctxnx8 - (JS3 + 1) + 2) / wxzcnfh8
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(190, k).Value = Nothing
                            End If
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(190, k).Value = Nothing
                        End If
                    Next
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
                    '——————————————————————————————————————————————————————————————————————————————————————————————
                    '直线法计算折旧摊销
                    '第9次投资发生后，折旧计算系数
                    For i = 3 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(303, i).Value <= (gdzczjnx9 + tznf9)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(303, i).Value > tznf9 And tznf9 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(304, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(304, i).Value = 0
                        End If
                    Next
                    For i = 2 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(306, i).Value <= (gdzczjnx9 + tznf9)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(306, i).Value > tznf9 And tznf9 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(307, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(307, i).Value = 0
                        End If
                    Next
                    '第9次投资发生后，无形资产摊销计算系数
                    For i = 3 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(423, i).Value <= (wxzctxnx9 + tznf9)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(423, i).Value > tznf9 And tznf9 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(424, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(424, i).Value = 0
                        End If
                    Next
                    For i = 2 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(426, i).Value <= (wxzctxnx9 + tznf9)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(426, i).Value > tznf9 And tznf9 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(427, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(427, i).Value = 0
                        End If
                    Next
                    '——————————————————————————————————————————————————————————————————————————————————————————————
                    '年数总和法计算折旧摊销
                    '求年份之和
                    Dim gdzcnfh9 = gdzczjnx9 * (gdzczjnx9 + 1) / 2 '固定资产年份和
                    Dim wxzcnfh9 = wxzctxnx9 * (wxzctxnx9 + 1) / 2 '无形资产年份和
                    '固定资产折旧逐年系数
                    Dim JS2 As Integer = 0
                    For k = 3 To 33 '投资计划与资金筹措表列号
                        If (ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(166, k).Value >= tznf9 + 1 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(166, k).Value <= tznf9 + gdzczjnx9) Then
                            JS2 = JS2 + 1
                            If JS2 <= gdzczjnx9 Then '仅写入大于0的值
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(176, k).Value = (gdzczjnx9 - (JS2 + 1) + 2) * (1 - gdzcczl9) / gdzcnfh9
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(176, k).Value = Nothing
                            End If
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(176, k).Value = Nothing
                        End If
                    Next
                    '无形资产摊销逐年系数
                    Dim JS3 As Integer = 0
                    For k = 3 To 33 '投资计划与资金筹措表列号
                        If (ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(181, k).Value >= tznf9 + 1 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(181, k).Value <= tznf9 + wxzctxnx9) Then
                            JS3 = JS3 + 1
                            If JS3 <= wxzctxnx9 Then
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(191, k).Value = (wxzctxnx9 - (JS3 + 1) + 2) / wxzcnfh9
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(191, k).Value = Nothing
                            End If
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(191, k).Value = Nothing
                        End If
                    Next
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
                    '——————————————————————————————————————————————————————————————————————————————————————————————
                    '直线法计算折旧摊销
                    '第10次投资发生后，折旧计算系数
                    For i = 3 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(311, i).Value <= (gdzczjnx10 + tznf10)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(311, i).Value > tznf10 And tznf10 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(312, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(312, i).Value = 0
                        End If
                    Next
                    For i = 2 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(314, i).Value <= (gdzczjnx10 + tznf10)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(314, i).Value > tznf10 And tznf10 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(315, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(315, i).Value = 0
                        End If
                    Next
                    '第10次投资发生后，无形资产摊销计算系数
                    For i = 3 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(431, i).Value <= (wxzctxnx10 + tznf10)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(431, i).Value > tznf10 And tznf10 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(432, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(432, i).Value = 0
                        End If
                    Next
                    For i = 2 To 17
                        If (ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(434, i).Value <= (wxzctxnx10 + tznf10)) And ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(434, i).Value > tznf10 And tznf10 <> 0 Then
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(435, i).Value = 1
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(435, i).Value = 0
                        End If
                    Next
                    '——————————————————————————————————————————————————————————————————————————————————————————————
                    '年数总和法计算折旧摊销
                    '求年份之和
                    Dim gdzcnfh10 = gdzczjnx10 * (gdzczjnx10 + 1) / 2 '固定资产年份和
                    Dim wxzcnfh10 = wxzctxnx10 * (wxzctxnx10 + 1) / 2 '无形资产年份和
                    '固定资产折旧逐年系数
                    Dim JS2 As Integer = 0
                    For k = 3 To 33 '投资计划与资金筹措表列号
                        If (ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(166, k).Value >= tznf10 + 1 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(166, k).Value <= tznf10 + gdzczjnx10) Then
                            JS2 = JS2 + 1
                            If JS2 <= gdzczjnx10 Then '仅写入大于0的值
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(177, k).Value = (gdzczjnx10 - (JS2 + 1) + 2) * (1 - gdzcczl10) / gdzcnfh10
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(177, k).Value = Nothing
                            End If
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(177, k).Value = Nothing
                        End If
                    Next
                    '无形资产摊销逐年系数
                    Dim JS3 As Integer = 0
                    For k = 3 To 33 '投资计划与资金筹措表列号
                        If (ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(181, k).Value >= tznf10 + 1 And ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(181, k).Value <= tznf10 + wxzctxnx10) Then
                            JS3 = JS3 + 1
                            If JS3 <= wxzctxnx10 Then
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(192, k).Value = (wxzctxnx10 - (JS3 + 1) + 2) / wxzcnfh10
                            Else
                                ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(192, k).Value = Nothing
                            End If
                        Else
                            ExcelApp.ThisWorkbook.Worksheets("投资计划与资金筹措表").Cells(192, k).Value = Nothing
                        End If
                    Next
                End If
            End If
            '写入计算模式
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(166, 7).Value = "不相同"
            MsgBox("设置完成，每次投资的固定资产折旧和无形资产摊销计算系数就均进行了分别设置，请将计算模式改为方法三或者方法四！")
        End If
        '计算一次Excel
        ExcelApp.Calculate()
        '重新打开excel自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '重新计算投资回收期
        Dim mainprogram As New Com技术经济分析计算程序
        '计算流动资金
        Call mainprogram.流动资金相关计算(ExcelApp)
        '计算回收期
        Call mainprogram.投资回收期计算(ExcelApp)
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
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
            '解锁表格
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").unProtect(Password:="wscjc")
            ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").unProtect(Password:="wscjc")
            '写入各种系数默认值
            For i = 4 To 13
                '固定资产折旧年限
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 22).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(6, 7).Value
                '固定资产残值率
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 24).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(7, 5).Value
                '无形资产摊销年限
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(i, 26).Value = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(8, 7).Value
            Next
            ExcelApp.ThisWorkbook.Worksheets("建设期时间计划表").Cells(166, 7).Value = "相同"
            MsgBox("设置完成，每次投资的固定资产折旧和无形资产摊销计算系数就均相同！")
        End If
        '————————————————————————————————————————————————————————————————————————————————————————————
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        Dim mainprogram As New Com技术经济分析计算程序
        '折旧年数改变后改变的计算用系数
        Call mainprogram.折旧计算年限变化后改变相关系数(ExcelApp)
        '无形资产摊销年限改变后改变的计算用系数
        Call mainprogram.无形资产摊销年限变化后改变相关系数(ExcelApp)
        '年数总和法逐年折旧摊销系数
        Call mainprogram.年数总和法逐年折旧摊销系数(ExcelApp)
        '————————————————————————————————————————————————————————————————————————————————————————————
        '计算一次Excel
        ExcelApp.Calculate()
        '重新打开excel自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '计算流动资金
        Call mainprogram.流动资金相关计算(ExcelApp)
        '重新计算投资回收期
        Call mainprogram.投资回收期计算(ExcelApp)
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
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
            '解锁表格
            ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").unProtect(Password:="wscjc")
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").unProtect(Password:="wscjc")
            '读取项目计算年限
            Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
            '读取固定资产残值
            Dim gdzccz = ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(26, 22).Value
            '清空已有的数据
            '前15年
            For i = 5 To 19
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(1078, i).Value = 0
            Next
            '后16年
            For i = 4 To 19
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(1080, i).Value = 0
            Next
            '将已经有的回收固定资产残值数据写入Excel（计算年限最后一年的固定资产净值）
            '前15年
            For i = 5 To 19
                If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(1077, i).Value = jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(1078, i).Value = gdzccz
                Else
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(1078, i).Value = 0
                End If
            Next
            '后16年
            For i = 4 To 19
                If ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(1079, i).Value = jsnx Then
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(1080, i).Value = gdzccz
                Else
                    ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(1080, i).Value = 0
                End If
            Next
            '将设置状态写入表格
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(53, 18).Value = "期末不回收残值"
            MsgBox("设置完毕！")
        End If
        '————————————————————————————————————————————————————————————————————————————————————————————
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        Dim mainprogram As New Com技术经济分析计算程序
        '计算一次Excel
        ExcelApp.Calculate()
        '重新打开excel自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '计算流动资金
        Call mainprogram.流动资金相关计算(ExcelApp)
        '重新计算投资回收期
        Call mainprogram.投资回收期计算(ExcelApp)
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
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
            '解锁表格
            ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").unProtect(Password:="wscjc")
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").unProtect(Password:="wscjc")
            '数据全部设置为0
            '前15年
            For i = 5 To 19
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(1078, i).Value = 0
            Next
            '后16年
            For i = 4 To 19
                ExcelApp.ThisWorkbook.Worksheets("折旧摊销表").Cells(1080, i).Value = 0
            Next
            '将设置状态写入表格
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(53, 18).Value = "期末回收残值"
            MsgBox("设置完毕！")
        End If
        '————————————————————————————————————————————————————————————————————————————————————————————
        '屏蔽ctrl+break
        ExcelApp.Application.EnableCancelKey = XlEnableCancelKey.xlDisabled
        Dim mainprogram As New Com技术经济分析计算程序
        '计算一次Excel
        ExcelApp.Calculate()
        '重新打开excel自动计算
        ExcelApp.Application.Calculation = XlCalculation.xlCalculationAutomatic
        '计算流动资金
        Call mainprogram.流动资金相关计算(ExcelApp)
        '重新计算投资回收期
        Call mainprogram.投资回收期计算(ExcelApp)
        '重新打开屏幕更新
        ExcelApp.Application.ScreenUpdating = True
        Me.Close()
    End Sub
End Class