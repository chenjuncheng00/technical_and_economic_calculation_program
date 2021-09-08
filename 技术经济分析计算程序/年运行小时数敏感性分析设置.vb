Imports Microsoft.Office.Interop.Excel
Public Class 年运行小时数敏感性分析设置
    Private Sub 年运行小时数敏感性分析设置_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————
        'Checkbox载入收入成本文字
        '收入
        Me.sr1.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(13, 2).Value, String)
        Me.sr2.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 2).Value, String)
        Me.sr3.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 2).Value, String)
        Me.sr4.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 2).Value, String)
        Me.sr5.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 2).Value, String)
        Me.sr6.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(18, 2).Value, String)
        Me.sr7.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 2).Value, String)
        Me.sr8.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(20, 2).Value, String)
        Me.sr9.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 2).Value, String)
        Me.sr10.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 2).Value, String)
        Me.sr11.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 2).Value, String)
        Me.sr12.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 2).Value, String)
        Me.sr13.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 2).Value, String)
        '成本
        Me.cb1.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 9).Value, String)
        Me.cb2.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 9).Value, String)
        Me.cb3.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 9).Value, String)
        Me.cb4.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 9).Value, String)
        Me.cb5.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(18, 9).Value, String)
        Me.cb6.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 9).Value, String)
        Me.cb7.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 9).Value, String)
        Me.cb8.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 9).Value, String)
        Me.cb9.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 9).Value, String)
        Me.cb10.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 9).Value, String)
        Me.cb11.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 9).Value, String)
        Me.cb12.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 9).Value, String)
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '根据输入的收入成本等内容，判断checkbox是否可以被勾选        
        '收入
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(13, 7).Value > 0 Then
            Me.sr1.Enabled = True
        Else
            Me.sr1.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 7).Value > 0 Then
            Me.sr2.Enabled = True
        Else
            Me.sr2.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 7).Value > 0 Then
            Me.sr3.Enabled = True
        Else
            Me.sr3.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 7).Value > 0 Then
            Me.sr4.Enabled = True
        Else
            Me.sr4.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 7).Value > 0 Then
            Me.sr5.Enabled = True
        Else
            Me.sr5.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(18, 7).Value > 0 Then
            Me.sr6.Enabled = True
        Else
            Me.sr6.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 7).Value > 0 Then
            Me.sr7.Enabled = True
        Else
            Me.sr7.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(20, 7).Value > 0 Then
            Me.sr8.Enabled = True
        Else
            Me.sr8.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 7).Value > 0 Then
            Me.sr9.Enabled = True
        Else
            Me.sr9.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 7).Value > 0 Then
            Me.sr10.Enabled = True
        Else
            Me.sr10.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 7).Value > 0 Then
            Me.sr11.Enabled = True
        Else
            Me.sr11.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 7).Value > 0 Then
            Me.sr12.Enabled = True
        Else
            Me.sr12.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 7).Value > 0 Then
            Me.sr13.Enabled = True
        Else
            Me.sr13.Enabled = False
        End If
        '成本
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 14).Value > 0 Then
            Me.cb1.Enabled = True
        Else
            Me.cb1.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 14).Value > 0 Then
            Me.cb2.Enabled = True
        Else
            Me.cb2.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 14).Value > 0 Then
            Me.cb3.Enabled = True
        Else
            Me.cb3.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 14).Value > 0 Then
            Me.cb4.Enabled = True
        Else
            Me.cb4.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(18, 14).Value > 0 Then
            Me.cb5.Enabled = True
        Else
            Me.cb5.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 14).Value > 0 Then
            Me.cb6.Enabled = True
        Else
            Me.cb6.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 14).Value > 0 Then
            Me.cb7.Enabled = True
        Else
            Me.cb7.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 14).Value > 0 Then
            Me.cb8.Enabled = True
        Else
            Me.cb8.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 14).Value > 0 Then
            Me.cb9.Enabled = True
        Else
            Me.cb9.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 14).Value > 0 Then
            Me.cb10.Enabled = True
        Else
            Me.cb10.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 14).Value > 0 Then
            Me.cb11.Enabled = True
        Else
            Me.cb11.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 14).Value > 0 Then
            Me.cb12.Enabled = True
        Else
            Me.cb12.Enabled = False
        End If
    End Sub

    Private Sub 清空输入_Click(sender As Object, e As EventArgs) Handles 清空输入.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否清空已选择的各项内容？", vbOKCancel)
        If XZ = vbOK Then
            '清空Excel内已有的输入
            '收入
            For i = 13 To 25
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 30).Value = Nothing
            Next
            '成本
            For i = 14 To 19
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 31).Value = Nothing
            Next
            For i = 21 To 25
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 31).Value = Nothing
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 31).Value = Nothing
            '————————————————————————————————————————————————————————————————————
            '收入
            Me.sr1.Checked = False
            Me.sr2.Checked = False
            Me.sr3.Checked = False
            Me.sr4.Checked = False
            Me.sr5.Checked = False
            Me.sr6.Checked = False
            Me.sr7.Checked = False
            Me.sr8.Checked = False
            Me.sr9.Checked = False
            Me.sr10.Checked = False
            Me.sr11.Checked = False
            Me.sr12.Checked = False
            Me.sr13.Checked = False
            '成本
            Me.cb1.Checked = False
            Me.cb2.Checked = False
            Me.cb3.Checked = False
            Me.cb4.Checked = False
            Me.cb5.Checked = False
            Me.cb6.Checked = False
            Me.cb7.Checked = False
            Me.cb8.Checked = False
            Me.cb9.Checked = False
            Me.cb10.Checked = False
            Me.cb11.Checked = False
            Me.cb12.Checked = False
        End If
    End Sub

    Private Sub 确定参数_Click(sender As Object, e As EventArgs) Handles 确定参数.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否确认选择的各项内容？", vbOKCancel)
        If XZ = vbOK Then
            '清空Excel内已有的输入
            '收入
            For i = 13 To 25
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 30).Value = Nothing
            Next
            '成本
            For i = 14 To 19
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 31).Value = Nothing
            Next
            For i = 21 To 25
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(i, 31).Value = Nothing
            Next
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 31).Value = Nothing
            '隐藏窗体
            Me.Hide()
            '————————————————————————————————————————————————————————————————————
            '根据复选框的内容，将系数写入Excel，1代表计算，0代表不计算
            '收入
            If Me.sr1.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(13, 30).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(13, 30).Value = 0
            End If
            If Me.sr2.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 30).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 30).Value = 0
            End If
            If Me.sr3.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 30).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 30).Value = 0
            End If
            If Me.sr4.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 30).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 30).Value = 0
            End If
            If Me.sr5.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 30).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 30).Value = 0
            End If
            If Me.sr6.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(18, 30).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(18, 30).Value = 0
            End If
            If Me.sr7.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 30).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 30).Value = 0
            End If
            If Me.sr8.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(20, 30).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(20, 30).Value = 0
            End If
            If Me.sr9.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 30).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 30).Value = 0
            End If
            If Me.sr10.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 30).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 30).Value = 0
            End If
            If Me.sr11.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 30).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 30).Value = 0
            End If
            If Me.sr12.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 30).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 30).Value = 0
            End If
            If Me.sr13.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 30).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 30).Value = 0
            End If
            '成本
            If Me.cb1.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 31).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 31).Value = 0
            End If
            If Me.cb2.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 31).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 31).Value = 0
            End If
            If Me.cb3.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 31).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 31).Value = 0
            End If
            If Me.cb4.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 31).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 31).Value = 0
            End If
            If Me.cb5.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(18, 31).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(18, 31).Value = 0
            End If
            If Me.cb6.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 31).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 31).Value = 0
            End If
            If Me.cb7.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 31).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 31).Value = 0
            End If
            If Me.cb8.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 31).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 31).Value = 0
            End If
            If Me.cb9.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 31).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 31).Value = 0
            End If
            If Me.cb10.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 31).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 31).Value = 0
            End If
            If Me.cb11.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 31).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 31).Value = 0
            End If
            If Me.cb12.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 31).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 31).Value = 0
            End If
            '————————————————————————————————————————————————————————————————————
            '关闭窗体
            Me.Close()
        End If
    End Sub
End Class